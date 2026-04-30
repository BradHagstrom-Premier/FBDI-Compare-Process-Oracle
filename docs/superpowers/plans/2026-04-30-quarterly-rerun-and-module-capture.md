# Quarterly Rerun + Module Capture Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Capture each FBDI file's source Oracle module during download, persist as `baselines/<ver>/file_modules.json`, surgically populate the Module column in `FBDI_to_ApplaudTables_Mapping.xlsx` without disturbing the 639 hand-edited rows, then validate by running the full `fbdi-compare-release` skill against fresh 26A and 26B baselines.

**Architecture:** Six independent code components (URL classifier, download-loop capture, column-F updater, CLI subcommand, post-run validator, skill-stage update). TDD-first for all new code. The rerun itself is a single execution event after Phase 1 is green and committed.

**Tech Stack:** Python 3.14+, openpyxl, Selenium (existing), pytest, argparse.

**Spec:** `docs/superpowers/specs/2026-04-30-quarterly-rerun-and-module-capture-design.md`.

---

## File Structure

**Files to create:**
- `tests/test_module_classifier.py` — URL classifier tests
- `tests/test_populate_module.py` — Module column updater tests
- `tests/test_verify_rerun.py` — Post-run validator tests
- `fbdi/populate_module.py` — surgical column-F updater for the mapping xlsx
- `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py` — macro-signal validator (catalog row delta, compare changes delta, module pct populated)

**Files to modify:**
- `tools/download_and_clear.py` — add `URL_TO_MODULE`, `module_from_base_url()`, capture loop, JSON write
- `fbdi/cli.py` — add `populate-module` subcommand
- `.claude/skills/fbdi-compare-release/SKILL.md` — add Stage 6.5 + HITL #7

**Important:** the existing `verify_run.py` already covers NO_HEADER and Issues regression checks. The new `verify_rerun.py` only adds the **non-overlapping** signals: catalog row count delta, compare changes delta, and module column population %. Both scripts run in Stage 8.

---

## Phase 1 — Code (TDD)

### Task 1: Module URL classifier

Pure function `module_from_base_url(url: str) -> str` mapping the four `MODULE_URL_TEMPLATES` URL patterns to canonical module names.

**Files:**
- Create: `tests/test_module_classifier.py`
- Modify: `tools/download_and_clear.py:50` (insert `URL_TO_MODULE` and helper after `MODULE_URL_TEMPLATES`)

- [ ] **Step 1.1: Write failing tests**

Create `tests/test_module_classifier.py`:

```python
"""Tests for module_from_base_url — Oracle docs URL → canonical module name."""

import pytest

from tools.download_and_clear import module_from_base_url


class TestModuleFromBaseUrl:
    def test_financials_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/financials/26b/oefbf/index.html"
        assert module_from_base_url(url) == "Financials"

    def test_procurement_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/procurement/26b/oefbp/index.html"
        assert module_from_base_url(url) == "Procurement"

    def test_supply_chain_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/supply-chain-and-manufacturing/26b/oefsc/index.html"
        assert module_from_base_url(url) == "Supply Chain & Manufacturing"

    def test_project_management_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/project-management/26b/oefpp/index.html"
        assert module_from_base_url(url) == "Project Management"

    def test_unknown_url_raises(self):
        with pytest.raises(ValueError, match="Unknown Oracle module URL"):
            module_from_base_url("https://docs.oracle.com/en/cloud/saas/hcm/26b/x/y.html")
```

- [ ] **Step 1.2: Run tests and verify they fail**

Run: `python -m pytest tests/test_module_classifier.py -v`
Expected: 5 errors with `ImportError: cannot import name 'module_from_base_url'`.

- [ ] **Step 1.3: Implement classifier**

In `tools/download_and_clear.py`, immediately after the `MODULE_URL_TEMPLATES` block (around line 49), insert:

```python
# Maps URL slug → canonical module name. Keep keys aligned with
# MODULE_URL_TEMPLATES; values match the existing taxonomy used in
# fbdi/build_mapping.py KNOWN_MAPPINGS (`&` not "and" for SCM).
URL_TO_MODULE = {
    "project-management": "Project Management",
    "financials": "Financials",
    "procurement": "Procurement",
    "supply-chain-and-manufacturing": "Supply Chain & Manufacturing",
}


def module_from_base_url(url: str) -> str:
    """Extract the Oracle module name from a base URL.

    Example: 'https://docs.oracle.com/en/cloud/saas/financials/26b/oefbf/index.html'
             → 'Financials'

    Raises ValueError for URLs that don't match any known module slug.
    The `/saas/<slug>/` guard avoids false positives where the slug
    happens to appear elsewhere in the URL (e.g., as a query param).
    """
    for slug, module in URL_TO_MODULE.items():
        if f"/saas/{slug}/" in url:
            return module
    raise ValueError(f"Unknown Oracle module URL: {url}")
```

- [ ] **Step 1.4: Run tests and verify they pass**

Run: `python -m pytest tests/test_module_classifier.py -v`
Expected: 5 passed.

- [ ] **Step 1.5: Commit**

```bash
git add tests/test_module_classifier.py tools/download_and_clear.py
git commit -m "$(cat <<'EOF'
feat(downloader): module URL classifier

Pure function mapping the four MODULE_URL_TEMPLATES URLs to canonical
Oracle module names (Financials, Procurement, Supply Chain &
Manufacturing, Project Management). Used by the download loop to
capture per-file module metadata.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 2: Per-release module capture in download loop

Wire the classifier into the existing `download_files()` loop. Accumulate `{filename: module}` and write `baselines/<ver>/file_modules.json` after a successful download pass.

**Files:**
- Modify: `tools/download_and_clear.py:113-222` (the `download_files()` function)

This task does NOT add unit tests — Selenium-mocking is high-effort, low-value. The integration test is the actual rerun in Phase 2. We do extract the JSON-writing helper so it could be tested separately if ever needed.

- [ ] **Step 2.1: Add JSON-write helper**

In `tools/download_and_clear.py`, after the `module_from_base_url` definition, add:

```python
def write_module_map(file_modules: dict, version: str, baselines_root: str) -> str:
    """Write {filename: module} to baselines/<ver>/file_modules.json.

    Adds the hardcoded entry for the FSM-distributed file
    RapidImplementationForCashManagement.xlsm (manually placed by user, so
    not seen during scrape). Returns the absolute path written.
    """
    import json
    file_modules = dict(file_modules)
    file_modules.setdefault("RapidImplementationForCashManagement.xlsm", "Financials")
    out_path = os.path.join(baselines_root, version.lower(), "file_modules.json")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(file_modules, f, indent=2, sort_keys=True)
    return out_path
```

- [ ] **Step 2.2: Wire capture into download_files()**

Modify `download_files()` in `tools/download_and_clear.py:113`. Add `file_modules = {}` initialization above the `for base_url in base_urls:` loop, capture per-file module after each successful download, and write the JSON at the end.

Replace the function signature and add capture/write. The diff (showing only changed lines):

```python
def download_files(driver, download_path, version):
    """Scrape Oracle docs for the given version and download all .xlsm files.

    Returns: dict mapping filename -> module string (for module map).
    """
    base_urls = [t.format(ver=version) for t in MODULE_URL_TEMPLATES]
    session = requests.Session()
    seen_filenames = set(os.listdir(download_path))
    file_modules: dict[str, str] = {}  # NEW

    for base_url in base_urls:
        # Resolve module once per base_url; default to "Unknown" if Oracle
        # restructures URL patterns so we keep going rather than crash mid-run.
        try:
            module = module_from_base_url(base_url)
        except ValueError as e:
            print(f"  WARNING: {e} — module will be 'Unknown' for files from this URL")
            module = "Unknown"

        print(f"\nNavigating to {base_url}")
        driver.get(base_url)
        # ... (existing nav/expand logic unchanged) ...

        # Inside the inner-most download success block, after the
        # `time.sleep(1)` that follows the file write (~line 210), add:
        file_modules[local_filename] = module

    # After the for-loop closes, before the function returns, return the dict
    return file_modules
```

The exact insertion points:
- Line ~117: add `file_modules: dict[str, str] = {}` after `seen_filenames = ...`
- Line ~120 (top of `for base_url` loop body, before `print(f"\nNavigating ..."`): add the module-resolution try/except block above
- Line ~210 (immediately after `time.sleep(1)` that follows the chunk-write loop): add `file_modules[local_filename] = module`
- End of function: add `return file_modules`

- [ ] **Step 2.3: Wire JSON write into the main entry point**

Find the call site of `download_files(...)` in `tools/download_and_clear.py` (search for `download_files(`). It's invoked from the `main()` block. After the call, add:

```python
file_modules = download_files(driver, download_path, version)
# Write the per-release module map. Only happens on a successful download
# pass — --clear-only and --skip-clear paths don't reach this line.
modules_path = write_module_map(file_modules, version, "baselines")
print(f"Wrote module map for {len(file_modules)} files to {modules_path}")
```

If the existing call to `download_files` doesn't capture its return value, update it to do so.

- [ ] **Step 2.4: Smoke test the JSON helper**

Add a quick standalone smoke test inline in `tests/test_module_classifier.py`:

```python
def test_write_module_map_round_trip(tmp_path, monkeypatch):
    """write_module_map produces well-formed JSON with FSM file added."""
    from tools.download_and_clear import write_module_map
    import json

    file_modules = {
        "AutoInvoiceImportTemplate.xlsm": "Financials",
        "ItemImportTemplate.xlsm": "Supply Chain & Manufacturing",
    }
    out_path = write_module_map(file_modules, "26C", str(tmp_path))

    with open(out_path) as f:
        data = json.load(f)

    # FSM file auto-added
    assert data["RapidImplementationForCashManagement.xlsm"] == "Financials"
    # User-supplied entries preserved
    assert data["AutoInvoiceImportTemplate.xlsm"] == "Financials"
    assert data["ItemImportTemplate.xlsm"] == "Supply Chain & Manufacturing"
    # Sorted keys (deterministic output)
    keys = list(data.keys())
    assert keys == sorted(keys)
```

Run: `python -m pytest tests/test_module_classifier.py -v`
Expected: 6 passed (5 from Task 1 + 1 new).

- [ ] **Step 2.5: Commit**

```bash
git add tools/download_and_clear.py tests/test_module_classifier.py
git commit -m "$(cat <<'EOF'
feat(downloader): capture per-file module during download

download_files() now accumulates {filename: module} as it scrapes,
defaulting to 'Unknown' if module_from_base_url raises (defensive
against Oracle URL restructuring). write_module_map() persists to
baselines/<ver>/file_modules.json with the FSM-distributed
RapidImplementationForCashManagement.xlsm seeded as Financials.

The map is only written on a successful download pass; --clear-only
and --skip-clear leave any existing JSON untouched.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 3: Module column updater (`fbdi/populate_module.py`)

Surgical updater for column F of the working mapping spreadsheet. Opens with full-mode openpyxl, writes only column F per row, leaves all other cells (including formatting, formulas, validations, freeze-panes) untouched.

**Files:**
- Create: `fbdi/populate_module.py`
- Create: `tests/test_populate_module.py`

- [ ] **Step 3.1: Write the failing tests**

Create `tests/test_populate_module.py`:

```python
"""Tests for fbdi.populate_module — surgical column-F updater for the mapping xlsx."""

import pytest
from pathlib import Path
from openpyxl import Workbook, load_workbook

from fbdi.populate_module import populate_module_column


# Mirrors the actual columns of FBDI_to_ApplaudTables_Mapping.xlsx
# 'FBDI Mapping' sheet: A FBDI Template, B FBDI Tab, C Applaud Table,
# D Prefix, E Status, F Module, G In Base System?
HEADER_ROW = ["FBDI Template", "FBDI Tab", "Applaud Table",
              "Prefix", "Status", "Module", "In Base System?"]


def _make_mapping_workbook(path: Path, rows: list[list]):
    """Build a synthetic mapping xlsx with the production sheet structure."""
    wb = Workbook()
    ws = wb.active
    ws.title = "FBDI Mapping"
    for col_idx, val in enumerate(HEADER_ROW, start=1):
        ws.cell(row=1, column=col_idx, value=val)
    for r_idx, row_vals in enumerate(rows, start=2):
        for c_idx, val in enumerate(row_vals, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)
    # Add the second sheet that exists in production so the updater
    # has to find the right one by name.
    wb.create_sheet("Applaud Tables Reference")
    wb.save(path)
    wb.close()


def _read_module_col(path: Path) -> list:
    wb = load_workbook(path, read_only=True)
    ws = wb["FBDI Mapping"]
    out = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue  # skip header
        out.append(row[5])  # column F = index 5
    wb.close()
    return out


class TestPopulateModuleColumn:
    def test_populated_from_new_release(self, tmp_path):
        """Happy path: every row's FBDI Template appears in NEW; all populated."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL", "T_RA", "TA4", "", "", ""],
            ["ItemImportTemplate", "EGP_ITEMS_INTERFACE", "T_EGP", "T91", "", "", ""],
        ])
        new_modules = {
            "AutoInvoiceImportTemplate.xlsm": "Financials",
            "ItemImportTemplate.xlsm": "Supply Chain & Manufacturing",
        }
        result = populate_module_column(mapping, new_modules, old_modules={})

        assert result == {"populated": 2, "blank": 0, "overwritten": 0}
        assert _read_module_col(mapping) == [
            "Financials", "Supply Chain & Manufacturing",
        ]

    def test_falls_back_to_old_when_missing_from_new(self, tmp_path):
        """File only in OLD release: OLD module is used."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["LegacyTemplate", "LEGACY_TAB", "", "", "", "", ""],
        ])
        result = populate_module_column(
            mapping,
            new_modules={},
            old_modules={"LegacyTemplate.xlsm": "Procurement"},
        )
        assert result == {"populated": 1, "blank": 0, "overwritten": 0}
        assert _read_module_col(mapping) == ["Procurement"]

    def test_new_wins_when_present_in_both(self, tmp_path):
        """When file is in BOTH releases, NEW takes precedence."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["DualTemplate", "DUAL_TAB", "", "", "", "", ""],
        ])
        result = populate_module_column(
            mapping,
            new_modules={"DualTemplate.xlsm": "Financials"},
            old_modules={"DualTemplate.xlsm": "Procurement"},
        )
        assert result == {"populated": 1, "blank": 0, "overwritten": 0}
        assert _read_module_col(mapping) == ["Financials"]

    def test_blank_when_in_neither(self, tmp_path):
        """File in neither release: Module stays blank."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["GhostTemplate", "GHOST_TAB", "", "", "", "", ""],
        ])
        result = populate_module_column(mapping, new_modules={}, old_modules={})
        assert result == {"populated": 0, "blank": 1, "overwritten": 0}
        assert _read_module_col(mapping) == [None]

    def test_other_columns_preserved(self, tmp_path):
        """Manually-edited columns (A, B, C, D, E, G) survive the update."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL",
             "T_RA_INTERFACE_LINES_ALL", "TA4", "Mapped", "", "Yes"],
        ])
        populate_module_column(
            mapping,
            new_modules={"AutoInvoiceImportTemplate.xlsm": "Financials"},
            old_modules={},
        )
        wb = load_workbook(mapping, read_only=True)
        ws = wb["FBDI Mapping"]
        row = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
        wb.close()
        assert row == (
            "AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL",
            "T_RA_INTERFACE_LINES_ALL", "TA4", "Mapped", "Financials", "Yes",
        )

    def test_idempotency(self, tmp_path):
        """Running twice produces identical output and bumps overwritten count."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["AutoInvoiceImportTemplate", "RA_TAB", "", "", "", "", ""],
        ])
        new = {"AutoInvoiceImportTemplate.xlsm": "Financials"}
        first = populate_module_column(mapping, new_modules=new, old_modules={})
        second = populate_module_column(mapping, new_modules=new, old_modules={})
        assert first == {"populated": 1, "blank": 0, "overwritten": 0}
        assert second == {"populated": 1, "blank": 0, "overwritten": 1}
        assert _read_module_col(mapping) == ["Financials"]

    def test_xlsm_suffix_normalized(self, tmp_path):
        """JSON keys may have .xlsm; spreadsheet col A may not. Both should match."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["AutoInvoiceImportTemplate", "RA_TAB", "", "", "", "", ""],   # no suffix
            ["ItemImportTemplate.xlsm", "EGP_TAB", "", "", "", "", ""],     # with suffix
        ])
        new = {
            "AutoInvoiceImportTemplate.xlsm": "Financials",
            "ItemImportTemplate.xlsm": "Supply Chain & Manufacturing",
        }
        result = populate_module_column(mapping, new_modules=new, old_modules={})
        assert result == {"populated": 2, "blank": 0, "overwritten": 0}
        assert _read_module_col(mapping) == [
            "Financials", "Supply Chain & Manufacturing",
        ]

    def test_blank_template_row_skipped(self, tmp_path):
        """Rows with empty FBDI Template (col A) are skipped, not counted."""
        mapping = tmp_path / "mapping.xlsx"
        _make_mapping_workbook(mapping, [
            ["AutoInvoiceImportTemplate", "RA_TAB", "", "", "", "", ""],
            [None, None, None, None, None, None, None],  # blank row in middle
            ["ItemImportTemplate", "EGP_TAB", "", "", "", "", ""],
        ])
        new = {
            "AutoInvoiceImportTemplate.xlsm": "Financials",
            "ItemImportTemplate.xlsm": "Supply Chain & Manufacturing",
        }
        result = populate_module_column(mapping, new_modules=new, old_modules={})
        # 2 rows have a non-blank template; both populated. Blank row not counted.
        assert result == {"populated": 2, "blank": 0, "overwritten": 0}
```

- [ ] **Step 3.2: Run tests and verify they fail**

Run: `python -m pytest tests/test_populate_module.py -v`
Expected: 8 errors with `ModuleNotFoundError: No module named 'fbdi.populate_module'`.

- [ ] **Step 3.3: Implement `fbdi/populate_module.py`**

Create the file:

```python
"""Surgical Module-column updater for FBDI_to_ApplaudTables_Mapping.xlsx.

Reads file_modules.json from a NEW and OLD release, looks up each row's
FBDI Template (col A) against the merged dict (NEW wins), writes column
F (Module) only. All other cells, formatting, formulas, merged cells,
validations, and freeze-panes are preserved by openpyxl's full-mode load.
"""

from __future__ import annotations

import json
from pathlib import Path

from openpyxl import load_workbook


SHEET_NAME = "FBDI Mapping"
TEMPLATE_COL = 1  # A
MODULE_COL = 6    # F


def _stem(name) -> str:
    """Normalize an FBDI Template identifier: strip .xlsm suffix and whitespace."""
    if name is None:
        return ""
    s = str(name).strip()
    if s.lower().endswith(".xlsm"):
        s = s[:-5]
    return s


def _load_modules_json(path: Path) -> dict[str, str]:
    """Load a file_modules.json. Returns {} if path is missing."""
    if not path.is_file():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def populate_module_column(
    mapping_path: Path,
    new_modules: dict[str, str],
    old_modules: dict[str, str],
) -> dict[str, int]:
    """Update the Module column (F) in place. Returns counts dict.

    Lookup order: NEW release wins; OLD fills only when NEW lacks the file.
    Files in neither release leave the cell blank.

    Returns: {'populated': N, 'blank': M, 'overwritten': K}
      - populated: rows that ended with a non-blank Module value
      - blank: rows with non-blank FBDI Template that found no match
      - overwritten: rows whose pre-existing Module value was changed
    """
    # Merge: new_modules takes precedence — Python dict merge semantics
    # mean the right-hand operand wins for duplicate keys.
    merged = {_stem(k): v for k, v in old_modules.items()}
    merged.update({_stem(k): v for k, v in new_modules.items()})

    wb = load_workbook(mapping_path)  # full mode preserves everything
    if SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Sheet '{SHEET_NAME}' not found in {mapping_path}")

    ws = wb[SHEET_NAME]
    populated = 0
    blank = 0
    overwritten = 0

    # Iterate data rows (skip header at row 1)
    for row_idx in range(2, ws.max_row + 1):
        template_cell = ws.cell(row=row_idx, column=TEMPLATE_COL)
        template = _stem(template_cell.value)
        if not template:
            continue  # blank rows don't count

        module_cell = ws.cell(row=row_idx, column=MODULE_COL)
        previous = module_cell.value
        new_value = merged.get(template)

        if new_value:
            if previous and previous != new_value:
                overwritten += 1
            elif previous == new_value:
                # Idempotent re-write: count as overwritten so callers can
                # detect re-runs.
                overwritten += 1
            module_cell.value = new_value
            populated += 1
        else:
            blank += 1

    wb.save(mapping_path)
    wb.close()

    return {"populated": populated, "blank": blank, "overwritten": overwritten}
```

- [ ] **Step 3.4: Run tests and verify they pass**

Run: `python -m pytest tests/test_populate_module.py -v`
Expected: 8 passed.

If `test_idempotency` fails because the first run's `populated=1, overwritten=0` doesn't match expectation, recheck the logic: on first run there's no `previous` value (or it equals new_value? — actually first run previous=None so `if previous and previous != new_value` is False, and `elif previous == new_value` also False since None != "Financials"). So overwritten=0 on first run. On second run previous="Financials", equals new_value, so the `elif` branch fires → overwritten=1. Matches the test.

- [ ] **Step 3.5: Commit**

```bash
git add fbdi/populate_module.py tests/test_populate_module.py
git commit -m "$(cat <<'EOF'
feat(populate-module): surgical Module-column updater

New module fbdi.populate_module surgically writes column F (Module) of
FBDI_to_ApplaudTables_Mapping.xlsx based on file_modules.json from the
NEW and OLD releases (NEW wins). Uses openpyxl full mode so formatting,
formulas, merged cells, validations, and freeze-panes are preserved.
Tolerates .xlsm suffix mismatches between the JSON and the spreadsheet.

Returns counts: {populated, blank, overwritten} for caller telemetry.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 4: CLI subcommand for populate-module

Wire `python -m fbdi populate-module --new 26B --old 26A` into the existing argparse-based CLI.

**Files:**
- Modify: `fbdi/cli.py:84-107` (insert new subparser after `catalog_parser` config)
- Modify: `fbdi/cli.py:115-120` (add dispatch branch)
- Modify: `fbdi/cli.py:291` (append `_run_populate_module()` function)
- Modify: `tests/test_cli.py` (add 1 smoke test)

- [ ] **Step 4.1: Write failing CLI smoke test**

Append to `tests/test_cli.py`:

```python
def test_populate_module_subcommand_invocation(tmp_path, monkeypatch, capsys):
    """`python -m fbdi populate-module` invokes populate_module_column with the
    right args and prints the summary."""
    import fbdi.cli as cli_mod
    from openpyxl import Workbook

    # Build minimal artifacts in tmp_path
    (tmp_path / "baselines" / "26a").mkdir(parents=True)
    (tmp_path / "baselines" / "26b").mkdir(parents=True)
    (tmp_path / "baselines" / "26a" / "file_modules.json").write_text(
        '{"AutoInvoiceImportTemplate.xlsm": "Financials"}'
    )
    (tmp_path / "baselines" / "26b" / "file_modules.json").write_text(
        '{"AutoInvoiceImportTemplate.xlsm": "Financials"}'
    )

    mapping_path = tmp_path / "mapping.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "FBDI Mapping"
    headers = ["FBDI Template", "FBDI Tab", "Applaud Table", "Prefix",
               "Status", "Module", "In Base System?"]
    for c_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c_idx, value=h)
    ws.cell(row=2, column=1, value="AutoInvoiceImportTemplate")
    ws.cell(row=2, column=2, value="RA_TAB")
    wb.save(mapping_path)
    wb.close()

    monkeypatch.chdir(tmp_path)
    cli_mod.main(["populate-module", "--new", "26b", "--old", "26a",
                  "--mapping", str(mapping_path)])

    out = capsys.readouterr().out
    assert "populated" in out.lower()
    assert "1" in out  # one row populated
```

- [ ] **Step 4.2: Run the test and verify it fails**

Run: `python -m pytest tests/test_cli.py::test_populate_module_subcommand_invocation -v`
Expected: FAIL with `argparse.ArgumentError` or `SystemExit: 2` (subcommand unknown).

- [ ] **Step 4.3: Add the subparser**

In `fbdi/cli.py`, after the catalog subparser block (around line 107) and before `args = parser.parse_args(argv)`, insert:

```python
    populate_parser = subparsers.add_parser(
        "populate-module",
        help="Populate the Module column in FBDI_to_ApplaudTables_Mapping.xlsx",
    )
    populate_parser.add_argument(
        "--new", required=True, type=str,
        help="Newer release label (e.g. 26B) — reads baselines/<new>/file_modules.json",
    )
    populate_parser.add_argument(
        "--old", required=True, type=str,
        help="Older release label (e.g. 26A) — reads baselines/<old>/file_modules.json as fallback",
    )
    populate_parser.add_argument(
        "--mapping", type=Path,
        default=Path("FBDI_to_ApplaudTables_Mapping.xlsx"),
        help="Path to the mapping spreadsheet (default: ./FBDI_to_ApplaudTables_Mapping.xlsx)",
    )
```

- [ ] **Step 4.4: Add the dispatch branch**

In `fbdi/cli.py`, locate the `if args.command == "compare":` block (around line 115) and add:

```python
    elif args.command == "populate-module":
        _run_populate_module(args)
```

- [ ] **Step 4.5: Add the runner function**

At the end of `fbdi/cli.py`, append:

```python
def _run_populate_module(args: argparse.Namespace) -> None:
    """Surgically populate the Module column in the mapping spreadsheet."""
    import json

    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(name)s: %(message)s",
    )

    from fbdi.populate_module import populate_module_column

    new_path = Path("baselines") / args.new.lower() / "file_modules.json"
    old_path = Path("baselines") / args.old.lower() / "file_modules.json"

    if not new_path.is_file():
        print(f"Error: {new_path} not found. Run downloader for {args.new} first.")
        sys.exit(2)
    if not old_path.is_file():
        print(f"Error: {old_path} not found. Run downloader for {args.old} first.")
        sys.exit(2)

    if not args.mapping.is_file():
        print(f"Notice: mapping file {args.mapping} not present — skipping populate-module.")
        return  # not an error; mapping may not be checked out

    with open(new_path, "r", encoding="utf-8") as f:
        new_modules = json.load(f)
    with open(old_path, "r", encoding="utf-8") as f:
        old_modules = json.load(f)

    try:
        result = populate_module_column(args.mapping, new_modules, old_modules)
    except PermissionError:
        print(f"Error: {args.mapping} is open in Excel — close it and re-run.")
        sys.exit(3)

    print(json.dumps({
        "mapping": str(args.mapping),
        "new_release": args.new.upper(),
        "old_release": args.old.upper(),
        **result,
    }, indent=2))
```

- [ ] **Step 4.6: Run the test and verify it passes**

Run: `python -m pytest tests/test_cli.py::test_populate_module_subcommand_invocation -v`
Expected: PASS.

- [ ] **Step 4.7: Run the full test suite to confirm nothing broke**

Run: `python -m pytest tests/ -q`
Expected: 270 passed (255 existing + 6 from Tasks 1–2 + 8 from Task 3 + 1 from Task 4). Actual baseline may differ slightly if the existing suite has evolved — verify zero failures and that count is monotonically increasing from prior task.

- [ ] **Step 4.8: Commit**

```bash
git add fbdi/cli.py tests/test_cli.py
git commit -m "$(cat <<'EOF'
feat(cli): add populate-module subcommand

Wires fbdi.populate_module into the python -m fbdi entry point.
Reads baselines/<new>/file_modules.json and baselines/<old>/file_modules.json,
calls populate_module_column on the mapping spreadsheet, prints a JSON
summary. Exits 2 if either JSON is missing, 3 if the mapping file is
locked by Excel.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 5: Post-run validator (`verify_rerun.py`)

Add the macro signals not already covered by `verify_run.py`: catalog row count delta vs the existing catalog, compare-report changes-count delta vs a configurable baseline, and Module column population %.

**Files:**
- Create: `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py`
- Create: `tests/test_verify_rerun.py`

- [ ] **Step 5.1: Write failing tests**

Create `tests/test_verify_rerun.py`:

```python
"""Tests for the verify_rerun.py macro-signal validator."""

import importlib.util
import json
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook


SKILL_SCRIPT = Path(__file__).resolve().parent.parent / ".claude" / "skills" / \
    "fbdi-compare-release" / "scripts" / "verify_rerun.py"


def _load_module():
    spec = importlib.util.spec_from_file_location("verify_rerun", SKILL_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["verify_rerun"] = mod
    spec.loader.exec_module(mod)
    return mod


def _make_catalog(path: Path, release_rows: dict[str, int]):
    """Build a synthetic FBDI_Master_Catalog with N rows on each release sheet."""
    wb = Workbook()
    wb.remove(wb.active)
    for release, n_rows in release_rows.items():
        ws = wb.create_sheet(release)
        ws.cell(row=1, column=1, value="file")
        ws.cell(row=1, column=2, value="tab")
        for i in range(n_rows):
            ws.cell(row=i + 2, column=1, value=f"file{i}.xlsm")
            ws.cell(row=i + 2, column=2, value=f"tab{i}")
    wb.save(path)
    wb.close()


def _make_compare_report(path: Path, n_changes: int):
    wb = Workbook()
    ws = wb.active
    ws.title = "Changes"
    ws.cell(row=1, column=1, value="file")
    for i in range(n_changes):
        ws.cell(row=i + 2, column=1, value=f"row{i}")
    wb.save(path)
    wb.close()


def _make_mapping(path: Path, total_rows: int, populated_rows: int):
    wb = Workbook()
    ws = wb.active
    ws.title = "FBDI Mapping"
    headers = ["FBDI Template", "FBDI Tab", "Applaud Table", "Prefix",
               "Status", "Module", "In Base System?"]
    for c_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c_idx, value=h)
    for i in range(total_rows):
        ws.cell(row=i + 2, column=1, value=f"Template{i}")
        if i < populated_rows:
            ws.cell(row=i + 2, column=6, value="Financials")
    wb.save(path)
    wb.close()


class TestVerifyRerun:
    def test_all_green(self, tmp_path):
        mod = _load_module()
        # post = pre (no delta)
        _make_catalog(tmp_path / "post.xlsx", {"26A": 12000, "26B": 12000})
        _make_catalog(tmp_path / "pre.xlsx",  {"26A": 12000, "26B": 12000})
        _make_compare_report(tmp_path / "report.xlsx", 706)
        _make_mapping(tmp_path / "mapping.xlsx", 100, 96)

        result = mod.run_checks(
            new_catalog=tmp_path / "post.xlsx",
            baseline_catalog=tmp_path / "pre.xlsx",
            compare_report=tmp_path / "report.xlsx",
            mapping=tmp_path / "mapping.xlsx",
            release="26B",
        )
        assert result["regressions"] == []
        assert result["catalog_delta_pct"] == pytest.approx(0.0, abs=0.01)
        assert result["module_pct_populated"] == pytest.approx(96.0, abs=0.01)

    def test_catalog_row_delta_exceeds_threshold(self, tmp_path):
        mod = _load_module()
        _make_catalog(tmp_path / "post.xlsx", {"26B": 11000})  # -8.3% vs 12000
        _make_catalog(tmp_path / "pre.xlsx",  {"26B": 12000})
        _make_compare_report(tmp_path / "report.xlsx", 706)
        _make_mapping(tmp_path / "mapping.xlsx", 100, 96)

        result = mod.run_checks(
            new_catalog=tmp_path / "post.xlsx",
            baseline_catalog=tmp_path / "pre.xlsx",
            compare_report=tmp_path / "report.xlsx",
            mapping=tmp_path / "mapping.xlsx",
            release="26B",
        )
        assert any("catalog" in r.lower() for r in result["regressions"])

    def test_compare_changes_delta_exceeds_threshold(self, tmp_path):
        mod = _load_module()
        _make_catalog(tmp_path / "post.xlsx", {"26B": 12000})
        _make_catalog(tmp_path / "pre.xlsx",  {"26B": 12000})
        _make_compare_report(tmp_path / "report.xlsx", 900)  # 706 ± 50 → fail
        _make_mapping(tmp_path / "mapping.xlsx", 100, 96)

        result = mod.run_checks(
            new_catalog=tmp_path / "post.xlsx",
            baseline_catalog=tmp_path / "pre.xlsx",
            compare_report=tmp_path / "report.xlsx",
            mapping=tmp_path / "mapping.xlsx",
            release="26B",
            expected_compare_changes=706,
        )
        assert any("compare" in r.lower() for r in result["regressions"])

    def test_module_pct_below_threshold(self, tmp_path):
        mod = _load_module()
        _make_catalog(tmp_path / "post.xlsx", {"26B": 12000})
        _make_catalog(tmp_path / "pre.xlsx",  {"26B": 12000})
        _make_compare_report(tmp_path / "report.xlsx", 706)
        _make_mapping(tmp_path / "mapping.xlsx", 100, 80)  # 80% < 95%

        result = mod.run_checks(
            new_catalog=tmp_path / "post.xlsx",
            baseline_catalog=tmp_path / "pre.xlsx",
            compare_report=tmp_path / "report.xlsx",
            mapping=tmp_path / "mapping.xlsx",
            release="26B",
        )
        assert any("module" in r.lower() for r in result["regressions"])

    def test_baseline_catalog_missing_skips_delta(self, tmp_path):
        """If pre-rerun catalog isn't available (first run), skip the delta check."""
        mod = _load_module()
        _make_catalog(tmp_path / "post.xlsx", {"26B": 12000})
        _make_compare_report(tmp_path / "report.xlsx", 706)
        _make_mapping(tmp_path / "mapping.xlsx", 100, 96)

        result = mod.run_checks(
            new_catalog=tmp_path / "post.xlsx",
            baseline_catalog=tmp_path / "missing.xlsx",
            compare_report=tmp_path / "report.xlsx",
            mapping=tmp_path / "mapping.xlsx",
            release="26B",
        )
        # No regression from the missing baseline; the result records that the
        # check was skipped instead.
        assert result["catalog_delta_pct"] is None
        assert "regressions" in result
```

- [ ] **Step 5.2: Run tests and verify they fail**

Run: `python -m pytest tests/test_verify_rerun.py -v`
Expected: 5 errors with `FileNotFoundError` for `verify_rerun.py`.

- [ ] **Step 5.3: Implement `verify_rerun.py`**

Create `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py`:

```python
"""Stage 8 macro-signal validator for fbdi-compare-release.

Adds checks not already covered by verify_run.py:
- Catalog row count delta (post-rerun vs pre-rerun catalog)
- Compare-report changes count vs expected baseline
- Module column population % in the working mapping spreadsheet

Never blocks. Exit 0 = clean, 1 = regression detected.
"""

from __future__ import annotations

import argparse
import json
import sys
from contextlib import closing
from pathlib import Path

from openpyxl import load_workbook


# Tunable thresholds — bump in future quarters as macro signals shift
CATALOG_DELTA_PCT_THRESHOLD = 5.0          # ±5% on per-release row count
COMPARE_CHANGES_DELTA_THRESHOLD = 50       # absolute delta around expected
DEFAULT_EXPECTED_COMPARE_CHANGES = 706     # baseline 26A→26B ground truth
MODULE_PCT_THRESHOLD = 95.0                # ≥95% rows with col A populated


def _count_release_rows(catalog_path: Path, release: str) -> int:
    """Return the number of data rows on the per-release sheet of the catalog."""
    with closing(load_workbook(catalog_path, read_only=True, data_only=True)) as wb:
        if release not in wb.sheetnames:
            return 0
        ws = wb[release]
        return max((ws.max_row or 1) - 1, 0)


def _count_compare_changes(report_path: Path) -> int:
    with closing(load_workbook(report_path, read_only=True, data_only=True)) as wb:
        ws = wb.active
        return max((ws.max_row or 1) - 1, 0)


def _module_pct(mapping_path: Path) -> tuple[float, int, int]:
    """Compute Module column population %: rows with col A non-blank
    are the denominator; rows with col F non-blank are the numerator.
    Returns (pct, populated, total)."""
    with closing(load_workbook(mapping_path, read_only=True, data_only=True)) as wb:
        if "FBDI Mapping" not in wb.sheetnames:
            return 0.0, 0, 0
        ws = wb["FBDI Mapping"]
        total = 0
        populated = 0
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            if row[0]:  # col A non-blank
                total += 1
                if len(row) >= 6 and row[5]:
                    populated += 1
    pct = (populated / total * 100.0) if total > 0 else 0.0
    return pct, populated, total


def run_checks(
    new_catalog: Path,
    baseline_catalog: Path,
    compare_report: Path,
    mapping: Path,
    release: str,
    expected_compare_changes: int = DEFAULT_EXPECTED_COMPARE_CHANGES,
) -> dict:
    """Run all macro checks. Returns a JSON-serializable dict."""
    regressions: list[str] = []

    # Catalog row delta — only if baseline exists
    delta_pct = None
    new_rows = _count_release_rows(new_catalog, release)
    if baseline_catalog.is_file():
        baseline_rows = _count_release_rows(baseline_catalog, release)
        if baseline_rows > 0:
            delta_pct = (new_rows - baseline_rows) / baseline_rows * 100.0
            if abs(delta_pct) > CATALOG_DELTA_PCT_THRESHOLD:
                regressions.append(
                    f"Catalog row count for {release}: {new_rows} vs baseline "
                    f"{baseline_rows} ({delta_pct:+.1f}%, threshold ±{CATALOG_DELTA_PCT_THRESHOLD}%)"
                )

    # Compare changes delta
    changes = _count_compare_changes(compare_report) if compare_report.is_file() else None
    if changes is not None:
        delta = abs(changes - expected_compare_changes)
        if delta > COMPARE_CHANGES_DELTA_THRESHOLD:
            regressions.append(
                f"Compare report changes: {changes} vs expected ~{expected_compare_changes} "
                f"(±{COMPARE_CHANGES_DELTA_THRESHOLD})"
            )

    # Module pct populated
    module_pct, populated, total = _module_pct(mapping) if mapping.is_file() else (None, 0, 0)
    if module_pct is not None and module_pct < MODULE_PCT_THRESHOLD:
        regressions.append(
            f"Module column populated: {module_pct:.1f}% ({populated}/{total}) "
            f"vs threshold ≥{MODULE_PCT_THRESHOLD}%"
        )

    return {
        "release": release,
        "catalog_rows_new": new_rows,
        "catalog_delta_pct": delta_pct,
        "compare_changes": changes,
        "expected_compare_changes": expected_compare_changes,
        "module_pct_populated": module_pct,
        "regressions": regressions,
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Stage 8 macro-signal validator")
    parser.add_argument("--release", required=True)
    parser.add_argument("--new-catalog", type=Path,
                        default=Path("FBDI_Master_Catalog.xlsx"))
    parser.add_argument("--baseline-catalog", type=Path,
                        default=Path("FBDI_Master_Catalog.bak.xlsx"),
                        help="Pre-rerun catalog snapshot for delta check")
    parser.add_argument("--compare-report", type=Path,
                        help="e.g. Comparison_Report_26A_26B.xlsx")
    parser.add_argument("--mapping", type=Path,
                        default=Path("FBDI_to_ApplaudTables_Mapping.xlsx"))
    parser.add_argument("--expected-compare-changes", type=int,
                        default=DEFAULT_EXPECTED_COMPARE_CHANGES)
    args = parser.parse_args(argv)

    report_path = args.compare_report or Path(f"Comparison_Report_*_{args.release}.xlsx")

    result = run_checks(
        new_catalog=args.new_catalog,
        baseline_catalog=args.baseline_catalog,
        compare_report=report_path,
        mapping=args.mapping,
        release=args.release.upper(),
        expected_compare_changes=args.expected_compare_changes,
    )
    print(json.dumps(result, indent=2))
    return 1 if result["regressions"] else 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 5.4: Run tests and verify they pass**

Run: `python -m pytest tests/test_verify_rerun.py -v`
Expected: 5 passed.

- [ ] **Step 5.5: Run the full suite**

Run: `python -m pytest tests/ -q`
Expected: all green, ~275 passed.

- [ ] **Step 5.6: Commit**

```bash
git add .claude/skills/fbdi-compare-release/scripts/verify_rerun.py tests/test_verify_rerun.py
git commit -m "$(cat <<'EOF'
feat(skill): add verify_rerun.py macro-signal validator

Adds Stage 8 checks not already covered by verify_run.py:
- catalog row count delta vs pre-rerun snapshot (±5%)
- compare-report changes count vs expected baseline (706 ± 50)
- Module column population % in the working mapping (≥95%)

Never blocks; surfaces regressions as warnings. Thresholds are
configurable constants at the top of the script for future tuning.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

### Task 6: Skill Stage 6.5 + verify_rerun integration

Wedge a new Stage 6.5 (populate-module) into the skill, with HITL #7 (backup mapping file) and the verify_rerun.py invocation in Stage 8.

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/SKILL.md` (insert Stage 6.5 after Stage 6, extend Stage 8)

- [ ] **Step 6.1: Insert Stage 6.5 between Stage 6 and Stage 7**

In `.claude/skills/fbdi-compare-release/SKILL.md`, locate the line `## Stage 7 — Summary` (around line 244). Immediately before it, insert:

````markdown
## Stage 6.5 — Populate Module column in mapping spreadsheet

If `FBDI_to_ApplaudTables_Mapping.xlsx` is absent at the repo root, skip
this stage with the notice "Module column population skipped — mapping
file not present" and proceed to Stage 7. This is a feature, not a
failure: the file may not be checked out locally.

**HITL #7 — backup before overwrite:** Ask the user:

> "About to update the Module column in
> `FBDI_to_ApplaudTables_Mapping.xlsx` based on
> `baselines/<NEW>/file_modules.json` and
> `baselines/<OLD>/file_modules.json`. Backup first?
>   (a) Yes, copy to `FBDI_to_ApplaudTables_Mapping.bak.xlsx` [default]
>   (b) No, just go (the file is git-tracked, you can revert)"

On (a): `cp FBDI_to_ApplaudTables_Mapping.xlsx FBDI_to_ApplaudTables_Mapping.bak.xlsx`.
If a backup with that name already exists, append a timestamp:
`FBDI_to_ApplaudTables_Mapping.bak.<YYYYMMDD-HHMMSS>.xlsx`.

Then run:

```
python -m fbdi populate-module --new <NEW> --old <OLD>
```

Expected exit codes:
- `0` → JSON summary printed (populated/blank/overwritten counts).
  Capture this for Stage 7's summary.
- `2` → `file_modules.json` missing for one or both releases. This means
  Stage 3 didn't complete cleanly for that release. Halt; surface the
  error to the user.
- `3` → mapping spreadsheet is open in Excel. Ask the user to close it,
  then retry once.
````

- [ ] **Step 6.2: Extend Stage 7 summary template**

In Stage 7 of `SKILL.md`, find the current summary template (the heredoc-style example after "Render the JSON as a human-readable summary"). Add a new section **after** "Catalog: ..." and before the "Stage 4 timeouts" section:

```
Module column update (Stage 6.5):
  populated: <populated>, blank: <blank>, overwritten: <overwritten>
  mapping file: FBDI_to_ApplaudTables_Mapping.xlsx
  backup:       FBDI_to_ApplaudTables_Mapping.bak.xlsx
```

If Stage 6.5 was skipped (mapping file absent), omit this whole section.

- [ ] **Step 6.3: Add verify_rerun.py invocation to Stage 8**

In Stage 8 of `SKILL.md`, after the existing `verify_run.py` invocation block, add:

````markdown
Then run the macro-signal validator:

```
python .claude/skills/fbdi-compare-release/scripts/verify_rerun.py \
  --release <NEW> \
  --compare-report Comparison_Report_<OLD>_<NEW>.xlsx \
  --baseline-catalog FBDI_Master_Catalog.bak.xlsx
```

(`FBDI_Master_Catalog.bak.xlsx` is created at the start of Stage 6 by
copying the existing catalog before regeneration. If absent, the
catalog-delta check is skipped — not a failure.)

If `verify_rerun.py` exits 1, append the regression list to the Stage 7
summary as warnings (do not fail the skill).
````

- [ ] **Step 6.4: Add the catalog-backup step to Stage 6**

In Stage 6 of `SKILL.md`, before the `python -m fbdi catalog` invocation, add a setup line:

```
# Snapshot the existing catalog so verify_rerun.py can compare deltas.
cp FBDI_Master_Catalog.xlsx FBDI_Master_Catalog.bak.xlsx
```

If `FBDI_Master_Catalog.xlsx` doesn't exist, skip this line silently.

- [ ] **Step 6.5: Verify SKILL.md is well-formed**

Run: `python -m pytest tests/test_skill_scripts.py -q` (the skill scripts have a smoke test suite).
Expected: still passing.

Visual check: open `.claude/skills/fbdi-compare-release/SKILL.md` and confirm:
- Stage 6.5 appears between Stage 6 and Stage 7.
- HITL #7 numbering is consistent.
- The new verify_rerun.py invocation in Stage 8 appears after the existing verify_run.py block.

- [ ] **Step 6.6: Commit**

```bash
git add .claude/skills/fbdi-compare-release/SKILL.md
git commit -m "$(cat <<'EOF'
feat(skill): add Stage 6.5 (populate Module column) + verify_rerun

New Stage 6.5 wedged between catalog and summary: prompts for backup
(HITL #7), runs python -m fbdi populate-module, captures the summary
for Stage 7. Stage 8 now also runs verify_rerun.py for macro-signal
checks (catalog row delta, compare changes delta, module column %).

Stage 6 backs up the existing catalog as FBDI_Master_Catalog.bak.xlsx
so the delta check has a baseline.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

---

## Phase 1 wrap-up

- [ ] **Run the full test suite one final time**

Run: `python -m pytest tests/ -q`
Expected: all tests pass. Test count should be 255 (existing) + 6 (Task 1+2) + 8 (Task 3) + 1 (Task 4) + 5 (Task 5) = ~275.

- [ ] **Push Phase 1 to master**

```bash
git push origin master
```

This commits the code changes before the rerun, so any rollback during Phase 2 can `git revert` cleanly.

---

## Phase 2 — Rerun

This is a single execution event, not a code change. Walk through the steps below interactively. Estimated wall-clock time: 2–3 hours, mostly unattended downloads.

### Task 7: Wipe baselines, run skill, validate, commit

**Files:**
- Wipe: `baselines/26A/` and `baselines/26B/` (originals + blanks)
- Created during run: `baselines/26A/file_modules.json`, `baselines/26B/file_modules.json`
- Modified during run: `Comparison_Report_26A_26B.xlsx`, `FBDI_Master_Catalog.xlsx`, `FBDI_to_ApplaudTables_Mapping.xlsx`
- Created during run: `FBDI_to_ApplaudTables_Mapping.bak.xlsx`, `FBDI_Master_Catalog.bak.xlsx`

- [ ] **Step 7.1: Confirm clean working tree and Phase 1 pushed**

Run: `git status` → working tree clean.
Run: `git log origin/master..HEAD` → empty (Phase 1 already pushed).

If either fails, resolve before proceeding.

- [ ] **Step 7.2: Snapshot the existing FBDI_to_ApplaudTables_Mapping.xlsx outside the rerun flow**

This is belt-and-suspenders — Stage 6.5 will also offer a backup, but having one *before* the rerun even starts means the Phase 2 commit can be reverted to a known-good state if anything catastrophic happens.

```bash
cp FBDI_to_ApplaudTables_Mapping.xlsx FBDI_to_ApplaudTables_Mapping.pre-rerun.xlsx
```

- [ ] **Step 7.3: Wipe baselines**

Confirm with the user before destructive action:

> "About to delete baselines/26A/originals/, baselines/26A/blanks/, baselines/26B/originals/, and baselines/26B/blanks/. ~425 .xlsm files will be re-downloaded over the next ~70-100 minutes. Confirm? (y/N)"

On confirmation:

```bash
rm -rf baselines/26A/originals baselines/26A/blanks
rm -rf baselines/26B/originals baselines/26B/blanks
mkdir -p baselines/26A/originals baselines/26A/blanks
mkdir -p baselines/26B/originals baselines/26B/blanks
```

- [ ] **Step 7.4: Disable Windows sleep/lock if user is stepping away**

Tell the user: "Selenium will run foreground for the next ~2 hours. If you're stepping away, disable sleep/lock now: Settings → System → Power → Screen and sleep → set both to Never."

- [ ] **Step 7.5: Invoke the skill**

The user types: `compare 26A to 26B`

This triggers `fbdi-compare-release`. The skill walks through 8 stages plus Stage 6.5. Use HITL prompts as designed:
- HITL #1 (26A baseline missing): user chooses "download both"
- HITL #2 if `RapidImplementationForCashManagement.xlsm` absent after 26B download: user copies from 26A or downloads via FSM
- HITL #7 (backup mapping?): user picks (a) yes
- Other HITLs as they arise

- [ ] **Step 7.6: Walk through stages with active monitoring**

Watch for:
- **Stage 3 retries** (HITL #5 trips at 3 attempts) — if downloads keep failing for one module, halt and debug `tools/download_and_clear.py` directly.
- **Stage 4 timeouts** — captured by the skill, surfaced in Stage 7 summary; manual clear required for any timed-out files. Not a blocker.
- **Stage 6.5 result** — confirm `populated` count is high (>95% of mapped rows). If `blank` count is suspiciously high (>30), pause and check `baselines/<ver>/file_modules.json` actually has entries for those files.
- **Stage 8 verify_rerun output** — review JSON. If any regression flagged, decide whether to investigate or accept.

- [ ] **Step 7.7: Sanity-check outputs by hand**

```bash
ls -la baselines/26A/file_modules.json baselines/26B/file_modules.json
python -c "import json; d=json.load(open('baselines/26B/file_modules.json')); print(f'{len(d)} entries; first 3: {dict(list(d.items())[:3])}')"
```

Expected: ~213 entries for 26B, ~212 for 26A.

```bash
python -c "
from openpyxl import load_workbook
wb = load_workbook('FBDI_to_ApplaudTables_Mapping.xlsx', read_only=True)
ws = wb['FBDI Mapping']
total = 0; populated = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    if row[0]:
        total += 1
        if len(row) >= 6 and row[5]: populated += 1
print(f'mapping rows with FBDI Template: {total}, with Module: {populated} ({populated*100/total:.1f}%)')"
```

Expected: total ~639, populated ≥607 (95%).

- [ ] **Step 7.8: Commit Phase 2**

First check what's actually changed and tracked. `baselines/` is gitignored, so it won't show up. The interesting modified files are the regenerated artifacts:

```bash
git status
```

Stage only the files that are tracked AND modified (not all of these may be tracked — `Comparison_Report_*.xlsx` in particular may or may not be tracked depending on `.gitignore`):

```bash
# Stage whichever of these git status reported as modified or untracked
git add FBDI_Master_Catalog.xlsx FBDI_to_ApplaudTables_Mapping.xlsx
# Only add the comparison report if git status shows it as modified/untracked
# (skip if it's gitignored)
git add Comparison_Report_26A_26B.xlsx 2>/dev/null || true
git status   # confirm staged file list looks right
```

```bash
git commit -m "$(cat <<'EOF'
chore(rerun): regenerate 26A/26B baselines, comparison, catalog, mapping

Full skill-driven rerun against fresh 26A and 26B downloads to validate
the fbdi-compare-release skill end-to-end and the detect_header.py fix.
Module column populated for the first time via the new Stage 6.5
populate-module path.

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 7.9: Push Phase 2**

```bash
git push origin master
```

- [ ] **Step 7.10: Clean up scratch backups**

The pre-rerun safety copy is no longer needed once Phase 2 is pushed:

```bash
rm -f FBDI_to_ApplaudTables_Mapping.pre-rerun.xlsx
rm -f FBDI_to_ApplaudTables_Mapping.bak.xlsx
rm -f FBDI_Master_Catalog.bak.xlsx
```

These files are gitignored already; deleting them just keeps the working tree tidy.

---

## Acceptance criteria checklist

Before declaring complete, confirm all are true:

- [ ] All ~275 tests pass (`python -m pytest tests/ -q`)
- [ ] `git log` shows Phase 1 and Phase 2 commits pushed to origin/master
- [ ] `baselines/26A/file_modules.json` exists with ≥210 entries
- [ ] `baselines/26B/file_modules.json` exists with ≥210 entries
- [ ] `FBDI_to_ApplaudTables_Mapping.xlsx` Module column populated for ≥95% of rows where col A is non-blank
- [ ] `verify_rerun.py` reports `regressions: []` (or only acceptable warnings)
- [ ] `verify_run.py` reports `overall_regression: false` (NO_HEADER == 0, Issues count not regressed)
- [ ] No unrelated cells in `FBDI_to_ApplaudTables_Mapping.xlsx` were modified (visual spot-check 5–10 rows pre-vs-post — `git diff HEAD~1 -- FBDI_to_ApplaudTables_Mapping.xlsx` won't show cell-level diffs since it's a binary, so eyeball with Excel and the `pre-rerun` backup if needed)
