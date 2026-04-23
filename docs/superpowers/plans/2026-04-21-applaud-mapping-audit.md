# Applaud Mapping Audit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `fbdi/audit.py` — a deterministic Python audit engine that consumes an `applaud_snapshot.json` (MDB extract), the 26B FBDI catalog, and the prior `fbdi_applaud_mapping.xlsx` to produce `Claude_fbdi_applaud_mapping.xlsx` (3 sheets) and `Claude_fbdi_applaud_mapping_audit.md` with a full evidence-backed re-verdict on all 183 Applaud T_ tables.

**Architecture:** Two-step pipeline: Step A (Claude + applaud-mcp extracts the MDB to `applaud_snapshot.json` — agent-driven, one-time) → Step B (deterministic `fbdi/audit.py` reads snapshot + catalog + prior mapping and runs two passes: Pass 1 builds a candidate index from signal scores, Pass 2 adjudicates each Applaud table against the rubric and emits rows with verdict, confidence, and rationale).

**Tech Stack:** Python 3.14+, openpyxl (read catalog + prior mapping, write output xlsx), json (snapshot), pytest (139 → ~175 tests), applaud-mcp (Step A only).

---

## Pre-Flight: Inspect Prior Mapping Sheet2 Column Headers

Before writing any loader code, inspect the actual column headers in `fbdi_applaud_mapping.xlsx` Sheet2 so all column-index assumptions in the plan are verified. Do this first — it is 2 minutes and prevents wasted implementation.

```python
from openpyxl import load_workbook
wb = load_workbook("fbdi_applaud_mapping.xlsx", read_only=True, data_only=True)
print(wb.sheetnames)
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    first_row = next(ws.iter_rows(values_only=True))
    print(f"\n{sheet_name}: {list(first_row)}")
wb.close()
```

Run: `python -c "from openpyxl import load_workbook; wb = load_workbook('fbdi_applaud_mapping.xlsx', read_only=True, data_only=True); [print(n, list(next(wb[n].iter_rows(values_only=True)))) for n in wb.sheetnames]; wb.close()"`

Note the exact column names. The plan uses these assumed headers for Sheet2; adjust if actuals differ:
- `applaud_table` (col B)
- `status` (col C) — may be an XLOOKUP formula; data_only=True returns computed value
- `prefix` (col D)
- `fbdi_template_mappings` (col E)
- `module` (col F)
- `notes` (col G)

---

## File Structure

| Action | Path | Responsibility |
|---|---|---|
| Create | `fbdi/audit.py` | Data classes, loaders, signal functions, Pass 1, Pass 2, output writers, `run_audit` |
| Create | `tests/test_audit.py` | All unit + integration tests (~40 new) |
| Produce | `applaud_snapshot.json` | MDB extract from Step A (agent-driven; checked into repo) |
| Produce | `Claude_fbdi_applaud_mapping.xlsx` | 3-sheet audit output |
| Produce | `Claude_fbdi_applaud_mapping_audit.md` | Prose sidecar for NEEDS_REVIEW / changed rows |

No existing files are modified. `fbdi/audit.py` is a new standalone module alongside `build_mapping.py`.

---

## Task 1: Scaffold — data classes + empty test file

**Files:**
- Create: `fbdi/audit.py`
- Create: `tests/test_audit.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit.py
from fbdi.audit import (
    SnapshotField, SnapshotKeySeq, SnapshotTable, ApplaudSnapshot,
    Candidate, EvidenceBundle, PriorRow, AuditRow,
)

def test_data_classes_importable():
    field = SnapshotField(
        name="TA4INVOICE_ID",
        bare_name="INVOICE_ID",
        is_legacy_tracking=False,
        data_type="N",
        length=15,
    )
    assert field.bare_name == "INVOICE_ID"
    assert not field.is_legacy_tracking

def test_legacy_tracking_field():
    field = SnapshotField(
        name="@TA4SITE",
        bare_name="SITE",
        is_legacy_tracking=True,
        data_type="X",
        length=10,
    )
    assert field.is_legacy_tracking
    assert field.bare_name == "SITE"
```

- [ ] **Step 2: Run test — verify it fails**

```
python -m pytest tests/test_audit.py -v
```
Expected: `ImportError` (module doesn't exist yet).

- [ ] **Step 3: Create `fbdi/audit.py` with data classes**

```python
"""
audit.py — FBDI ↔ Applaud mapping audit engine.

Consumes applaud_snapshot.json + FBDI_Master_Catalog.xlsx (26B) +
fbdi_applaud_mapping.xlsx and produces Claude_fbdi_applaud_mapping.xlsx
(3 sheets) + Claude_fbdi_applaud_mapping_audit.md.

Run: python -m fbdi.audit
"""
from __future__ import annotations

import json
import re
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


REPO_ROOT = Path(__file__).parent.parent
SNAPSHOT_PATH = REPO_ROOT / "applaud_snapshot.json"
CATALOG_PATH = REPO_ROOT / "FBDI_Master_Catalog.xlsx"
PRIOR_MAPPING_PATH = REPO_ROOT / "fbdi_applaud_mapping.xlsx"
OUTPUT_MAPPING_PATH = REPO_ROOT / "Claude_fbdi_applaud_mapping.xlsx"
OUTPUT_AUDIT_PATH = REPO_ROOT / "Claude_fbdi_applaud_mapping_audit.md"
CATALOG_RELEASE = "26B"
SNAPSHOT_MAX_AGE_DAYS = 30

_STRIP_SUFFIXES = ("_ALL", "_INT", "_INTERFACE")


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class SnapshotField:
    name: str
    bare_name: str
    is_legacy_tracking: bool
    data_type: str
    length: int


@dataclass
class SnapshotKeySeq:
    seq: str
    keys: list[str]


@dataclass
class SnapshotTable:
    name: str
    prefix: str | None
    description: str
    type: str
    key_sequences: list[SnapshotKeySeq]
    fields: list[SnapshotField]

    def business_fields(self) -> list[SnapshotField]:
        return [f for f in self.fields if not f.is_legacy_tracking]

    def key_bare_names(self) -> set[str]:
        bare: set[str] = set()
        for seq in self.key_sequences:
            for k in seq.keys:
                # Keys are stored as full prefixed names; strip prefix
                if self.prefix and k.upper().startswith(self.prefix.upper()):
                    bare.add(k[len(self.prefix):].upper())
                else:
                    bare.add(k.upper())
        return bare


@dataclass
class ApplaudSnapshot:
    mdb_path: str
    extracted_at: str
    extractor_version: str
    tables: list[SnapshotTable]
    missing_tables: list[dict]

    def table_by_name(self) -> dict[str, SnapshotTable]:
        return {t.name: t for t in self.tables}

    def missing_set(self) -> set[str]:
        return {m["name"] for m in self.missing_tables}


@dataclass
class Candidate:
    fbdi_file: str
    fbdi_tab: str
    name_alignment: str           # EXACT | PARTIAL | NONE
    key_coverage: float
    column_overlap: float
    prefix_conformance: bool
    applaud_key_fields_matched: list[str]
    applaud_fields_matched: list[str]
    applaud_fields_missing: list[str]


@dataclass
class EvidenceBundle:
    candidates_evaluated: list[Candidate] = field(default_factory=list)
    rejected_alternatives: list[Candidate] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass
class PriorRow:
    applaud_table: str
    prior_status: str
    prefix: str
    mapping_text: str
    module: str
    notes: str


@dataclass
class AuditRow:
    applaud_table: str
    prefix: str
    verdict: str                   # YES | UNMAPPED | NEEDS_REVIEW | FILE_TOO_LARGE | FILE_ERROR
    fbdi_mapping: str
    confidence: str                # H | M | L | ""
    rationale: str
    prior_verdict: str
    changed: bool
    needs_deep_rationale: bool
    evidence: EvidenceBundle


# Type aliases
CatalogIndex = dict[tuple[str, str], set[str]]   # {(file_name, tab_name): set[column_technical]}
CandidateIndex = dict[str, list[Candidate]]       # {applaud_table_name: sorted candidates}
```

- [ ] **Step 4: Run test — verify it passes**

```
python -m pytest tests/test_audit.py::test_data_classes_importable tests/test_audit.py::test_legacy_tracking_field -v
```
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): scaffold data classes for applaud mapping audit"
```

---

## Task 2: Snapshot + catalog + prior-mapping loaders

**Files:**
- Modify: `fbdi/audit.py` — add `load_snapshot`, `load_catalog`, `load_prior_mapping`
- Modify: `tests/test_audit.py` — add loader tests

- [ ] **Step 1: Write failing loader tests**

```python
# Add to tests/test_audit.py
import json
import tempfile
from pathlib import Path
from openpyxl import Workbook

from fbdi.audit import load_snapshot, load_catalog, load_prior_mapping, ApplaudSnapshot, CatalogIndex


def _make_snapshot_dict(tables=None, missing=None) -> dict:
    return {
        "mdb_path": "C:/test/AP0STE.mdb",
        "extracted_at": "2026-04-21T12:00:00Z",
        "extractor_version": "1",
        "tables": tables or [],
        "missing_tables": missing or [],
    }


def test_load_snapshot_basic(tmp_path):
    snap_data = _make_snapshot_dict(tables=[{
        "name": "T_RA_INTERFACE_LINES_ALL",
        "prefix": "TA4",
        "description": "T_RA_INTERFACE_LINES_ALL (TA4)",
        "type": "1",
        "key_sequences": [{"seq": "1", "keys": ["TA4INVOICE_ID"]}],
        "fields": [
            {"name": "TA4INVOICE_ID", "bare_name": "INVOICE_ID",
             "is_legacy_tracking": False, "data_type": "N", "length": 15},
            {"name": "@TA4SITE", "bare_name": "SITE",
             "is_legacy_tracking": True, "data_type": "X", "length": 10},
        ],
    }])
    snap_file = tmp_path / "applaud_snapshot.json"
    snap_file.write_text(json.dumps(snap_data))
    snap = load_snapshot(snap_file)
    assert isinstance(snap, ApplaudSnapshot)
    assert len(snap.tables) == 1
    t = snap.tables[0]
    assert t.name == "T_RA_INTERFACE_LINES_ALL"
    assert t.prefix == "TA4"
    assert len(t.fields) == 2
    biz = t.business_fields()
    assert len(biz) == 1
    assert biz[0].bare_name == "INVOICE_ID"


def test_load_snapshot_missing_file(tmp_path):
    import pytest
    with pytest.raises(FileNotFoundError):
        load_snapshot(tmp_path / "missing.json")


def _make_catalog_xlsx(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "26B"
    ws.append(["release", "file_name", "tab_name", "position",
                "column_label", "column_technical",
                "data_type", "length", "scale", "data_type_raw", "required"])
    ws.append(["26B", "AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL",
                1, "Invoice ID", "INVOICE_ID", "N", 15, None, "NUMBER(15)", "TRUE"])
    ws.append(["26B", "AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL",
                2, "Batch Source Name", "BATCH_SOURCE_NAME", "X", 50, None, "VARCHAR2(50)", "FALSE"])
    path = tmp_path / "FBDI_Master_Catalog.xlsx"
    wb.save(path)
    return path


def test_load_catalog_basic(tmp_path):
    path = _make_catalog_xlsx(tmp_path)
    index = load_catalog(path, release="26B")
    key = ("AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL")
    assert key in index
    assert "INVOICE_ID" in index[key]
    assert "BATCH_SOURCE_NAME" in index[key]


def test_load_catalog_missing_release(tmp_path):
    import pytest
    path = _make_catalog_xlsx(tmp_path)
    with pytest.raises(ValueError, match="25D"):
        load_catalog(path, release="25D")


def _make_prior_mapping_xlsx(tmp_path: Path) -> Path:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "FBDI Mapping"
    ws2 = wb.create_sheet("Applaud Tables")
    ws2.append(["#", "applaud_table", "status", "prefix",
                 "fbdi_template_mappings", "module", "notes"])
    ws2.append([1, "T_RA_INTERFACE_LINES_ALL", "YES", "TA4",
                 "AutoInvoiceImportTemplate / RA_INTERFACE_LINES_ALL", "Financials", ""])
    ws2.append([2, "T_GHOST_TABLE", "UNMAPPED", "", "", "HR", "no match found"])
    path = tmp_path / "fbdi_applaud_mapping.xlsx"
    wb.save(path)
    return path


def test_load_prior_mapping_basic(tmp_path):
    path = _make_prior_mapping_xlsx(tmp_path)
    mapping = load_prior_mapping(path)
    assert "T_RA_INTERFACE_LINES_ALL" in mapping
    row = mapping["T_RA_INTERFACE_LINES_ALL"]
    assert row.prior_status == "YES"
    assert row.prefix == "TA4"
    assert "AutoInvoiceImportTemplate" in row.mapping_text
    assert "T_GHOST_TABLE" in mapping
    assert mapping["T_GHOST_TABLE"].prior_status == "UNMAPPED"
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py -k "loader or snapshot or catalog or prior" -v
```
Expected: `ImportError` — functions not defined yet.

- [ ] **Step 3: Add loader functions to `fbdi/audit.py`**

After the data classes block, append:

```python
# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_snapshot(path: Path = SNAPSHOT_PATH) -> ApplaudSnapshot:
    if not path.exists():
        raise FileNotFoundError(f"Snapshot missing — run Step A first: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    tables = []
    for t in data["tables"]:
        fields = [
            SnapshotField(
                name=f["name"],
                bare_name=f["bare_name"],
                is_legacy_tracking=f["is_legacy_tracking"],
                data_type=f["data_type"],
                length=f["length"],
            )
            for f in t["fields"]
        ]
        key_seqs = [
            SnapshotKeySeq(seq=k["seq"], keys=k["keys"])
            for k in t["key_sequences"]
        ]
        tables.append(SnapshotTable(
            name=t["name"],
            prefix=t.get("prefix"),
            description=t.get("description", ""),
            type=t.get("type", ""),
            key_sequences=key_seqs,
            fields=fields,
        ))
    return ApplaudSnapshot(
        mdb_path=data["mdb_path"],
        extracted_at=data["extracted_at"],
        extractor_version=data["extractor_version"],
        tables=tables,
        missing_tables=data.get("missing_tables", []),
    )


def load_catalog(
    path: Path = CATALOG_PATH, release: str = CATALOG_RELEASE
) -> CatalogIndex:
    if not path.exists():
        raise FileNotFoundError(f"Catalog missing: {path}")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if release not in wb.sheetnames:
            raise ValueError(f"No '{release}' tab in catalog. Available: {wb.sheetnames}")
        ws = wb[release]
        rows_iter = ws.iter_rows(values_only=True)
        raw_headers = next(rows_iter)
        headers = [str(h).strip().lower() if h else "" for h in raw_headers]
        try:
            file_col = headers.index("file_name")
            tab_col = headers.index("tab_name")
            tech_col = headers.index("column_technical")
        except ValueError as exc:
            raise ValueError(f"Catalog missing expected header: {exc}. Got: {headers}")
        index: CatalogIndex = {}
        for row in rows_iter:
            fname = str(row[file_col]).strip() if row[file_col] else ""
            tab = str(row[tab_col]).strip() if row[tab_col] else ""
            tech = str(row[tech_col]).strip() if row[tech_col] else ""
            if fname and tab:
                key = (fname, tab)
                index.setdefault(key, set())
                if tech:
                    index[key].add(tech.upper())
        return index
    finally:
        wb.close()


def load_prior_mapping(path: Path = PRIOR_MAPPING_PATH) -> dict[str, PriorRow]:
    if not path.exists():
        raise FileNotFoundError(f"Prior mapping missing: {path}")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        # Find Applaud Tables sheet by name (case-insensitive), fall back to index 1
        ws = None
        for name in wb.sheetnames:
            if "applaud" in name.lower():
                ws = wb[name]
                break
        if ws is None:
            if len(wb.sheetnames) >= 2:
                ws = wb.worksheets[1]
            else:
                raise ValueError(
                    f"No 'Applaud Tables' sheet found. Sheets: {wb.sheetnames}"
                )
        rows_iter = ws.iter_rows(values_only=True)
        raw_headers = next(rows_iter)
        headers = [
            str(h).strip().lower().replace(" ", "_") if h else ""
            for h in raw_headers
        ]

        def _col(name: str) -> int:
            return headers.index(name) if name in headers else -1

        col_table = _col("applaud_table")
        col_status = _col("status")
        col_prefix = _col("prefix")
        col_mapping = _col("fbdi_template_mappings")
        col_module = _col("module")
        col_notes = _col("notes")

        if col_table == -1:
            raise ValueError(
                f"Sheet2 missing 'applaud_table' column. Headers found: {headers}"
            )

        result: dict[str, PriorRow] = {}
        for row in rows_iter:
            def _val(idx: int) -> str:
                if idx == -1 or idx >= len(row):
                    return ""
                v = row[idx]
                return str(v).strip() if v is not None else ""

            table_name = _val(col_table)
            if not table_name or table_name.startswith("#"):
                continue
            result[table_name] = PriorRow(
                applaud_table=table_name,
                prior_status=_val(col_status),
                prefix=_val(col_prefix),
                mapping_text=_val(col_mapping),
                module=_val(col_module),
                notes=_val(col_notes),
            )
        return result
    finally:
        wb.close()
```

- [ ] **Step 4: Run — verify passing**

```
python -m pytest tests/test_audit.py -v
```
Expected: all tests pass (≥7 tests so far).

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): add snapshot/catalog/prior-mapping loaders"
```

---

## Task 3: Prefix extraction + bare_name derivation

**Files:**
- Modify: `fbdi/audit.py` — add `extract_prefix`, `derive_bare_name`
- Modify: `tests/test_audit.py` — add 5 tests

- [ ] **Step 1: Write failing tests**

```python
from fbdi.audit import extract_prefix, derive_bare_name

def test_extract_prefix_standard():
    assert extract_prefix("T_RA_INTERFACE_LINES_ALL (TA4)") == "TA4"

def test_extract_prefix_alphanumeric():
    assert extract_prefix("T_EGP_COMPONENTS_INTERFACE (T91)") == "T91"

def test_extract_prefix_no_parens():
    assert extract_prefix("T_GHOST_TABLE") is None

def test_derive_bare_name_regular():
    bare, is_legacy = derive_bare_name("TA4INVOICE_ID", "TA4")
    assert bare == "INVOICE_ID"
    assert not is_legacy

def test_derive_bare_name_at_prefix():
    bare, is_legacy = derive_bare_name("@TA4SITE", "TA4")
    assert bare == "SITE"
    assert is_legacy

def test_derive_bare_name_no_prefix_match():
    # Field that doesn't start with prefix keeps full name
    bare, is_legacy = derive_bare_name("SOMETHING_ELSE", "TA4")
    assert bare == "SOMETHING_ELSE"
    assert not is_legacy
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py -k "prefix or bare_name" -v
```
Expected: `ImportError`.

- [ ] **Step 3: Add functions to `fbdi/audit.py`**

After the loaders block:

```python
# ---------------------------------------------------------------------------
# Prefix + bare_name utilities
# ---------------------------------------------------------------------------

_PREFIX_RE = re.compile(r'\(([A-Z0-9]+)\)\s*$')


def extract_prefix(description: str) -> str | None:
    m = _PREFIX_RE.search(description.strip())
    return m.group(1) if m else None


def derive_bare_name(field_name: str, prefix: str) -> tuple[str, bool]:
    """Return (bare_name, is_legacy_tracking)."""
    name = field_name
    is_legacy = False
    if name.startswith("@"):
        is_legacy = True
        name = name[1:]  # strip @
    upper_prefix = prefix.upper()
    if name.upper().startswith(upper_prefix):
        return name[len(prefix):], is_legacy
    return name, is_legacy
```

- [ ] **Step 4: Run — verify passing**

```
python -m pytest tests/test_audit.py -k "prefix or bare_name" -v
```
Expected: 6 tests pass.

- [ ] **Step 5: Run full suite**

```
python -m pytest tests/ -v --tb=short
```
Expected: all 139 + new tests pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): prefix extraction and bare_name derivation"
```

---

## Task 4: Signal computation functions

**Files:**
- Modify: `fbdi/audit.py` — add `compute_name_alignment`, `compute_key_coverage`, `compute_column_overlap`, `check_prefix_conformance`
- Modify: `tests/test_audit.py` — add 12 tests

- [ ] **Step 1: Write failing tests**

```python
from fbdi.audit import (
    compute_name_alignment, compute_key_coverage,
    compute_column_overlap, check_prefix_conformance,
    SnapshotField,
)

# --- name_alignment ---

def test_name_alignment_exact():
    # Applaud: T_RA_INTERFACE_LINES_ALL → strip T_ → RA_INTERFACE_LINES_ALL
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "RA_INTERFACE_LINES_ALL") == "EXACT"

def test_name_alignment_partial_strip_all():
    # FBDI tab has _ALL suffix that is stripped for comparison
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "RA_INTERFACE_LINES") == "PARTIAL"

def test_name_alignment_partial_strip_interface():
    assert compute_name_alignment("T_RCV_HEADERS_INTERFACE", "RCV_HEADERS") == "PARTIAL"

def test_name_alignment_none():
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "TOTALLY_DIFFERENT_TAB") == "NONE"

def test_name_alignment_case_insensitive():
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "ra_interface_lines_all") == "EXACT"

# --- key_coverage ---

def test_key_coverage_full():
    keys = {"INVOICE_ID", "LINE_NUMBER"}
    fbdi_cols = {"INVOICE_ID", "LINE_NUMBER", "BATCH_SOURCE_NAME"}
    assert compute_key_coverage(keys, fbdi_cols) == 1.0

def test_key_coverage_partial():
    keys = {"INVOICE_ID", "LINE_NUMBER"}
    fbdi_cols = {"INVOICE_ID", "BATCH_SOURCE_NAME"}
    assert compute_key_coverage(keys, fbdi_cols) == 0.5

def test_key_coverage_empty_keys():
    assert compute_key_coverage(set(), {"INVOICE_ID"}) == 0.0

# --- column_overlap ---

def _biz_field(bare_name: str) -> SnapshotField:
    return SnapshotField(name=f"TA4{bare_name}", bare_name=bare_name,
                         is_legacy_tracking=False, data_type="X", length=30)

def _legacy_field(bare_name: str) -> SnapshotField:
    return SnapshotField(name=f"@TA4{bare_name}", bare_name=bare_name,
                         is_legacy_tracking=True, data_type="X", length=10)

def test_column_overlap_excludes_legacy():
    fields = [_biz_field("INVOICE_ID"), _biz_field("LINE_NUM"), _legacy_field("SITE")]
    fbdi_cols = {"INVOICE_ID", "LINE_NUM"}
    # denominator = 2 (exclude legacy); numerator = 2 matched
    assert compute_column_overlap(fields, fbdi_cols) == 1.0

def test_column_overlap_partial():
    fields = [_biz_field("INVOICE_ID"), _biz_field("LINE_NUM"), _biz_field("BATCH_NAME")]
    fbdi_cols = {"INVOICE_ID", "LINE_NUM"}
    assert abs(compute_column_overlap(fields, fbdi_cols) - 2/3) < 0.001

def test_column_overlap_all_legacy():
    fields = [_legacy_field("SITE"), _legacy_field("LEGACY_HEADER")]
    fbdi_cols = {"INVOICE_ID"}
    # No business fields → 0.0, no division by zero
    assert compute_column_overlap(fields, fbdi_cols) == 0.0

def test_column_overlap_case_insensitive():
    fields = [_biz_field("INVOICE_ID")]
    fbdi_cols = {"invoice_id"}
    assert compute_column_overlap(fields, fbdi_cols) == 1.0

# --- prefix conformance ---

def test_prefix_conformance_true():
    # T_RA_INTERFACE_LINES_ALL, prefix=TA4, FBDI tab=RA_INTERFACE_LINES_ALL
    # "expected" prefix would be derived from T_ + tab → this is diagnostic only
    # Convention: prefix_conformance=True when Applaud prefix matches the T_<tab> convention
    # We check: stripping T_ from table name and comparing to tab name gives exact match
    assert check_prefix_conformance("T_RA_INTERFACE_LINES_ALL", "TA4", "RA_INTERFACE_LINES_ALL") is True

def test_prefix_conformance_false():
    # Prefix TA4 but tab is RA_INTERFACE_LINES (not RA_INTERFACE_LINES_ALL)
    assert check_prefix_conformance("T_RA_INTERFACE_LINES_ALL", "TA4", "RA_INTERFACE_LINES") is False
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py -k "alignment or coverage or overlap or conformance" -v
```
Expected: `ImportError`.

- [ ] **Step 3: Add signal functions to `fbdi/audit.py`**

```python
# ---------------------------------------------------------------------------
# Signal computation
# ---------------------------------------------------------------------------

def compute_name_alignment(applaud_table: str, fbdi_tab: str) -> str:
    """Compare Applaud table name (strip T_) against FBDI tab name."""
    stripped = applaud_table.upper().removeprefix("T_")
    tab_upper = fbdi_tab.upper()

    if stripped == tab_upper:
        return "EXACT"

    # Try stripping suffixes from both sides
    def _base(s: str) -> str:
        for suffix in _STRIP_SUFFIXES:
            if s.endswith(suffix):
                return s[: -len(suffix)]
        return s

    if _base(stripped) == _base(tab_upper):
        return "PARTIAL"
    if _base(stripped) == tab_upper or stripped == _base(tab_upper):
        return "PARTIAL"

    return "NONE"


def compute_key_coverage(
    applaud_key_bare_names: set[str], fbdi_columns: set[str]
) -> float:
    if not applaud_key_bare_names:
        return 0.0
    fbdi_upper = {c.upper() for c in fbdi_columns}
    matched = sum(1 for k in applaud_key_bare_names if k.upper() in fbdi_upper)
    return matched / len(applaud_key_bare_names)


def compute_column_overlap(
    applaud_fields: list[SnapshotField], fbdi_columns: set[str]
) -> float:
    biz_fields = [f for f in applaud_fields if not f.is_legacy_tracking]
    if not biz_fields:
        return 0.0
    fbdi_upper = {c.upper() for c in fbdi_columns}
    matched = sum(1 for f in biz_fields if f.bare_name.upper() in fbdi_upper)
    return matched / len(biz_fields)


def check_prefix_conformance(
    applaud_table: str, prefix: str, fbdi_tab: str
) -> bool:
    """True when Applaud table name minus T_ exactly equals the FBDI tab name."""
    return applaud_table.upper().removeprefix("T_") == fbdi_tab.upper()
```

- [ ] **Step 4: Run — verify passing**

```
python -m pytest tests/test_audit.py -k "alignment or coverage or overlap or conformance" -v
```
Expected: 12 tests pass.

- [ ] **Step 5: Run full suite**

```
python -m pytest tests/ -v --tb=short
```
Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): signal computation functions (name alignment, key coverage, column overlap)"
```

---

## Task 5: Pass 1 — candidate index

**Files:**
- Modify: `fbdi/audit.py` — add `_score_candidate`, `build_candidate_index`
- Modify: `tests/test_audit.py` — add 5 tests

- [ ] **Step 1: Write failing tests**

```python
from fbdi.audit import build_candidate_index, ApplaudSnapshot, SnapshotTable, SnapshotField, SnapshotKeySeq

def _make_snapshot_table(
    name: str, prefix: str, fields: list[SnapshotField], key_fields: list[str]
) -> SnapshotTable:
    return SnapshotTable(
        name=name, prefix=prefix,
        description=f"{name} ({prefix})", type="1",
        key_sequences=[SnapshotKeySeq(seq="1", keys=key_fields)],
        fields=fields,
    )

def _make_snap(*tables: SnapshotTable) -> ApplaudSnapshot:
    return ApplaudSnapshot(
        mdb_path="", extracted_at="", extractor_version="1",
        tables=list(tables), missing_tables=[],
    )


def test_pass1_exact_name_match_kept():
    fields = [SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)]
    table = _make_snapshot_table("T_RA_INTERFACE_LINES_ALL", "TA4", fields, ["TA4INVOICE_ID"])
    catalog = {("AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL"): {"INVOICE_ID"}}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    assert "T_RA_INTERFACE_LINES_ALL" in idx
    candidates = idx["T_RA_INTERFACE_LINES_ALL"]
    assert len(candidates) == 1
    c = candidates[0]
    assert c.fbdi_file == "AutoInvoiceImportTemplate"
    assert c.fbdi_tab == "RA_INTERFACE_LINES_ALL"
    assert c.name_alignment == "EXACT"


def test_pass1_below_threshold_dropped():
    # No name match, key_coverage=0, column_overlap=0 → should not appear
    fields = [SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)]
    table = _make_snapshot_table("T_RA_INTERFACE_LINES_ALL", "TA4", fields, ["TA4INVOICE_ID"])
    # Catalog has a completely different tab
    catalog = {("SomeTemplate", "UNRELATED_TAB"): {"UNRELATED_COL"}}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    # No candidate meets any threshold
    assert idx.get("T_RA_INTERFACE_LINES_ALL", []) == []


def test_pass1_high_column_overlap_kept():
    # 5 biz fields; fbdi has 4 → overlap=0.8 ≥ 0.3 → kept even with no name match
    fields = [SnapshotField(f"TA4F{i}", f"F{i}", False, "X", 10) for i in range(5)]
    table = _make_snapshot_table("T_SOME_TABLE", "TA4", fields, [])
    fbdi_cols = {f"F{i}" for i in range(4)}
    catalog = {("SomeTemplate", "UNRELATED_TAB"): fbdi_cols}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    candidates = idx.get("T_SOME_TABLE", [])
    assert len(candidates) == 1
    assert candidates[0].column_overlap >= 0.3


def test_pass1_legacy_fields_excluded_from_overlap():
    biz = SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)
    leg = SnapshotField("@TA4SITE", "SITE", True, "X", 10)
    table = _make_snapshot_table("T_RA", "TA4", [biz, leg], [])
    # FBDI has INVOICE_ID but not SITE (SITE is legacy — shouldn't hurt overlap)
    catalog = {("AnyTemplate", "RA"): {"INVOICE_ID"}}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    candidates = idx.get("T_RA", [])
    assert candidates  # kept because name PARTIAL match or high overlap
    assert candidates[0].column_overlap == 1.0  # 1/1 biz field matched


def test_pass1_sorted_strongest_first():
    # Two candidates: one with EXACT name, one with PARTIAL
    fields = [SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)]
    table = _make_snapshot_table("T_RA_INTERFACE_LINES_ALL", "TA4", fields, ["TA4INVOICE_ID"])
    catalog = {
        ("TemplateA", "RA_INTERFACE_LINES_ALL"): {"INVOICE_ID"},   # EXACT + key 1.0
        ("TemplateB", "RA_INTERFACE_LINES"): {"INVOICE_ID"},       # PARTIAL + key 1.0
    }
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    candidates = idx["T_RA_INTERFACE_LINES_ALL"]
    assert len(candidates) == 2
    assert candidates[0].name_alignment == "EXACT"
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py -k "pass1" -v
```
Expected: `ImportError`.

- [ ] **Step 3: Add Pass 1 to `fbdi/audit.py`**

```python
# ---------------------------------------------------------------------------
# Pass 1 — candidate index
# ---------------------------------------------------------------------------

_PASS1_MIN_NAME_ALIGNMENT = {"EXACT", "PARTIAL"}
_PASS1_MIN_KEY_COVERAGE = 0.5
_PASS1_MIN_COLUMN_OVERLAP = 0.3


def _sort_key(c: Candidate) -> tuple:
    align_order = {"EXACT": 0, "PARTIAL": 1, "NONE": 2}
    return (align_order[c.name_alignment], -c.key_coverage, -c.column_overlap)


def build_candidate_index(
    snapshot: ApplaudSnapshot, catalog: CatalogIndex
) -> CandidateIndex:
    index: CandidateIndex = {}
    table_by_name = snapshot.table_by_name()

    for applaud_table_name, snap_table in table_by_name.items():
        candidates: list[Candidate] = []
        key_bare = snap_table.key_bare_names()

        for (fbdi_file, fbdi_tab), fbdi_cols in catalog.items():
            name_align = compute_name_alignment(applaud_table_name, fbdi_tab)
            key_cov = compute_key_coverage(key_bare, fbdi_cols)
            col_ovlp = compute_column_overlap(snap_table.fields, fbdi_cols)
            prefix_ok = check_prefix_conformance(
                applaud_table_name, snap_table.prefix or "", fbdi_tab
            )

            # Pass-1 threshold: keep if any signal clears its floor
            if (
                name_align in _PASS1_MIN_NAME_ALIGNMENT
                or key_cov >= _PASS1_MIN_KEY_COVERAGE
                or col_ovlp >= _PASS1_MIN_COLUMN_OVERLAP
            ):
                fbdi_upper = {c.upper() for c in fbdi_cols}
                biz_fields = snap_table.business_fields()
                matched = [f.bare_name for f in biz_fields if f.bare_name.upper() in fbdi_upper]
                missing = [f.bare_name for f in biz_fields if f.bare_name.upper() not in fbdi_upper]
                key_matched = [k for k in key_bare if k.upper() in fbdi_upper]

                candidates.append(Candidate(
                    fbdi_file=fbdi_file,
                    fbdi_tab=fbdi_tab,
                    name_alignment=name_align,
                    key_coverage=key_cov,
                    column_overlap=col_ovlp,
                    prefix_conformance=prefix_ok,
                    applaud_key_fields_matched=key_matched,
                    applaud_fields_matched=matched,
                    applaud_fields_missing=missing,
                ))

        candidates.sort(key=_sort_key)
        index[applaud_table_name] = candidates

    return index
```

- [ ] **Step 4: Run — verify passing**

```
python -m pytest tests/test_audit.py -k "pass1" -v
```
Expected: 5 tests pass.

- [ ] **Step 5: Run full suite**

```
python -m pytest tests/ --tb=short
```
Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): pass 1 candidate index with signal scoring and threshold filter"
```

---

## Task 6: Prior-mapping parser + confidence evaluator

**Files:**
- Modify: `fbdi/audit.py` — add `parse_prior_mapping`, `evaluate_confidence`
- Modify: `tests/test_audit.py` — add 9 tests

- [ ] **Step 1: Write failing tests**

```python
from fbdi.audit import parse_prior_mapping, evaluate_confidence, Candidate


def _cand(name_align: str, key_cov: float, col_ovlp: float) -> Candidate:
    return Candidate(
        fbdi_file="F", fbdi_tab="T",
        name_alignment=name_align,
        key_coverage=key_cov,
        column_overlap=col_ovlp,
        prefix_conformance=True,
        applaud_key_fields_matched=[],
        applaud_fields_matched=[],
        applaud_fields_missing=[],
    )


# --- parse_prior_mapping ---

def test_parse_prior_mapping_single():
    result = parse_prior_mapping("AutoInvoiceImportTemplate / RA_INTERFACE_LINES_ALL")
    assert result == [("AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL")]

def test_parse_prior_mapping_multi():
    result = parse_prior_mapping(
        "AutoInvoiceImportTemplate / RA_INTERFACE_LINES_ALL; "
        "ItemStructureImportTemplate / EGP_COMPONENTS_INTERFACE"
    )
    assert len(result) == 2
    assert ("AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL") in result
    assert ("ItemStructureImportTemplate", "EGP_COMPONENTS_INTERFACE") in result

def test_parse_prior_mapping_blank():
    assert parse_prior_mapping("") == []
    assert parse_prior_mapping("   ") == []

def test_parse_prior_mapping_malformed(caplog):
    import logging
    with caplog.at_level(logging.WARNING):
        result = parse_prior_mapping("NoSlashHere")
    # Returns empty, doesn't crash
    assert result == []

# --- evaluate_confidence ---

def test_evaluate_confidence_high():
    # EXACT + key=1.0 + overlap>=0.7
    c = _cand("EXACT", 1.0, 0.85)
    assert evaluate_confidence(c) == "H"

def test_evaluate_confidence_high_no_keys():
    # EXACT + key=1.0 (vacuously, no keys) + overlap>=0.7
    c = _cand("EXACT", 0.0, 0.75)
    # key_coverage=0 means no keys; "key_coverage==1.0" is False but spec says
    # High = EXACT AND (key_coverage==1.0 OR column_overlap>=0.7)
    # With zero keys, key_coverage=0.0 so only overlap carries it
    assert evaluate_confidence(c) == "H"

def test_evaluate_confidence_medium_partial():
    c = _cand("PARTIAL", 0.8, 0.5)
    assert evaluate_confidence(c) == "M"

def test_evaluate_confidence_medium_key_coverage():
    c = _cand("NONE", 0.6, 0.45)
    assert evaluate_confidence(c) == "M"

def test_evaluate_confidence_low():
    # Only clears pass-1 floor but doesn't meet High or Medium
    c = _cand("PARTIAL", 0.0, 0.1)
    assert evaluate_confidence(c) == "L"
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py -k "parse_prior or confidence" -v
```
Expected: `ImportError`.

- [ ] **Step 3: Add functions to `fbdi/audit.py`**

```python
# ---------------------------------------------------------------------------
# Prior-mapping text parser
# ---------------------------------------------------------------------------

import logging
_log = logging.getLogger(__name__)


def parse_prior_mapping(mapping_text: str) -> list[tuple[str, str]]:
    """Parse "Template / Tab[; Template / Tab]" → [(file, tab), ...]."""
    result: list[tuple[str, str]] = []
    if not mapping_text or not mapping_text.strip():
        return result
    for segment in mapping_text.split(";"):
        segment = segment.strip()
        if not segment:
            continue
        parts = segment.split(" / ", maxsplit=1)
        if len(parts) != 2 or not parts[0].strip() or not parts[1].strip():
            _log.warning("Malformed prior mapping segment (skipping): %r", segment)
            continue
        result.append((parts[0].strip(), parts[1].strip()))
    return result


# ---------------------------------------------------------------------------
# Confidence tier evaluator
# ---------------------------------------------------------------------------

def evaluate_confidence(candidate: Candidate) -> str:
    """Return H, M, or L per spec §6.2. Evaluated in order; first match wins."""
    if (
        candidate.name_alignment == "EXACT"
        and (candidate.key_coverage == 1.0 or candidate.column_overlap >= 0.7)
    ):
        return "H"
    if candidate.name_alignment == "PARTIAL" or (
        0 < candidate.key_coverage < 1.0 and candidate.column_overlap >= 0.4
    ):
        return "M"
    return "L"
```

- [ ] **Step 4: Run — verify passing**

```
python -m pytest tests/test_audit.py -k "parse_prior or confidence" -v
```
Expected: 9 tests pass.

- [ ] **Step 5: Run full suite**

```
python -m pytest tests/ --tb=short
```
Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): prior-mapping parser and confidence tier evaluator"
```

---

## Task 7: Pass 2 — adjudication engine

**Files:**
- Modify: `fbdi/audit.py` — add `adjudicate_table`
- Modify: `tests/test_audit.py` — add 11 tests covering every branch in §6.1

- [ ] **Step 1: Write failing tests**

```python
from fbdi.audit import (
    adjudicate_table, AuditRow, Candidate, EvidenceBundle,
    SnapshotTable, SnapshotField, SnapshotKeySeq, PriorRow,
)


def _pr(status: str, mapping: str = "", prefix: str = "TA4",
        module: str = "Fin", notes: str = "") -> PriorRow:
    return PriorRow("T_TEST", status, prefix, mapping, module, notes)


def _cand(file: str, tab: str, align: str, key: float, overlap: float) -> Candidate:
    return Candidate(
        fbdi_file=file, fbdi_tab=tab,
        name_alignment=align, key_coverage=key, column_overlap=overlap,
        prefix_conformance=True,
        applaud_key_fields_matched=[], applaud_fields_matched=[], applaud_fields_missing=[],
    )


# Branch 1: NOT_IN_APPLAUD → UNMAPPED High
def test_adjudicate_not_in_applaud():
    row = adjudicate_table("T_GHOST", None, [], _pr("UNMAPPED"))
    assert row.verdict == "UNMAPPED"
    assert row.confidence == "H"
    assert "not present" in row.rationale.lower()


# Branch 2: FILE_TOO_LARGE carry-through
def test_adjudicate_file_too_large_carrythrough():
    row = adjudicate_table("T_TEST", None, [], _pr("FILE_TOO_LARGE"))
    assert row.verdict == "FILE_TOO_LARGE"
    assert row.confidence == ""


# Branch 3: Single prior, High signals → YES High
def test_adjudicate_single_prior_high():
    c = _cand("AutoInvoice", "RA_INTERFACE_LINES_ALL", "EXACT", 1.0, 0.85)
    prior = _pr("YES", "AutoInvoice / RA_INTERFACE_LINES_ALL")
    snap_table = SnapshotTable("T_RA", "TA4", "T_RA (TA4)", "1", [], [])
    row = adjudicate_table("T_RA", snap_table, [c], prior)
    assert row.verdict == "YES"
    assert row.confidence == "H"
    assert not row.changed


# Branch 4: Single prior, low signals → NEEDS_REVIEW
def test_adjudicate_single_prior_low():
    c = _cand("AutoInvoice", "WRONG_TAB", "NONE", 0.0, 0.1)
    prior = _pr("YES", "AutoInvoice / WRONG_TAB")
    snap_table = SnapshotTable("T_RA", "TA4", "T_RA (TA4)", "1", [], [])
    row = adjudicate_table("T_RA", snap_table, [c], prior)
    assert row.verdict == "NEEDS_REVIEW"
    assert row.needs_deep_rationale


# Branch 5: Multi prior, both High → multi retained
def test_adjudicate_multi_both_high():
    c1 = _cand("TemplA", "TAB_X", "EXACT", 1.0, 0.85)
    c2 = _cand("TemplB", "TAB_X", "EXACT", 1.0, 0.90)
    prior = _pr("YES", "TemplA / TAB_X; TemplB / TAB_X")
    snap_table = SnapshotTable("T_TAB_X", "TXX", "T_TAB_X (TXX)", "1", [], [])
    row = adjudicate_table("T_TAB_X", snap_table, [c1, c2], prior)
    assert row.verdict == "YES"
    assert ";" in row.fbdi_mapping  # multi retained


# Branch 6: Multi prior, one High + one Low → collapsed
def test_adjudicate_multi_collapse():
    c1 = _cand("TemplA", "TAB_X", "EXACT", 1.0, 0.85)
    # c2 not in candidates list (low signal, filtered by pass 1)
    prior = _pr("YES", "TemplA / TAB_X; TemplB / TAB_MISSING")
    snap_table = SnapshotTable("T_TAB_X", "TXX", "T_TAB_X (TXX)", "1", [], [])
    row = adjudicate_table("T_TAB_X", snap_table, [c1], prior)
    assert row.verdict == "YES"
    assert "TemplB" not in row.fbdi_mapping
    assert row.changed  # collapsed from multi


# Branch 7: UNMAPPED + High candidate → promoted to YES
def test_adjudicate_unmapped_promoted():
    c = _cand("AutoInvoice", "RA_INTERFACE_LINES_ALL", "EXACT", 1.0, 0.85)
    prior = _pr("UNMAPPED")
    snap_table = SnapshotTable("T_RA", "TA4", "T_RA (TA4)", "1", [], [])
    row = adjudicate_table("T_RA", snap_table, [c], prior)
    assert row.verdict == "YES"
    assert row.confidence == "H"
    assert row.changed


# Branch 8: UNMAPPED + Medium candidate → NEEDS_REVIEW
def test_adjudicate_unmapped_medium_candidate():
    c = _cand("AutoInvoice", "RA_INTERFACE_LINES", "PARTIAL", 0.5, 0.5)
    prior = _pr("UNMAPPED")
    snap_table = SnapshotTable("T_RA_INTERFACE_LINES_ALL", "TA4", "T_RA (TA4)", "1", [], [])
    row = adjudicate_table("T_RA_INTERFACE_LINES_ALL", snap_table, [c], prior)
    assert row.verdict == "NEEDS_REVIEW"
    assert row.confidence == "M"


# Branch 9: UNMAPPED + no viable candidate → stays UNMAPPED High
def test_adjudicate_unmapped_no_candidate():
    prior = _pr("UNMAPPED")
    snap_table = SnapshotTable("T_GHOST", "TGH", "T_GHOST (TGH)", "1", [], [])
    row = adjudicate_table("T_GHOST", snap_table, [], prior)
    assert row.verdict == "UNMAPPED"
    assert row.confidence == "H"
    assert not row.changed


# Branch 10: prefix mismatch surfaces in notes, doesn't change verdict
def test_adjudicate_prefix_mismatch_noted():
    c = _cand("AutoInvoice", "RA_INTERFACE_LINES_ALL", "EXACT", 1.0, 0.85)
    c.prefix_conformance = False
    prior = _pr("YES", "AutoInvoice / RA_INTERFACE_LINES_ALL")
    snap_table = SnapshotTable("T_RA_INTERFACE_LINES_ALL", "TA4", "T_RA (TA4)", "1", [], [])
    row = adjudicate_table("T_RA_INTERFACE_LINES_ALL", snap_table, [c], prior)
    assert row.verdict == "YES"
    assert "prefix" in row.rationale.lower() or any("prefix" in n.lower() for n in row.evidence.notes)


# Branch 11: deep_rationale trigger — changed from prior
def test_adjudicate_deep_rationale_on_change():
    c = _cand("AutoInvoice", "RA_INTERFACE_LINES_ALL", "EXACT", 1.0, 0.85)
    prior = _pr("UNMAPPED")  # was UNMAPPED, now promoted to YES
    snap_table = SnapshotTable("T_RA", "TA4", "T_RA (TA4)", "1", [], [])
    row = adjudicate_table("T_RA", snap_table, [c], prior)
    assert row.changed
    assert row.needs_deep_rationale
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py -k "adjudicate" -v
```
Expected: `ImportError`.

- [ ] **Step 3: Add `adjudicate_table` to `fbdi/audit.py`**

```python
# ---------------------------------------------------------------------------
# Pass 2 — adjudication
# ---------------------------------------------------------------------------

_CARRYTHROUGH_VERDICTS = {"FILE_TOO_LARGE", "FILE_ERROR"}


def _find_candidate(
    candidates: list[Candidate], fbdi_file: str, fbdi_tab: str
) -> Candidate | None:
    for c in candidates:
        if c.fbdi_file == fbdi_file and c.fbdi_tab == fbdi_tab:
            return c
    return None


def adjudicate_table(
    applaud_table: str,
    snap_table: SnapshotTable | None,
    candidates: list[Candidate],
    prior: PriorRow,
) -> AuditRow:
    evidence = EvidenceBundle(candidates_evaluated=list(candidates))
    prefix = snap_table.prefix or prior.prefix if snap_table else prior.prefix

    # ── PREFLIGHT ────────────────────────────────────────────────────────────
    if snap_table is None:
        return AuditRow(
            applaud_table=applaud_table, prefix=prefix,
            verdict="UNMAPPED", fbdi_mapping="",
            confidence="H", rationale="Applaud table not present in MDB snapshot",
            prior_verdict=prior.prior_status, changed=False,
            needs_deep_rationale=False, evidence=evidence,
        )

    if prior.prior_status in _CARRYTHROUGH_VERDICTS:
        return AuditRow(
            applaud_table=applaud_table, prefix=prefix,
            verdict=prior.prior_status, fbdi_mapping=prior.mapping_text,
            confidence="", rationale="Sized out / unreadable in 26B — unchanged from prior",
            prior_verdict=prior.prior_status, changed=False,
            needs_deep_rationale=False, evidence=evidence,
        )

    # ── PRIOR MAPPING PARSE ──────────────────────────────────────────────────
    prior_claims = parse_prior_mapping(prior.mapping_text)
    best_candidate = candidates[0] if candidates else None

    verdict: str
    fbdi_mapping: str
    confidence: str
    rationale: str

    # ── UNMAPPED / blank ─────────────────────────────────────────────────────
    if prior.prior_status in ("UNMAPPED", "") or (
        prior.prior_status == "YES" and not prior.mapping_text.strip()
    ):
        if best_candidate:
            conf = evaluate_confidence(best_candidate)
            if conf == "H":
                verdict = "YES"
                fbdi_mapping = f"{best_candidate.fbdi_file} / {best_candidate.fbdi_tab}"
                confidence = "H"
                rationale = (
                    f"Promoted from UNMAPPED — EXACT name match, "
                    f"key={best_candidate.key_coverage:.0%}, overlap={best_candidate.column_overlap:.0%}"
                )
            elif conf == "M":
                verdict = "NEEDS_REVIEW"
                fbdi_mapping = f"{best_candidate.fbdi_file} / {best_candidate.fbdi_tab}"
                confidence = "M"
                rationale = "Potential new mapping — Medium confidence; verify with Brad"
            else:
                verdict = "UNMAPPED"
                fbdi_mapping = ""
                confidence = "H"
                rationale = "No FBDI tab in 26B catalog scores above threshold"
        else:
            verdict = "UNMAPPED"
            fbdi_mapping = ""
            confidence = "H"
            rationale = "No FBDI tab in 26B catalog scores above threshold"

    # ── SINGLE prior claim ───────────────────────────────────────────────────
    elif len(prior_claims) == 1:
        file, tab = prior_claims[0]
        matched_c = _find_candidate(candidates, file, tab)
        if matched_c:
            conf = evaluate_confidence(matched_c)
            if conf in ("H", "M"):
                verdict = "YES"
                fbdi_mapping = f"{file} / {tab}"
                confidence = conf
                rationale = (
                    f"name={matched_c.name_alignment}, "
                    f"key={matched_c.key_coverage:.0%}, "
                    f"overlap={matched_c.column_overlap:.0%}"
                )
            else:
                verdict = "NEEDS_REVIEW"
                fbdi_mapping = f"{file} / {tab}"
                confidence = "L"
                rationale = "Prior claim scores Low against 26B catalog — verify"
        else:
            verdict = "NEEDS_REVIEW"
            fbdi_mapping = f"{file} / {tab}"
            confidence = "L" if best_candidate else "H"
            rationale = (
                "Prior references file/tab not found in 26B catalog or below all thresholds"
            )

    # ── MULTI prior claims ───────────────────────────────────────────────────
    else:
        high_or_med: list[tuple[str, str, Candidate, str]] = []
        low_or_absent: list[tuple[str, str]] = []
        for file, tab in prior_claims:
            c = _find_candidate(candidates, file, tab)
            if c:
                conf = evaluate_confidence(c)
                if conf in ("H", "M"):
                    high_or_med.append((file, tab, c, conf))
                else:
                    low_or_absent.append((file, tab))
                    evidence.rejected_alternatives.append(c)
            else:
                low_or_absent.append((file, tab))

        if len(high_or_med) == len(prior_claims):
            # All claims score High or Medium → keep multi
            verdict = "YES"
            fbdi_mapping = "; ".join(f"{f} / {t}" for f, t, _, _ in high_or_med)
            confidence = "H" if all(conf == "H" for _, _, _, conf in high_or_med) else "M"
            rationale = f"Multi-mapping retained — {len(high_or_med)} legs verified"
        elif len(high_or_med) == 1:
            # One good leg — collapse to single
            file, tab, c, conf = high_or_med[0]
            verdict = "YES"
            fbdi_mapping = f"{file} / {tab}"
            confidence = conf
            rationale = (
                f"Collapsed from multi — 1/{len(prior_claims)} legs scored {conf}; "
                f"rest below threshold"
            )
        else:
            verdict = "NEEDS_REVIEW"
            fbdi_mapping = "; ".join(f"{f} / {t}" for f, t in prior_claims)
            confidence = "M" if high_or_med else "L"
            rationale = "Multi-mapping contested — see audit.md for per-leg evidence"

    # ── PREFIX AUDIT (all verdicts) ──────────────────────────────────────────
    if verdict == "YES" and fbdi_mapping:
        # Check prefix_conformance on the first/primary chosen candidate
        first_claim = parse_prior_mapping(fbdi_mapping)
        if first_claim:
            chosen_c = _find_candidate(candidates, first_claim[0][0], first_claim[0][1])
            if chosen_c and not chosen_c.prefix_conformance:
                evidence.notes.append(
                    f"Prefix mismatch — expected T_<tab> convention, "
                    f"got prefix={prefix} for tab={first_claim[0][1]}"
                )

    changed = verdict != prior.prior_status
    needs_deep = (
        verdict == "NEEDS_REVIEW"
        or changed
        or confidence == "L"
        or bool(evidence.notes)
    )

    return AuditRow(
        applaud_table=applaud_table, prefix=prefix,
        verdict=verdict, fbdi_mapping=fbdi_mapping,
        confidence=confidence, rationale=rationale,
        prior_verdict=prior.prior_status, changed=changed,
        needs_deep_rationale=needs_deep, evidence=evidence,
    )
```

- [ ] **Step 4: Run — verify passing**

```
python -m pytest tests/test_audit.py -k "adjudicate" -v
```
Expected: 11 tests pass.

- [ ] **Step 5: Run full suite**

```
python -m pytest tests/ --tb=short
```
Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): pass 2 adjudication engine with all §6.1 branches"
```

---

## Task 8: Output writers — xlsx (3 sheets) + markdown sidecar

**Files:**
- Modify: `fbdi/audit.py` — add `write_output_xlsx`, `write_audit_md`
- Modify: `tests/test_audit.py` — add 4 tests

- [ ] **Step 1: Write failing tests**

```python
from fbdi.audit import write_output_xlsx, write_audit_md, AuditRow, EvidenceBundle
from openpyxl import load_workbook


def _make_audit_row(
    table: str, verdict: str, confidence: str, mapping: str,
    prior_verdict: str = "YES", changed: bool = False,
    needs_deep: bool = False, prefix: str = "TA4",
) -> AuditRow:
    return AuditRow(
        applaud_table=table, prefix=prefix,
        verdict=verdict, fbdi_mapping=mapping,
        confidence=confidence,
        rationale=f"{verdict} because signals",
        prior_verdict=prior_verdict, changed=changed,
        needs_deep_rationale=needs_deep,
        evidence=EvidenceBundle(),
    )


def test_write_output_xlsx_sheets(tmp_path):
    rows = [
        _make_audit_row("T_RA", "YES", "H", "AutoInvoice / RA_TAB"),
        _make_audit_row("T_GHOST", "UNMAPPED", "H", "", prior_verdict="UNMAPPED"),
        _make_audit_row("T_PROBLEM", "NEEDS_REVIEW", "M", "SomeFile / SomeTab",
                        prior_verdict="YES", changed=True, needs_deep=True),
    ]
    catalog: CatalogIndex = {
        ("AutoInvoice", "RA_TAB"): {"INVOICE_ID"},
    }
    out = tmp_path / "Claude_test.xlsx"
    write_output_xlsx(rows, catalog, out)
    wb = load_workbook(out, read_only=True)
    assert "FBDI Mapping" in wb.sheetnames
    assert "Applaud Tables" in wb.sheetnames
    assert "Needs Review" in wb.sheetnames
    wb.close()


def test_write_output_xlsx_sheet2_rows(tmp_path):
    rows = [
        _make_audit_row("T_RA", "YES", "H", "AutoInvoice / RA_TAB"),
        _make_audit_row("T_GHOST", "UNMAPPED", "H", "", prior_verdict="UNMAPPED"),
    ]
    catalog: CatalogIndex = {}
    out = tmp_path / "Claude_test.xlsx"
    write_output_xlsx(rows, catalog, out)
    wb = load_workbook(out, read_only=True, data_only=True)
    ws2 = wb["Applaud Tables"]
    data = list(ws2.iter_rows(values_only=True))
    # header + 2 data rows
    assert len(data) == 3
    table_names = [r[1] for r in data[1:]]
    assert "T_RA" in table_names
    assert "T_GHOST" in table_names
    wb.close()


def test_write_output_xlsx_needs_review_sheet(tmp_path):
    rows = [
        _make_audit_row("T_NEEDS", "NEEDS_REVIEW", "M", "SomeFile / SomeTab",
                        needs_deep=True),
        _make_audit_row("T_OK", "YES", "H", "AutoInvoice / RA_TAB"),
    ]
    catalog: CatalogIndex = {}
    out = tmp_path / "Claude_test.xlsx"
    write_output_xlsx(rows, catalog, out)
    wb = load_workbook(out, read_only=True, data_only=True)
    ws3 = wb["Needs Review"]
    data = list(ws3.iter_rows(values_only=True))
    # header + 1 needs-review row (T_OK not in there)
    assert len(data) == 2
    wb.close()


def test_write_audit_md(tmp_path):
    rows = [
        _make_audit_row("T_NEEDS", "NEEDS_REVIEW", "M", "SomeFile / SomeTab",
                        needs_deep=True, changed=True),
        _make_audit_row("T_OK", "YES", "H", "AutoInvoice / RA_TAB"),
    ]
    out = tmp_path / "audit.md"
    write_audit_md(rows, {"extracted_at": "2026-04-21T12:00:00Z"}, out)
    content = out.read_text()
    assert "T_NEEDS" in content
    assert "T_OK" not in content  # High confidence unchanged → no entry
    assert "NEEDS_REVIEW" in content
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py -k "write_output or write_audit" -v
```
Expected: `ImportError`.

- [ ] **Step 3: Add output writers to `fbdi/audit.py`**

```python
# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------

_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_DATA_FONT = Font(name="Calibri", size=11)

_VERDICT_FILLS = {
    "YES":          PatternFill("solid", fgColor="E2EFDA"),
    "UNMAPPED":     PatternFill("solid", fgColor="FCE4D6"),
    "NEEDS_REVIEW": PatternFill("solid", fgColor="FFF2CC"),
    "FILE_TOO_LARGE": PatternFill("solid", fgColor="F4B942"),
    "FILE_ERROR":   PatternFill("solid", fgColor="F4B942"),
}

_S1_HEADERS = [
    "FBDI Template", "FBDI Tab", "Applaud Table", "Prefix",
    "Status", "Module", "Notes", "Match Type", "Confidence",
]
_S2_HEADERS = [
    "#", "Applaud Table", "Status", "Prefix", "FBDI Template Mappings",
    "Confidence", "Rationale", "Changed From Prior", "Prior Status",
]


def _style_header_row(ws, n_cols: int) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _write_sheet1(ws, audit_rows: list[AuditRow], catalog: CatalogIndex) -> None:
    """FBDI Mapping — one row per (file, tab) in the 26B catalog."""
    # Build reverse lookup: (file, tab) → AuditRow
    tab_to_row: dict[tuple[str, str], AuditRow] = {}
    for ar in audit_rows:
        for file, tab in parse_prior_mapping(ar.fbdi_mapping):
            tab_to_row[(file, tab)] = ar

    ws.append(_S1_HEADERS)
    _style_header_row(ws, len(_S1_HEADERS))

    for (fbdi_file, fbdi_tab) in sorted(catalog):
        ar = tab_to_row.get((fbdi_file, fbdi_tab))
        if ar:
            match_type = "EXACT" if ar.confidence == "H" else "PARTIAL" if ar.confidence == "M" else "PRIOR-CARRYOVER"
            row = [fbdi_file, fbdi_tab, ar.applaud_table, ar.prefix,
                   ar.verdict, "", "", match_type, ar.confidence]
        else:
            row = [fbdi_file, fbdi_tab, "", "", "UNMAPPED", "", "", "", ""]
        ws.append(row)
        fill = _VERDICT_FILLS.get(row[4])
        if fill:
            for col in range(1, len(_S1_HEADERS) + 1):
                ws.cell(row=ws.max_row, column=col).fill = fill

    ws.freeze_panes = "A2"


def _write_sheet2(ws, audit_rows: list[AuditRow]) -> None:
    """Applaud Tables — 183 rows keyed by Applaud table."""
    ws.append(_S2_HEADERS)
    _style_header_row(ws, len(_S2_HEADERS))

    for i, ar in enumerate(audit_rows, start=1):
        changed_mark = "✓" if ar.changed else ""
        row = [i, ar.applaud_table, ar.verdict, ar.prefix, ar.fbdi_mapping,
               ar.confidence, ar.rationale, changed_mark, ar.prior_verdict]
        ws.append(row)
        fill = _VERDICT_FILLS.get(ar.verdict)
        if fill:
            for col in range(1, len(_S2_HEADERS) + 1):
                ws.cell(row=ws.max_row, column=col).fill = fill

    ws.freeze_panes = "A2"


def _write_sheet3(ws, audit_rows: list[AuditRow]) -> None:
    """Needs Review — filtered subset, sorted by priority."""
    ws.append(_S2_HEADERS)
    _style_header_row(ws, len(_S2_HEADERS))

    deep_rows = [ar for ar in audit_rows if ar.needs_deep_rationale]
    # Sort: NEEDS_REVIEW first, then changed, then Low confidence
    def _sort(ar: AuditRow) -> tuple:
        return (ar.verdict != "NEEDS_REVIEW", not ar.changed, ar.confidence != "L")
    deep_rows.sort(key=_sort)

    for i, ar in enumerate(deep_rows, start=1):
        changed_mark = "✓" if ar.changed else ""
        rationale = ar.rationale + " → see audit.md"
        row = [i, ar.applaud_table, ar.verdict, ar.prefix, ar.fbdi_mapping,
               ar.confidence, rationale, changed_mark, ar.prior_verdict]
        ws.append(row)
        fill = _VERDICT_FILLS.get(ar.verdict)
        if fill:
            for col in range(1, len(_S2_HEADERS) + 1):
                ws.cell(row=ws.max_row, column=col).fill = fill

    ws.freeze_panes = "A2"


def write_output_xlsx(
    audit_rows: list[AuditRow],
    catalog: CatalogIndex,
    output_path: Path = OUTPUT_MAPPING_PATH,
) -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "FBDI Mapping"
    _write_sheet1(ws1, audit_rows, catalog)

    ws2 = wb.create_sheet("Applaud Tables")
    _write_sheet2(ws2, audit_rows)

    ws3 = wb.create_sheet("Needs Review")
    _write_sheet3(ws3, audit_rows)

    wb.save(output_path)
    print(f"Wrote: {output_path}")


def write_audit_md(
    audit_rows: list[AuditRow],
    snapshot_meta: dict,
    output_path: Path = OUTPUT_AUDIT_PATH,
) -> None:
    deep_rows = [ar for ar in audit_rows if ar.needs_deep_rationale]
    needs_review = [ar for ar in deep_rows if ar.verdict == "NEEDS_REVIEW"]
    changed = [ar for ar in deep_rows if ar.changed and ar.verdict != "NEEDS_REVIEW"]
    other = [ar for ar in deep_rows if not ar.needs_deep_rationale or (
        ar.verdict != "NEEDS_REVIEW" and not ar.changed
    )]

    total = len(audit_rows)
    yes_count = sum(1 for ar in audit_rows if ar.verdict == "YES")
    unmapped_count = sum(1 for ar in audit_rows if ar.verdict == "UNMAPPED")
    nr_count = len(needs_review)
    changed_count = sum(1 for ar in audit_rows if ar.changed)

    lines: list[str] = [
        "# FBDI ↔ Applaud Mapping Audit — 26B",
        "",
        f"**Generated:** {datetime.now(timezone.utc).isoformat()}",
        f"**Snapshot:** applaud_snapshot.json @ {snapshot_meta.get('extracted_at', 'unknown')}",
        "**Catalog:** FBDI_Master_Catalog.xlsx 26B tab",
        "**Prior mapping:** fbdi_applaud_mapping.xlsx",
        "",
        "## Summary",
        "",
        f"Of {total} Applaud tables audited: "
        f"{yes_count} YES, {unmapped_count} UNMAPPED, {nr_count} NEEDS_REVIEW. "
        f"{changed_count} rows changed from prior.",
        "",
    ]

    if needs_review:
        lines += [f"## Needs Review ({len(needs_review)} rows)", ""]
        for ar in needs_review:
            lines += _md_section(ar)

    if changed:
        lines += ["## Changed From Prior", ""]
        for ar in changed:
            lines += _md_section(ar)

    prefix_mismatches = [ar for ar in audit_rows if ar.evidence.notes]
    if prefix_mismatches:
        lines += ["## Prefix Mismatches", ""]
        lines += ["| Applaud Table | Prefix | Notes |", "|---|---|---|"]
        for ar in prefix_mismatches:
            for note in ar.evidence.notes:
                lines.append(f"| {ar.applaud_table} | {ar.prefix} | {note} |")
        lines.append("")

    output_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote: {output_path}")


def _md_section(ar: AuditRow) -> list[str]:
    lines = [
        f"### {ar.applaud_table} (prefix: {ar.prefix}) — {ar.verdict}",
        f"- **Prior:** {ar.prior_verdict} → `{ar.fbdi_mapping or '(none)'}`",
        f"- **Decision:** {ar.rationale}",
    ]
    if ar.evidence.candidates_evaluated:
        lines.append("- **Candidates evaluated:**")
        for c in ar.evidence.candidates_evaluated[:5]:
            conf = evaluate_confidence(c)
            lines.append(
                f"  - `{c.fbdi_file} / {c.fbdi_tab}` — "
                f"name={c.name_alignment}, "
                f"keys={c.key_coverage:.0%}, "
                f"cols={c.column_overlap:.0%} → {conf}"
            )
    for note in ar.evidence.notes:
        lines.append(f"- **Note:** {note}")
    lines.append("")
    return lines
```

- [ ] **Step 4: Run — verify passing**

```
python -m pytest tests/test_audit.py -k "write_output or write_audit" -v
```
Expected: 4 tests pass.

- [ ] **Step 5: Run full suite**

```
python -m pytest tests/ --tb=short
```
Expected: all pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): output writers — 3-sheet xlsx and markdown sidecar"
```

---

## Task 9: `run_audit` orchestration + integration test

**Files:**
- Modify: `fbdi/audit.py` — add `run_audit`, `__main__` block
- Modify: `tests/test_audit.py` — add 1 end-to-end integration test

- [ ] **Step 1: Write failing integration test**

```python
from fbdi.audit import run_audit, CatalogIndex
from openpyxl import Workbook
import json
from pathlib import Path


def _make_e2e_snapshot(tmp_path: Path) -> Path:
    """5 tables covering: EXACT, PARTIAL, UNMAPPED, multi-collapse, needs-review, legacy-tracking."""
    tables = [
        # 1. EXACT + High → YES H
        {
            "name": "T_RA_INTERFACE_LINES_ALL",
            "prefix": "TA4",
            "description": "T_RA_INTERFACE_LINES_ALL (TA4)",
            "type": "1",
            "key_sequences": [{"seq": "1", "keys": ["TA4INVOICE_ID"]}],
            "fields": [
                {"name": "TA4INVOICE_ID", "bare_name": "INVOICE_ID",
                 "is_legacy_tracking": False, "data_type": "N", "length": 15},
                {"name": "@TA4SITE", "bare_name": "SITE",
                 "is_legacy_tracking": True, "data_type": "X", "length": 10},
            ],
        },
        # 2. PARTIAL match, prior=UNMAPPED → NEEDS_REVIEW M
        {
            "name": "T_RCV_HEADERS_INTERFACE",
            "prefix": "TH7",
            "description": "T_RCV_HEADERS_INTERFACE (TH7)",
            "type": "1",
            "key_sequences": [],
            "fields": [
                {"name": "TH7RECEIPT_NUM", "bare_name": "RECEIPT_NUM",
                 "is_legacy_tracking": False, "data_type": "X", "length": 30},
            ],
        },
        # 3. No candidate → UNMAPPED H (re-confirmed)
        {
            "name": "T_GHOST_TABLE",
            "prefix": "TGG",
            "description": "T_GHOST_TABLE (TGG)",
            "type": "1",
            "key_sequences": [],
            "fields": [],
        },
        # 4. Multi prior, both High → multi retained
        {
            "name": "T_EGP_COMPONENTS_INTERFACE",
            "prefix": "T91",
            "description": "T_EGP_COMPONENTS_INTERFACE (T91)",
            "type": "1",
            "key_sequences": [],
            "fields": [
                {"name": "T91COMPONENT_ITEM_ID", "bare_name": "COMPONENT_ITEM_ID",
                 "is_legacy_tracking": False, "data_type": "N", "length": 18},
            ],
        },
        # 5. Multi prior, one leg missing → collapse
        {
            "name": "T_DOO_ORDER_HEADERS_ALL",
            "prefix": "TC4",
            "description": "T_DOO_ORDER_HEADERS_ALL (TC4)",
            "type": "1",
            "key_sequences": [],
            "fields": [
                {"name": "TC4ORDER_NUMBER", "bare_name": "ORDER_NUMBER",
                 "is_legacy_tracking": False, "data_type": "X", "length": 50},
            ],
        },
    ]
    data = {
        "mdb_path": "test", "extracted_at": "2026-04-21T12:00:00Z",
        "extractor_version": "1", "tables": tables, "missing_tables": [],
    }
    p = tmp_path / "applaud_snapshot.json"
    p.write_text(json.dumps(data))
    return p


def _make_e2e_catalog(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "26B"
    ws.append(["release", "file_name", "tab_name", "position", "column_label",
                "column_technical", "data_type", "length", "scale", "data_type_raw", "required"])
    rows = [
        ("26B", "AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL", 1, "Invoice ID", "INVOICE_ID", "N", 15, None, "NUMBER(15)", "TRUE"),
        ("26B", "ReceivingReceiptImportTemplate", "RCV_HEADERS", 1, "Receipt Num", "RECEIPT_NUM", "X", 30, None, "VARCHAR2(30)", "FALSE"),
        ("26B", "ChangeOrderImportTemplate", "EGP_COMPONENTS_INTERFACE", 1, "Comp Item", "COMPONENT_ITEM_ID", "N", 18, None, "NUMBER(18)", "FALSE"),
        ("26B", "ItemStructureImportTemplate", "EGP_COMPONENTS_INTERFACE", 1, "Comp Item", "COMPONENT_ITEM_ID", "N", 18, None, "NUMBER(18)", "FALSE"),
        ("26B", "SourceSalesOrderImportTemplate", "DOO_ORDER_HEADERS_ALL_INT", 1, "Order Num", "ORDER_NUMBER", "X", 50, None, "VARCHAR2(50)", "FALSE"),
    ]
    for r in rows:
        ws.append(r)
    p = tmp_path / "FBDI_Master_Catalog.xlsx"
    wb.save(p)
    return p


def _make_e2e_prior(tmp_path: Path) -> Path:
    wb = Workbook()
    wb.active.title = "FBDI Mapping"
    ws2 = wb.create_sheet("Applaud Tables")
    ws2.append(["#", "applaud_table", "status", "prefix",
                 "fbdi_template_mappings", "module", "notes"])
    rows = [
        (1, "T_RA_INTERFACE_LINES_ALL", "YES", "TA4",
         "AutoInvoiceImportTemplate / RA_INTERFACE_LINES_ALL", "Financials", ""),
        (2, "T_RCV_HEADERS_INTERFACE", "UNMAPPED", "TH7", "", "Procurement", ""),
        (3, "T_GHOST_TABLE", "UNMAPPED", "TGG", "", "Unknown", ""),
        (4, "T_EGP_COMPONENTS_INTERFACE", "YES", "T91",
         "ChangeOrderImportTemplate / EGP_COMPONENTS_INTERFACE; "
         "ItemStructureImportTemplate / EGP_COMPONENTS_INTERFACE",
         "SCM", ""),
        (5, "T_DOO_ORDER_HEADERS_ALL", "YES", "TC4",
         "SourceSalesOrderImportTemplate / DOO_ORDER_HEADERS_ALL_INT; "
         "ItemStructureImportTemplate / NONEXISTENT_TAB",
         "SCM", ""),
    ]
    for r in rows:
        ws2.append(r)
    p = tmp_path / "fbdi_applaud_mapping.xlsx"
    wb.save(p)
    return p


def test_audit_end_to_end(tmp_path):
    snap_path = _make_e2e_snapshot(tmp_path)
    cat_path = _make_e2e_catalog(tmp_path)
    prior_path = _make_e2e_prior(tmp_path)
    out_xlsx = tmp_path / "Claude_test.xlsx"
    out_md = tmp_path / "audit.md"

    audit_rows = run_audit(snap_path, cat_path, prior_path, out_xlsx, out_md)

    assert len(audit_rows) == 5

    by_table = {ar.applaud_table: ar for ar in audit_rows}

    # T_RA: EXACT + High → YES H, unchanged
    ra = by_table["T_RA_INTERFACE_LINES_ALL"]
    assert ra.verdict == "YES"
    assert ra.confidence == "H"
    assert not ra.changed

    # T_RCV: PARTIAL → NEEDS_REVIEW
    rcv = by_table["T_RCV_HEADERS_INTERFACE"]
    assert rcv.verdict in ("NEEDS_REVIEW", "YES")  # depends on overlap score

    # T_GHOST: no candidates → UNMAPPED H
    ghost = by_table["T_GHOST_TABLE"]
    assert ghost.verdict == "UNMAPPED"
    assert ghost.confidence == "H"

    # T_EGP: multi, both legs High → multi retained YES
    egp = by_table["T_EGP_COMPONENTS_INTERFACE"]
    assert egp.verdict == "YES"
    assert ";" in egp.fbdi_mapping

    # T_DOO: multi, one leg missing → collapsed YES
    doo = by_table["T_DOO_ORDER_HEADERS_ALL"]
    assert doo.verdict == "YES"
    assert "NONEXISTENT_TAB" not in doo.fbdi_mapping
    assert doo.changed

    # Outputs exist
    assert out_xlsx.exists()
    assert out_md.exists()

    # Legacy tracking exclusion: T_RA overlap denominator = 1 (INVOICE_ID only), not 2
    ra_candidates = ra.evidence.candidates_evaluated
    if ra_candidates:
        assert ra_candidates[0].column_overlap == 1.0
```

- [ ] **Step 2: Run — verify failure**

```
python -m pytest tests/test_audit.py::test_audit_end_to_end -v
```
Expected: `ImportError` — `run_audit` not defined.

- [ ] **Step 3: Add `run_audit` and `__main__` block to `fbdi/audit.py`**

```python
# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def run_audit(
    snapshot_path: Path = SNAPSHOT_PATH,
    catalog_path: Path = CATALOG_PATH,
    prior_mapping_path: Path = PRIOR_MAPPING_PATH,
    output_xlsx_path: Path = OUTPUT_MAPPING_PATH,
    output_md_path: Path = OUTPUT_AUDIT_PATH,
) -> list[AuditRow]:
    # Check snapshot freshness (warn only)
    snap = load_snapshot(snapshot_path)
    try:
        extracted = datetime.fromisoformat(snap.extracted_at.replace("Z", "+00:00"))
        age_days = (datetime.now(timezone.utc) - extracted).days
        if age_days > SNAPSHOT_MAX_AGE_DAYS:
            warnings.warn(
                f"Snapshot is {age_days} days old (>{SNAPSHOT_MAX_AGE_DAYS}). "
                "Re-run Step A if the MDB has changed.",
                stacklevel=2,
            )
    except Exception:
        pass

    catalog = load_catalog(catalog_path)
    prior_mapping = load_prior_mapping(prior_mapping_path)

    print(f"Snapshot: {len(snap.tables)} tables, {len(snap.missing_set())} missing")
    print(f"Catalog: {len(catalog)} (file, tab) pairs")
    print(f"Prior mapping: {len(prior_mapping)} Applaud tables")

    # Pass 1
    candidate_index = build_candidate_index(snap, catalog)

    # Pass 2 — iterate over prior mapping row order (preserves Brad's ordering)
    table_by_name = snap.table_by_name()
    missing_set = snap.missing_set()
    audit_rows: list[AuditRow] = []

    for table_name, prior_row in prior_mapping.items():
        snap_table = table_by_name.get(table_name)
        if table_name in missing_set:
            snap_table = None
        candidates = candidate_index.get(table_name, [])
        row = adjudicate_table(table_name, snap_table, candidates, prior_row)
        audit_rows.append(row)

    # Emit summary
    verdicts = {}
    for ar in audit_rows:
        verdicts[ar.verdict] = verdicts.get(ar.verdict, 0) + 1
    changed_count = sum(1 for ar in audit_rows if ar.changed)
    print(f"\nResults: {verdicts}")
    print(f"Changed from prior: {changed_count}")
    print(f"Needs deep rationale: {sum(1 for ar in audit_rows if ar.needs_deep_rationale)}")

    write_output_xlsx(audit_rows, catalog, output_xlsx_path)
    write_audit_md(audit_rows, {"extracted_at": snap.extracted_at}, output_md_path)

    return audit_rows


if __name__ == "__main__":
    run_audit()
```

- [ ] **Step 4: Run integration test**

```
python -m pytest tests/test_audit.py::test_audit_end_to_end -v
```
Expected: PASS.

- [ ] **Step 5: Run full suite**

```
python -m pytest tests/ --tb=short
```
Expected: all 139 + new tests pass (target: ~175 total).

- [ ] **Step 6: Count tests**

```
python -m pytest tests/ --collect-only -q | tail -5
```
Confirm test count is ~175.

- [ ] **Step 7: Commit**

```bash
git add fbdi/audit.py tests/test_audit.py
git commit -m "feat(audit): run_audit orchestration + end-to-end integration test"
```

---

## Task 10: Step A — Snapshot extraction (agent-driven)

> **This task is not Python code.** It is a Claude agent run using applaud-mcp to query the live MDB and write `applaud_snapshot.json`. Execute in this session using the MCP tools directly.

- [ ] **Step 1: Schema probe — verify DataDictionary columns**

Run:
```
query_table(table_name="DataDictionary", where_clause="Name LIKE 'TA4%'", limit=3)
```
Expected: returns rows with `Name`, `DataType`, `Size` columns and values for prefix `TA4`.
If columns differ from expected, abort and report actual columns before proceeding.

- [ ] **Step 2: Load the 183 Applaud table names from Sheet2**

Read `fbdi_applaud_mapping.xlsx` Sheet2 column B (Applaud Table). This is the working set — do not add tables that appear in the MDB but are not in Sheet2.

- [ ] **Step 3: Per-table loop — extract each of the 183 tables**

For each table name:

```
get_table_definition(name="T_RA_INTERFACE_LINES_ALL")
```
→ Parse description to extract prefix using: `re.search(r'\(([A-Z0-9]+)\)\s*$', description)`

```
query_table(table_name="DataDictionary", where_clause="Name LIKE 'TA4%'")
```
→ For each row: 
  - `is_legacy = name.startswith("@")`
  - `clean_name = name.lstrip("@")`
  - `bare_name = clean_name[len(prefix):]` if `clean_name.upper().startswith(prefix.upper())`
  - `data_type` from `DataType` column
  - `length` from `Size` column

Tables not found in DatabaseTable → add to `missing_tables` list.

- [ ] **Step 4: Write `applaud_snapshot.json`**

Write to repo root. Schema must match `load_snapshot` expectations exactly (see Task 2 Step 3).

- [ ] **Step 5: Verify snapshot loads cleanly**

```bash
python -c "from fbdi.audit import load_snapshot; s = load_snapshot(); print(f'{len(s.tables)} tables, {len(s.missing_set())} missing')"
```
Expected: `183 tables, N missing` (where N ≥ 0).

- [ ] **Step 6: Commit snapshot**

```bash
git add applaud_snapshot.json
git commit -m "feat(audit): applaud MDB snapshot for 26B audit (183 tables)"
```

---

## Task 11: Full audit run + verification

> Run the complete audit against real data and verify output counts are sensible.

- [ ] **Step 1: Run the audit**

```bash
python -m fbdi.audit
```
Expected output includes:
```
Snapshot: 183 tables, N missing
Catalog: NNN (file, tab) pairs
Prior mapping: 183 Applaud tables
Results: {'YES': N, 'UNMAPPED': N, 'NEEDS_REVIEW': N, ...}
Changed from prior: N
Needs deep rationale: N
Wrote: Claude_fbdi_applaud_mapping.xlsx
Wrote: Claude_fbdi_applaud_mapping_audit.md
```

- [ ] **Step 2: Spot-check Sheet2**

Open `Claude_fbdi_applaud_mapping.xlsx`, Sheet2. Verify:
- Row count = 183
- `T_RA_INTERFACE_LINES_ALL` row shows `YES`, `H`, `AutoInvoiceImportTemplate / RA_INTERFACE_LINES_ALL`
- At least one NEEDS_REVIEW row exists (expected 15-40)
- Changed From Prior column has checkmarks where expected

- [ ] **Step 3: Spot-check Needs Review sheet**

Open Sheet3. Verify:
- NEEDS_REVIEW rows appear first
- Rationale column ends with "→ see audit.md"
- Count is in expected range (15-40)

- [ ] **Step 4: Spot-check audit.md**

Open `Claude_fbdi_applaud_mapping_audit.md`. Verify:
- Summary counts match xlsx
- Each NEEDS_REVIEW row has a markdown section with candidates listed
- Prefix Mismatches section present if any mismatches detected

- [ ] **Step 5: Run full test suite one final time**

```bash
python -m pytest tests/ -v --tb=short
```
Expected: all tests pass.

- [ ] **Step 6: Final commit**

```bash
git add Claude_fbdi_applaud_mapping.xlsx Claude_fbdi_applaud_mapping_audit.md
git commit -m "feat(audit): full 26B applaud mapping audit output — N changed, N needs review"
```

---

## Self-Review Against Spec

| Spec Section | Covered By |
|---|---|
| §4 Snapshot extraction | Task 10 (agent-driven) + `load_snapshot` in Task 2 |
| §4.1 `@`-prefix fields + `is_legacy_tracking` | Task 1 data classes + Task 3 `derive_bare_name` + Task 4 `compute_column_overlap` |
| §4.2 Schema probe | Task 10 Step 1 |
| §4.3 183-table working set | Task 10 Step 2 |
| §5 Pass 1 candidate index | Task 5 `build_candidate_index` |
| §5.2 Four signals | Task 4 signal functions |
| §5.3 Threshold filter | Task 5 (≥0.5 key, ≥0.3 overlap, or name match) |
| §6 Pass 2 adjudication | Task 7 `adjudicate_table` |
| §6.1 All branches (PREFLIGHT, SINGLE, MULTI, UNMAPPED, PREFIX) | Task 7 (11 branch tests) |
| §6.2 Confidence tiers H/M/L | Task 6 `evaluate_confidence` |
| §6.2.1 Allowed confidence per verdict | Task 7 (LOW → NEEDS_REVIEW, not YES) |
| §6.3 AuditRow + EvidenceBundle | Task 1 data classes |
| §6.4 Deep-rationale trigger | Task 7 `needs_deep_rationale` logic |
| §6.5 Change tracking | Task 7 `changed` field |
| §7.1 Sheet 1 FBDI Mapping | Task 8 `_write_sheet1` |
| §7.1 Sheet 2 Applaud Tables | Task 8 `_write_sheet2` |
| §7.1 Sheet 3 Needs Review | Task 8 `_write_sheet3` |
| §7.2 audit.md sidecar | Task 8 `write_audit_md` |
| §8.1 Unit tests (all 10 categories) | Tasks 1-8 tests |
| §8.2 Integration test | Task 9 `test_audit_end_to_end` |
| §9.2 Snapshot freshness warning | Task 9 `run_audit` |
| §9.3 Hard errors (missing files) | Task 2 loaders |
| §11 DataDictionary column names | Task 10 schema probe + `query_table` |
| Module standalone (not wired to CLI) | Task 9 `__main__` block |
