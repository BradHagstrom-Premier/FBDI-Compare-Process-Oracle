# Smart Clearing & 26A-26B Comparison Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the FBDI clearing process to preserve all headers using dynamic detection, restructure baselines to a clean directory convention, and run the 26A-to-26B comparison.

**Architecture:** The clearing logic moves into `fbdi/clear.py` (new module) so it can import `detect_header_row` cleanly and be unit-tested without pulling in Selenium dependencies from `tools/download_and_clear.py`. The tools script delegates to `fbdi.clear` for the actual clearing, keeping subprocess orchestration and download logic separate. The CLI gets a `_resolve_dir()` helper so `--old 26A` resolves to `baselines/26A/originals/`.

**Tech Stack:** Python 3, openpyxl, pytest, existing `fbdi` package

**Spec:** `docs/superpowers/specs/2026-04-14-smart-clear-and-26a-26b-comparison-design.md`

---

### Task 1: Restructure baseline directories

Rename `Originals/` to `originals/` and remove empty `Blank Copies/` folders for 26A and 26B. Windows NTFS is case-insensitive, so renaming requires a two-step move through a temp name.

**Files:**
- Modify: `baselines/26A/` directory structure
- Modify: `baselines/26B/` directory structure

- [ ] **Step 1: Rename 26A directories**

```bash
cd "C:/Users/10193/Definian/FBDI-Compare-Process-Oracle"
mv "baselines/26A/Originals" "baselines/26A/originals_tmp"
mv "baselines/26A/originals_tmp" "baselines/26A/originals"
rmdir "baselines/26A/Blank Copies"
mkdir "baselines/26A/blanks"
```

- [ ] **Step 2: Rename 26B directories**

```bash
mv "baselines/26B/Originals" "baselines/26B/originals_tmp"
mv "baselines/26B/originals_tmp" "baselines/26B/originals"
rmdir "baselines/26B/Blank Copies"
mkdir "baselines/26B/blanks"
```

- [ ] **Step 3: Verify structure**

```bash
ls baselines/26A/
# Expected: blanks  originals
ls baselines/26B/
# Expected: blanks  originals
ls baselines/26A/originals/ | wc -l
# Expected: 211
ls baselines/26B/originals/ | wc -l
# Expected: 211
```

- [ ] **Step 4: Commit**

```bash
git add -A baselines/
git commit -m "chore: restructure baselines — lowercase dirs, remove Blank Copies"
```

---

### Task 2: Write failing tests for smart clearing

Create `tests/test_clear.py` with tests for the new `fbdi/clear.py` module (which doesn't exist yet — tests should fail).

**Files:**
- Create: `tests/test_clear.py`
- Will test: `fbdi/clear.py` (created in Task 3)

- [ ] **Step 1: Create test file**

```python
"""Tests for fbdi.clear — smart FBDI template clearing."""

import pytest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from fbdi.clear import clear_workbook


UPPER_SNAKE_HEADERS = [
    "TRANSACTION_TYPE", "BATCH_ID", "ORGANIZATION_CODE",
    "CHANGE_NOTICE", "CHANGE_NAME", "DESCRIPTION",
]

HUMAN_HEADERS = [
    "*Transaction Type", "*Batch ID", "Organization Code",
    "Change Number", "Change Name", "Description",
]

SAMPLE_DATA_ROW = ["CREATE", "10017", "V1", "DR_ECO_1", "Test Change", "A description"]


def _make_fbdi_workbook(
    path: Path,
    header_row: int,
    headers: list[str],
    data_rows: list[list] | None = None,
    extra_rows: dict[int, list] | None = None,
    include_instructions_sheet: bool = True,
):
    """Create a test FBDI workbook with headers and sample data."""
    wb = Workbook()
    if include_instructions_sheet:
        ws_instr = wb.active
        ws_instr.title = "Instructions and CSV Generation"
        ws_instr.cell(row=1, column=1, value="Instructions")
        ws_instr.cell(row=5, column=1, value="This row should NOT be cleared")
        ws_instr.cell(row=6, column=1, value="Neither should this one")
    else:
        wb.remove(wb.active)

    ws = wb.create_sheet(title="DATA_INTERFACE")
    if extra_rows:
        for row_num, values in extra_rows.items():
            for col_idx, val in enumerate(values, start=1):
                ws.cell(row=row_num, column=col_idx, value=val)
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col_idx, value=header)
    if data_rows:
        for offset, row_data in enumerate(data_rows, start=1):
            for col_idx, val in enumerate(row_data, start=1):
                ws.cell(row=header_row + offset, column=col_idx, value=val)

    wb.save(path)
    wb.close()


class TestClearWorkbook:
    def test_headers_at_row_4_preserved(self, tmp_path):
        """Standard pattern: headers at row 4, data at row 5+."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=4,
            headers=HUMAN_HEADERS,
            data_rows=[SAMPLE_DATA_ROW, SAMPLE_DATA_ROW],
            extra_rows={2: ["Invoice Headers"], 3: ["* Required"]},
        )

        skipped = clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["DATA_INTERFACE"]
        # Headers at row 4 preserved
        assert ws.cell(row=4, column=1).value == "*Transaction Type"
        assert ws.cell(row=4, column=2).value == "*Batch ID"
        # Metadata above headers preserved
        assert ws.cell(row=2, column=1).value == "Invoice Headers"
        assert ws.cell(row=3, column=1).value == "* Required"
        # Data rows cleared
        assert ws.cell(row=5, column=1).value is None
        assert ws.cell(row=6, column=1).value is None
        assert skipped == []
        wb.close()

    def test_headers_at_row_5_preserved(self, tmp_path):
        """Pattern 1: technical UPPER_SNAKE headers at row 5."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=5,
            headers=UPPER_SNAKE_HEADERS,
            data_rows=[SAMPLE_DATA_ROW, SAMPLE_DATA_ROW],
            extra_rows={
                1: ["Name", "Transaction Type", "Batch ID", "Change Type", "Change Number", "Description"],
                2: ["Description", "Indicates the...", "Indicates the...", "Value that...", "Value that...", "Description of..."],
                3: ["Data Type", "VARCHAR2(10)", "NUMBER(18)", "VARCHAR2(80)", "VARCHAR2(50)", "VARCHAR2(2000)"],
                4: ["Required", "Required", "Required", "Required", "Optional", "Optional"],
            },
        )

        skipped = clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["DATA_INTERFACE"]
        # Technical headers at row 5 preserved
        assert ws.cell(row=5, column=1).value == "TRANSACTION_TYPE"
        assert ws.cell(row=5, column=2).value == "BATCH_ID"
        # Metadata rows 1-4 preserved
        assert ws.cell(row=1, column=1).value == "Name"
        assert ws.cell(row=4, column=1).value == "Required"
        # Data rows cleared
        assert ws.cell(row=6, column=1).value is None
        assert ws.cell(row=7, column=1).value is None
        assert skipped == []
        wb.close()

    def test_headers_at_row_8_preserved(self, tmp_path):
        """Pattern 2: deeper headers at row 8 (e.g., ChangeOrderImportTemplate)."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=8,
            headers=UPPER_SNAKE_HEADERS,
            data_rows=[SAMPLE_DATA_ROW],
            extra_rows={
                4: HUMAN_HEADERS,
                5: ["Description"] * 6,
                6: ["VARCHAR2(10)"] * 6,
                7: ["Reserved for Future Use"],
            },
        )

        skipped = clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["DATA_INTERFACE"]
        # Technical headers at row 8 preserved
        assert ws.cell(row=8, column=1).value == "TRANSACTION_TYPE"
        # All metadata rows 1-7 preserved
        assert ws.cell(row=4, column=1).value == "*Transaction Type"
        assert ws.cell(row=7, column=1).value == "Reserved for Future Use"
        # Data row 9 cleared
        assert ws.cell(row=9, column=1).value is None
        assert skipped == []
        wb.close()

    def test_instructions_sheet_untouched(self, tmp_path):
        """Sheet 1 (Instructions) should never be modified."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=4,
            headers=UPPER_SNAKE_HEADERS,
            data_rows=[SAMPLE_DATA_ROW],
        )

        clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws_instr = wb["Instructions and CSV Generation"]
        assert ws_instr.cell(row=1, column=1).value == "Instructions"
        assert ws_instr.cell(row=5, column=1).value == "This row should NOT be cleared"
        assert ws_instr.cell(row=6, column=1).value == "Neither should this one"
        wb.close()

    def test_detection_fails_sheet_skipped_not_destroyed(self, tmp_path):
        """If detect_header_row returns None, leave the sheet untouched."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"

        wb = Workbook()
        ws_instr = wb.active
        ws_instr.title = "Instructions and CSV Generation"
        ws_data = wb.create_sheet(title="WeirdSheet")
        # Only 1 cell — below MIN_CELLS=2, detection will return None
        ws_data.cell(row=3, column=1, value="SOLO_HEADER")
        ws_data.cell(row=4, column=1, value="solo_data")
        wb.save(src)
        wb.close()

        skipped = clear_workbook(str(src), str(dst))

        assert "WeirdSheet" in skipped

        wb = load_workbook(dst, read_only=True)
        ws = wb["WeirdSheet"]
        # Data should be untouched since detection failed
        assert ws.cell(row=3, column=1).value == "SOLO_HEADER"
        assert ws.cell(row=4, column=1).value == "solo_data"
        wb.close()

    def test_multiple_data_sheets_cleared_independently(self, tmp_path):
        """Each sheet gets its own header detection and clearing."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"

        wb = Workbook()
        ws_instr = wb.active
        ws_instr.title = "Instructions and CSV Generation"

        # Sheet 2: headers at row 4
        ws1 = wb.create_sheet(title="SHEET_A")
        for col, val in enumerate(UPPER_SNAKE_HEADERS, start=1):
            ws1.cell(row=4, column=col, value=val)
        for col, val in enumerate(SAMPLE_DATA_ROW, start=1):
            ws1.cell(row=5, column=col, value=val)

        # Sheet 3: headers at row 8
        ws2 = wb.create_sheet(title="SHEET_B")
        for col, val in enumerate(HUMAN_HEADERS, start=1):
            ws2.cell(row=4, column=col, value=val)
        for col, val in enumerate(["Description"] * 6, start=1):
            ws2.cell(row=5, column=col, value=val)
        for col, val in enumerate(["VARCHAR2"] * 6, start=1):
            ws2.cell(row=6, column=col, value=val)
        ws2.cell(row=7, column=1, value="Reserved")
        for col, val in enumerate(UPPER_SNAKE_HEADERS, start=1):
            ws2.cell(row=8, column=col, value=val)
        for col, val in enumerate(SAMPLE_DATA_ROW, start=1):
            ws2.cell(row=9, column=col, value=val)

        wb.save(src)
        wb.close()

        clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        # Sheet A: headers at row 4, data at row 5 cleared
        ws_a = wb["SHEET_A"]
        assert ws_a.cell(row=4, column=1).value == "TRANSACTION_TYPE"
        assert ws_a.cell(row=5, column=1).value is None

        # Sheet B: headers at row 8, data at row 9 cleared
        ws_b = wb["SHEET_B"]
        assert ws_b.cell(row=4, column=1).value == "*Transaction Type"
        assert ws_b.cell(row=8, column=1).value == "TRANSACTION_TYPE"
        assert ws_b.cell(row=9, column=1).value is None
        wb.close()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_clear.py -v`
Expected: ImportError — `cannot import name 'clear_workbook' from 'fbdi.clear'`

- [ ] **Step 3: Commit failing tests**

```bash
git add tests/test_clear.py
git commit -m "test: add failing tests for smart FBDI clearing"
```

---

### Task 3: Implement fbdi/clear.py

Create the `fbdi/clear.py` module with `clear_workbook()` that uses `detect_header_row()` to find headers dynamically.

**Files:**
- Create: `fbdi/clear.py`
- Test: `tests/test_clear.py` (from Task 2)

- [ ] **Step 1: Create fbdi/clear.py**

```python
"""Smart clearing for Oracle FBDI template workbooks.

Clears sample data from FBDI templates while preserving all header rows
and metadata. Uses detect_header_row() to dynamically find the header
boundary per sheet, then clears everything below it.

Used by tools/download_and_clear.py to produce blank copies for client
Oracle projects. NOT used by the comparison engine (which reads originals).
"""

import logging

from openpyxl import load_workbook
from openpyxl.styles.fonts import Font

from fbdi.detect_header import detect_header_row

logger = logging.getLogger(__name__)


def clear_workbook(src_path: str, dst_path: str) -> list[str]:
    """Clear sample data from an FBDI workbook, preserving all headers.

    For each data sheet (skipping sheet 1 / instructions):
      1. Run detect_header_row() to find the actual header row
      2. Clear all cell values from header_row + 1 onward
      3. Strip legacy_drawing references (prevents openpyxl save errors)

    If header detection fails for a sheet, that sheet is left untouched.

    Args:
        src_path: Path to the original FBDI .xlsm file.
        dst_path: Path to write the cleared copy.

    Returns:
        List of sheet names where header detection failed (skipped).
    """
    # Oracle FBDI files use font family values >14 which openpyxl rejects.
    Font.family.max = 255

    wb = load_workbook(src_path, keep_vba=True)
    skipped_sheets: list[str] = []

    for sheet in wb.worksheets[1:]:
        header_row = detect_header_row(sheet)
        if header_row is not None:
            for row in sheet.iter_rows(min_row=header_row + 1):
                for cell in row:
                    try:
                        cell.value = None
                    except AttributeError:
                        pass  # MergedCell — read-only, skip
        else:
            logger.warning(
                "Header detection failed for sheet '%s' — skipping (data preserved)",
                sheet.title,
            )
            skipped_sheets.append(sheet.title)

        # Strip malformed VML XML references on ALL sheets to prevent save errors.
        sheet.legacy_drawing = None

    wb.save(dst_path)
    wb.close()
    return skipped_sheets
```

- [ ] **Step 2: Run tests to verify they pass**

Run: `python -m pytest tests/test_clear.py -v`
Expected: All 6 tests PASS

- [ ] **Step 3: Run full test suite to check for regressions**

Run: `python -m pytest tests/ -v`
Expected: All 33 existing tests + 6 new tests PASS

- [ ] **Step 4: Commit**

```bash
git add fbdi/clear.py tests/test_clear.py
git commit -m "feat: add smart clearing module using detect_header_row"
```

---

### Task 4: Update tools/download_and_clear.py

Wire up the new `fbdi.clear.clear_workbook` and update paths to the new `baselines/<ver>/originals/` and `baselines/<ver>/blanks/` convention. Remove the `BLANK_` filename prefix.

**Files:**
- Modify: `tools/download_and_clear.py:197-218` (`_clear_single_file`)
- Modify: `tools/download_and_clear.py:221-245` (`clear_files_python` — filename convention)
- Modify: `tools/download_and_clear.py:341-442` (`main` — path convention)

- [ ] **Step 1: Replace `_clear_single_file` to delegate to `fbdi.clear`**

Replace lines 197-218:

```python
def _clear_single_file(src, dst):
    """Worker: clear one FBDI file. Runs in a subprocess for timeout support."""
    from fbdi.clear import clear_workbook
    clear_workbook(src, dst)
```

- [ ] **Step 2: Update `clear_files_python` — remove BLANK_ prefix**

On line 245, change the destination filename:

```python
        dst = os.path.join(blanks_path, filename)
```

(Was: `dst = os.path.join(blanks_path, f"BLANK_{filename}")`)

Update the docstring (lines 222-228) to:

```python
    """
    Clear FBDI templates using smart header detection.

    For each file in originals_path:
      - Detect the header row per sheet using detect_header_row()
      - Clear all cell values below the header row
      - Save to blanks_path with the same filename (no BLANK_ prefix)

    Each file is processed in a subprocess with a timeout so that huge
    files don't block the entire batch.
    """
```

- [ ] **Step 3: Update `main()` paths**

Replace lines 381-383:

```python
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    originals_path = os.path.join(repo_root, "baselines", version, "originals")
    blanks_path = os.path.join(repo_root, "baselines", version, "blanks")
```

Replace lines 385-386 (print labels):

```python
    print(f"=== FBDI Download & Clear: {version} ===")
    print(f"  Originals: {originals_path}")
    print(f"  Blanks:    {blanks_path}")
```

- [ ] **Step 4: Update docstring at top of file**

Replace lines 1-11:

```python
"""
Download Oracle FBDI templates for a given release version and create
blank (cleared) copies with headers preserved.

Usage:
    python tools/download_and_clear.py 26a
    python tools/download_and_clear.py 26b
    python tools/download_and_clear.py 26a --skip-clear       # download only
    python tools/download_and_clear.py 26a --clear-only        # clear only (already downloaded)
    python tools/download_and_clear.py 26a --use-vba-macro     # use Clear_FBDIs VBA instead of Python
"""
```

- [ ] **Step 5: Commit**

```bash
git add tools/download_and_clear.py
git commit -m "feat: update download_and_clear to use smart clearing + new paths"
```

---

### Task 5: Write failing test for CLI `_resolve_dir`

**Files:**
- Create: `tests/test_cli.py`

- [ ] **Step 1: Create test file**

```python
"""Tests for fbdi.cli — CLI helpers."""

from pathlib import Path
from fbdi.cli import _resolve_dir


class TestResolveDir:
    def test_existing_directory_passes_through(self, tmp_path):
        """A path that is already a directory is returned unchanged."""
        assert _resolve_dir(tmp_path) == tmp_path

    def test_release_label_resolves_to_originals(self, tmp_path, monkeypatch):
        """A non-directory path like '26A' resolves to baselines/26A/originals/."""
        baselines = tmp_path / "baselines" / "26A" / "originals"
        baselines.mkdir(parents=True)
        monkeypatch.chdir(tmp_path)
        result = _resolve_dir(Path("26A"))
        assert result == Path("baselines") / "26A" / "originals"
        assert result.is_dir()

    def test_nonexistent_path_passes_through(self, tmp_path, monkeypatch):
        """A path that doesn't exist and has no baselines match passes through."""
        monkeypatch.chdir(tmp_path)
        result = _resolve_dir(Path("nonexistent"))
        # Should return the original path for downstream error handling
        assert result == Path("nonexistent")

    def test_explicit_originals_path_passes_through(self, tmp_path):
        """An explicit path to originals/ is returned unchanged."""
        originals = tmp_path / "baselines" / "26A" / "originals"
        originals.mkdir(parents=True)
        assert _resolve_dir(originals) == originals
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_cli.py -v`
Expected: ImportError — `cannot import name '_resolve_dir' from 'fbdi.cli'`

- [ ] **Step 3: Commit failing tests**

```bash
git add tests/test_cli.py
git commit -m "test: add failing tests for CLI _resolve_dir helper"
```

---

### Task 6: Implement `_resolve_dir` in cli.py

**Files:**
- Modify: `fbdi/cli.py:7-8` (add helper function)
- Modify: `fbdi/cli.py:81-95` (apply in `_run_compare`)
- Test: `tests/test_cli.py` (from Task 5)

- [ ] **Step 1: Add `_resolve_dir` helper**

Add after the imports (after line 8):

```python
def _resolve_dir(path: Path) -> Path:
    """Resolve a release label to its baselines originals directory.

    If path is already a directory, return it unchanged.
    Otherwise, try baselines/<path>/originals/ as a convenience shorthand.
    Falls through to the original path if no match (caller handles the error).
    """
    if path.is_dir():
        return path
    candidate = Path("baselines") / str(path) / "originals"
    if candidate.is_dir():
        return candidate
    return path
```

- [ ] **Step 2: Apply `_resolve_dir` in `_run_compare`**

In `_run_compare`, replace lines 87-88:

```python
    old_dir = _resolve_dir(args.old)
    new_dir = _resolve_dir(args.new)
```

(Was: `old_dir = args.old` / `new_dir = args.new`)

- [ ] **Step 3: Run CLI tests to verify they pass**

Run: `python -m pytest tests/test_cli.py -v`
Expected: All 4 tests PASS

- [ ] **Step 4: Run full test suite**

Run: `python -m pytest tests/ -v`
Expected: All tests PASS (33 original + 6 clear + 4 CLI = 43)

- [ ] **Step 5: Commit**

```bash
git add fbdi/cli.py tests/test_cli.py
git commit -m "feat: add --release shorthand to compare CLI"
```

---

### Task 7: Regenerate blanks for 26A and 26B

Run the updated clearing script on both releases. This is a long-running task (~75 min per release for originals, likely faster for the clearing step).

**Files:**
- Reads: `baselines/26A/originals/` (211 files)
- Reads: `baselines/26B/originals/` (211 files)
- Writes: `baselines/26A/blanks/` (output)
- Writes: `baselines/26B/blanks/` (output)

- [ ] **Step 1: Clear 26A**

```bash
cd "C:/Users/10193/Definian/FBDI-Compare-Process-Oracle"
python tools/download_and_clear.py 26A --clear-only
```

Expected output: `Results: N/211 cleared` with N close to 211. Note any timeouts or failures.

- [ ] **Step 2: Clear 26B**

```bash
python tools/download_and_clear.py 26B --clear-only
```

Expected output: similar to 26A.

- [ ] **Step 3: Verify file counts**

```bash
ls baselines/26A/blanks/ | wc -l
# Expected: ~211 (minus any timeouts)
ls baselines/26B/blanks/ | wc -l
# Expected: ~211
```

---

### Task 8: Validate blanks — spot-check header preservation

Verify that the smart clearing preserved headers that the old `min_row=5` approach destroyed.

**Files:**
- Reads: `baselines/26A/blanks/AttachmentsImportTemplate.xlsm`
- Reads: `baselines/26A/blanks/ChangeOrderImportTemplate.xlsm`

- [ ] **Step 1: Check AttachmentsImportTemplate (headers were at row 5)**

```bash
python -c "
from openpyxl import load_workbook
from openpyxl.styles.fonts import Font
Font.family.max = 255
wb = load_workbook('baselines/26A/blanks/AttachmentsImportTemplate.xlsm', read_only=True, data_only=True)
ws = wb['Attachment Details']
print('Row 5 (should have UPPER_SNAKE headers):')
for col in range(1, 9):
    v = ws.cell(row=5, column=col).value
    if v: print(f'  C{col}: {v}')
print('Row 6 (should be empty):')
for col in range(1, 9):
    v = ws.cell(row=6, column=col).value
    if v: print(f'  C{col}: {v}')
wb.close()
"
```

Expected: Row 5 shows `ATTACHMENT_TYPE`, `ATTACHMENT_NAME`, etc. Row 6 is empty.

- [ ] **Step 2: Check ChangeOrderImportTemplate (headers were at row 8)**

```bash
python -c "
from openpyxl import load_workbook
from openpyxl.styles.fonts import Font
Font.family.max = 255
wb = load_workbook('baselines/26A/blanks/ChangeOrderImportTemplate.xlsm', read_only=True, data_only=True)
ws = wb['EGO_CHANGES_INT']
print('Row 8 (should have technical headers):')
for col in range(1, 9):
    v = ws.cell(row=8, column=col).value
    if v: print(f'  C{col}: {v}')
print('Row 9 (should be empty):')
for col in range(1, 9):
    v = ws.cell(row=9, column=col).value
    if v: print(f'  C{col}: {v}')
wb.close()
"
```

Expected: Row 8 shows `TRANSACTION_TYPE`, `BATCH_ID`, etc. Row 9 is empty.

---

### Task 9: Run 26A to 26B comparison

Run the comparison using originals via the new CLI shorthand.

**Files:**
- Reads: `baselines/26A/originals/` and `baselines/26B/originals/`
- Writes: `Comparison_Report_26A_26B.xlsx`

- [ ] **Step 1: Run comparison**

```bash
cd "C:/Users/10193/Definian/FBDI-Compare-Process-Oracle"
python -m fbdi compare --old 26A --new 26B --output Comparison_Report_26A_26B.xlsx --verbose 2>&1 | tee comparison_log.txt
```

Expected: `Comparing 211 file pairs...` followed by progress, then `Changes found: N`.

This will take ~75 minutes. Let it run to completion.

- [ ] **Step 2: Check output**

```bash
python -c "
from openpyxl import load_workbook
wb = load_workbook('Comparison_Report_26A_26B.xlsx', read_only=True)
ws = wb.active
print(f'Total change rows: {ws.max_row - 1}')
# Spot check first few rows
for row in range(1, min(ws.max_row, 11)):
    vals = [ws.cell(row=row, column=c).value for c in range(1, 8)]
    print(vals)
wb.close()
"
```

---

### Task 10: Verify results with diagnostics

Run the diagnostic report on both releases to confirm header detection health.

**Files:**
- Reads: `baselines/26A/originals/` and `baselines/26B/originals/`
- Writes: `Diagnostic_Report_26A_26B.xlsx`

- [ ] **Step 1: Run diagnostics**

```bash
python -m fbdi diagnose --old baselines/26A/originals --new baselines/26B/originals --output Diagnostic_Report_26A_26B.xlsx
```

Expected output:
```
  DETECTED:       ~600+ (vast majority)
  NO_HEADER:      0
  SKIPPED_TAB:    ~50-80
  FILE_TOO_LARGE: ~6
  FILE_ERROR:     ~2
```

Key check: `NO_HEADER` should be 0 (all resolved in Phase 3).

- [ ] **Step 2: Commit results**

```bash
git add Comparison_Report_26A_26B.xlsx
git commit -m "feat: 26A to 26B comparison report"
```
