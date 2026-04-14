# Smart Clearing & 26A-26B Comparison Pipeline

**Date:** 2026-04-14
**Status:** Approved

---

## Problem

The `download_and_clear.py` script clears FBDI templates from a hard-coded `min_row=5`, which destroys technical UPPER_SNAKE_CASE headers (e.g., `TRANSACTION_TYPE`, `INTERFACE_BATCH_CODE`) that live at row 5 or deeper. Analysis showed **227 out of ~387 data tabs** lose their technical headers to this clearing.

The goal is to:
1. Fix the clearing process to preserve all headers
2. Restructure baseline directories to a clean convention
3. Run the 26A to 26B comparison using originals

## Design Decisions

### Comparison uses originals, not blanks

The comparison engine only reads the detected header row. Sample data below headers is never touched. Originals are always complete (211 = 211), while blanks can have missing files from timeouts, crashes, or load failures. Originals eliminate any risk of silent misses.

### Smart clearing uses `detect_header_row()`

Instead of hard-coded `min_row=5`, the clear process calls `detect_header_row()` from the `fbdi` package per sheet to find the actual header row, then clears from `header_row + 1` onward. If detection fails, the sheet is skipped (not destroyed) and a warning is logged.

### Blank copies are for client projects only

Blanks serve Oracle import workflows on client projects. They are not used for comparison. This separates the two concerns cleanly.

### Script stays standalone

`download_and_clear.py` moves to `tools/` but is not integrated into the `fbdi` package. It imports `detect_header_row` from `fbdi.detect_header`. This avoids adding Selenium/webdriver dependencies to the comparison engine.

## Directory Convention

```
baselines/
  26A/
    originals/          <- as-downloaded from Oracle (untouched)
    blanks/             <- smart-cleared copies (same filenames, no BLANK_ prefix)
  26B/
    originals/
    blanks/
```

- Lowercase folder names, no spaces
- Same filenames in originals/ and blanks/ (no `BLANK_` prefix)
- The folder is the differentiator, not the filename

## Changes

### 1. `tools/download_and_clear.py`

**Smart clearing logic** (replaces `_clear_single_file`):

```python
from fbdi.detect_header import detect_header_row

def _clear_single_file(src, dst):
    from openpyxl import load_workbook
    from openpyxl.styles.fonts import Font
    Font.family.max = 255

    wb = load_workbook(src, keep_vba=True)
    for sheet in wb.worksheets[1:]:
        header_row = detect_header_row(sheet)
        if header_row is None:
            # Can't find header -- skip sheet, don't destroy
            continue
        for row in sheet.iter_rows(min_row=header_row + 1):
            for cell in row:
                try:
                    cell.value = None
                except AttributeError:
                    pass  # MergedCell
        sheet.legacy_drawing = None
    wb.save(dst)
    wb.close()
```

**Output path updates:**

```python
originals_path = os.path.join(repo_root, "baselines", version, "originals")
blanks_path = os.path.join(repo_root, "baselines", version, "blanks")
```

**Filename convention:** save as `filename.xlsm` (not `BLANK_filename.xlsm`).

### 2. Directory restructure

- Rename `baselines/26A/Originals/` to `baselines/26A/originals/`
- Rename `baselines/26B/Originals/` to `baselines/26B/originals/`
- Delete existing `Blank Copies/` folders (already done by Brad)
- Regenerate `blanks/` using fixed clearing
- Rename `baselines/25d/Originals/` to `baselines/25d/originals/` if present

### 3. `fbdi/cli.py` - `--release` shorthand

Add `_resolve_dir()` helper:

```python
def _resolve_dir(path: Path) -> Path:
    if path.is_dir():
        return path
    candidate = Path("baselines") / str(path) / "originals"
    if candidate.is_dir():
        return candidate
    return path  # let existing error handling catch it
```

Apply to both `--old` and `--new` in `_run_compare()`.

Enables: `python -m fbdi compare --old 26A --new 26B`

### 4. No changes to

- `fbdi/detect_header.py` -- used as-is
- `fbdi/compare.py` -- used as-is
- `fbdi/config.py` -- used as-is
- Download logic in `tools/download_and_clear.py` -- Selenium code unchanged

## Execution Sequence

1. Restructure directories (rename Originals -> originals)
2. Update `tools/download_and_clear.py` with smart clearing + new paths + no BLANK_ prefix
3. Update `fbdi/cli.py` with `_resolve_dir()` shorthand
4. Regenerate blanks: `python tools/download_and_clear.py 26A --clear-only` then 26B
5. Validate blanks: spot-check that technical headers at row 5+ are preserved
6. Run comparison: `python -m fbdi compare --old 26A --new 26B --output Comparison_Report_26A_26B.xlsx`
7. Verify results: diagnose + spot checks

## Edge Cases

- **`detect_header_row()` returns None during clearing**: sheet skipped, logged as warning. Sheet retains all original data in the blank copy.
- **File >5MB times out during clearing**: file skipped, no blank created. Comparison uses originals so this doesn't affect comparison results. Blank is missing for client use -- logged in the clearing summary.
- **Corrupt XML/zip**: openpyxl can't load. File skipped during clearing. Same file would also fail during comparison (existing behavior, logged as FILE_ERROR).
- **`legacy_drawing` strip**: still applied after clearing to prevent openpyxl save errors on Oracle's malformed VML XML.
- **Windows case-insensitive filesystem**: renaming `Originals` to `originals` may need a two-step rename (`Originals` -> `originals_tmp` -> `originals`) on Windows.
