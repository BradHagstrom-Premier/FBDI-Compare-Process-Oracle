# Handoff: Build FBDI Comparison Engine (`compare.py`)

## Context

This project replaces a VBA macro (`FBDI_Compare.xlsm`) with a Python-based comparison engine. The macro compares Oracle FBDI (File-Based Data Import) template files between two release versions (e.g., 25D → 26A) and produces a `Comparison_Report.xlsx` showing every field-level change.

The VBA macro works but has critical brittleness: it determines which row contains field headers using a massive hardcoded if/else block mapping ~80+ filenames to specific row numbers (1, 3, 4, 5, 8, 9, 10, 16). The Python replacement must detect header rows dynamically. Additionally, the VBA has known bugs (documented below) that the Python version should fix.

### Key decisions already made:
- Full Python replacement of VBA — no fallback to Excel/VBA
- `openpyxl` for reading `.xlsm` files
- Dynamic header row detection using content-based heuristics
- Output must at minimum be a `.xlsx` file with the same 7-column structure as the VBA produces. If there is clear oppurtunity for improvement, then exceed the expectations.
- The existing `Comparison_Report_25D_26A.xlsx` is the validation target

### Known bugs in the VBA macro to fix:
1. **Typo: `rowRow = 10` instead of `FieldRow = 10`** — affects ~10 Scp* templates in the last ElseIf block (ScpCalendarImportTemplate, ScpRealTimeSupplyUpdatesImportTemplate, ScpRegionZoneMappingImportTemplate, ScpReservationImportTemplate, ScpSafetyStockLevelImportTemplate, ScpUOMImportTemplate, ScpWIPComponentDemandsImportTemplate, ScpWIPOperationResourceImportTemplate, ScpWorkOrderSuppliesImportTemplate, ScpZonesImportTemplate). These templates silently fall through to the default `FieldRow = 4`, which may be wrong.
2. **Typo: `fbid_name` instead of `fbdi_name`** — in the row-8 block for ChangeOrderImportTemplate. This template never matches and falls to default `FieldRow = 4`.
3. **Single FieldRow for both old and new files** — the VBA determines FieldRow from the old filename and uses it for both. If Oracle moves the header row between releases, the comparison breaks silently. Python version must detect header row independently per file per tab.
4. **`Range("A4").End(xlToRight).Column`** — the VBA always uses row 4 to find the last column, regardless of where the actual header row is. Should use the detected header row.

---

## Scope

### Files to create:
| File | Purpose |
|------|---------|
| `fbdi/__init__.py` | Package init |
| `fbdi/__main__.py` | Entry point for `python -m fbdi` |
| `fbdi/cli.py` | CLI argument parsing and dispatch |
| `fbdi/config.py` | Configuration constants (skip-tab names, output column headers) |
| `fbdi/detect_header.py` | Header row detection logic (separate for testability) |
| `fbdi/compare.py` | Core comparison engine module |
| `fbdi/utils.py` | Shared utilities (column letter generation, file matching) |
| `tests/__init__.py` | Test package init |
| `tests/test_utils.py` | Unit tests for column letter generation and file matching |
| `tests/test_detect_header.py` | Unit tests for header detection |
| `tests/test_compare.py` | Integration tests for comparison engine |
| `tests/validate_against_vba.py` | Acceptance test: diff Python output against VBA output |
| `tests/vba_fieldrow_map.json` | Extracted VBA FieldRow mapping for validation |

### Files to modify:
| File | Change |
|------|--------|
| `.gitignore` | Add baseline folders and report files |

### Files NOT to modify:
- `test.py` (Dan's downloader — wrapping comes later)
- `FBDI_Compare.xlsm` (reference only)
- `Comparison_Report_25D_26A.xlsx` (validation target, read-only)

---

## Step-by-Step Instructions

### Step 0: Environment setup, branch, and planning

Use the `claude-code-setup` plugin to set up the development environment:

```bash
pip install openpyxl pytest oletools
```

Use the `github` plugin to create a feature branch:

```bash
git checkout -b feat/compare-engine
```

Before writing any code:

1. Read `test.py` to understand the download pipeline this will eventually integrate with
2. Read this entire handoff file
3. Use the `brainstorming` skill to think through module architecture, edge cases, and potential failure modes before implementation
4. Use the `writing-plans` skill to create an execution plan for all steps below
5. Use the `context7` plugin to review the `openpyxl` API documentation, specifically: `load_workbook` with `.xlsm` files, accessing sheets by name and index, reading cell values, `ws.max_column`, `ws.max_row`, iterating rows and columns, `MergedCell` handling

Use the `feature-dev` plugin to scaffold the `fbdi/` package structure (all files listed in scope above, initially empty or with docstrings only).

Commit with `commit-commands` plugin: `chore: scaffold fbdi package structure`

### Step 1: Create `fbdi/config.py`

Create the configuration module with constants.

```python
# Tabs to skip during comparison (case-sensitive match against sheet names)
SKIP_TABS = {
    "Instructions and CSV Generation",
    "Instructions",
    "Options",
    "Create CSV",
    "reference",
    "Validation Report",
    "LOV",
    "XDO_METADATA",
    "Lookups",
}

# Output column headers for Comparison_Report.xlsx
REPORT_HEADERS = [
    "FBDI File",
    "FBDI Tab",
    "Column Letter",
    "Column Number",
    "Old FBDI Field Name",
    "New FBDI Field Name",
    "Difference?",
]
```

After creating this file, run `pyright-lsp` plugin to verify.

### Step 2: Create `fbdi/utils.py`

Use the `test-driven-development` skill: write tests in `tests/test_utils.py` first, then implement.

**Column letter generation:**
Convert a 1-based column index to an Excel column letter (1->A, 27->AA, 703->AAA). Use the standard algorithm, not the VBA's convoluted nested loop.

```python
def col_index_to_letter(index: int) -> str:
    """Convert 1-based column index to Excel column letter. 1->A, 27->AA, etc."""
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result
```

**File matching:**
Match old and new FBDI files by filename stem (case-insensitive), supporting both `.xlsm` and `.xlsx` extensions.

```python
def match_fbdi_files(old_dir: Path, new_dir: Path) -> tuple[list[tuple[Path, Path]], list[Path], list[Path]]:
    """Match FBDI files between old and new directories by filename stem.
    Returns: (matched_pairs, old_only, new_only)
    """
```

**Tests** (`tests/test_utils.py`):
- `test_col_index_to_letter`: verify A=1, Z=26, AA=27, AZ=52, BA=53, AAA=703
- `test_col_index_to_letter_edge_cases`: verify behavior for index=0 and negative values
- `test_match_fbdi_files`: create temp directories with known files, verify correct matching, old-only, and new-only lists
- `test_match_fbdi_files_case_insensitive`: verify case-insensitive stem matching

Run `pyright-lsp` plugin after implementation.
Commit with `commit-commands` plugin: `feat: add fbdi utility functions with tests`

### Step 3: Create `fbdi/detect_header.py`

This is the critical module. Use the `test-driven-development` skill.

**Algorithm:**

```python
import re
from openpyxl.worksheet.worksheet import Worksheet

# Oracle FBDI column names follow this pattern: uppercase letters, digits, underscores
HEADER_PATTERN = re.compile(r'^[A-Z][A-Z0-9_]+$')

def detect_header_row(ws: Worksheet, max_scan: int = 20) -> int | None:
    """Detect the header row in an FBDI worksheet.
    
    Scans rows 1 through max_scan. Scores each row on:
    - pattern_ratio (weight 0.5): fraction of non-empty cells matching Oracle's
      UPPER_SNAKE_CASE naming convention
    - fill_ratio (weight 0.3): fraction of columns that are non-empty
    - str_ratio (weight 0.2): fraction of non-empty cells that are strings
    
    Returns 1-indexed row number, or None if no confident match found.
    """
```

**Scoring details:**
- `pattern_ratio`: Count cells matching `^[A-Z][A-Z0-9_]+$` divided by non-empty cell count. This is the strongest signal. Oracle headers look like `INVOICE_NUMBER`, `PO_LINE_ID`, `ATTRIBUTE_CATEGORY`.
- `fill_ratio`: Non-empty cells divided by `ws.max_column`. Header rows are typically fully populated; instruction rows and title rows tend to span only a few columns.
- `str_ratio`: String-type cells divided by non-empty cell count. Headers are always strings; data rows might contain dates, numbers, or formulas.
- **Minimum thresholds**: Skip rows with fewer than 3 non-empty cells. Require final score > 0.4 to return a match. If no row passes threshold, return `None` and log a warning.
- **Confidence logging**: Log the detected row number and score for every file/tab processed using Python's `logging` module at DEBUG level. This makes debugging straightforward when a detection is wrong.

**Edge cases to handle:**
- Merged cells in instruction rows (common in Oracle templates; openpyxl returns `MergedCell` objects for these; treat as empty)
- Rows where all cells are None (skip immediately)
- Templates where the header row is row 1 (some Scp* templates, UploadCreditDataTemplate)
- Templates with a header row on row 16 (XlaMappingsImportTemplate)

**Unit tests** (`tests/test_detect_header.py`):
- Create mock worksheets using openpyxl's in-memory `Workbook()` with headers at known rows
- Test with headers at rows 1, 3, 4, 5, 8, 10, 16
- Test with merged cells above the header row
- Test with empty worksheets (should return None)
- Test with instruction text in rows above the header
- Test that the highest-scoring row wins when multiple rows contain some uppercase text

Run `pyright-lsp` plugin after implementation.

**Then validate against real FBDI files:**

First, extract the complete VBA FieldRow mapping into `tests/vba_fieldrow_map.json`. The VBA source is in `FBDI_Compare.xlsm`. Extract it with oletools:

```bash
pip install oletools
python -c "
import zipfile, subprocess
with zipfile.ZipFile('FBDI_Compare.xlsm', 'r') as z:
    z.extract('xl/vbaProject.bin', '/tmp/fbdi_vba')
subprocess.run(['olevba', '/tmp/fbdi_vba/xl/vbaProject.bin'])
"
```

Parse the entire FieldRow if/elseif block from the VBA output. Produce a JSON object mapping every template name to its assigned FieldRow integer. Include as many entries as you deem necessary. Templates not in the map default to FieldRow=4. Mark entries affected by VBA bugs:

```json
{
    "_metadata": {
        "source": "FBDI_Compare.xlsm VBA FieldRow block",
        "default_row": 4,
        "bugs": {
            "rowRow_typo": "ScpCalendar* etc. assigned rowRow=10 instead of FieldRow=10, effective FieldRow=4",
            "fbid_name_typo": "ChangeOrderImportTemplate uses fbid_name, never matches, effective FieldRow=4"
        }
    },
    "WorkDefinitionTemplate": 5,
    "CseInstalledBaseAssetImport": 5,
    "...": "EXTRACT ALL ~80+ ENTRIES FROM THE VBA SOURCE"
}
```

**Known VBA bugs to account for during validation:**
- Templates in the `rowRow = 10` block (ScpCalendarImportTemplate, ScpRealTimeSupplyUpdatesImportTemplate, ScpRegionZoneMappingImportTemplate, ScpReservationImportTemplate, ScpSafetyStockLevelImportTemplate, ScpUOMImportTemplate, ScpWIPComponentDemandsImportTemplate, ScpWIPOperationResourceImportTemplate, ScpWorkOrderSuppliesImportTemplate, ScpZonesImportTemplate) actually got FieldRow=4 in VBA due to the `rowRow` typo. If the Python detector finds row 10, that is the CORRECT answer.
- ChangeOrderImportTemplate (`fbid_name` typo) got FieldRow=4 in VBA. If the Python detector finds row 8, that is likely correct.

Write a validation script that runs `detect_header_row` against every non-skipped tab in every `.xlsm` file in the `25D/` folder. Compare results against the JSON map. Print a structured report: matches, expected mismatches (VBA bugs), and unexpected mismatches. Use the `systematic-debugging` skill for any unexpected mismatches.

Commit with `commit-commands` plugin: `feat: add dynamic header row detection with real-file validation`

### Step 4: Create `fbdi/compare.py`

The core comparison engine. This is the largest module. Use the `test-driven-development` skill.

**Data structures and function signatures:**

```python
from pathlib import Path
from dataclasses import dataclass

@dataclass
class ComparisonRow:
    fbdi_file: str          # Template filename (no extension)
    fbdi_tab: str           # Sheet name
    column_letter: str      # Excel column letter
    column_number: int      # 1-based column index
    old_field_name: str | None   # Header value in old file (None if column is new)
    new_field_name: str | None   # Header value in new file (None if column was removed)
    difference: str         # "YES" or "NO"

def compare_fbdi_pair(
    old_path: Path,
    new_path: Path,
) -> list[ComparisonRow]:
    """Compare one old/new FBDI template pair across all non-skipped tabs.
    Returns all rows (including unchanged) -- filtering happens at output stage.
    """

def compare_all(
    old_dir: Path,
    new_dir: Path,
    output_path: Path,
    changes_only: bool = True,
) -> Path:
    """Compare all matched FBDI pairs and write Comparison_Report.xlsx.
    If changes_only=True (default), only rows where difference="YES" are written.
    Returns path to the output file.
    """
```

**Comparison logic per tab:**

1. Detect header row independently for old and new file using `detect_header_row`
2. Read all header values from the detected row in both files
3. Find the last populated column using the detected header row (not hardcoded row 4 like the VBA does)
4. Align by position (column index), NOT by name. This matches the VBA behavior. Column 1 old vs column 1 new, column 2 old vs column 2 new, etc.
5. Handle length differences: if new file has more columns, old_field_name is None for the extras (additions). If old file has more columns, new_field_name is None (removals).
6. Generate column letter and column number for each position
7. Mark difference as "YES" if old != new, "NO" if they match

**Tab matching between old and new files:**
The VBA iterates tabs from the OLD file and looks up the same tab name in the NEW file. Do the same. If a tab exists in old but not new, log a warning. If a tab exists in new but not old, log it as well (the VBA silently drops this information).

**Writing the output:**
Use `openpyxl` to create `Comparison_Report.xlsx`:
- Sheet name: "Sheet1"
- Row 1: Headers (bold, with autofilter)
- Font: Calibri 11pt (matching VBA output formatting)
- If `changes_only=True`, only write rows where difference="YES"
- Autofit column widths (approximate: set each column width to max content length + 2)

**Tests** (`tests/test_compare.py`):
- Test `compare_fbdi_pair` with two in-memory workbooks: identical headers -> all "NO"; different headers -> "YES" where changed; different lengths -> additions/removals marked correctly
- Test skip-tab filtering: tabs in SKIP_TABS should be excluded from output
- Test that tab matching logs warnings for tabs only in old or only in new

Run `pyright-lsp` plugin after implementation.
Commit with `commit-commands` plugin: `feat: add FBDI comparison engine`

### Step 5: Create `fbdi/cli.py`, `fbdi/__main__.py`, and `fbdi/__init__.py`

**`fbdi/__main__.py`** enables `python -m fbdi`:
```python
from fbdi.cli import main

if __name__ == "__main__":
    main()
```

**`fbdi/cli.py`** uses `argparse`:

```
python -m fbdi compare --old 25D --new 26A [--output Comparison_Report.xlsx] [--all-rows] [--verbose]
```

Arguments:
- `--old`: Path to directory containing old FBDI templates (required)
- `--new`: Path to directory containing new FBDI templates (required)
- `--output`: Output file path (default: `Comparison_Report.xlsx` in current directory)
- `--all-rows`: Include unchanged rows in output (default: changes only)
- `--verbose`: Set logging to DEBUG (shows header detection scores for every tab)

The CLI should print a summary on completion:
- Count of matched file pairs
- Count of old-only and new-only files (with filenames listed)
- Count of total changes found
- Output file path

Run `pyright-lsp` plugin after implementation.
Commit with `commit-commands` plugin: `feat: add CLI entry point for FBDI comparison`

### Step 6: Validation against VBA output

This is the acceptance test. Use the `executing-plans` skill from the `superpowers` plugin to systematically work through validation.

**Validation script** (`tests/validate_against_vba.py`):

1. Load `Comparison_Report_25D_26A.xlsx` (VBA output) into a list of tuples
2. Run `compare_all("25D", "26A", "test_output.xlsx", changes_only=True)`
3. Load the Python output into a list of tuples
4. Compare row-by-row:
   - Sort both datasets by (FBDI File, FBDI Tab, Column Number) for stable comparison
   - For each row, check if FBDI File, FBDI Tab, Column Number, Old Field Name, New Field Name, and Difference match
5. Print a structured report:
   - **Exact matches**: count and percentage of rows identical in both outputs
   - **Python-only rows**: rows in Python output but not VBA (likely VBA bugs we're fixing)
   - **VBA-only rows**: rows in VBA output but not Python (potential Python bugs to investigate)
   - **Value mismatches**: same file/tab/column but different field names (header row detection disagreement)

**Expected discrepancies** (VBA bugs, not Python bugs):
- Templates affected by `rowRow = 10` typo should show different results if header is actually at row 10 vs the VBA's effective row 4
- ChangeOrderImportTemplate may show different results due to `fbid_name` typo

For any unexpected discrepancies, use the `systematic-debugging` skill to investigate root cause. Document every discrepancy: if the Python version is correct and the VBA was wrong, note it explicitly. If the Python header detection is wrong for a specific template, fix the detection and re-run.

**Target**: >95% row match with all discrepancies explained.

After validation passes:
1. Use `coderabbit` plugin to review all code in `fbdi/` and `tests/` holistically
2. Use `autofix` skill to apply any fixes from the coderabbit review
3. Use `code-simplifier` plugin to reduce complexity where possible
4. Run `pyright-lsp` plugin one final time across all `.py` files
5. Use the `verification-before-completion` skill to confirm all verification criteria are met

Commit with `commit-commands` plugin: `test: add VBA validation and apply review fixes`

### Step 7: Update `.gitignore` and finalize

Add these lines to `.gitignore`:
```
25D/
26A/
baselines/
Comparison_Report*.xlsx
test_output.xlsx
```

Commit with `commit-commands` plugin: `chore: update gitignore for baseline folders and reports`

### Step 8: Branch close-out

Use the `finishing-a-development-branch` skill to prepare the branch for merge:
1. Ensure all tests pass: `pytest tests/`
2. Ensure validation passes: `python tests/validate_against_vba.py`
3. Use the `pr-review-toolkit` plugin to review the branch before merge
4. Use the `github` plugin to push the branch and create a PR to main

Do NOT merge the PR. Brad will review and merge manually.

---

## Verification Criteria

The implementation is complete when:

1. **`python -m fbdi compare --old 25D --new 26A`** runs without errors and produces `Comparison_Report.xlsx`
2. **Header detection** correctly identifies header rows for all templates in 25D and 26A folders, validated against the extracted VBA FieldRow map, with known VBA bugs documented
3. **Unit tests pass**: `pytest tests/` with all tests green in test_detect_header.py, test_compare.py, test_utils.py
4. **Validation test**: `python tests/validate_against_vba.py` shows >95% row match with `Comparison_Report_25D_26A.xlsx`, with every discrepancy explained as a VBA bug or documented improvement
5. **Output format**: The generated .xlsx has 7 columns with correct headers, bold header row, autofilter, Calibri 11pt font, and only changed rows when `changes_only=True`
6. **No regressions**: Column letters are correct (A through AZ+ range), column numbers are 1-based sequential, file/tab names match the VBA output format
7. **`pyright-lsp`** reports no type errors across `fbdi/` package
8. **`coderabbit`** code review completed with no critical findings unresolved
9. **PR created** on `feat/compare-engine` branch, ready for Brad's review

---

## Migration Notes

```bash
# Install dependencies
pip install openpyxl pytest oletools

# Run comparison
python -m fbdi compare --old 25D --new 26A

# Run with verbose logging (shows header detection details)
python -m fbdi compare --old 25D --new 26A --verbose

# Run tests
pytest tests/

# Run VBA validation
python tests/validate_against_vba.py
```

---

## Architecture Notes for Future Steps

This handoff builds only the comparison engine. Future handoffs will add:
- **`fbdi/download.py`** — wrapping Dan's `test.py` downloader
- **`fbdi/report.py`** — generating the condensed Audrey-facing docx from the comparison report
- **`python -m fbdi run`** — full pipeline combining download + compare + report

The package structure is designed to accommodate these additions. Do not build them now.