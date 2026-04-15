# FBDI Master Catalog Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate `FBDI_Master_Catalog.xlsx` — a per-release snapshot of every FBDI file, tab, and column (in order) with label, technical name, data type, length, scale, and required flag. One release tab per release, plus consolidated `Issues` and `Drift` tabs.

**Architecture:** Decoupled new module (`fbdi/catalog.py`) with two pure helpers (`fbdi/catalog_normalize.py`, `fbdi/type_parser.py`). Zero changes to `compare.py`, `clear.py`, `diagnose.py`. CLI adds a `catalog` subcommand. Subprocess-per-file isolation with 120s timeout mirrors `compare.py`'s pattern.

**Tech Stack:** Python 3.14, openpyxl, multiprocessing, pytest. No new dependencies.

**Reference spec:** `docs/superpowers/specs/2026-04-15-fbdi-master-catalog-design.md`

**Graphify consultation:** `graphify-out/GRAPH_REPORT.md` was reviewed and confirmed the plan's pattern choices are aligned with the existing codebase communities:
- Task 7 subprocess worker mirrors Community 10 (`compare_all` / `_compare_worker`).
- Task 11 CLI `_run_catalog` mirrors Community 16 (`_run_compare` / `_run_diagnose`).
- Task 3 dataclasses follow the `ComparisonRow` pattern (god node, 9 edges).
- Task 5's reuse of `detect_header_row` stays inside Community 15 (header detection).
- Test helper names (`_make_thin_tab`, `_make_rich_tab`, `_make_rich_xlsm`) don't collide with existing helpers (`_create_fbdi_workbook`, `_make_ws_with_headers`, `_make_fbdi_workbook`, `_make_workbook`).

One stale entry noted in the graph — "FILE_ERROR Tables — Unreadable xlsm (2)" reflects an older snapshot; the current count is 6 and is listed by name in the spec. No plan change needed. The graph will be rebuilt in Task 14 to capture the new catalog modules.

---

## File Structure

**Create:**
- `fbdi/catalog_normalize.py` — `normalize_label()` helper
- `fbdi/type_parser.py` — `parse_data_type()` + `ParsedType` dataclass
- `fbdi/catalog.py` — dataclasses, extraction, drift computation, workbook writer, orchestrator
- `tests/test_catalog_normalize.py`
- `tests/test_type_parser.py`
- `tests/test_catalog.py`

**Modify:**
- `fbdi/cli.py` — add `catalog` subcommand
- `fbdi/config.py` — add `CATALOG_TIMEOUT = 120`

Each file has one responsibility. `catalog_normalize.py` is a single pure function (testable in isolation). `type_parser.py` is a single pure function + dataclass. `catalog.py` does everything that needs `openpyxl` + subprocess orchestration. `cli.py` stays a thin dispatch layer.

---

## Task Overview

1. Pure helper: `normalize_label`
2. Pure helper: `parse_data_type`
3. Dataclasses + config constant
4. `extract_tab_rows` — thin-tab path (Tier 2)
5. `extract_tab_rows` — rich-tab path (Tier 1 with metadata row detection)
6. `extract_file` — file-level orchestration, FILE_ERROR → IssueRow
7. Subprocess worker + timeout handling (`_catalog_worker`)
8. `_compute_drift` — position-aligned diff with `change_type` classification
9. Workbook writer — per-release tab, Issues tab, Drift tab
10. `generate_catalog` — end-to-end idempotent orchestration
11. CLI subcommand `python -m fbdi catalog --release 26B`
12. End-to-end synthetic test (two mini "releases", every change_type classification)
13. Real-world smoke test against `baselines/26A` and `baselines/26B`
14. Documentation updates (`CLAUDE.md`, `NEXT_STEPS.md`)

---

## Task 1: Pure helper — `normalize_label`

**Files:**
- Create: `fbdi/catalog_normalize.py`
- Test: `tests/test_catalog_normalize.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/test_catalog_normalize.py
"""Tests for fbdi.catalog_normalize."""

from fbdi.catalog_normalize import normalize_label


class TestNormalizeLabel:
    def test_strips_leading_asterisk(self):
        assert normalize_label("*Source Budget Type") == "Source Budget Type"

    def test_strips_punctuation_keeps_alphanumeric_and_underscore(self):
        assert normalize_label("$Weird, Chars!") == "Weird Chars"

    def test_preserves_underscore(self):
        assert normalize_label("COLUMN_NAME") == "COLUMN_NAME"

    def test_preserves_mixed_snake_case(self):
        assert normalize_label("my_col_name") == "my_col_name"

    def test_collapses_runs_of_whitespace(self):
        assert normalize_label("  *Foo  Bar  ") == "Foo Bar"

    def test_empty_string(self):
        assert normalize_label("") == ""

    def test_none_returns_empty(self):
        assert normalize_label(None) == ""

    def test_only_punctuation_returns_empty(self):
        assert normalize_label("!!!") == ""

    def test_digits_preserved(self):
        assert normalize_label("Column123") == "Column123"

    def test_unicode_alphanumerics_pass_through(self):
        # Python's str.isalnum() returns True for Unicode letters
        assert normalize_label("Café Name") == "Café Name"

    def test_collapses_tabs_and_newlines(self):
        assert normalize_label("Foo\tBar\nBaz") == "Foo Bar Baz"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_catalog_normalize.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'fbdi.catalog_normalize'`

- [ ] **Step 3: Implement**

```python
# fbdi/catalog_normalize.py
"""Normalize user-facing FBDI column labels for the master catalog.

Strips characters Applaud doesn't handle well (asterisks, punctuation,
symbols) while preserving alphanumerics, underscores, and whitespace.
Applied only to labels — technical UPPER_SNAKE_CASE names are untouched
because they are already canonical by construction.
"""


def normalize_label(s: str | None) -> str:
    """Strip non-alphanumeric/underscore/whitespace, collapse whitespace, trim.

    "Alphanumeric" uses Python's Unicode-aware str.isalnum(), so non-ASCII
    letters (e.g., accented characters) pass through; only punctuation and
    symbols are stripped. Whitespace runs collapse to a single space.
    """
    if not s:
        return ""
    kept = [ch for ch in s if ch.isalnum() or ch == "_" or ch.isspace()]
    return " ".join("".join(kept).split())
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_catalog_normalize.py -v`
Expected: PASS — 11 tests

- [ ] **Step 5: Commit**

```bash
git add fbdi/catalog_normalize.py tests/test_catalog_normalize.py
git commit -m "feat(catalog): add normalize_label helper

Strips non-alphanumeric/underscore/whitespace characters from FBDI
column labels, collapses whitespace, and trims. Prepares labels for
Applaud MDB compatibility.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: Pure helper — `parse_data_type`

**Files:**
- Create: `fbdi/type_parser.py`
- Test: `tests/test_type_parser.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/test_type_parser.py
"""Tests for fbdi.type_parser."""

from fbdi.type_parser import parse_data_type, ParsedType


class TestParseDataType:
    def test_varchar2_with_char_suffix(self):
        result = parse_data_type("VARCHAR2(5 CHAR)")
        assert result == ParsedType("VARCHAR2", 5, None, False)

    def test_varchar2_large_with_char_suffix(self):
        result = parse_data_type("VARCHAR2(2048 CHAR)")
        assert result == ParsedType("VARCHAR2", 2048, None, False)

    def test_varchar2_without_char_suffix(self):
        result = parse_data_type("VARCHAR2(80)")
        assert result == ParsedType("VARCHAR2", 80, None, False)

    def test_lowercase_varchar2_normalizes(self):
        result = parse_data_type("Varchar2(250)")
        assert result == ParsedType("VARCHAR2", 250, None, False)

    def test_number_precision_only(self):
        result = parse_data_type("NUMBER(18)")
        assert result == ParsedType("NUMBER", 18, None, False)

    def test_number_with_scale(self):
        result = parse_data_type("NUMBER(18,4)")
        assert result == ParsedType("NUMBER", 18, 4, False)

    def test_number_with_scale_and_spaces(self):
        result = parse_data_type("NUMBER(18, 4)")
        assert result == ParsedType("NUMBER", 18, 4, False)

    def test_date_no_parens(self):
        result = parse_data_type("DATE")
        assert result == ParsedType("DATE", None, None, False)

    def test_clob_no_parens(self):
        result = parse_data_type("CLOB")
        assert result == ParsedType("CLOB", None, None, False)

    def test_blob_no_parens(self):
        result = parse_data_type("BLOB")
        assert result == ParsedType("BLOB", None, None, False)

    def test_varchar2_with_byte_suffix(self):
        # Some templates use BYTE instead of CHAR
        result = parse_data_type("VARCHAR2(100 BYTE)")
        assert result == ParsedType("VARCHAR2", 100, None, False)

    def test_empty_string_no_warning(self):
        # Empty input is a legitimate blank, not a parse failure
        result = parse_data_type("")
        assert result == ParsedType("", None, None, False)

    def test_none_no_warning(self):
        result = parse_data_type(None)
        assert result == ParsedType("", None, None, False)

    def test_whitespace_only_no_warning(self):
        result = parse_data_type("   ")
        assert result == ParsedType("", None, None, False)

    def test_garbage_string_sets_warning(self):
        result = parse_data_type("???weird junk???")
        assert result.parse_warning is True
        assert result.data_type == ""
        assert result.length is None
        assert result.scale is None

    def test_extra_text_sets_warning(self):
        result = parse_data_type("VARCHAR2(50) NOT NULL DEFAULT 'x'")
        assert result.parse_warning is True
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_type_parser.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'fbdi.type_parser'`

- [ ] **Step 3: Implement**

```python
# fbdi/type_parser.py
"""Parse Oracle data-type strings from FBDI templates into structured fields.

FBDI templates store types in a 'Data Type' row as strings like:
  VARCHAR2(5 CHAR), VARCHAR2(2048 CHAR), VARCHAR2(80), Varchar2(250),
  NUMBER(18), NUMBER(18,4), DATE, CLOB, BLOB

This module parses those strings once so downstream comparison to Applaud
doesn't re-parse on every run.
"""

import re
from dataclasses import dataclass


@dataclass
class ParsedType:
    """Result of parsing a data-type string.

    data_type is uppercase ('VARCHAR2', 'NUMBER', 'DATE'). Empty string
    means the input was blank/None. length and scale are None when
    absent. parse_warning is True only for non-empty inputs that couldn't
    be decoded; blank inputs are not warnings.
    """
    data_type: str
    length: int | None
    scale: int | None
    parse_warning: bool


# Shape:
#   TYPENAME
#   TYPENAME(length)
#   TYPENAME(length CHAR|BYTE)
#   TYPENAME(length,scale)
# Case-insensitive; leading/trailing whitespace tolerated.
_TYPE_RE = re.compile(
    r"^\s*"
    r"([A-Za-z][A-Za-z0-9]*)"              # type name
    r"\s*"
    r"(?:"
        r"\(\s*"
        r"(\d+)"                           # length / precision
        r"(?:\s*,\s*(\d+))?"               # optional scale
        r"(?:\s+(?:CHAR|BYTE))?"           # optional CHAR|BYTE suffix
        r"\s*\)"
    r")?"
    r"\s*$",
    re.IGNORECASE,
)


def parse_data_type(raw: str | None) -> ParsedType:
    """Parse an Oracle data-type string into (data_type, length, scale).

    Returns ParsedType with parse_warning=True when raw is non-empty but
    doesn't match any known shape. Blank/None returns an empty ParsedType
    with parse_warning=False (blank is legitimate, not a failure).
    """
    if raw is None or not str(raw).strip():
        return ParsedType("", None, None, False)

    m = _TYPE_RE.match(str(raw))
    if not m:
        return ParsedType("", None, None, True)

    dtype = m.group(1).upper()
    length = int(m.group(2)) if m.group(2) else None
    scale = int(m.group(3)) if m.group(3) else None
    return ParsedType(dtype, length, scale, False)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_type_parser.py -v`
Expected: PASS — 16 tests

- [ ] **Step 5: Commit**

```bash
git add fbdi/type_parser.py tests/test_type_parser.py
git commit -m "feat(catalog): add parse_data_type helper

Parses Oracle data-type strings (VARCHAR2(N CHAR), NUMBER(p,s), DATE,
CLOB, etc.) into structured fields. Flags unrecognized strings with
parse_warning=True for downstream Issues reporting.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: Dataclasses + config constant

**Files:**
- Create: `fbdi/catalog.py` (just the dataclasses for now)
- Modify: `fbdi/config.py` — add `CATALOG_TIMEOUT`

- [ ] **Step 1: Add timeout constant to config**

Edit `fbdi/config.py`. Append at the end of the file:

```python
# Per-file timeout (seconds) for catalog subprocess workers.
# Mirrors COMPARE_TIMEOUT in compare.py; isolates openpyxl resource leaks.
CATALOG_TIMEOUT = 120
```

- [ ] **Step 2: Create catalog.py with dataclasses**

```python
# fbdi/catalog.py
"""FBDI Master Catalog — per-release snapshot generator.

Generates FBDI_Master_Catalog.xlsx with:
  - One tab per Oracle release (e.g., 26A, 26B): flat (file, tab, position,
    label, technical, type, length, scale, required) snapshot.
  - Issues tab: consolidated coverage gaps across all releases.
  - Drift tab: position-aligned diff between the two most-recent releases.

Uses subprocess-per-file isolation (mirroring compare.py) with a 120s
timeout to handle openpyxl resource accumulation. Re-running for an
existing release regenerates only that release's tab plus Issues/Drift.
"""

from dataclasses import dataclass


@dataclass
class CatalogRow:
    """One row per (release, file, tab, column position)."""
    release: str
    file_name: str
    tab_name: str
    position: int               # 1-based column index
    column_label: str           # normalized user-friendly label
    column_technical: str       # UPPER_SNAKE_CASE; blank for thin tabs
    data_type: str              # uppercase; blank for thin tabs or parse failures
    length: int | None          # None when absent; blank in output
    scale: int | None           # None when absent; blank in output
    data_type_raw: str          # original string; blank for thin tabs
    required: bool | None       # True/False; None when unknown


@dataclass
class IssueRow:
    """One row per coverage gap or error condition."""
    release: str
    file: str
    tab: str                    # empty for file-level issues
    issue_type: str             # FILE_ERROR | TIMEOUT | SUBPROCESS_FAILED | NO_HEADER | TYPE_PARSE_WARNING
    detail: str


@dataclass
class DriftRow:
    """One row per position where two releases differ."""
    file: str
    tab: str
    position: int
    col_label_old: str
    col_label_new: str
    col_technical_old: str
    col_technical_new: str
    data_type_old: str
    data_type_new: str
    length_old: str
    length_new: str
    required_old: str
    required_new: str
    change_type: str            # ADDED | REMOVED | RENAMED | TYPE_CHANGED | LENGTH_CHANGED | REQUIRED_CHANGED | MULTI
```

- [ ] **Step 3: Quick sanity check — import works**

Run: `python -c "from fbdi.catalog import CatalogRow, IssueRow, DriftRow; from fbdi.config import CATALOG_TIMEOUT; print(CATALOG_TIMEOUT)"`
Expected: `120`

- [ ] **Step 4: Commit**

```bash
git add fbdi/catalog.py fbdi/config.py
git commit -m "feat(catalog): add catalog dataclasses and timeout constant

Adds CatalogRow, IssueRow, DriftRow dataclasses as the skeleton for
fbdi/catalog.py, and CATALOG_TIMEOUT=120 to config.py. No behavior yet.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: `extract_tab_rows` — thin-tab (Tier 2) path

**Files:**
- Modify: `fbdi/catalog.py` — add `extract_tab_rows`
- Test: `tests/test_catalog.py`

Thin tabs have only a user-label row (sometimes with `*` prefixes for required). No inline data type, length, or technical name. Simpler path — implement first.

- [ ] **Step 1: Write failing tests**

```python
# tests/test_catalog.py
"""Tests for fbdi.catalog — master catalog generation."""

import pytest
from pathlib import Path
from openpyxl import Workbook, load_workbook

from fbdi.catalog import (
    CatalogRow,
    IssueRow,
    DriftRow,
    extract_tab_rows,
)


def _make_thin_tab(ws, labels: list[str], header_row: int = 4):
    """Build a thin-tab workbook: just a title/legend and a label row."""
    ws.cell(row=2, column=1, value="Some Import")
    ws.cell(row=3, column=1, value="* Required")
    for col_idx, label in enumerate(labels, start=1):
        ws.cell(row=header_row, column=col_idx, value=label)


class TestExtractTabRowsThin:
    def test_thin_tab_labels_only(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "XCC_BUDGET_INTERFACE"
        _make_thin_tab(ws, [
            "*Source Budget Type",
            "*Source Budget Name",
            "Line Number",
            "Amount",
        ])
        rows, issues = extract_tab_rows(
            ws, file_stem="BudgetImportTemplate", release="26B"
        )

        assert issues == []
        assert len(rows) == 4
        # Position 1: required (had asterisk), normalized label
        assert rows[0].position == 1
        assert rows[0].column_label == "Source Budget Type"
        assert rows[0].column_technical == ""
        assert rows[0].data_type == ""
        assert rows[0].length is None
        assert rows[0].scale is None
        assert rows[0].data_type_raw == ""
        assert rows[0].required is True
        # Position 2: required
        assert rows[1].column_label == "Source Budget Name"
        assert rows[1].required is True
        # Position 3: not required (no asterisk)
        assert rows[2].column_label == "Line Number"
        assert rows[2].required is False
        # Position 4: not required
        assert rows[3].column_label == "Amount"
        assert rows[3].required is False

    def test_thin_tab_sets_release_and_file_and_tab(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "MY_TAB"
        _make_thin_tab(ws, ["*Field A", "Field B"])
        rows, _ = extract_tab_rows(ws, file_stem="MyTemplate", release="26A")
        assert rows[0].release == "26A"
        assert rows[0].file_name == "MyTemplate"
        assert rows[0].tab_name == "MY_TAB"

    def test_thin_tab_no_header_emits_issue(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "EmptyTab"
        # Only a title — no detectable header row
        ws.cell(row=1, column=1, value="Just a title")

        rows, issues = extract_tab_rows(
            ws, file_stem="Tpl", release="26A"
        )

        assert rows == []
        assert len(issues) == 1
        assert issues[0].issue_type == "NO_HEADER"
        assert issues[0].tab == "EmptyTab"
        assert issues[0].release == "26A"
        assert issues[0].file == "Tpl"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_catalog.py -v`
Expected: FAIL with `ImportError: cannot import name 'extract_tab_rows'`

- [ ] **Step 3: Implement thin-tab path**

Append to `fbdi/catalog.py`:

```python
# (imports to add at top of file)
import logging
import re
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from fbdi.catalog_normalize import normalize_label
from fbdi.detect_header import UPPER_SNAKE_PATTERN, detect_header_row
from fbdi.type_parser import parse_data_type

logger = logging.getLogger(__name__)

# Max columns scanned per row — same cap as compare.py (avoids phantom max_col=16384).
_MAX_COL = 500


def _read_row_values(ws: Worksheet, row_idx: int) -> list[str | None]:
    """Read one row into a list of trimmed string values (or None), trailing blanks trimmed."""
    row_cells = next(
        iter(ws.iter_rows(
            min_row=row_idx, max_row=row_idx,
            max_col=min(ws.max_column or 1, _MAX_COL),
        )),
        None,
    )
    if row_cells is None:
        return []
    raw = []
    for cell in row_cells:
        if isinstance(cell, MergedCell):
            raw.append(None)
        elif cell.value is not None and str(cell.value).strip() != "":
            raw.append(str(cell.value).strip())
        else:
            raw.append(None)
    # Trim trailing None
    last = 0
    for i, v in enumerate(raw, start=1):
        if v is not None:
            last = i
    return raw[:last]


def _is_tier1_header(values: list[str | None]) -> bool:
    """True if the row is dominated by UPPER_SNAKE_CASE technical names."""
    non_empty = [v for v in values if v]
    if not non_empty:
        return False
    snake = sum(
        1 for v in non_empty
        if isinstance(v, str) and UPPER_SNAKE_PATTERN.match(v.strip())
    )
    return (snake / len(non_empty)) >= 0.5


def extract_tab_rows(
    ws: Worksheet,
    file_stem: str,
    release: str,
) -> tuple[list[CatalogRow], list[IssueRow]]:
    """Extract catalog rows + any issues for one worksheet.

    Returns (rows, issues). On NO_HEADER, returns ([], [IssueRow]).
    Dispatches to thin-tab or rich-tab extraction based on the detected
    header row's content.
    """
    header_row = detect_header_row(ws)
    if header_row is None:
        return [], [IssueRow(
            release=release,
            file=file_stem,
            tab=ws.title,
            issue_type="NO_HEADER",
            detail=f"no confident header row in '{ws.title}'",
        )]

    header_values = _read_row_values(ws, header_row)
    if _is_tier1_header(header_values):
        # TODO in next task: rich-tab extraction
        return _extract_rich(ws, file_stem, release, header_row, header_values)
    return _extract_thin(ws, file_stem, release, header_row, header_values)


def _extract_thin(
    ws: Worksheet,
    file_stem: str,
    release: str,
    header_row: int,
    header_values: list[str | None],
) -> tuple[list[CatalogRow], list[IssueRow]]:
    """Thin-tab extraction: header row is a list of user-friendly labels
    (possibly asterisk-prefixed for required). No type/length/technical info."""
    rows: list[CatalogRow] = []
    for idx, raw in enumerate(header_values, start=1):
        if raw is None:
            continue
        raw_str = str(raw)
        required = raw_str.lstrip().startswith("*")
        rows.append(CatalogRow(
            release=release,
            file_name=file_stem,
            tab_name=ws.title,
            position=idx,
            column_label=normalize_label(raw_str),
            column_technical="",
            data_type="",
            length=None,
            scale=None,
            data_type_raw="",
            required=required,
        ))
    return rows, []


def _extract_rich(
    ws: Worksheet,
    file_stem: str,
    release: str,
    header_row: int,
    header_values: list[str | None],
) -> tuple[list[CatalogRow], list[IssueRow]]:
    """Rich-tab extraction — implemented in the next task.

    Temporary: returns empty to satisfy the Tier 1 path while tests for
    Tier 2 are green.
    """
    raise NotImplementedError("Rich-tab extraction lands in Task 5")
```

- [ ] **Step 4: Run tests to verify thin-tab tests pass**

Run: `python -m pytest tests/test_catalog.py::TestExtractTabRowsThin -v`
Expected: PASS — 3 tests

- [ ] **Step 5: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "feat(catalog): implement thin-tab extraction path

extract_tab_rows dispatches on Tier 1/Tier 2 header content. Thin tabs
(user-friendly labels only) produce rows with label+required populated
and type/technical/length blank. NO_HEADER tabs emit an IssueRow.

Rich-tab extraction raises NotImplementedError for now; lands in the
next commit.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: `extract_tab_rows` — rich-tab (Tier 1) path

**Files:**
- Modify: `fbdi/catalog.py` — implement `_extract_rich` + metadata row detection
- Test: `tests/test_catalog.py` — add rich-tab tests

Rich tabs have metadata rows above the header. Column A of each metadata row tags its role. We scan rows `1..header_row` and keyword-match column A.

- [ ] **Step 1: Write failing tests**

Append to `tests/test_catalog.py`:

```python
def _make_rich_tab(
    ws,
    labels: list[str],
    descriptions: list[str] | None = None,
    data_types: list[str] | None = None,
    required_flags: list[str] | None = None,
    technicals: list[str] | None = None,
    header_row: int = 5,
    table_name: str = "RCS_ATTACHMENTS_INT",
):
    """Build a rich-tab workbook with metadata rows + technical header."""
    def _row(row_idx, label_a, values):
        ws.cell(row=row_idx, column=1, value=label_a)
        for col_idx, v in enumerate(values, start=2):
            ws.cell(row=row_idx, column=col_idx, value=v)
    # Header row: "Column name of the Table X" in col A, then tech names col B..
    _row(header_row, f"Column name of the Table {table_name}", technicals or labels)
    # Name row above, then Description, Data Type, Required
    if header_row >= 2:
        _row(header_row - 1, "Required or Optional", required_flags or ["Optional"] * len(labels))
    if header_row >= 3:
        _row(header_row - 2, "Data Type", data_types or [""] * len(labels))
    if header_row >= 4:
        _row(header_row - 3, "Description", descriptions or [""] * len(labels))
    if header_row >= 5:
        _row(header_row - 4, "Name", labels)


class TestExtractTabRowsRich:
    def test_rich_tab_all_metadata_rows(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Attachment Details"
        _make_rich_tab(
            ws,
            labels=["Attachment Type", "Attachment Name", "Document ID"],
            data_types=["VARCHAR2(5 CHAR)", "VARCHAR2(2048 CHAR)", "NUMBER(18)"],
            required_flags=["Required", "Required", "Optional"],
            technicals=["ATTACHMENT_TYPE", "ATTACHMENT_NAME", "DOCUMENT_ID"],
            table_name="RCS_ATTACHMENTS_INT",
            header_row=5,
        )

        rows, issues = extract_tab_rows(ws, file_stem="AttachmentsImportTemplate", release="26B")

        assert issues == []
        assert len(rows) == 3
        r0 = rows[0]
        assert r0.position == 1
        assert r0.column_label == "Attachment Type"
        assert r0.column_technical == "ATTACHMENT_TYPE"
        assert r0.data_type == "VARCHAR2"
        assert r0.length == 5
        assert r0.scale is None
        assert r0.data_type_raw == "VARCHAR2(5 CHAR)"
        assert r0.required is True
        assert rows[1].length == 2048
        assert rows[2].data_type == "NUMBER"
        assert rows[2].length == 18
        assert rows[2].required is False

    def test_rich_tab_with_bom_on_required_row(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "TabWithBOM"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["VARCHAR2(80)"],
            technicals=["COL_A"],
        )
        # Overwrite the 'Required or Optional' col-A label with BOM-prefixed variant
        # In _make_rich_tab, required is at header_row - 1 = row 4
        ws.cell(row=4, column=1, value="\ufeffRequired or Optional")
        ws.cell(row=4, column=2, value="Required")

        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        assert rows[0].required is True

    def test_rich_tab_case_insensitive_col_a_match(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "MixedCase"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["VARCHAR2(80)"],
            technicals=["COL_A"],
        )
        # Lowercase the 'Data Type' label
        ws.cell(row=3, column=1, value="DATA type")  # row 3 = Data Type row
        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        assert rows[0].data_type == "VARCHAR2"
        assert rows[0].length == 80

    def test_rich_tab_missing_data_type_row(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "NoDataType"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["VARCHAR2(80)"],
            technicals=["COL_A"],
        )
        # Clear the Data Type row's column A label so it's not recognized
        ws.cell(row=3, column=1, value="Reserved for Future Use")

        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        # Type fields blank; other fields still populate
        assert rows[0].column_label == "Col A"
        assert rows[0].column_technical == "COL_A"
        assert rows[0].data_type == ""
        assert rows[0].length is None
        assert rows[0].data_type_raw == ""

    def test_rich_tab_unparseable_type_emits_warning_issue(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "WeirdType"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["???junk???"],
            technicals=["COL_A"],
        )
        rows, issues = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        # Row still emitted, raw preserved, parsed fields blank
        assert rows[0].data_type == ""
        assert rows[0].length is None
        assert rows[0].data_type_raw == "???junk???"
        # One warning issue for that raw string
        warnings = [i for i in issues if i.issue_type == "TYPE_PARSE_WARNING"]
        assert len(warnings) == 1
        assert warnings[0].detail == "???junk???"

    def test_rich_tab_asterisk_in_label_stripped(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "StarLabel"
        _make_rich_tab(
            ws, labels=["*Required Label"],
            data_types=["VARCHAR2(80)"],
            technicals=["REQ_LABEL"],
        )
        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        # Asterisk stripped by normalize_label; required comes from R4 row anyway
        assert rows[0].column_label == "Required Label"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_catalog.py::TestExtractTabRowsRich -v`
Expected: FAIL — `NotImplementedError`

- [ ] **Step 3: Implement rich-tab path**

Replace the `_extract_rich` stub in `fbdi/catalog.py` with:

```python
# Column-A keyword -> metadata role. Case-insensitive match after BOM strip.
# Note: header row (tier 1) has col A starting with "Column name of the Table";
# that's already handled by detect_header_row, we don't re-discover it here.
_COL_A_ROLE_KEYWORDS = {
    "name": "label",
    "data type": "type",
    "required or optional": "required",
}


def _find_metadata_rows(
    ws: Worksheet, header_row: int
) -> dict[str, int]:
    """Scan col A of rows 1..header_row-1. Return {role: row_idx} for matched rows.

    Match is case-insensitive with BOM prefix stripped. Unmatched rows are
    silently ignored (e.g., 'Description', 'Reserved for Future Use').
    """
    found: dict[str, int] = {}
    for r in range(1, header_row):
        cell = ws.cell(row=r, column=1).value
        if cell is None:
            continue
        key = str(cell).lstrip("\ufeff").strip().lower()
        role = _COL_A_ROLE_KEYWORDS.get(key)
        if role and role not in found:
            found[role] = r
    return found


def _extract_rich(
    ws: Worksheet,
    file_stem: str,
    release: str,
    header_row: int,
    header_values: list[str | None],
) -> tuple[list[CatalogRow], list[IssueRow]]:
    """Rich-tab extraction. header_values is the tier-1 (technical) row.

    Uses col-A keyword matching to locate the label/type/required rows
    above header_row. Missing role rows leave the corresponding field
    blank; other fields still populate. Unparseable data types emit
    TYPE_PARSE_WARNING issues but the row still emits (raw preserved).
    """
    roles = _find_metadata_rows(ws, header_row)

    label_values = _read_row_values(ws, roles["label"]) if "label" in roles else []
    type_values = _read_row_values(ws, roles["type"]) if "type" in roles else []
    required_values = _read_row_values(ws, roles["required"]) if "required" in roles else []

    rows: list[CatalogRow] = []
    issues: list[IssueRow] = []

    # Header values from column B onward in real templates, but _read_row_values
    # returns starting from column 1. For rich tabs, col A is the role tag;
    # actual data starts at col B (index 1 in the list = position 2 in the sheet).
    # We treat position = 1-based sheet column; position 1 (col A) is the role
    # tag, not a data column, so we skip it.

    def _val_at(values: list[str | None], sheet_col: int) -> str:
        """Return the cell value at sheet column `sheet_col` (1-based) or ''."""
        idx = sheet_col - 1
        if 0 <= idx < len(values):
            v = values[idx]
            return str(v) if v is not None else ""
        return ""

    # Iterate data columns (col B onward) in the header row
    for sheet_col in range(2, len(header_values) + 1):
        tech_raw = header_values[sheet_col - 1]
        if not tech_raw:
            continue
        technical = str(tech_raw).strip()
        label_raw = _val_at(label_values, sheet_col)
        type_raw = _val_at(type_values, sheet_col)
        req_raw = _val_at(required_values, sheet_col)

        parsed = parse_data_type(type_raw)
        if parsed.parse_warning:
            issues.append(IssueRow(
                release=release,
                file=file_stem,
                tab=ws.title,
                issue_type="TYPE_PARSE_WARNING",
                detail=type_raw,
            ))

        # Position is 1-based sheet column. We include col A as position 1 to be
        # internally consistent, but col A is the role-tag column so its column_technical
        # will be the header-row tag string — callers can filter if needed.
        # For simplicity and alignment with compare.py (which uses col_index_to_letter
        # starting at col A), we emit one row per data column starting at position 2.
        rows.append(CatalogRow(
            release=release,
            file_name=file_stem,
            tab_name=ws.title,
            position=sheet_col - 1,  # renumber data columns starting from 1
            column_label=normalize_label(label_raw),
            column_technical=technical,
            data_type=parsed.data_type,
            length=parsed.length,
            scale=parsed.scale,
            data_type_raw=type_raw,
            required=_parse_required_flag(req_raw),
        ))

    return rows, issues


def _parse_required_flag(raw: str) -> bool | None:
    """Parse 'Required'/'Optional' (case-insensitive) to bool. Unknown -> None."""
    if not raw:
        return None
    v = raw.strip().lower()
    if v.startswith("required"):
        return True
    if v.startswith("optional"):
        return False
    return None
```

Note the position renumbering: in rich tabs, sheet column A is Oracle's role-tag column (`Name`, `Data Type`, etc.), not a data column. So the catalog's `position=1` refers to the first *data* column (sheet column B). This matches how a downstream consumer would logically count columns in the template. Thin tabs have no role-tag column, so `position=1` already means sheet column A there. Both cases end up labeling the first data column as `position=1`.

- [ ] **Step 4: Run rich-tab tests and all catalog tests**

Run: `python -m pytest tests/test_catalog.py -v`
Expected: PASS — all existing thin + rich tests.

- [ ] **Step 5: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "feat(catalog): implement rich-tab extraction with metadata rows

Rich tabs (Tier 1, UPPER_SNAKE_CASE headers) use col-A keyword matching
to locate label/data_type/required rows above the header. Missing roles
leave fields blank; unparseable types emit TYPE_PARSE_WARNING issues
while preserving the raw string. Position numbering starts at the first
data column (col B in rich tabs, col A in thin tabs).

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: `extract_file` — file-level orchestration

**Files:**
- Modify: `fbdi/catalog.py`
- Test: `tests/test_catalog.py`

`extract_file` opens one workbook, iterates non-skipped tabs, delegates to `extract_tab_rows`, and converts load/read failures into FILE_ERROR issues.

- [ ] **Step 1: Write failing tests**

Append to `tests/test_catalog.py`:

```python
from fbdi.catalog import extract_file


class TestExtractFile:
    def test_extract_file_multiple_tabs(self, tmp_path):
        path = tmp_path / "MultiTab.xlsm"
        wb = Workbook()
        wb.remove(wb.active)

        # Thin tab
        ws1 = wb.create_sheet("THIN_TAB")
        _make_thin_tab(ws1, ["*Field A", "Field B"])

        # Rich tab
        ws2 = wb.create_sheet("RICH_TAB")
        _make_rich_tab(
            ws2,
            labels=["Col A", "Col B"],
            data_types=["VARCHAR2(50)", "NUMBER(10)"],
            required_flags=["Required", "Optional"],
            technicals=["COL_A", "COL_B"],
        )
        wb.save(path)

        rows, issues = extract_file(path, release="26B")
        tabs = {r.tab_name for r in rows}
        assert tabs == {"THIN_TAB", "RICH_TAB"}
        # Thin tab contributes 2 rows, rich tab contributes 2 rows
        assert len(rows) == 4
        assert issues == []

    def test_extract_file_skips_instruction_tabs(self, tmp_path):
        from fbdi.config import SKIP_TABS
        path = tmp_path / "WithInstructions.xlsm"
        wb = Workbook()
        wb.remove(wb.active)
        for name in list(SKIP_TABS)[:2]:
            ws = wb.create_sheet(name)
            ws.cell(row=1, column=1, value="Instruction content")
        data_ws = wb.create_sheet("DATA_TAB")
        _make_thin_tab(data_ws, ["*F1"])
        wb.save(path)

        rows, issues = extract_file(path, release="26B")
        tabs = {r.tab_name for r in rows}
        assert tabs == {"DATA_TAB"}
        assert issues == []

    def test_extract_file_load_error_yields_issue(self, tmp_path):
        path = tmp_path / "Corrupt.xlsm"
        path.write_bytes(b"not a real xlsx file")

        rows, issues = extract_file(path, release="26B")
        assert rows == []
        assert len(issues) == 1
        assert issues[0].issue_type == "FILE_ERROR"
        assert issues[0].file == "Corrupt"
        assert issues[0].tab == ""
        assert issues[0].release == "26B"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_catalog.py::TestExtractFile -v`
Expected: FAIL — `ImportError: cannot import name 'extract_file'`

- [ ] **Step 3: Implement `extract_file`**

Append to `fbdi/catalog.py`:

```python
from pathlib import Path
from openpyxl import load_workbook

from fbdi.config import SKIP_TABS


def extract_file(
    path: Path, release: str
) -> tuple[list[CatalogRow], list[IssueRow]]:
    """Open one .xlsm and extract catalog rows for every non-skipped tab.

    Returns (rows, issues). FILE_ERROR on load failure produces one
    IssueRow with tab="". Each data tab that extract_tab_rows flags with
    issues contributes its issues to the combined list.
    """
    file_stem = path.stem
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        return [], [IssueRow(
            release=release,
            file=file_stem,
            tab="",
            issue_type="FILE_ERROR",
            detail=f"{type(e).__name__}: {str(e)[:200]}",
        )]

    all_rows: list[CatalogRow] = []
    all_issues: list[IssueRow] = []
    try:
        for sheet_name in wb.sheetnames:
            if sheet_name in SKIP_TABS:
                continue
            ws = wb[sheet_name]
            rows, issues = extract_tab_rows(ws, file_stem=file_stem, release=release)
            all_rows.extend(rows)
            all_issues.extend(issues)
    finally:
        wb.close()
    return all_rows, all_issues
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_catalog.py -v`
Expected: PASS — all tests including new TestExtractFile

- [ ] **Step 5: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "feat(catalog): add extract_file for workbook-level orchestration

Opens one .xlsm in read-only mode, iterates non-skipped tabs, and
aggregates per-tab rows and issues. Converts openpyxl load failures
into FILE_ERROR IssueRows.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 7: Subprocess worker + timeout handling

**Files:**
- Modify: `fbdi/catalog.py`

Mirrors `_compare_worker` pattern. The worker runs in a subprocess and sends pickled result tuples through a `multiprocessing.Queue`.

- [ ] **Step 1: Add worker function and runner**

Append to `fbdi/catalog.py`:

```python
import multiprocessing

from fbdi.config import CATALOG_TIMEOUT


def _rows_to_tuples(rows: list[CatalogRow]) -> list[tuple]:
    return [(
        r.release, r.file_name, r.tab_name, r.position,
        r.column_label, r.column_technical,
        r.data_type, r.length, r.scale, r.data_type_raw,
        r.required,
    ) for r in rows]


def _issues_to_tuples(issues: list[IssueRow]) -> list[tuple]:
    return [(i.release, i.file, i.tab, i.issue_type, i.detail) for i in issues]


def _tuples_to_rows(tuples: list[tuple]) -> list[CatalogRow]:
    return [CatalogRow(*t) for t in tuples]


def _tuples_to_issues(tuples: list[tuple]) -> list[IssueRow]:
    return [IssueRow(*t) for t in tuples]


def _catalog_worker(path_str: str, release: str, queue: multiprocessing.Queue) -> None:
    """Subprocess entry point. Mirrors _compare_worker for resource isolation."""
    # Some templates hit openpyxl's Font.family.max=14 cap; 255 matches compare.
    from openpyxl.styles.fonts import Font as WorkerFont
    WorkerFont.family.max = 255

    try:
        rows, issues = extract_file(Path(path_str), release=release)
        queue.put((_rows_to_tuples(rows), _issues_to_tuples(issues)))
    except Exception as e:
        queue.put(f"ERROR: {type(e).__name__}: {e}")


def _run_file_in_subprocess(
    path: Path, release: str, timeout: int
) -> tuple[list[CatalogRow], list[IssueRow]]:
    """Run extract_file in a fresh subprocess with timeout. Returns issue rows on failure."""
    queue: multiprocessing.Queue = multiprocessing.Queue()
    proc = multiprocessing.Process(
        target=_catalog_worker, args=(str(path), release, queue)
    )
    proc.start()
    proc.join(timeout=timeout)

    file_stem = path.stem
    if proc.is_alive():
        proc.terminate()
        proc.join(5)
        return [], [IssueRow(
            release=release, file=file_stem, tab="",
            issue_type="TIMEOUT", detail=f"exceeded {timeout}s",
        )]

    if proc.exitcode != 0:
        return [], [IssueRow(
            release=release, file=file_stem, tab="",
            issue_type="SUBPROCESS_FAILED",
            detail=f"exit code {proc.exitcode}",
        )]

    try:
        result = queue.get_nowait()
    except Exception:
        return [], [IssueRow(
            release=release, file=file_stem, tab="",
            issue_type="SUBPROCESS_FAILED",
            detail="no result on queue",
        )]

    if isinstance(result, str) and result.startswith("ERROR:"):
        return [], [IssueRow(
            release=release, file=file_stem, tab="",
            issue_type="SUBPROCESS_FAILED",
            detail=result,
        )]

    row_tuples, issue_tuples = result
    return _tuples_to_rows(row_tuples), _tuples_to_issues(issue_tuples)
```

- [ ] **Step 2: Quick smoke test — worker runs without error on a real file**

Run:

```bash
python -c "
from pathlib import Path
from fbdi.catalog import _run_file_in_subprocess
rows, issues = _run_file_in_subprocess(
    Path('baselines/26B/originals/BudgetImportTemplate.xlsm'), '26B', 120
)
print(f'rows: {len(rows)}  issues: {len(issues)}')
print(f'first row: {rows[0] if rows else None}')
print(f'first issue: {issues[0] if issues else None}')
"
```

Expected output shape: `rows: <some number >0>  issues: 0` (or low), first row populated with BudgetImportTemplate data.

- [ ] **Step 3: Commit**

```bash
git add fbdi/catalog.py
git commit -m "feat(catalog): add subprocess worker with timeout handling

Mirrors compare.py subprocess-per-file isolation pattern. _catalog_worker
runs extract_file in a fresh process; _run_file_in_subprocess applies
CATALOG_TIMEOUT (120s) and converts TIMEOUT/SUBPROCESS_FAILED conditions
into IssueRows. No tests added for this thin wrapper — end-to-end tests
in Task 12 exercise the subprocess path.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 8: `_compute_drift` — diff classification

**Files:**
- Modify: `fbdi/catalog.py`
- Test: `tests/test_catalog.py`

Computes the Drift tab from two lists of `CatalogRow` representing the two most-recent releases. Position-aligned diff with `change_type` classification per the spec.

- [ ] **Step 1: Write failing tests**

Append to `tests/test_catalog.py`:

```python
from fbdi.catalog import _compute_drift


def _row(**kwargs) -> CatalogRow:
    defaults = dict(
        release="26A", file_name="F", tab_name="T", position=1,
        column_label="L", column_technical="T", data_type="VARCHAR2",
        length=50, scale=None, data_type_raw="VARCHAR2(50)", required=False,
    )
    defaults.update(kwargs)
    return CatalogRow(**defaults)


class TestComputeDrift:
    def test_unchanged_rows_not_in_drift(self):
        old = [_row(release="26A")]
        new = [_row(release="26B")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift == []

    def test_added_column(self):
        old = []
        new = [_row(release="26B")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].change_type == "ADDED"
        assert drift[0].col_label_old == ""
        assert drift[0].col_label_new == "L"

    def test_removed_column(self):
        old = [_row(release="26A")]
        new = []
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].change_type == "REMOVED"
        assert drift[0].col_label_new == ""

    def test_renamed_label_only(self):
        old = [_row(release="26A", column_label="Old Name")]
        new = [_row(release="26B", column_label="New Name")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].change_type == "RENAMED"

    def test_renamed_technical_only(self):
        old = [_row(release="26A", column_technical="OLD_NAME")]
        new = [_row(release="26B", column_technical="NEW_NAME")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "RENAMED"

    def test_type_changed_only(self):
        old = [_row(release="26A", data_type="VARCHAR2")]
        new = [_row(release="26B", data_type="NUMBER")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "TYPE_CHANGED"

    def test_length_changed_only(self):
        old = [_row(release="26A", length=50)]
        new = [_row(release="26B", length=100)]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "LENGTH_CHANGED"

    def test_scale_changed_classified_as_length(self):
        old = [_row(release="26A", data_type="NUMBER", length=18, scale=None, data_type_raw="NUMBER(18)")]
        new = [_row(release="26B", data_type="NUMBER", length=18, scale=4, data_type_raw="NUMBER(18,4)")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        # data_type same, length same, only scale differs — classified as LENGTH_CHANGED
        assert drift[0].change_type == "LENGTH_CHANGED"

    def test_required_changed_only(self):
        old = [_row(release="26A", required=False)]
        new = [_row(release="26B", required=True)]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "REQUIRED_CHANGED"

    def test_multi_change(self):
        old = [_row(release="26A", data_type="VARCHAR2", length=50, required=False)]
        new = [_row(release="26B", data_type="NUMBER", length=18, required=True)]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "MULTI"

    def test_aligns_by_file_tab_position(self):
        old = [
            _row(release="26A", file_name="F1", tab_name="T1", position=1),
            _row(release="26A", file_name="F2", tab_name="T2", position=1),
        ]
        new = [
            _row(release="26B", file_name="F1", tab_name="T1", position=1, column_label="NEW"),
            _row(release="26B", file_name="F2", tab_name="T2", position=1),
        ]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].file == "F1"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_catalog.py::TestComputeDrift -v`
Expected: FAIL — `ImportError: cannot import name '_compute_drift'`

- [ ] **Step 3: Implement `_compute_drift`**

Append to `fbdi/catalog.py`:

```python
def _compute_drift(
    old_rows: list[CatalogRow],
    new_rows: list[CatalogRow],
    release_old: str,
    release_new: str,
) -> list[DriftRow]:
    """Position-aligned diff between two release row sets.

    Aligns by (file_name, tab_name, position). Emits DriftRows only for
    positions where something changed. change_type classified by the
    rules in the spec: ADDED, REMOVED, RENAMED (name only),
    TYPE_CHANGED, LENGTH_CHANGED (length or scale), REQUIRED_CHANGED,
    MULTI when more than one axis changes.
    """
    def key(r: CatalogRow) -> tuple:
        return (r.file_name, r.tab_name, r.position)

    old_by_key = {key(r): r for r in old_rows}
    new_by_key = {key(r): r for r in new_rows}
    all_keys = sorted(set(old_by_key.keys()) | set(new_by_key.keys()))

    drift: list[DriftRow] = []
    for k in all_keys:
        old = old_by_key.get(k)
        new = new_by_key.get(k)
        if old is None:
            drift.append(_drift_row(None, new, "ADDED"))
            continue
        if new is None:
            drift.append(_drift_row(old, None, "REMOVED"))
            continue

        name_changed = (
            old.column_label != new.column_label
            or old.column_technical != new.column_technical
        )
        type_changed = old.data_type != new.data_type
        length_changed = (old.length != new.length) or (old.scale != new.scale)
        required_changed = old.required != new.required

        changed_axes = sum([name_changed, type_changed, length_changed, required_changed])
        if changed_axes == 0:
            continue
        if changed_axes > 1:
            ctype = "MULTI"
        elif name_changed:
            ctype = "RENAMED"
        elif type_changed:
            ctype = "TYPE_CHANGED"
        elif length_changed:
            ctype = "LENGTH_CHANGED"
        else:
            ctype = "REQUIRED_CHANGED"

        drift.append(_drift_row(old, new, ctype))
    return drift


def _drift_row(
    old: CatalogRow | None, new: CatalogRow | None, change_type: str
) -> DriftRow:
    """Build a DriftRow from optional old/new CatalogRows."""
    ref = new if new is not None else old
    assert ref is not None  # at least one side must exist
    return DriftRow(
        file=ref.file_name,
        tab=ref.tab_name,
        position=ref.position,
        col_label_old=old.column_label if old else "",
        col_label_new=new.column_label if new else "",
        col_technical_old=old.column_technical if old else "",
        col_technical_new=new.column_technical if new else "",
        data_type_old=_fmt_type(old),
        data_type_new=_fmt_type(new),
        length_old=_fmt_length(old),
        length_new=_fmt_length(new),
        required_old=_fmt_required(old),
        required_new=_fmt_required(new),
        change_type=change_type,
    )


def _fmt_type(r: CatalogRow | None) -> str:
    return r.data_type if r else ""


def _fmt_length(r: CatalogRow | None) -> str:
    if r is None or r.length is None:
        return ""
    if r.scale is not None:
        return f"{r.length},{r.scale}"
    return str(r.length)


def _fmt_required(r: CatalogRow | None) -> str:
    if r is None or r.required is None:
        return ""
    return "TRUE" if r.required else "FALSE"
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_catalog.py -v`
Expected: PASS — all tests including TestComputeDrift

- [ ] **Step 5: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "feat(catalog): add _compute_drift with change_type classification

Position-aligned diff between two release row sets keyed by (file, tab,
position). Classifies change_type as ADDED / REMOVED / RENAMED /
TYPE_CHANGED / LENGTH_CHANGED / REQUIRED_CHANGED / MULTI per the spec,
with first-single-axis match for single-axis changes and MULTI when
two or more axes differ.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 9: Workbook writer

**Files:**
- Modify: `fbdi/catalog.py`
- Test: `tests/test_catalog.py`

Writes the per-release data tab, Issues tab, and Drift tab into a workbook. Idempotent: regenerating the workbook from the same inputs produces content-identical output.

- [ ] **Step 1: Write failing tests**

Append to `tests/test_catalog.py`:

```python
from fbdi.catalog import _write_master_workbook


class TestWriteMasterWorkbook:
    def test_writes_release_tab_with_correct_headers(self, tmp_path):
        out = tmp_path / "Master.xlsx"
        rows_by_release = {
            "26A": [_row(release="26A", file_name="F", tab_name="T", position=1)],
        }
        _write_master_workbook(
            out,
            rows_by_release=rows_by_release,
            issues=[],
            drift=[],
            release_old=None,
            release_new="26A",
        )
        assert out.exists()
        wb = load_workbook(out)
        assert "26A" in wb.sheetnames
        assert "Issues" in wb.sheetnames
        assert "Drift" in wb.sheetnames
        ws = wb["26A"]
        headers = [c.value for c in ws[1]]
        assert headers == [
            "release", "file_name", "tab_name", "position",
            "column_label", "column_technical",
            "data_type", "length", "scale", "data_type_raw",
            "required",
        ]
        # Data row
        row2 = [c.value for c in ws[2]]
        assert row2[0] == "26A"
        assert row2[1] == "F"

    def test_writes_issues_tab(self, tmp_path):
        out = tmp_path / "Master.xlsx"
        issues = [IssueRow("26B", "F", "T", "FILE_ERROR", "boom")]
        _write_master_workbook(
            out, rows_by_release={}, issues=issues, drift=[],
            release_old=None, release_new=None,
        )
        wb = load_workbook(out)
        ws = wb["Issues"]
        headers = [c.value for c in ws[1]]
        assert headers == ["release", "file", "tab", "issue_type", "detail"]
        assert [c.value for c in ws[2]] == ["26B", "F", "T", "FILE_ERROR", "boom"]

    def test_writes_drift_tab(self, tmp_path):
        out = tmp_path / "Master.xlsx"
        drift = [DriftRow(
            file="F", tab="T", position=1,
            col_label_old="A", col_label_new="B",
            col_technical_old="A1", col_technical_new="B1",
            data_type_old="VARCHAR2", data_type_new="VARCHAR2",
            length_old="50", length_new="100",
            required_old="FALSE", required_new="FALSE",
            change_type="LENGTH_CHANGED",
        )]
        _write_master_workbook(
            out, rows_by_release={}, issues=[], drift=drift,
            release_old="26A", release_new="26B",
        )
        wb = load_workbook(out)
        ws = wb["Drift"]
        headers = [c.value for c in ws[1]]
        assert "col_label_26A" in headers
        assert "col_label_26B" in headers
        assert "change_type" in headers

    def test_idempotent_content(self, tmp_path):
        out1 = tmp_path / "M1.xlsx"
        out2 = tmp_path / "M2.xlsx"
        rows_by_release = {"26A": [_row(release="26A")]}
        for out in (out1, out2):
            _write_master_workbook(
                out, rows_by_release=rows_by_release, issues=[], drift=[],
                release_old=None, release_new="26A",
            )
        wb1 = load_workbook(out1)
        wb2 = load_workbook(out2)
        assert wb1.sheetnames == wb2.sheetnames
        for sn in wb1.sheetnames:
            r1 = [[c.value for c in row] for row in wb1[sn].iter_rows()]
            r2 = [[c.value for c in row] for row in wb2[sn].iter_rows()]
            assert r1 == r2, f"Tab {sn} content differs"

    def test_preserves_existing_release_tabs(self, tmp_path):
        """Writing release X shouldn't wipe release Y if Y was already present in the file."""
        out = tmp_path / "Master.xlsx"
        # First run: writes 26A
        _write_master_workbook(
            out, rows_by_release={"26A": [_row(release="26A")]},
            issues=[], drift=[],
            release_old=None, release_new="26A",
        )
        # Second run: writes 26B but must preserve 26A
        # (caller has loaded 26A rows from existing workbook and passes both)
        _write_master_workbook(
            out, rows_by_release={
                "26A": [_row(release="26A")],
                "26B": [_row(release="26B")],
            },
            issues=[], drift=[],
            release_old="26A", release_new="26B",
        )
        wb = load_workbook(out)
        assert "26A" in wb.sheetnames
        assert "26B" in wb.sheetnames
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_catalog.py::TestWriteMasterWorkbook -v`
Expected: FAIL — `ImportError: cannot import name '_write_master_workbook'`

- [ ] **Step 3: Implement `_write_master_workbook`**

Append to `fbdi/catalog.py`:

```python
from openpyxl import Workbook
from openpyxl.styles import Font

_RELEASE_TAB_HEADERS = [
    "release", "file_name", "tab_name", "position",
    "column_label", "column_technical",
    "data_type", "length", "scale", "data_type_raw",
    "required",
]

_ISSUES_TAB_HEADERS = ["release", "file", "tab", "issue_type", "detail"]


def _drift_tab_headers(release_old: str | None, release_new: str | None) -> list[str]:
    """Build Drift tab headers with release names substituted."""
    old = release_old or "OLD"
    new = release_new or "NEW"
    return [
        "file", "tab", "position",
        f"col_label_{old}", f"col_label_{new}",
        f"col_technical_{old}", f"col_technical_{new}",
        f"data_type_{old}", f"data_type_{new}",
        f"length_{old}", f"length_{new}",
        f"required_{old}", f"required_{new}",
        "change_type",
    ]


def _write_master_workbook(
    output_path: Path,
    rows_by_release: dict[str, list[CatalogRow]],
    issues: list[IssueRow],
    drift: list[DriftRow],
    release_old: str | None,
    release_new: str | None,
) -> None:
    """Write the master workbook. Release tabs, Issues, Drift.

    The caller is responsible for providing *all* release data (merged
    from any existing workbook + fresh run). This function writes
    atomically via .tmp + rename.
    """
    wb = Workbook()
    # Remove the default empty sheet
    wb.remove(wb.active)

    bold = Font(name="Calibri", size=11, bold=True)
    plain = Font(name="Calibri", size=11)

    # Release tabs, in lexicographic order (26A < 26B < 26C < 27A)
    for release in sorted(rows_by_release.keys()):
        rows = rows_by_release[release]
        ws = wb.create_sheet(title=release)
        for col_idx, h in enumerate(_RELEASE_TAB_HEADERS, start=1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.font = bold
        for row_idx, r in enumerate(rows, start=2):
            ws.cell(row=row_idx, column=1, value=r.release).font = plain
            ws.cell(row=row_idx, column=2, value=r.file_name).font = plain
            ws.cell(row=row_idx, column=3, value=r.tab_name).font = plain
            ws.cell(row=row_idx, column=4, value=r.position).font = plain
            ws.cell(row=row_idx, column=5, value=r.column_label).font = plain
            ws.cell(row=row_idx, column=6, value=r.column_technical).font = plain
            ws.cell(row=row_idx, column=7, value=r.data_type).font = plain
            ws.cell(row=row_idx, column=8, value=r.length).font = plain
            ws.cell(row=row_idx, column=9, value=r.scale).font = plain
            ws.cell(row=row_idx, column=10, value=r.data_type_raw).font = plain
            ws.cell(
                row=row_idx, column=11,
                value="" if r.required is None else ("TRUE" if r.required else "FALSE"),
            ).font = plain
        ws.auto_filter.ref = f"A1:K{max(len(rows) + 1, 1)}"

    # Issues tab
    ws = wb.create_sheet(title="Issues")
    for col_idx, h in enumerate(_ISSUES_TAB_HEADERS, start=1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = bold
    for row_idx, i in enumerate(issues, start=2):
        ws.cell(row=row_idx, column=1, value=i.release).font = plain
        ws.cell(row=row_idx, column=2, value=i.file).font = plain
        ws.cell(row=row_idx, column=3, value=i.tab).font = plain
        ws.cell(row=row_idx, column=4, value=i.issue_type).font = plain
        ws.cell(row=row_idx, column=5, value=i.detail).font = plain
    ws.auto_filter.ref = f"A1:E{max(len(issues) + 1, 1)}"

    # Drift tab
    ws = wb.create_sheet(title="Drift")
    drift_headers = _drift_tab_headers(release_old, release_new)
    for col_idx, h in enumerate(drift_headers, start=1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = bold
    for row_idx, d in enumerate(drift, start=2):
        values = [
            d.file, d.tab, d.position,
            d.col_label_old, d.col_label_new,
            d.col_technical_old, d.col_technical_new,
            d.data_type_old, d.data_type_new,
            d.length_old, d.length_new,
            d.required_old, d.required_new,
            d.change_type,
        ]
        for col_idx, v in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=v).font = plain
    ws.auto_filter.ref = f"A1:N{max(len(drift) + 1, 1)}"

    # Atomic save
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    wb.save(tmp_path)
    wb.close()
    tmp_path.replace(output_path)
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_catalog.py -v`
Expected: PASS — all tests including TestWriteMasterWorkbook

- [ ] **Step 5: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "feat(catalog): add master workbook writer

Writes release tabs, Issues, and Drift tabs. Atomic save via .tmp
rename. Release tabs ordered lexicographically; Drift headers
substitute release names into column titles.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 10: `generate_catalog` — end-to-end orchestrator

**Files:**
- Modify: `fbdi/catalog.py`
- Test: `tests/test_catalog.py`

Stitches subprocess fan-out + existing-workbook merge + drift computation + workbook write. Idempotent per-release: re-running for release X rewrites tab X without touching other release tabs.

- [ ] **Step 1: Write failing tests**

Append to `tests/test_catalog.py`:

```python
from fbdi.catalog import generate_catalog


def _make_rich_xlsm(path: Path, tab_name: str, labels, types, techs, required):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title=tab_name)
    _make_rich_tab(
        ws, labels=labels, data_types=types,
        required_flags=required, technicals=techs,
    )
    wb.save(path)


class TestGenerateCatalog:
    def test_end_to_end_single_release(self, tmp_path):
        # Build a tiny release with one rich file
        release_dir = tmp_path / "baselines" / "TESTA" / "originals"
        release_dir.mkdir(parents=True)
        _make_rich_xlsm(
            release_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A", "Col B"],
            types=["VARCHAR2(50)", "NUMBER(18)"],
            techs=["COL_A", "COL_B"],
            required=["Required", "Optional"],
        )
        master = tmp_path / "Catalog.xlsx"

        generate_catalog(
            release="TESTA",
            baselines_dir=release_dir,
            master_path=master,
            timeout=60,
        )

        assert master.exists()
        wb = load_workbook(master)
        assert "TESTA" in wb.sheetnames
        assert "Issues" in wb.sheetnames
        assert "Drift" in wb.sheetnames
        # Data rows
        ws = wb["TESTA"]
        data = [[c.value for c in row] for row in ws.iter_rows(min_row=2)]
        assert len(data) == 2
        # No drift with single release
        drift_ws = wb["Drift"]
        drift_rows = [r for r in drift_ws.iter_rows(min_row=2)]
        assert drift_rows == []

    def test_end_to_end_two_releases_drift_classifications(self, tmp_path):
        # Release TESTA: 3 columns
        testa_dir = tmp_path / "baselines" / "TESTA" / "originals"
        testa_dir.mkdir(parents=True)
        _make_rich_xlsm(
            testa_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A", "Col B", "Col C"],
            types=["VARCHAR2(50)", "NUMBER(18)", "DATE"],
            techs=["COL_A", "COL_B", "COL_C"],
            required=["Required", "Optional", "Optional"],
        )

        # Release TESTB: different cols — Col A renamed technical, B length grown,
        # C required flipped, D added
        testb_dir = tmp_path / "baselines" / "TESTB" / "originals"
        testb_dir.mkdir(parents=True)
        _make_rich_xlsm(
            testb_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A", "Col B", "Col C", "Col D"],
            types=["VARCHAR2(50)", "NUMBER(32)", "DATE", "VARCHAR2(10)"],
            techs=["COL_A_RENAMED", "COL_B", "COL_C", "COL_D"],
            required=["Required", "Optional", "Required", "Optional"],
        )

        master = tmp_path / "Catalog.xlsx"
        generate_catalog(
            release="TESTA", baselines_dir=testa_dir,
            master_path=master, timeout=60,
        )
        generate_catalog(
            release="TESTB", baselines_dir=testb_dir,
            master_path=master, timeout=60,
        )

        wb = load_workbook(master)
        # Both release tabs present
        assert "TESTA" in wb.sheetnames
        assert "TESTB" in wb.sheetnames
        # Drift captures changes
        drift_ws = wb["Drift"]
        drift = [[c.value for c in row] for row in drift_ws.iter_rows(min_row=2)]
        change_types = {r[-1] for r in drift}
        assert "RENAMED" in change_types      # Col A technical
        assert "LENGTH_CHANGED" in change_types  # Col B length
        assert "REQUIRED_CHANGED" in change_types  # Col C required
        assert "ADDED" in change_types         # Col D

    def test_end_to_end_file_error_in_issues(self, tmp_path):
        release_dir = tmp_path / "baselines" / "TESTA" / "originals"
        release_dir.mkdir(parents=True)
        # Write a broken file
        (release_dir / "Broken.xlsm").write_bytes(b"not a real xlsx file")

        master = tmp_path / "Catalog.xlsx"
        generate_catalog(
            release="TESTA", baselines_dir=release_dir,
            master_path=master, timeout=60,
        )

        wb = load_workbook(master)
        ws = wb["Issues"]
        issue_rows = [[c.value for c in row] for row in ws.iter_rows(min_row=2)]
        assert any(r[3] == "FILE_ERROR" and r[1] == "Broken" for r in issue_rows)

    def test_end_to_end_idempotent(self, tmp_path):
        release_dir = tmp_path / "baselines" / "TESTA" / "originals"
        release_dir.mkdir(parents=True)
        _make_rich_xlsm(
            release_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A"],
            types=["VARCHAR2(50)"],
            techs=["COL_A"],
            required=["Required"],
        )
        master = tmp_path / "Catalog.xlsx"

        generate_catalog(release="TESTA", baselines_dir=release_dir,
                         master_path=master, timeout=60)
        wb1 = load_workbook(master)
        snap1 = {sn: [[c.value for c in row] for row in wb1[sn].iter_rows()]
                 for sn in wb1.sheetnames}

        generate_catalog(release="TESTA", baselines_dir=release_dir,
                         master_path=master, timeout=60)
        wb2 = load_workbook(master)
        snap2 = {sn: [[c.value for c in row] for row in wb2[sn].iter_rows()]
                 for sn in wb2.sheetnames}

        assert snap1 == snap2
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_catalog.py::TestGenerateCatalog -v`
Expected: FAIL — `ImportError: cannot import name 'generate_catalog'`

- [ ] **Step 3: Implement `generate_catalog`**

Append to `fbdi/catalog.py`:

```python
def _load_existing_release_rows(
    master_path: Path, current_release: str
) -> dict[str, list[CatalogRow]]:
    """Read rows for every release tab in master_path except `current_release`.

    Returns a dict mapping release name to reconstructed CatalogRows. If
    the file doesn't exist or has no release tabs, returns {}.
    """
    if not master_path.exists():
        return {}
    try:
        wb = load_workbook(master_path, read_only=True, data_only=True)
    except Exception as e:
        logger.warning(
            "Could not load existing master at %s: %s — starting fresh",
            master_path, e,
        )
        return {}

    result: dict[str, list[CatalogRow]] = {}
    try:
        for sn in wb.sheetnames:
            # Skip non-release tabs and the current release (will be rewritten)
            if sn in {"Issues", "Drift"} or sn == current_release:
                continue
            rows = _read_release_tab_rows(wb[sn])
            if rows:
                result[sn] = rows
    finally:
        wb.close()
    return result


def _read_release_tab_rows(ws) -> list[CatalogRow]:
    """Reconstruct CatalogRows from an existing release tab in the master workbook."""
    rows: list[CatalogRow] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v is not None for v in row):
            continue
        # Pad if short
        padded = list(row) + [None] * (11 - len(row))
        (release_v, file_v, tab_v, pos_v, label_v, tech_v,
         dtype_v, length_v, scale_v, raw_v, required_v) = padded[:11]
        required: bool | None
        if required_v in (None, ""):
            required = None
        elif str(required_v).upper() == "TRUE":
            required = True
        else:
            required = False
        rows.append(CatalogRow(
            release=str(release_v) if release_v else "",
            file_name=str(file_v) if file_v else "",
            tab_name=str(tab_v) if tab_v else "",
            position=int(pos_v) if pos_v is not None else 0,
            column_label=str(label_v) if label_v else "",
            column_technical=str(tech_v) if tech_v else "",
            data_type=str(dtype_v) if dtype_v else "",
            length=int(length_v) if isinstance(length_v, (int, float)) else None,
            scale=int(scale_v) if isinstance(scale_v, (int, float)) else None,
            data_type_raw=str(raw_v) if raw_v else "",
            required=required,
        ))
    return rows


def _load_existing_issues_excluding(
    master_path: Path, current_release: str
) -> list[IssueRow]:
    """Read Issues tab, excluding rows for current_release. Used for preservation."""
    if not master_path.exists():
        return []
    try:
        wb = load_workbook(master_path, read_only=True, data_only=True)
    except Exception:
        return []
    try:
        if "Issues" not in wb.sheetnames:
            return []
        ws = wb["Issues"]
        out: list[IssueRow] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(v is not None for v in row):
                continue
            padded = list(row) + [None] * (5 - len(row))
            release_v, file_v, tab_v, itype_v, detail_v = padded[:5]
            if release_v == current_release:
                continue
            out.append(IssueRow(
                release=str(release_v) if release_v else "",
                file=str(file_v) if file_v else "",
                tab=str(tab_v) if tab_v else "",
                issue_type=str(itype_v) if itype_v else "",
                detail=str(detail_v) if detail_v else "",
            ))
        return out
    finally:
        wb.close()


def generate_catalog(
    release: str,
    baselines_dir: Path,
    master_path: Path,
    timeout: int = CATALOG_TIMEOUT,
) -> None:
    """Generate / update the master catalog for one release.

    Reads every .xlsm in baselines_dir via subprocess-per-file isolation.
    Merges results with existing release tabs in master_path, regenerates
    the Issues tab (preserves issues from other releases, replaces this
    release's), recomputes Drift between the two most-recent releases,
    and writes the master workbook atomically.
    """
    baselines_dir = Path(baselines_dir)
    master_path = Path(master_path)

    # Fresh rows for this release
    new_rows: list[CatalogRow] = []
    new_issues: list[IssueRow] = []
    xlsm_files = sorted(baselines_dir.glob("*.xlsm"))
    for i, path in enumerate(xlsm_files, 1):
        logger.info("[%d/%d] Cataloging: %s", i, len(xlsm_files), path.stem)
        rows, issues = _run_file_in_subprocess(path, release=release, timeout=timeout)
        new_rows.extend(rows)
        new_issues.extend(issues)

    # Merge with existing releases (preserves other release tabs)
    rows_by_release = _load_existing_release_rows(master_path, current_release=release)
    rows_by_release[release] = new_rows

    # Merge issues: keep all non-current-release issues, add new ones
    preserved_issues = _load_existing_issues_excluding(master_path, current_release=release)
    all_issues = preserved_issues + new_issues

    # Compute drift between the two most-recent releases lexicographically
    sorted_releases = sorted(rows_by_release.keys())
    if len(sorted_releases) >= 2:
        release_old = sorted_releases[-2]
        release_new = sorted_releases[-1]
        drift = _compute_drift(
            rows_by_release[release_old],
            rows_by_release[release_new],
            release_old=release_old,
            release_new=release_new,
        )
    else:
        release_old = None
        release_new = sorted_releases[0] if sorted_releases else None
        drift = []

    _write_master_workbook(
        master_path,
        rows_by_release=rows_by_release,
        issues=all_issues,
        drift=drift,
        release_old=release_old,
        release_new=release_new,
    )
    logger.info(
        "Catalog written: %s (%d releases, %d rows, %d issues, %d drift)",
        master_path, len(rows_by_release),
        sum(len(v) for v in rows_by_release.values()),
        len(all_issues), len(drift),
    )
```

- [ ] **Step 4: Run all catalog tests**

Run: `python -m pytest tests/test_catalog.py -v`
Expected: PASS — all tests

- [ ] **Step 5: Run full test suite — nothing else should break**

Run: `python -m pytest tests/ -v`
Expected: All 54 existing tests still pass, plus the new catalog tests.

- [ ] **Step 6: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "feat(catalog): add generate_catalog end-to-end orchestrator

Ties subprocess fan-out to existing-workbook merge and drift
computation. Per-release regeneration preserves other release tabs
and cross-release issues. Drift is always computed between the two
most-recent releases lexicographically.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 11: CLI subcommand

**Files:**
- Modify: `fbdi/cli.py`
- Test: `tests/test_cli.py`

Add `python -m fbdi catalog --release 26B [--master PATH]` dispatching to `generate_catalog`.

- [ ] **Step 1: Inspect existing CLI tests for style**

Run: `python -m pytest tests/test_cli.py -v`
Expected: PASS — confirms existing CLI tests still work.

- [ ] **Step 2: Modify `fbdi/cli.py` — add catalog subparser**

In `fbdi/cli.py`, inside `main()`, after the `diagnose_parser` definition block (around line 82) and before `args = parser.parse_args(argv)`, add:

```python
    catalog_parser = subparsers.add_parser(
        "catalog",
        help="Generate or update the FBDI master catalog for a release",
    )
    catalog_parser.add_argument(
        "--release", required=True, type=str,
        help="Release label (e.g. 26B) — looks in baselines/<release>/originals",
    )
    catalog_parser.add_argument(
        "--baselines-dir", type=Path, default=None,
        help="Explicit path to release originals dir (overrides --release resolution)",
    )
    catalog_parser.add_argument(
        "--master", type=Path, default=Path("FBDI_Master_Catalog.xlsx"),
        help="Output master workbook path (default: FBDI_Master_Catalog.xlsx)",
    )
    catalog_parser.add_argument(
        "--timeout", type=int, default=120,
        help="Per-file subprocess timeout in seconds (default: 120)",
    )
    catalog_parser.add_argument(
        "--verbose", action="store_true",
        help="Set logging to DEBUG",
    )
```

In the `if args.command == "compare":` chain below, add:

```python
    elif args.command == "catalog":
        _run_catalog(args)
```

At the bottom of the file, add:

```python
def _run_catalog(args: argparse.Namespace) -> None:
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(name)s: %(message)s",
    )

    from fbdi.catalog import generate_catalog

    # Resolve baselines dir
    if args.baselines_dir:
        baselines_dir = args.baselines_dir
    else:
        candidate = Path("baselines") / args.release / "originals"
        baselines_dir = candidate

    if not baselines_dir.is_dir():
        print(f"Error: baselines directory not found: {baselines_dir}")
        sys.exit(1)

    xlsm_count = len(list(baselines_dir.glob("*.xlsm")))
    if xlsm_count == 0:
        print(f"Error: no .xlsm files found in {baselines_dir}")
        sys.exit(1)

    print(f"Cataloging release {args.release.upper()} from {baselines_dir}")
    print(f"  {xlsm_count} .xlsm files")
    print(f"  Output: {args.master}")
    print()

    generate_catalog(
        release=args.release.upper(),
        baselines_dir=baselines_dir,
        master_path=args.master,
        timeout=args.timeout,
    )

    # Summary from the written workbook
    from openpyxl import load_workbook as _lw
    wb = _lw(args.master, read_only=True)
    release_tabs = [sn for sn in wb.sheetnames if sn not in {"Issues", "Drift"}]
    issue_count = max(0, (wb["Issues"].max_row or 1) - 1)
    drift_count = max(0, (wb["Drift"].max_row or 1) - 1)
    wb.close()

    print(f"\nCatalog updated: {args.master}")
    print(f"  Release tabs: {', '.join(release_tabs)}")
    print(f"  Issues: {issue_count}")
    print(f"  Drift rows: {drift_count}")
```

- [ ] **Step 3: Write a CLI test**

Append to `tests/test_cli.py`:

```python
class TestCatalogCLI:
    def test_catalog_cli_requires_release(self, tmp_path, capsys):
        from fbdi.cli import main
        with pytest.raises(SystemExit):
            main(["catalog"])

    def test_catalog_cli_missing_baselines_errors(self, tmp_path, capsys):
        from fbdi.cli import main
        # Use --baselines-dir pointing at nonexistent path
        with pytest.raises(SystemExit):
            main([
                "catalog", "--release", "99Z",
                "--baselines-dir", str(tmp_path / "does-not-exist"),
                "--master", str(tmp_path / "M.xlsx"),
            ])
        captured = capsys.readouterr()
        assert "not found" in captured.out.lower()

    def test_catalog_cli_end_to_end(self, tmp_path):
        """Build a tiny release dir + run catalog CLI + verify file written."""
        from openpyxl import Workbook
        from fbdi.cli import main

        baselines = tmp_path / "baselines" / "TESTZ" / "originals"
        baselines.mkdir(parents=True)
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet("MY_TAB")
        # Thin tab
        ws.cell(row=4, column=1, value="*Only Field")
        wb.save(baselines / "Tpl.xlsm")

        master = tmp_path / "Catalog.xlsx"
        main([
            "catalog", "--release", "TESTZ",
            "--baselines-dir", str(baselines),
            "--master", str(master),
            "--timeout", "30",
        ])
        assert master.exists()
```

- [ ] **Step 4: Run CLI tests**

Run: `python -m pytest tests/test_cli.py -v`
Expected: PASS — including new TestCatalogCLI

- [ ] **Step 5: Run full test suite**

Run: `python -m pytest tests/ -v`
Expected: All existing + new tests pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/cli.py tests/test_cli.py
git commit -m "feat(catalog): wire up python -m fbdi catalog subcommand

Adds 'catalog' to the CLI. Resolves --release to
baselines/<release>/originals by default; --baselines-dir overrides.
--master defaults to FBDI_Master_Catalog.xlsx at the working-dir root.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 12: End-to-end verification on real 26A / 26B

**Files:**
- (no code changes; run verification only)

- [ ] **Step 1: Generate catalog for 26A**

Run:

```bash
python -m fbdi catalog --release 26A
```

Expected: run completes (may take several minutes due to 211 files). `FBDI_Master_Catalog.xlsx` now exists with tabs `26A`, `Issues`, `Drift` (Drift will be empty).

- [ ] **Step 2: Generate catalog for 26B**

Run:

```bash
python -m fbdi catalog --release 26B
```

Expected: run completes. `FBDI_Master_Catalog.xlsx` now has `26A`, `26B`, `Issues`, `Drift` tabs.

- [ ] **Step 3: Spot-check output**

Run:

```bash
python -c "
from openpyxl import load_workbook
wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True)
print('Tabs:', wb.sheetnames)
for sn in wb.sheetnames:
    ws = wb[sn]
    print(f'  {sn}: {(ws.max_row or 1) - 1} rows')
wb.close()
"
```

Expected: four tabs (`26A`, `26B`, `Issues`, `Drift`), with thousands of rows in each release tab and at least some rows in `Issues` (FILE_ERROR for the 6 known bad files) and `Drift` (the changes seen in `Comparison_Report_26A_26B.xlsx` plus any type/length drift).

- [ ] **Step 4: Cross-check Drift against existing Comparison_Report**

Run:

```bash
python -c "
from openpyxl import load_workbook

# Count YES rows in Comparison_Report_26A_26B
cr_wb = load_workbook('Comparison_Report_26A_26B.xlsx', read_only=True)
cr_ws = cr_wb.active
cr_yes = sum(1 for row in cr_ws.iter_rows(min_row=2, values_only=True) if row[-1] == 'YES')
cr_wb.close()

# Count drift rows in master catalog
mc_wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True)
drift_count = (mc_wb['Drift'].max_row or 1) - 1
mc_wb.close()

print(f'Comparison_Report YES rows: {cr_yes}')
print(f'Drift tab rows: {drift_count}')
print(f'Drift >= ComparisonReport YES: {drift_count >= cr_yes}')
"
```

Expected: Drift tab rows ≥ Comparison_Report YES count. The Drift tab is a strict superset — every name change surfaced in the existing report should appear as a `RENAMED` or `ADDED` or `REMOVED` or `MULTI` row in Drift, plus additional rows for pure type/length/required changes the old engine couldn't see.

- [ ] **Step 5: Spot-check a known rich tab and a known thin tab**

Run:

```bash
python -c "
from openpyxl import load_workbook
wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True)
ws = wb['26B']

# Print a few rows from AttachmentsImportTemplate (rich) and BudgetImportTemplate (thin)
def show(file_stem, tab_name, limit=5):
    print(f'\n--- {file_stem} / {tab_name} (first {limit}) ---')
    count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] == file_stem and row[2] == tab_name:
            print(f'  pos {row[3]}: label={row[4]!r} tech={row[5]!r} type={row[6]!r} length={row[7]} required={row[10]}')
            count += 1
            if count >= limit:
                break

show('AttachmentsImportTemplate', 'Attachment Details')
show('BudgetImportTemplate', 'XCC_BUDGET_INTERFACE')
wb.close()
"
```

Expected: `AttachmentsImportTemplate / Attachment Details` rows show populated `tech`, `type`, `length`, `required`. `BudgetImportTemplate / XCC_BUDGET_INTERFACE` rows show normalized label, blank `tech/type/length`, `required=TRUE` for asterisk-prefixed columns.

- [ ] **Step 6: If anything looks wrong, diagnose and fix before committing**

If rich-tab rows have blank `column_label` or `data_type`, likely cause: col-A keyword in 26A/26B differs from what's hardcoded. Add the observed keyword to `_COL_A_ROLE_KEYWORDS` and add a regression test. Re-run.

- [ ] **Step 7: Add the generated file to .gitignore (if not already)**

Check: `cat .gitignore | grep -i 'catalog'`. If `FBDI_Master_Catalog.xlsx` isn't covered, add it:

```
FBDI_Master_Catalog.xlsx
```

- [ ] **Step 8: Commit (gitignore only — the xlsx is gitignored output)**

```bash
git add .gitignore
git commit -m "chore: gitignore FBDI_Master_Catalog.xlsx

Consistent with Comparison_Report and Diagnostic_Report handling —
the master catalog is a build output, not source.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

(Skip this commit if `.gitignore` already covers the file — check before committing.)

---

## Task 13: Documentation updates

**Files:**
- Modify: `CLAUDE.md`
- Modify: `NEXT_STEPS.md`

- [ ] **Step 1: Update `CLAUDE.md` — add catalog to the Quick Start section**

In `CLAUDE.md`, find the `## Quick Start` section and add after the `diagnose` command example (around line 18):

```bash
# Generate/update the FBDI master catalog for a release
python -m fbdi catalog --release 26B
```

- [ ] **Step 2: Update `CLAUDE.md` — add `fbdi/catalog.py` et al. to the "Active Pipeline" list**

In the bullet list under `## Active Pipeline`, before `cli.py / __main__.py`, add:

```
- `catalog.py` — generates `FBDI_Master_Catalog.xlsx` with per-release snapshots (file × tab × position × label × technical × type × length × scale × required) + `Issues` + `Drift` tabs. Subprocess-isolated like `compare.py`.
- `type_parser.py` — parses Oracle data-type strings (`VARCHAR2(N CHAR)`, `NUMBER(p,s)`, `DATE`) into structured fields. Emits `TYPE_PARSE_WARNING` issues for unrecognized forms.
- `catalog_normalize.py` — normalizes FBDI labels (strips non-alphanumeric/underscore/whitespace) for Applaud MDB compatibility.
```

And update the "Output" line to mention both artifacts:

```
- **Outputs:**
  - `Comparison_Report_<OLD>_<NEW>.xlsx` — 7-column diff for VBA validation (unchanged)
  - `FBDI_Master_Catalog.xlsx` — per-release snapshots + Issues + Drift tabs
```

- [ ] **Step 3: Update `NEXT_STEPS.md` — mark the catalog work complete and note follow-ups**

In `NEXT_STEPS.md`, add a new numbered section at the top (or near the "Current Frontier" reference), with wording consistent with the Phase 2/3 resolution pattern:

```markdown
## N. ~~Build the FBDI Master Catalog~~ — RESOLVED

**Resolution (2026-04-15):** `python -m fbdi catalog --release <label>` generates
`FBDI_Master_Catalog.xlsx` with per-release snapshot tabs (file × tab × position ×
label × technical × type × length × scale × required), plus `Issues` and `Drift`
tabs. Rich tabs (~349 tabs in 26B) yield full metadata; thin tabs (~303) yield
label-only rows with blanks honestly flagged. Subprocess-isolated per-file with
120s timeout.

**Follow-ups still open:**
- 6 FILE_ERROR templates remain unreadable (corrupt stylesheets). Could be addressed
  by unzipping, dropping styles.xml, re-zipping. Separate workstream.
- CSV companion exports alongside `FBDI_Master_Catalog.xlsx` once applaud-mcp's
  preferred ingestion format is known.
- Shared `template_reader.py` refactor if duplication between `compare.py` and
  `catalog.py` becomes painful.
```

(Keep the existing sections in `NEXT_STEPS.md` as-is; this is an addition.)

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md NEXT_STEPS.md
git commit -m "docs: add FBDI Master Catalog to CLAUDE.md and NEXT_STEPS.md

CLAUDE.md gets the new CLI in Quick Start and new modules in the
Active Pipeline list. NEXT_STEPS.md gets a resolution section for the
catalog work and notes three still-open follow-ups (FILE_ERROR repair,
CSV companion, shared reader refactor).

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 14: Rebuild graphify + final push

- [ ] **Step 1: Verify working tree is clean and all tests pass**

Run:

```bash
git status && python -m pytest tests/ -v
```

Expected: `working tree clean`, all tests pass (54 existing + ~30 new catalog tests = ~84).

- [ ] **Step 2: Rebuild the graphify knowledge graph**

Per `CLAUDE.md`: "After modifying code files in this session, run … to keep the graph current".

Run:

```bash
python3 -c "from graphify.watch import _rebuild_code; from pathlib import Path; _rebuild_code(Path('.'))"
```

Expected: graph rebuilds without errors. New nodes should appear for `fbdi/catalog.py`, `fbdi/type_parser.py`, `fbdi/catalog_normalize.py`, and the new test files.

- [ ] **Step 3: Commit the regenerated graph outputs**

Run:

```bash
git add graphify-out/
git commit -m "chore: rebuild graphify knowledge graph after catalog work

Regenerates graphify-out/ to reflect fbdi.catalog, fbdi.type_parser,
and fbdi.catalog_normalize modules and their tests.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 4: Push to master (per Brad's handoff workflow — direct push, not PR)**

Run:

```bash
git push origin master
```

Expected: push succeeds, all catalog commits now on `origin/master`.

---

## Self-review against spec

**Spec coverage:**
- ✅ Master workbook layout (release tabs, Issues, Drift) — Tasks 9, 10
- ✅ Per-release tab schema (11 columns) — Task 9 (`_RELEASE_TAB_HEADERS`)
- ✅ Label normalization with Unicode-aware alphanumerics — Task 1
- ✅ Type parsing with parsed + raw fields — Task 2
- ✅ Tier 1/Tier 2 dispatch via `_is_tier1_header` — Task 4
- ✅ Metadata row detection via col-A keyword match — Task 5 (`_find_metadata_rows`)
- ✅ BOM stripping, case-insensitive col-A match — Task 5 tests
- ✅ Required flag universal — Task 4 (thin) + Task 5 (rich)
- ✅ Subprocess isolation with 120s timeout — Task 7
- ✅ Idempotent regeneration — Task 10 (`test_end_to_end_idempotent`)
- ✅ Preservation of other release tabs — Task 10 (`_load_existing_release_rows`)
- ✅ `change_type` classification with first-match ordering — Task 8
- ✅ Issues taxonomy (FILE_ERROR, TIMEOUT, SUBPROCESS_FAILED, NO_HEADER, TYPE_PARSE_WARNING) — Tasks 4, 5, 6, 7
- ✅ Atomic save via .tmp rename — Task 9
- ✅ CLI subcommand — Task 11
- ✅ Real-world smoke test — Task 12
- ✅ Documentation updates — Task 13
- ✅ Non-goal boundaries respected (no FILE_ERROR repair, no description column, no CSV exports) — explicit in plan

**Placeholder scan:** No "TBD" / "implement later" / "fill in". Every step has complete code or exact commands. One intentional `NotImplementedError` in Task 4 (rich-tab stub) is replaced with real code in Task 5 — documented explicitly so reviewers don't flag it.

**Type consistency:**
- `CatalogRow`, `IssueRow`, `DriftRow` fields consistent across Tasks 3, 9, 10.
- `ParsedType` fields consistent across Tasks 2, 5.
- Function signatures: `extract_tab_rows(ws, file_stem, release)`, `extract_file(path, release)`, `generate_catalog(release, baselines_dir, master_path, timeout)` — all consistent across tasks.
- `_compute_drift(old_rows, new_rows, release_old, release_new)` — consistent between Tasks 8 and 10.
- `_write_master_workbook(output_path, rows_by_release, issues, drift, release_old, release_new)` — consistent between Tasks 9 and 10.
