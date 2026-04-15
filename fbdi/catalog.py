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

import logging
import re
from dataclasses import dataclass

from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from fbdi.catalog_normalize import normalize_label
from fbdi.detect_header import UPPER_SNAKE_PATTERN, detect_header_row
from fbdi.type_parser import parse_data_type

logger = logging.getLogger(__name__)

_MAX_COL = 500


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
