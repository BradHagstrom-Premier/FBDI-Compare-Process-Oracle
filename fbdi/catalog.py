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
import multiprocessing
import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from fbdi.catalog_normalize import normalize_label
from fbdi.config import CATALOG_TIMEOUT, SKIP_TABS
from fbdi.detect_header import UPPER_SNAKE_PATTERN, detect_header_row
from fbdi.type_parser import parse_data_type

logger = logging.getLogger(__name__)

_MAX_COL = 500

# Column-A keyword -> metadata role. Case-insensitive match after BOM strip.
# Note: header row (tier 1) has col A starting with "Column name of the Table";
# that's already handled by detect_header_row, we don't re-discover it here.
_COL_A_ROLE_KEYWORDS = {
    "name": "label",
    "data type": "type",
    "required or optional": "required",
}


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
    path: Path, release: str, timeout: int = CATALOG_TIMEOUT
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
