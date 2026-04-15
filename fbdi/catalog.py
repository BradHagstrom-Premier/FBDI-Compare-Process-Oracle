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
