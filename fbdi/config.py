"""Configuration constants for the FBDI comparison engine."""

# Maximum file size in bytes before skipping in diagnostics / mapping builds.
# These tools load workbooks in non-read_only mode (everything in memory), so
# the limit is a memory safeguard. The comparison engine uses streaming
# read-only + iter_rows and is not bounded by this limit.
MAX_FILE_SIZE_BYTES = 5 * 1024 * 1024  # 5 MB

# Minimum non-empty cells for a row to be a header candidate
MIN_CELLS = 2

# Tabs to skip during comparison (case-sensitive match against sheet names)
SKIP_TABS = {
    "Instructions and CSV Generation",
    "Instructions and DAT Generation",
    "Instructions and ZIP Generation",
    "Instructions",
    "Options",
    "Create CSV",
    "reference",
    "Validation Report",
    "LOV",
    "XDO_METADATA",
    "Lookups",
    "Messages",  # Oracle error code lookup table — not an import field definition tab
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

# Per-file timeout (seconds) for catalog subprocess workers.
# Mirrors COMPARE_TIMEOUT in compare.py; isolates openpyxl resource leaks.
CATALOG_TIMEOUT = 120

# Applaud system alias is forwarded to applaud-mcp (which resolves the .mdb path via its
# MDB_SYSTEMS env) — MCP is the single source of truth for alias->path, so no path table
# is duplicated here. The alias is only used to name the per-system snapshot file.
def applaud_snapshot_path(system: str):
    """Per-system snapshot location, e.g. baselines/applaud/applaud_snapshot_ORACLE_MASTER.json."""
    from pathlib import Path
    return Path("baselines") / "applaud" / f"applaud_snapshot_{system}.json"
