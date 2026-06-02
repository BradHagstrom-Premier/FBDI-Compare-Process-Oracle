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

# Applaud system aliases -> .mdb path. Mirrors MDB_SYSTEMS in the applaud-mcp env;
# kept here so Step B can name-qualify snapshot files / output without the MCP up.
APPLAUD_SYSTEMS = {
    "ORACLE_MASTER": "C:/Users/10193/Definian/MDB_for_ApplaudMCP/ORACLE_MASTER/AP0STE.mdb",
    "AWC_MASTER":    "C:/Users/10193/Definian/MDB_for_ApplaudMCP/AWC_MASTER/AP0STE.mdb",
}
DEFAULT_APPLAUD_SYSTEM = "ORACLE_MASTER"


def applaud_snapshot_path(system: str):
    from pathlib import Path
    return Path("baselines") / "applaud" / f"applaud_snapshot_{system}.json"
