"""Configuration constants for the FBDI comparison engine."""

# Maximum file size in bytes before skipping (files larger than this are excluded
# from automated comparison and must be reviewed manually)
MAX_FILE_SIZE_BYTES = 5 * 1024 * 1024  # 5 MB

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
