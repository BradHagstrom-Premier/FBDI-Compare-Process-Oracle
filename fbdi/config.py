"""Configuration constants for the FBDI comparison engine."""

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
