"""Normalize user-facing FBDI column labels for the master catalog.

Produces the catalog's ``column_label`` value: strips punctuation and
symbols (asterisks, etc.) while preserving alphanumerics, underscores, and
whitespace, then collapses whitespace runs. Applied only to labels —
technical UPPER_SNAKE_CASE names are untouched because they are already
canonical by construction.
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
