"""Per-template post-clear patches for FBDI blanks.

Registry of deliberate, template-specific edits applied by
``fbdi.clear.clear_workbook`` after clearing, so every release refresh
(``tools/download_and_clear.py``, including ``--clear-only``) re-applies them.
"""

from __future__ import annotations

from fbdi.ooxml import Package


def apply(filename: str, pkg: Package, header_rows: dict[str, int]) -> None:
    """Apply the registered patches for ``filename`` (no-op if none)."""
    for patch in PATCHES.get(filename, ()):
        patch(pkg, header_rows)


PATCHES: dict[str, tuple] = {}
