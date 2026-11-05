"""Smart clearing for Oracle FBDI template workbooks.

Clears sample data from FBDI templates while preserving all header rows
and metadata. Uses detect_header_row() to dynamically find the header
boundary per sheet, then clears every cell value below it.

The blank is produced by **patching the original package in place**, not by
re-saving it through openpyxl: openpyxl is used read-only, only to detect
header rows. Each cleared sheet's XML has the cells below its header emptied
(style index kept); every other part — ActiveX controls (Oracle's Generate
button), x14 data validations, drawings, printer settings, VBA — is copied
byte-for-byte. Re-saving through openpyxl dropped all of those from every
26A–26C blank; don't reintroduce it.

Used by tools/download_and_clear.py to produce blank copies for client
Oracle projects. NOT used by the comparison engine (which reads originals).
"""

from __future__ import annotations

import logging
import os
import re

from openpyxl import load_workbook
from openpyxl.styles.fonts import Font

from fbdi import blank_patches
from fbdi.detect_header import detect_header_row
from fbdi.ooxml import Package, tag_attrs

logger = logging.getLogger(__name__)

# (filename, sheet name) pairs that are visible and header-detectable but are
# NOT load sheets, so clearing "below the header" would destroy content.
SKIP_CLEAR: frozenset[tuple[str, str]] = frozenset({
    # Yes/No feature-selection form; detection lands on its first option row.
    ("ProjectImportTemplate.xlsm", "Options"),
})

_SHEET_DATA_RE = re.compile(rb"<sheetData\b[^>]*?(?:/>|>(.*?)</sheetData>)", re.S)
_ROW_RE = re.compile(rb"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
_CELL_RE = re.compile(rb"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_ROW_NUM_RE = re.compile(rb'\br="(\d+)"')
_KEEP_CELL_ATTRS = (b"r", b"s")


def clear_workbook(src_path: str, dst_path: str) -> list[str]:
    """Clear sample data from an FBDI workbook, preserving all headers.

    For each **visible** data sheet (skipping sheet 1 / instructions, hidden
    LOV/reference sheets, and SKIP_CLEAR settings sheets):
      1. Run detect_header_row() to find the actual header row
      2. Empty every cell from header_row + 1 onward (style index kept)

    If header detection fails for a sheet, that sheet is left untouched.
    Per-template post-clear fixes (fbdi.blank_patches) are applied last.

    Args:
        src_path: Path to the original FBDI .xlsm file.
        dst_path: Path to write the cleared copy.

    Returns:
        List of sheet names where header detection failed (skipped).
    """
    filename = os.path.basename(src_path)
    skipped_sheets: list[str] = []
    header_rows: dict[str, int] = {}

    with Package(src_path) as pkg:
        sheets = pkg.worksheets()
        to_clear = [
            s for s in sheets[1:]
            if s.state == "visible" and (filename, s.name) not in SKIP_CLEAR
        ]

        if to_clear:
            # Oracle FBDI files use font family values >14 which openpyxl rejects.
            Font.family.max = 255
            wb = load_workbook(src_path, read_only=True)
            try:
                for sheet in to_clear:
                    header_row = detect_header_row(wb[sheet.name])
                    if header_row is None:
                        logger.warning(
                            "Header detection failed for sheet '%s' — skipping (data preserved)",
                            sheet.name,
                        )
                        skipped_sheets.append(sheet.name)
                    else:
                        header_rows[sheet.name] = header_row
            finally:
                wb.close()

        formula_cleared = False
        for sheet in to_clear:
            if sheet.name not in header_rows:
                continue
            xml = pkg.read(sheet.part)
            new_xml, had_formula = clear_rows_below(xml, header_rows[sheet.name])
            if new_xml != xml:
                pkg.write(sheet.part, new_xml)
            formula_cleared |= had_formula

        if formula_cleared:
            pkg.drop_calc_chain()

        blank_patches.apply(filename, pkg, header_rows)
        pkg.save(dst_path)

    return skipped_sheets


def clear_rows_below(sheet_xml: bytes, header_row: int) -> tuple[bytes, bool]:
    """Empty every cell in rows > header_row of one worksheet part.

    A cleared cell keeps only its ``r`` and ``s`` (style) attributes, so
    borders/fills/number formats survive; ``t``, ``<v>``, ``<f>`` and ``<is>``
    are dropped. Rows themselves (heights, styles) stay. Bytes outside
    ``<sheetData>`` are untouched.

    Returns (new_xml, whether any cleared cell carried a formula).
    """
    m = _SHEET_DATA_RE.search(sheet_xml)
    if m is None:
        raise ValueError("worksheet XML has no <sheetData> element")
    if m.group(1) is None:  # <sheetData/> — nothing to clear
        return sheet_xml, False

    had_formula = False
    current_row = 0

    def _clear_cell(cm: re.Match) -> bytes:
        nonlocal had_formula
        if cm.group(2) is None:  # already an empty, self-closing cell
            return cm.group(0)
        if b"<f" in cm.group(2):
            had_formula = True
        attrs = tag_attrs(b"<c" + cm.group(1) + b">")
        kept = b"".join(
            b' %s="%s"' % (k, attrs[k.decode()].encode("utf-8"))
            for k in _KEEP_CELL_ATTRS if k.decode() in attrs
        )
        return b"<c" + kept + b"/>"

    def _clear_row(rm: re.Match) -> bytes:
        nonlocal current_row
        num = _ROW_NUM_RE.search(rm.group(1))
        current_row = int(num.group(1)) if num else current_row + 1
        if rm.group(2) is None or current_row <= header_row:
            return rm.group(0)
        cleared = _CELL_RE.sub(_clear_cell, rm.group(2))
        return b"<row" + rm.group(1) + b">" + cleared + b"</row>"

    body = _ROW_RE.sub(_clear_row, m.group(1))
    start, end = m.span(1)
    return sheet_xml[:start] + body + sheet_xml[end:], had_formula
