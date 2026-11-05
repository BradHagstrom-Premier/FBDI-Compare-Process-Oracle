"""Per-template post-clear patches for FBDI blanks.

Registry of deliberate, template-specific edits that ``fbdi.clear.clear_workbook``
applies after clearing, so every release refresh (``tools/download_and_clear.py``,
including ``--clear-only``) re-applies them. Patches edit the package at byte
level (see ``fbdi.ooxml``) and raise ``BlankPatchError`` when the original no
longer looks the way the patch expects — a blank must never ship silently
without its fix.

RapidImplementationForCashManagement.xlsm (Oracle FSM rapid-implementation
template, not a standard FBDI file):

* **A7 paste anchor** on "Bank Account". The Conversion Dashboard pastes at
  (last non-empty cell in column A) + 1. This sheet's column A is empty (title
  B2, headers B7, data from row 8), so data landed on row 2 over the title and
  headers, and the Generate macro (reads from row 8) dropped the first rows.
  Generate reads only columns B..X, so a marker in A7 is harmless.
* **VBA Save As fix.** ``CommandButton1_Click`` set ``.FilterIndex = 5 'for xml``;
  entry 5 is "CSV UTF-8" in current Excel, so the generated XML was saved as
  .csv. The fixed ``vbaProject.bin`` is committed under ``fbdi/assets/`` (built
  by ``tools/build_vba_patch.py`` through Excel) and swapped in only when the
  original's ``vbaProject.bin`` hash matches the one the fix was built from.
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

from fbdi.ooxml import CELL_RE, ROW_NUM_RE, ROW_RE, Package, tag_attrs

VBA_PART = "xl/vbaProject.bin"

RAPID_IMPL = "RapidImplementationForCashManagement.xlsm"
RAPID_IMPL_SHEET = "Bank Account"
RAPID_IMPL_HEADER_ROW = 7
ANCHOR_TEXT = "Dashboard paste anchor - do not delete"

# sha256 of the Oracle-shipped xl/vbaProject.bin the patched asset was built
# from (identical in 26A, 26B and 26C).
RAPID_IMPL_ORIGINAL_VBA_SHA256 = (
    "86bc0a730110b66c2b8a64bcc3488eacb051b5e6845fe3fd15bea32951e59137"
)
PATCHED_VBA_ASSET = (
    Path(__file__).resolve().parent / "assets" / "RapidImplementationForCashManagement.vbaProject.bin"
)


class BlankPatchError(RuntimeError):
    """The original no longer matches what a registered patch expects."""


def apply(filename: str, pkg: Package, header_rows: dict[str, int]) -> None:
    """Apply the registered patches for ``filename`` (no-op if none)."""
    for patch in PATCHES.get(filename, ()):
        patch(pkg, header_rows)


# ---------------------------------------------------------------------------
# Patch primitives
# ---------------------------------------------------------------------------

def insert_anchor(
    pkg: Package,
    header_rows: dict[str, int],
    sheet: str = RAPID_IMPL_SHEET,
    row: int = RAPID_IMPL_HEADER_ROW,
    text: str = ANCHOR_TEXT,
) -> None:
    """Write ``text`` into column A of the header row of ``sheet``.

    Fails if the detected header row is not ``row``: if Oracle moved the
    header, the anchor row has to be re-derived, not guessed.
    """
    detected = header_rows.get(sheet)
    if detected != row:
        raise BlankPatchError(
            f"{sheet!r}: expected header row {row} for the paste anchor, "
            f"detected {detected}. Re-check the template and update blank_patches."
        )
    part = next((s.part for s in pkg.worksheets() if s.name == sheet), None)
    if part is None:
        raise BlankPatchError(f"sheet {sheet!r} not found")
    ref = f"A{row}".encode()
    xml = pkg.read(part)

    target = None
    for m in ROW_RE.finditer(xml):
        num = ROW_NUM_RE.search(m.group(1))
        if num and int(num.group(1)) == row:
            target = m
            break
    if target is None:
        raise BlankPatchError(f"{sheet!r}: row {row} not found in {part}")

    content = target.group(2) or b""
    style = b""
    existing = None
    for cm in CELL_RE.finditer(content):
        if tag_attrs(b"<c" + cm.group(1) + b">").get("r") == ref.decode():
            existing = cm
            break
    s = None
    if existing is not None:
        s = tag_attrs(b"<c" + existing.group(1) + b">").get("s")
        content = content[:existing.start()] + content[existing.end():]
    if s is None:
        # A new cell takes the style Excel would give a typed-in value: the row's
        # (if customFormat) else the column's. Unstyled = cellXfs[0], which wraps in
        # the Oracle template -> the anchor wraps letter-by-letter in the 0.75-wide
        # gutter column and Excel auto-fits row 7 to 408pt.
        s = _new_cell_style(xml, target.group(1), column=1)
    style = b' s="%s"' % s.encode() if s is not None else b""

    cell = (
        b'<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>'
        % (ref, style, _xml_text(text))
    )
    row_attrs = re.sub(
        rb'\bspans="\d+:(\d+)"', rb'spans="1:\1"', target.group(1), count=1
    )
    new_row = b"<row" + row_attrs + b">" + cell + content + b"</row>"
    xml = xml[:target.start()] + new_row + xml[target.end():]
    xml = re.sub(rb'(<dimension\b[^>]*\bref=")[A-Z]+(\d+:)', rb"\1A\2", xml, count=1)
    pkg.write(part, xml)


def swap_vba_project(pkg: Package, expected_sha256: str, patched: bytes) -> None:
    """Replace xl/vbaProject.bin with ``patched`` iff the original hash matches."""
    actual = hashlib.sha256(pkg.read(VBA_PART)).hexdigest()
    if actual != expected_sha256:
        raise BlankPatchError(
            f"{VBA_PART} sha256 {actual} != {expected_sha256}: Oracle changed the "
            "macro, so the committed VBA fix no longer applies. Rebuild it with "
            "`py tools/build_vba_patch.py <original.xlsm>` and update the hash in "
            "fbdi/blank_patches.py."
        )
    pkg.write(VBA_PART, patched)


def _new_cell_style(sheet_xml: bytes, row_attrs: bytes, column: int) -> str | None:
    """Style index Excel applies to a value typed into an empty cell."""
    row = tag_attrs(b"<row" + row_attrs + b">")
    if row.get("customFormat") in ("1", "true") and "s" in row:
        return row["s"]
    for tag in re.findall(rb"<col\b[^>]*>", sheet_xml):
        col = tag_attrs(tag)
        if int(col.get("min", 0)) <= column <= int(col.get("max", 0)):
            return col.get("style")
    return None


def _xml_text(text: str) -> bytes:
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    ).encode("utf-8")


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

def _rapid_impl_anchor(pkg: Package, header_rows: dict[str, int]) -> None:
    insert_anchor(pkg, header_rows)


def _rapid_impl_vba(pkg: Package, header_rows: dict[str, int]) -> None:
    if not PATCHED_VBA_ASSET.exists():
        raise BlankPatchError(
            f"missing {PATCHED_VBA_ASSET}; build it with `py tools/build_vba_patch.py`"
        )
    swap_vba_project(pkg, RAPID_IMPL_ORIGINAL_VBA_SHA256, PATCHED_VBA_ASSET.read_bytes())


PATCHES: dict[str, tuple] = {
    RAPID_IMPL: (_rapid_impl_anchor, _rapid_impl_vba),
}
