"""Minimal byte-level editing of an OOXML (.xlsx/.xlsm) package.

Why this exists: re-saving an Oracle FBDI template through openpyxl
re-serializes every part it understands and silently drops the rest — ActiveX
``<controls>`` (Oracle's "Generate CSV File" button), x14 data validations
(cross-sheet dropdowns), drawings, printer settings. ``Package`` instead copies
the original zip entry-for-entry and lets callers replace only the parts they
deliberately edit, so everything else ships byte-identical to what Oracle
published.
"""

from __future__ import annotations

import os
import posixpath
import re
import zipfile
from dataclasses import dataclass
from xml.sax.saxutils import unescape

WORKBOOK_PART = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
CALC_CHAIN = "xl/calcChain.xml"

_XML_ATTR_UNESCAPE = {"&quot;": '"', "&apos;": "'"}

_SHEET_TAG_RE = re.compile(rb"<sheet\b[^>]*>")
_REL_TAG_RE = re.compile(rb"<Relationship\b[^>]*>")
_ATTR_RE = re.compile(rb'([\w:]+)="([^"]*)"')


@dataclass(frozen=True)
class SheetRef:
    """One worksheet as listed in workbook.xml, resolved to its zip part."""

    name: str
    state: str  # "visible" | "hidden" | "veryHidden"
    part: str   # e.g. "xl/worksheets/sheet2.xml"


def tag_attrs(tag: bytes) -> dict[str, str]:
    """Attributes of a single XML start tag, XML entities unescaped."""
    return {
        k.decode("ascii"): unescape(v.decode("utf-8"), _XML_ATTR_UNESCAPE)
        for k, v in _ATTR_RE.findall(tag)
    }


class Package:
    """Read an OOXML zip, replace/delete selected parts, write a new zip.

    Parts that are never written are copied through unchanged on ``save``.
    """

    def __init__(self, path: str):
        self._zf = zipfile.ZipFile(path)
        self._infos = self._zf.infolist()
        self._changed: dict[str, bytes] = {}
        self._deleted: set[str] = set()

    def __enter__(self) -> "Package":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def close(self) -> None:
        self._zf.close()

    # -- part access -------------------------------------------------------

    def names(self) -> list[str]:
        return [i.filename for i in self._infos if i.filename not in self._deleted]

    def has(self, name: str) -> bool:
        return name in self.names()

    def read(self, name: str) -> bytes:
        if name in self._deleted:
            raise KeyError(name)
        if name in self._changed:
            return self._changed[name]
        return self._zf.read(name)

    def write(self, name: str, data: bytes) -> None:
        """Replace an existing part's bytes (adding new parts is not supported)."""
        if name not in self.names():
            raise KeyError(f"no such part to replace: {name}")
        self._changed[name] = data

    def delete(self, name: str) -> None:
        self._deleted.add(name)
        self._changed.pop(name, None)

    # -- workbook structure ------------------------------------------------

    def worksheets(self) -> list[SheetRef]:
        """Worksheets in workbook order (chartsheets/dialogsheets excluded)."""
        rels: dict[str, tuple[str, str]] = {}
        for tag in _REL_TAG_RE.findall(self.read(WORKBOOK_RELS)):
            a = tag_attrs(tag)
            rels[a["Id"]] = (a.get("Type", ""), _resolve_target(a["Target"]))
        sheets: list[SheetRef] = []
        for tag in _SHEET_TAG_RE.findall(self.read(WORKBOOK_PART)):
            a = tag_attrs(tag)
            rid = next(v for k, v in a.items() if k.endswith(":id"))
            rel_type, part = rels[rid]
            if not rel_type.endswith("/worksheet"):
                continue
            sheets.append(SheetRef(a["name"], a.get("state", "visible"), part))
        return sheets

    def drop_calc_chain(self) -> bool:
        """Remove xl/calcChain.xml plus its relationship and content-type override.

        Excel rebuilds the calc chain on open; a stale one that points at cells
        whose formulas were cleared triggers a "We found a problem" repair prompt.
        Returns True if a calc chain was present.
        """
        if not self.has(CALC_CHAIN):
            return False
        self.delete(CALC_CHAIN)
        rels = self.read(WORKBOOK_RELS)
        self.write(WORKBOOK_RELS, re.sub(rb'<Relationship\b[^>]*/calcChain"[^>]*/>', b"", rels))
        ct = self.read(CONTENT_TYPES)
        self.write(CONTENT_TYPES, re.sub(rb'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b"", ct))
        return True

    # -- output ------------------------------------------------------------

    def save(self, dst_path: str) -> None:
        """Write the package to dst_path (atomically via a temp file)."""
        tmp = dst_path + ".tmp"
        with zipfile.ZipFile(tmp, "w") as out:
            for info in self._infos:
                if info.filename in self._deleted:
                    continue
                data = self._changed.get(info.filename)
                if data is None:
                    data = self._zf.read(info)
                new = zipfile.ZipInfo(info.filename, info.date_time)
                new.compress_type = info.compress_type
                new.external_attr = info.external_attr
                new.create_system = info.create_system
                out.writestr(new, data)
        os.replace(tmp, dst_path)


def _resolve_target(target: str) -> str:
    """Workbook-relationship Target -> zip part name."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))
