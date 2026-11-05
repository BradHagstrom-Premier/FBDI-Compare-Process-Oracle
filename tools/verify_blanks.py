"""Verify a release's blanks against their originals (read-only).

Two checks per original/blank pair in baselines/<VER>/:

1. **Package survey** — the blank must be the original with only the cleared
   sheets' ``<sheetData>`` changed:
   * no package part added; none removed except ``xl/calcChain.xml`` (and only
     when a cleared cell held a formula);
   * per sheet, no loss of controls (ActiveX/form), data validations (incl. x14),
     drawings, legacy drawings or conditional formats;
   * hidden sheets (LOVs behind dropdowns), the instructions sheet (index 0)
     and ``fbdi.clear.SKIP_CLEAR`` sheets byte-identical;
   * on cleared sheets, bytes outside ``<sheetData>`` unchanged;
   * ``[Content_Types].xml`` / workbook rels differ only by the calcChain entry;
   * files with registered ``fbdi.blank_patches`` may also differ in
     ``xl/vbaProject.bin`` and a sheet's ``<dimension ref>`` (the paste anchor).
2. **Dashboard paste anchor** — on every visible data sheet (not the
   instructions sheet, not ``SKIP_CLEAR`` settings forms), the last non-empty
   column-A row of the blank equals the header row detected on the original.
   The Conversion Dashboard pastes at (last non-empty col-A row) + 1, so any
   other value means data would land on the header or leave a gap.

Missing blanks fail too. Header-detection failures are reported, not failed
(``fbdi.clear`` leaves those sheets untouched).

Usage:
    py tools/verify_blanks.py 26D            # summary; exit 1 on any violation
    py tools/verify_blanks.py 26D --json     # full report

Exit codes: 0 = clean, 1 = violations / mismatches / missing blanks.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import zipfile
from html import unescape
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from openpyxl import load_workbook  # noqa: E402
from openpyxl.styles.fonts import Font  # noqa: E402

from fbdi import blank_patches  # noqa: E402
from fbdi.clear import SKIP_CLEAR  # noqa: E402
from fbdi.detect_header import detect_header_row  # noqa: E402
from fbdi.ooxml import CALC_CHAIN, CELL_RE, SHEET_DATA_RE, Package, tag_attrs  # noqa: E402
from fbdi.utils import quiet_openpyxl  # noqa: E402

# Per-sheet features a blank must never lose (counted in the sheet XML).
TOKENS = {
    "control": rb"<control\b",
    "dataValidation": rb"<dataValidation\b",
    "x14_dataValidation": rb"<x14:dataValidation\b",
    "drawing": rb"<drawing\b",
    "legacyDrawing": rb"<legacyDrawing\b",
    "conditionalFormatting": rb"<conditionalFormatting\b",
    "x14_conditionalFormatting": rb"<x14:conditionalFormatting\b",
}
_CALC_REL = rb'<Relationship\b[^>]*/calcChain"[^>]*/>'
_CALC_CT = rb'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>'
_DIMENSION = rb"<dimension\b[^>]*/>"
_FORMULA = rb"<f\b"
_SI_RE = re.compile(rb"<si\b[^>]*>(.*?)</si>", re.S)
_T_RE = re.compile(rb"<t\b[^>]*>(.*?)</t>", re.S)
_V_RE = re.compile(rb"<v>(.*?)</v>", re.S)
_COL_A_REF = re.compile(r"A(\d+)")


def _outside_sheetdata(xml: bytes) -> bytes:
    m = SHEET_DATA_RE.search(xml)
    return xml[: m.start()] + xml[m.end():] if m else xml


def survey_pair(orig: Path, blank: Path) -> list[str]:
    """Violations (human-readable) for one original/blank pair; [] if clean."""
    fn = orig.name
    patched = fn in blank_patches.PATCHES
    violations: list[str] = []
    with zipfile.ZipFile(orig) as zo, zipfile.ZipFile(blank) as zb:
        no, nb = set(zo.namelist()), set(zb.namelist())
        with Package(str(orig)) as pkg:
            sheet_by_part = {s.part: (i, s) for i, s in enumerate(pkg.worksheets())}

        if nb - no:
            violations.append(f"{fn}: parts added {sorted(nb - no)}")
        removed = no - nb
        calc_dropped = CALC_CHAIN in removed
        if removed - {CALC_CHAIN}:
            violations.append(f"{fn}: parts removed {sorted(removed - {CALC_CHAIN})}")

        formulas_cleared = False
        for name in sorted(no & nb):
            a, b = zo.read(name), zb.read(name)
            sheet = sheet_by_part.get(name)
            if sheet:
                idx, s = sheet
                for token, pat in TOKENS.items():
                    ca, cb = len(re.findall(pat, a)), len(re.findall(pat, b))
                    if cb < ca:
                        violations.append(f"{fn}::{s.name}: {token} {ca} -> {cb}")
            if a == b:
                continue
            if sheet:
                idx, s = sheet
                if idx == 0 or s.state != "visible" or (fn, s.name) in SKIP_CLEAR:
                    kind = "hidden" if s.state != "visible" else "non-clearable"
                    violations.append(f"{fn}::{s.name}: {kind} sheet changed")
                    continue
                oa, ob = _outside_sheetdata(a), _outside_sheetdata(b)
                if oa != ob and not (
                    patched and re.sub(_DIMENSION, b"", oa) == re.sub(_DIMENSION, b"", ob)
                ):
                    violations.append(f"{fn}::{s.name}: bytes outside <sheetData> changed")
                if len(re.findall(_FORMULA, b)) < len(re.findall(_FORMULA, a)):
                    formulas_cleared = True
            elif name == blank_patches.VBA_PART:
                if not patched:
                    violations.append(f"{fn}: {name} changed (no registered blank patch)")
            elif name in ("[Content_Types].xml", "xl/_rels/workbook.xml.rels") and calc_dropped:
                pat = _CALC_CT if name.startswith("[") else _CALC_REL
                if re.sub(pat, b"", a) != b:
                    violations.append(f"{fn}: {name} differs beyond calcChain removal")
            else:
                violations.append(f"{fn}: unexpected part changed: {name}")

        if calc_dropped and not formulas_cleared:
            violations.append(f"{fn}: calcChain dropped but no cleared formula found")
    return violations


def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    try:
        xml = z.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    return [unescape(b"".join(_T_RE.findall(si)).decode("utf-8")) for si in _SI_RE.findall(xml)]


def last_col_a_row(sheet_xml: bytes, sst: list[str]) -> int | None:
    """Highest row number with a non-empty column-A value, or None."""
    last = None
    for m in CELL_RE.finditer(sheet_xml):
        body = m.group(2)
        if body is None:
            continue
        attrs = tag_attrs(b"<c" + m.group(1) + b">")
        ref = _COL_A_REF.fullmatch(attrs.get("r", ""))
        if not ref:
            continue
        t = attrs.get("t")
        if t == "inlineStr":
            val = unescape(b"".join(_T_RE.findall(body)).decode("utf-8"))
        else:
            v = _V_RE.search(body)
            if v is None:
                continue
            val = v.group(1).decode("utf-8")
            if t == "s":
                val = sst[int(val)]
        if val.strip():
            last = max(last or 0, int(ref.group(1)))
    return last


def col_a_anchor(orig: Path, blank: Path) -> tuple[list[dict], list[dict], int]:
    """(mismatches, no_header, sheets_checked) for one pair's visible data sheets."""
    with Package(str(blank)) as pkg:
        visible = [
            s for s in pkg.worksheets()[1:]
            if s.state == "visible" and (blank.name, s.name) not in SKIP_CLEAR
        ]
    if not visible:
        return [], [], 0
    mismatches, no_header = [], []
    Font.family.max = 255  # Oracle templates use font families openpyxl rejects
    with quiet_openpyxl():
        wb = load_workbook(str(orig), read_only=True)
        try:
            with zipfile.ZipFile(blank) as zb:
                sst = _shared_strings(zb)
                for s in visible:
                    header = detect_header_row(wb[s.name])
                    last = last_col_a_row(zb.read(s.part), sst)
                    rec = {"file": blank.name, "sheet": s.name, "header": header, "last_col_a": last}
                    if header is None:
                        no_header.append(rec)
                    elif last != header:
                        mismatches.append(rec)
        finally:
            wb.close()
    return mismatches, no_header, len(visible)


def verify_release(release_dir: Path) -> dict:
    originals = release_dir / "originals"
    blanks = release_dir / "blanks"
    origs = sorted(p for p in originals.glob("*.xls[mx]") if not p.name.startswith("~$"))
    report = {
        "release": release_dir.name,
        "pairs": 0,
        "sheets_checked": 0,
        "missing_blanks": [],
        "violations": [],
        "anchor_mismatches": [],
        "no_header": [],
    }
    for orig in origs:
        blank = blanks / orig.name
        if not blank.exists():
            report["missing_blanks"].append(orig.name)
            continue
        report["pairs"] += 1
        report["violations"] += survey_pair(orig, blank)
        mism, nohdr, n = col_a_anchor(orig, blank)
        report["anchor_mismatches"] += mism
        report["no_header"] += nohdr
        report["sheets_checked"] += n
    report["ok"] = not (report["missing_blanks"] or report["violations"] or report["anchor_mismatches"])
    return report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("release", help="release label, e.g. 26D")
    parser.add_argument("--baselines-root", default=str(REPO / "baselines"))
    parser.add_argument("--json", action="store_true", help="print the full JSON report")
    args = parser.parse_args(argv)
    # detect_header logs a WARNING per undetectable sheet; those are reported
    # in the summary as "header not detected" instead.
    logging.getLogger("fbdi.detect_header").setLevel(logging.ERROR)

    release_dir = Path(args.baselines_root) / args.release.upper()
    if not (release_dir / "originals").is_dir():
        print(f"ERROR: {release_dir / 'originals'} not found", file=sys.stderr)
        return 1
    report = verify_release(release_dir)
    if args.json:
        print(json.dumps(report, indent=2))
    else:
        print(f"{report['release']}: {report['pairs']} pairs, {report['sheets_checked']} data sheets")
        print(f"  missing blanks:     {len(report['missing_blanks'])} {report['missing_blanks'] or ''}")
        print(f"  package violations: {len(report['violations'])}")
        for v in report["violations"][:50]:
            print(f"    {v}")
        print(f"  paste-anchor mismatches: {len(report['anchor_mismatches'])}")
        for m in report["anchor_mismatches"][:50]:
            print(f"    {m['file']}::{m['sheet']} header={m['header']} last_col_a={m['last_col_a']}")
        print(f"  header not detected (left uncleared, informational): {len(report['no_header'])}")
        print("  OK" if report["ok"] else "  FAIL")
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
