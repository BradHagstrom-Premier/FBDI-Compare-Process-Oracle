"""Environment-gated integrity checks: every blank keeps what its original has.

Skips when baselines/ is absent (it is gitignored). Guards against the blanks
writer silently dropping package content — the openpyxl re-save it replaced
removed Oracle's ActiveX Generate buttons from every 26A-26C template and wiped
the hidden LOV sheets behind the dropdowns.
"""

import re
from pathlib import Path

import pytest

from fbdi.ooxml import Package

REPO = Path(__file__).resolve().parent.parent
PAIRS = sorted(
    (orig, orig.parent.parent / "blanks" / orig.name)
    for orig in REPO.glob("baselines/*/originals/*.xls[mx]")
    if not orig.name.startswith("~$")
    and (orig.parent.parent / "blanks" / orig.name).exists()
)

_CONTROL_NAME_RE = re.compile(rb'<control\b[^>]*\bname="([^"]+)"')
_VALUE_RE = re.compile(rb"<v>|<is>")

pytestmark = pytest.mark.skipif(not PAIRS, reason="no original/blank pairs under baselines/")


def _ids(pair):
    orig, _ = pair
    return f"{orig.parent.parent.name}/{orig.name}"


def _per_sheet(path: Path):
    with Package(str(path)) as pkg:
        return {s.name: (s.state, pkg.read(s.part)) for s in pkg.worksheets()}


@pytest.mark.parametrize("pair", PAIRS, ids=_ids)
def test_blank_keeps_every_control(pair):
    orig, blank = pair
    o, b = _per_sheet(orig), _per_sheet(blank)
    missing = {
        name: sorted(set(_CONTROL_NAME_RE.findall(xml)) - set(_CONTROL_NAME_RE.findall(b[name][1])))
        for name, (_, xml) in o.items()
        if name in b
    }
    missing = {k: v for k, v in missing.items() if v}
    assert not missing, f"blank lost ActiveX/form controls: {missing}"


@pytest.mark.parametrize("pair", PAIRS, ids=_ids)
def test_blank_keeps_hidden_sheet_data(pair):
    orig, blank = pair
    o, b = _per_sheet(orig), _per_sheet(blank)
    wiped = [
        name for name, (state, xml) in o.items()
        if state != "visible" and name in b
        and len(_VALUE_RE.findall(b[name][1])) < len(_VALUE_RE.findall(xml))
    ]
    assert not wiped, f"hidden (LOV/reference) sheets were cleared: {wiped}"
