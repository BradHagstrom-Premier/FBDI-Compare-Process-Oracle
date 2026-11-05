"""Environment-gated integrity checks: every blank is its original with only
the cleared sheets' <sheetData> changed, and its column-A paste anchor sits on
the header row.

Skips when baselines/ is absent (it is gitignored). Guards against the blanks
writer silently dropping package content — the openpyxl re-save it replaced
removed Oracle's ActiveX Generate buttons from every 26A-26C template and wiped
the hidden LOV sheets behind the dropdowns. The checks themselves live in
tools/verify_blanks.py (the skill's Stage 4.6 gate), so the test suite and the
pipeline enforce the same rules.
"""

import importlib.util
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
PAIRS = sorted(
    (orig, orig.parent.parent / "blanks" / orig.name)
    for orig in REPO.glob("baselines/*/originals/*.xls[mx]")
    if not orig.name.startswith("~$")
    and (orig.parent.parent / "blanks" / orig.name).exists()
)

pytestmark = pytest.mark.skipif(not PAIRS, reason="no original/blank pairs under baselines/")


@pytest.fixture(scope="module")
def vb():
    spec = importlib.util.spec_from_file_location("verify_blanks", REPO / "tools" / "verify_blanks.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules["verify_blanks"] = mod
    spec.loader.exec_module(mod)
    return mod


def _ids(pair):
    orig, _ = pair
    return f"{orig.parent.parent.name}/{orig.name}"


@pytest.mark.parametrize("pair", PAIRS, ids=_ids)
def test_blank_package_matches_original(vb, pair):
    violations = vb.survey_pair(*pair)
    assert not violations, "\n".join(violations)


@pytest.mark.parametrize("pair", PAIRS, ids=_ids)
def test_blank_paste_anchor_on_header_row(vb, pair):
    mismatches, _no_header, _checked = vb.col_a_anchor(*pair)
    assert not mismatches, mismatches
