"""Tests for tools/verify_blanks.py — original-vs-blank package survey and the
Conversion Dashboard paste-anchor check.

Blanks are produced by the real fbdi.clear.clear_workbook from synthetic
originals, then tampered with to prove each violation class is caught.
"""

import importlib.util
import sys
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

from fbdi.clear import clear_workbook

TOOL = Path(__file__).resolve().parent.parent / "tools" / "verify_blanks.py"


@pytest.fixture(scope="module")
def vb():
    spec = importlib.util.spec_from_file_location("verify_blanks", TOOL)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["verify_blanks"] = mod
    spec.loader.exec_module(mod)
    return mod


def _make_original(path: Path, *, header_col_a: bool = True, hidden_lov: bool = True) -> None:
    """Instructions sheet + one data sheet (header row 4, sample rows) + hidden LOV."""
    wb = Workbook()
    ins = wb.active
    ins.title = "Instructions"
    ins["B2"] = "Oracle Fusion Applications"
    ws = wb.create_sheet("DATA_TAB")
    ws["B1"] = "Some Import"
    start = 1 if header_col_a else 2
    for i, h in enumerate(["BATCH_ID", "ITEM_NUMBER", "ORG_CODE"], start=start):
        ws.cell(row=4, column=i, value=h)
        ws.cell(row=5, column=i, value=f"sample-{i}")
        ws.cell(row=6, column=i, value=f"sample-{i}b")
    dv = DataValidation(type="list", formula1='"Yes,No"')
    dv.add("B5:B100")
    ws.add_data_validation(dv)
    if hidden_lov:
        lov = wb.create_sheet("LOV")
        lov["A1"] = "Yes"
        lov["A2"] = "No"
        lov.sheet_state = "hidden"
    wb.save(path)


def _release(tmp_path: Path, **kw) -> tuple[Path, Path, Path]:
    rel = tmp_path / "26D"
    (rel / "originals").mkdir(parents=True)
    (rel / "blanks").mkdir()
    orig = rel / "originals" / "SomeImportTemplate.xlsx"
    blank = rel / "blanks" / orig.name
    _make_original(orig, **kw)
    clear_workbook(str(orig), str(blank))
    return rel, orig, blank


def _rewrite(blank: Path, edit) -> None:
    """Rewrite the blank zip, passing {name: bytes} through `edit` (mutate in place)."""
    with zipfile.ZipFile(blank) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    edit(parts)
    with zipfile.ZipFile(blank, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)


def _sheet_part(blank: Path, name: str) -> str:
    from fbdi.ooxml import Package
    with Package(str(blank)) as pkg:
        return next(s.part for s in pkg.worksheets() if s.name == name)


class TestSurveyPair:
    def test_real_clear_output_is_clean(self, vb, tmp_path):
        _, orig, blank = _release(tmp_path)
        assert vb.survey_pair(orig, blank) == []

    def test_unexpected_part_change_is_caught(self, vb, tmp_path):
        _, orig, blank = _release(tmp_path)
        _rewrite(blank, lambda p: p.__setitem__("xl/styles.xml", p["xl/styles.xml"] + b"<!-- x -->"))
        assert any("unexpected part changed: xl/styles.xml" in v for v in vb.survey_pair(orig, blank))

    def test_removed_part_is_caught(self, vb, tmp_path):
        _, orig, blank = _release(tmp_path)
        _rewrite(blank, lambda p: p.pop("docProps/app.xml"))
        assert any("parts removed" in v and "docProps/app.xml" in v for v in vb.survey_pair(orig, blank))

    def test_hidden_sheet_change_is_caught(self, vb, tmp_path):
        _, orig, blank = _release(tmp_path)
        part = _sheet_part(blank, "LOV")
        _rewrite(blank, lambda p: p.__setitem__(part, p[part].replace(b"<sheetData>", b"<sheetData> ")))
        assert any("LOV: hidden sheet changed" in v for v in vb.survey_pair(orig, blank))

    def test_lost_data_validation_is_caught(self, vb, tmp_path):
        import re
        _, orig, blank = _release(tmp_path)
        part = _sheet_part(blank, "DATA_TAB")
        strip = lambda p: p.__setitem__(  # noqa: E731
            part, re.sub(rb"<dataValidations\b.*?</dataValidations>", b"", p[part], flags=re.S)
        )
        _rewrite(blank, strip)
        violations = vb.survey_pair(orig, blank)
        assert any("DATA_TAB: dataValidation 1 -> 0" in v for v in violations)
        assert any("bytes outside <sheetData> changed" in v for v in violations)


class TestColAAnchor:
    def test_header_in_col_a_matches(self, vb, tmp_path):
        _, orig, blank = _release(tmp_path)
        mismatches, no_header, checked = vb.col_a_anchor(orig, blank)
        assert (mismatches, no_header, checked) == ([], [], 1)

    def test_empty_col_a_at_header_is_a_mismatch(self, vb, tmp_path):
        # RapidImpl's pre-fix shape: col A empty, headers from B. The Dashboard
        # would paste on row 1 (last col-A row None -> 0, +1), over the header.
        _, orig, blank = _release(tmp_path, header_col_a=False)
        mismatches, _, _ = vb.col_a_anchor(orig, blank)
        assert mismatches == [{"file": blank.name, "sheet": "DATA_TAB", "header": 4, "last_col_a": None}]


class TestVerifyRelease:
    def test_clean_release(self, vb, tmp_path):
        rel, _, _ = _release(tmp_path)
        report = vb.verify_release(rel)
        assert report["ok"] is True
        assert report["pairs"] == 1

    def test_missing_blank_fails(self, vb, tmp_path):
        rel, _, blank = _release(tmp_path)
        blank.unlink()
        report = vb.verify_release(rel)
        assert report["ok"] is False
        assert report["missing_blanks"] == ["SomeImportTemplate.xlsx"]

    def test_main_exit_codes(self, vb, tmp_path, capsys):
        rel, _, blank = _release(tmp_path)
        assert vb.main(["26D", "--baselines-root", str(tmp_path)]) == 0
        _rewrite(blank, lambda p: p.pop("docProps/app.xml"))
        assert vb.main(["26d", "--baselines-root", str(tmp_path)]) == 1
        assert "FAIL" in capsys.readouterr().out
