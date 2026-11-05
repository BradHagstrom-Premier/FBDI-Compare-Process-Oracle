"""Tests for fbdi.blank_patches — per-template post-clear fixes.

RapidImplementationForCashManagement.xlsm gets two fixes that were first made by
hand on a live engagement:
  * A7 paste anchor on "Bank Account": the Conversion Dashboard pastes at
    (last non-empty cell in column A) + 1; the template's column A is empty, so
    data landed on row 2 over the title and headers.
  * VBA Save As fix: CommandButton1_Click hard-coded ``.FilterIndex = 5 'for xml``,
    which is "CSV UTF-8" in current Excel, so the generated XML was saved as .csv.
"""

import hashlib
import re
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from fbdi import blank_patches
from fbdi.blank_patches import (
    ANCHOR_TEXT,
    PATCHED_VBA_ASSET,
    RAPID_IMPL,
    BlankPatchError,
    insert_anchor,
    swap_vba_project,
)
from fbdi.ooxml import Package

REPO = Path(__file__).resolve().parent.parent


def _make_bank_account_book(path: Path, with_vba: bytes | None = None) -> None:
    wb = Workbook()
    wb.active.title = "Instructions"
    ws = wb.create_sheet("Bank Account")
    ws["B2"] = "Banks, Branches, and Accounts"
    for col, label in enumerate(["*Country", "*Bank Name", "Bank Code"], start=2):
        ws.cell(row=7, column=col, value=label)
    wb.save(path)
    wb.close()
    if with_vba is not None:
        with zipfile.ZipFile(path, "a") as zf:
            zf.writestr("xl/vbaProject.bin", with_vba)


def _sheet_xml(path: Path, part: str = "xl/worksheets/sheet2.xml") -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read(part).decode("utf-8")


class TestInsertAnchor:
    def test_anchor_written_to_a7(self, tmp_path):
        src, dst = tmp_path / "src.xlsm", tmp_path / "dst.xlsm"
        _make_bank_account_book(src)
        with Package(str(src)) as pkg:
            insert_anchor(pkg, {"Bank Account": 7})
            pkg.save(str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["Bank Account"]
        assert ws["A7"].value == ANCHOR_TEXT
        assert ws["B7"].value == "*Country"
        wb.close()

    def test_anchor_is_first_cell_of_row(self, tmp_path):
        """Excel requires cells in column order within a row."""
        src, dst = tmp_path / "src.xlsm", tmp_path / "dst.xlsm"
        _make_bank_account_book(src)
        with Package(str(src)) as pkg:
            insert_anchor(pkg, {"Bank Account": 7})
            pkg.save(str(dst))

        row7 = re.search(r'<row r="7"[^>]*>(.*?)</row>', _sheet_xml(dst)).group(1)
        assert re.findall(r'<c r="([A-Z]+)7"', row7)[:2] == ["A", "B"]

    def test_dimension_widened_to_column_a(self, tmp_path):
        src, dst = tmp_path / "src.xlsm", tmp_path / "dst.xlsm"
        _make_bank_account_book(src)
        with Package(str(src)) as pkg:
            insert_anchor(pkg, {"Bank Account": 7})
            pkg.save(str(dst))

        assert re.search(r'<dimension ref="A\d+:', _sheet_xml(dst))

    def test_existing_empty_a7_replaced_keeping_style(self, tmp_path):
        src, dst = tmp_path / "src.xlsm", tmp_path / "dst.xlsm"
        _make_bank_account_book(src)
        part = "xl/worksheets/sheet2.xml"
        with Package(str(src)) as pkg:
            xml = pkg.read(part)
            pkg.write(part, xml.replace(b'<c r="B7"', b'<c r="A7" s="3"/><c r="B7"', 1))
            pkg.save(str(src) + ".2")
        with Package(str(src) + ".2") as pkg:
            insert_anchor(pkg, {"Bank Account": 7})
            pkg.save(str(dst))

        xml = _sheet_xml(dst)
        assert xml.count('r="A7"') == 1
        assert re.search(r'<c r="A7" s="3" t="inlineStr"><is><t>', xml)

    def test_header_row_moved_fails_loudly(self, tmp_path):
        """If Oracle moves the header, the anchor row must be re-derived by a human."""
        src = tmp_path / "src.xlsm"
        _make_bank_account_book(src)
        with Package(str(src)) as pkg, pytest.raises(BlankPatchError, match="header"):
            insert_anchor(pkg, {"Bank Account": 8})


class TestSwapVbaProject:
    def test_matching_hash_swaps_bytes(self, tmp_path):
        original, patched = b"original-vba", b"patched-vba"
        src, dst = tmp_path / "src.xlsm", tmp_path / "dst.xlsm"
        _make_bank_account_book(src, with_vba=original)
        with Package(str(src)) as pkg:
            swap_vba_project(pkg, hashlib.sha256(original).hexdigest(), patched)
            pkg.save(str(dst))

        with zipfile.ZipFile(dst) as zf:
            assert zf.read("xl/vbaProject.bin") == patched

    def test_hash_mismatch_fails_loudly(self, tmp_path):
        """Oracle changed the macro: never ship a blank silently missing the fix."""
        src = tmp_path / "src.xlsm"
        _make_bank_account_book(src, with_vba=b"oracle-changed-this")
        with Package(str(src)) as pkg, pytest.raises(BlankPatchError, match="build_vba_patch"):
            swap_vba_project(pkg, hashlib.sha256(b"something-else").hexdigest(), b"patched")


def _vba_source(path_or_bytes, filename="vbaProject.bin") -> str:
    olevba = pytest.importorskip("oletools.olevba")
    if isinstance(path_or_bytes, bytes):
        parser = olevba.VBA_Parser(filename, data=path_or_bytes)
    else:
        parser = olevba.VBA_Parser(str(path_or_bytes))
    try:
        return "\n".join(code for _, _, _, code in parser.extract_macros())
    finally:
        parser.close()


def _assert_save_as_fixed(code: str) -> None:
    assert 'InitialFileName = ThisWorkbook.Path & "\\CE_RAPID_IMPLEMENTATION.xml"' in code
    assert ".Filters(k).Extensions" in code
    assert 'gsFilePath = Left(gsFilePath, p) & nm & ".xml"' in code
    assert not re.search(r"FilterIndex\s*=\s*5", code)


class TestPatchedVbaAsset:
    def test_asset_registered_for_rapid_impl(self):
        assert RAPID_IMPL in blank_patches.PATCHES

    def test_asset_contains_save_as_fix(self):
        _assert_save_as_fixed(_vba_source(PATCHED_VBA_ASSET.read_bytes()))

    def test_asset_keeps_command_button_binding(self):
        """The sheet module must still declare the ActiveX button. Excel drops this
        attribute when it saves with macros/ActiveX force-disabled, which would
        leave CommandButton1_Click unbound to the button."""
        code = _vba_source(PATCHED_VBA_ASSET.read_bytes())
        assert 'Attribute VB_Control = "CommandButton1' in code


# ---------------------------------------------------------------------------
# Environment-gated: the real regenerated blanks (baselines/ is gitignored).
# ---------------------------------------------------------------------------

RAPID_BLANKS = sorted(REPO.glob(f"baselines/*/blanks/{RAPID_IMPL}"))


@pytest.mark.skipif(not RAPID_BLANKS, reason="no RapidImpl blanks under baselines/")
@pytest.mark.parametrize("blank", RAPID_BLANKS, ids=lambda p: p.parent.parent.name)
class TestRapidImplBlank:
    def test_a7_anchor_present(self, blank):
        wb = load_workbook(blank, read_only=True)
        try:
            assert wb["Bank Account"]["A7"].value == ANCHOR_TEXT
        finally:
            wb.close()

    def test_command_button_present(self, blank):
        with zipfile.ZipFile(blank) as zf:
            xml = zf.read("xl/worksheets/sheet2.xml").decode("utf-8")
        assert re.search(r'<control\b[^>]*name="CommandButton1"', xml)

    def test_vba_save_as_fixed(self, blank):
        _assert_save_as_fixed(_vba_source(blank))
