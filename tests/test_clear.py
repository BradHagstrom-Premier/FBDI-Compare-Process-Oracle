"""Tests for fbdi.clear — smart FBDI template clearing."""

import re
import zipfile

import pytest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from fbdi.clear import clear_workbook


UPPER_SNAKE_HEADERS = [
    "TRANSACTION_TYPE", "BATCH_ID", "ORGANIZATION_CODE",
    "CHANGE_NOTICE", "CHANGE_NAME", "DESCRIPTION",
]

HUMAN_HEADERS = [
    "*Transaction Type", "*Batch ID", "Organization Code",
    "Change Number", "Change Name", "Description",
]

SAMPLE_DATA_ROW = ["Create Order", "10017", "V1-org", "eco test 1", "Test Change", "A description"]


def _make_fbdi_workbook(
    path: Path,
    header_row: int,
    headers: list[str],
    data_rows: list[list] | None = None,
    extra_rows: dict[int, list] | None = None,
    include_instructions_sheet: bool = True,
):
    """Create a test FBDI workbook with headers and sample data."""
    wb = Workbook()
    if include_instructions_sheet:
        ws_instr = wb.active
        ws_instr.title = "Instructions and CSV Generation"
        ws_instr.cell(row=1, column=1, value="Instructions")
        ws_instr.cell(row=5, column=1, value="This row should NOT be cleared")
        ws_instr.cell(row=6, column=1, value="Neither should this one")
    else:
        wb.remove(wb.active)

    ws = wb.create_sheet(title="DATA_INTERFACE")
    if extra_rows:
        for row_num, values in extra_rows.items():
            for col_idx, val in enumerate(values, start=1):
                ws.cell(row=row_num, column=col_idx, value=val)
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col_idx, value=header)
    if data_rows:
        for offset, row_data in enumerate(data_rows, start=1):
            for col_idx, val in enumerate(row_data, start=1):
                ws.cell(row=header_row + offset, column=col_idx, value=val)

    wb.save(path)
    wb.close()


class TestClearWorkbook:
    def test_headers_at_row_4_preserved(self, tmp_path):
        """Standard pattern: headers at row 4, data at row 5+."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=4,
            headers=HUMAN_HEADERS,
            data_rows=[SAMPLE_DATA_ROW, SAMPLE_DATA_ROW],
            extra_rows={2: ["Invoice Headers"], 3: ["* Required"]},
        )

        skipped = clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["DATA_INTERFACE"]
        # Headers at row 4 preserved
        assert ws.cell(row=4, column=1).value == "*Transaction Type"
        assert ws.cell(row=4, column=2).value == "*Batch ID"
        # Metadata above headers preserved
        assert ws.cell(row=2, column=1).value == "Invoice Headers"
        assert ws.cell(row=3, column=1).value == "* Required"
        # Data rows cleared
        assert ws.cell(row=5, column=1).value is None
        assert ws.cell(row=6, column=1).value is None
        assert skipped == []
        wb.close()

    def test_headers_at_row_5_preserved(self, tmp_path):
        """Pattern 1: technical UPPER_SNAKE headers at row 5."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=5,
            headers=UPPER_SNAKE_HEADERS,
            data_rows=[SAMPLE_DATA_ROW, SAMPLE_DATA_ROW],
            extra_rows={
                1: ["Name", "Transaction Type", "Batch ID", "Change Type", "Change Number", "Description"],
                2: ["Description", "Indicates the...", "Indicates the...", "Value that...", "Value that...", "Description of..."],
                3: ["Data Type", "VARCHAR2(10)", "NUMBER(18)", "VARCHAR2(80)", "VARCHAR2(50)", "VARCHAR2(2000)"],
                4: ["Required", "Required", "Required", "Required", "Optional", "Optional"],
            },
        )

        skipped = clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["DATA_INTERFACE"]
        # Technical headers at row 5 preserved
        assert ws.cell(row=5, column=1).value == "TRANSACTION_TYPE"
        assert ws.cell(row=5, column=2).value == "BATCH_ID"
        # Metadata rows 1-4 preserved
        assert ws.cell(row=1, column=1).value == "Name"
        assert ws.cell(row=4, column=1).value == "Required"
        # Data rows cleared
        assert ws.cell(row=6, column=1).value is None
        assert ws.cell(row=7, column=1).value is None
        assert skipped == []
        wb.close()

    def test_headers_at_row_8_preserved(self, tmp_path):
        """Pattern 2: deeper headers at row 8 (e.g., ChangeOrderImportTemplate)."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=8,
            headers=UPPER_SNAKE_HEADERS,
            data_rows=[SAMPLE_DATA_ROW],
            extra_rows={
                4: HUMAN_HEADERS,
                5: ["Description"] * 6,
                6: ["VARCHAR2(10)"] * 6,
                7: ["Reserved for Future Use"],
            },
        )

        skipped = clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["DATA_INTERFACE"]
        # Technical headers at row 8 preserved
        assert ws.cell(row=8, column=1).value == "TRANSACTION_TYPE"
        # All metadata rows 1-7 preserved
        assert ws.cell(row=4, column=1).value == "*Transaction Type"
        assert ws.cell(row=7, column=1).value == "Reserved for Future Use"
        # Data row 9 cleared
        assert ws.cell(row=9, column=1).value is None
        assert skipped == []
        wb.close()

    def test_instructions_sheet_untouched(self, tmp_path):
        """Sheet 1 (Instructions) should never be modified."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"
        _make_fbdi_workbook(
            src,
            header_row=4,
            headers=UPPER_SNAKE_HEADERS,
            data_rows=[SAMPLE_DATA_ROW],
        )

        clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws_instr = wb["Instructions and CSV Generation"]
        assert ws_instr.cell(row=1, column=1).value == "Instructions"
        assert ws_instr.cell(row=5, column=1).value == "This row should NOT be cleared"
        assert ws_instr.cell(row=6, column=1).value == "Neither should this one"
        wb.close()

    def test_detection_fails_sheet_skipped_not_destroyed(self, tmp_path):
        """If detect_header_row returns None, leave the sheet untouched."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"

        wb = Workbook()
        ws_instr = wb.active
        ws_instr.title = "Instructions and CSV Generation"
        ws_data = wb.create_sheet(title="WeirdSheet")
        # Only 1 cell — below MIN_CELLS=2, detection will return None
        ws_data.cell(row=3, column=1, value="SOLO_HEADER")
        ws_data.cell(row=4, column=1, value="solo_data")
        wb.save(src)
        wb.close()

        skipped = clear_workbook(str(src), str(dst))

        assert "WeirdSheet" in skipped

        wb = load_workbook(dst, read_only=True)
        ws = wb["WeirdSheet"]
        # Data should be untouched since detection failed
        assert ws.cell(row=3, column=1).value == "SOLO_HEADER"
        assert ws.cell(row=4, column=1).value == "solo_data"
        wb.close()

    def test_multiple_data_sheets_cleared_independently(self, tmp_path):
        """Each sheet gets its own header detection and clearing."""
        src = tmp_path / "Template.xlsx"
        dst = tmp_path / "Template_cleared.xlsx"

        wb = Workbook()
        ws_instr = wb.active
        ws_instr.title = "Instructions and CSV Generation"

        # Sheet 2: headers at row 4
        ws1 = wb.create_sheet(title="SHEET_A")
        for col, val in enumerate(UPPER_SNAKE_HEADERS, start=1):
            ws1.cell(row=4, column=col, value=val)
        for col, val in enumerate(SAMPLE_DATA_ROW, start=1):
            ws1.cell(row=5, column=col, value=val)

        # Sheet 3: headers at row 8
        ws2 = wb.create_sheet(title="SHEET_B")
        for col, val in enumerate(HUMAN_HEADERS, start=1):
            ws2.cell(row=4, column=col, value=val)
        for col, val in enumerate(["Description"] * 6, start=1):
            ws2.cell(row=5, column=col, value=val)
        for col, val in enumerate(["VARCHAR2(10)"] * 6, start=1):
            ws2.cell(row=6, column=col, value=val)
        ws2.cell(row=7, column=1, value="Reserved")
        for col, val in enumerate(UPPER_SNAKE_HEADERS, start=1):
            ws2.cell(row=8, column=col, value=val)
        for col, val in enumerate(SAMPLE_DATA_ROW, start=1):
            ws2.cell(row=9, column=col, value=val)

        wb.save(src)
        wb.close()

        clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        # Sheet A: headers at row 4, data at row 5 cleared
        ws_a = wb["SHEET_A"]
        assert ws_a.cell(row=4, column=1).value == "TRANSACTION_TYPE"
        assert ws_a.cell(row=5, column=1).value is None

        # Sheet B: headers at row 8, data at row 9 cleared
        ws_b = wb["SHEET_B"]
        assert ws_b.cell(row=4, column=1).value == "*Transaction Type"
        assert ws_b.cell(row=8, column=1).value == "TRANSACTION_TYPE"
        assert ws_b.cell(row=9, column=1).value is None
        wb.close()


# ---------------------------------------------------------------------------
# Package preservation: clear_workbook must patch the original package in place
# (cleared cells only), never re-serialize it through openpyxl. openpyxl drops
# ActiveX <controls> (Oracle's Generate CSV button), x14 data validations,
# drawings and printer settings — that is how every 26A-26C blank lost its button.
# ---------------------------------------------------------------------------

CONTROLS_XML = (
    '<controls xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<control shapeId="1025" r:id="rIdCtl1" name="CommandButton1"/></controls>'
)
X14_EXTLST_XML = (
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>LOV!A2:A3</xm:f>'
    '</x14:formula1><xm:sqref>A5:A100</xm:sqref></x14:dataValidation></x14:dataValidations>'
    '</ext></extLst>'
)

CALC_CHAIN_CT = (
    b'<Override PartName="/xl/calcChain.xml" ContentType="application/'
    b'vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
)
CALC_CHAIN_REL = (
    b'<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/'
    b'officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>'
)


def _sheet_part(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """Resolve a sheet name to its worksheet part path via workbook.xml + rels."""
    wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    tag = re.search(r'<sheet\b[^>]*name="%s"[^>]*>' % re.escape(sheet_name), wb_xml).group(0)
    rid = re.search(r'r:id="([^"]+)"', tag).group(1)
    rel = re.search(r'<Relationship\b[^>]*Id="%s"[^>]*>' % re.escape(rid), rels).group(0)
    target = re.search(r'Target="([^"]+)"', rel).group(1)
    return target.lstrip("/") if target.startswith("/") else "xl/" + target


def _rewrite_parts(path: Path, edits: dict, add: dict | None = None) -> None:
    """Rewrite zip parts in place: edits maps part name -> fn(bytes) -> bytes."""
    with zipfile.ZipFile(path) as zin:
        items = [(i, zin.read(i)) for i in zin.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename in edits:
                data = edits[info.filename](data)
            zout.writestr(info, data)
        for name, data in (add or {}).items():
            zout.writestr(name, data)


def _make_template_with_extras(path: Path) -> None:
    """FBDI-like workbook + a hidden LOV sheet, then inject <controls> and x14
    validations into the data sheet at zip level (openpyxl cannot write them)."""
    _make_fbdi_workbook(
        path, header_row=4, headers=UPPER_SNAKE_HEADERS,
        data_rows=[SAMPLE_DATA_ROW, SAMPLE_DATA_ROW],
    )
    wb = load_workbook(path)
    lov = wb.create_sheet("LOV")
    lov.sheet_state = "hidden"
    lov.append(["TERRITORY_SHORT_NAME", "TERRITORY_CODE"])
    lov.append(["Afghanistan", "AF"])
    lov.append(["Albania", "AL"])
    wb.save(path)
    wb.close()
    with zipfile.ZipFile(path) as zf:
        part = _sheet_part(zf, "DATA_INTERFACE")
    extra = (CONTROLS_XML + X14_EXTLST_XML).encode("utf-8")
    _rewrite_parts(path, {part: lambda b: b.replace(b"</worksheet>", extra + b"</worksheet>")})


class TestPackagePreservation:
    def test_controls_and_x14_validations_survive(self, tmp_path):
        src = tmp_path / "Template.xlsm"
        dst = tmp_path / "Template_blank.xlsm"
        _make_template_with_extras(src)

        clear_workbook(str(src), str(dst))

        with zipfile.ZipFile(dst) as zf:
            xml = zf.read(_sheet_part(zf, "DATA_INTERFACE")).decode("utf-8")
        assert CONTROLS_XML in xml
        assert X14_EXTLST_XML in xml

    def test_data_rows_still_cleared(self, tmp_path):
        src = tmp_path / "Template.xlsm"
        dst = tmp_path / "Template_blank.xlsm"
        _make_template_with_extras(src)

        clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        ws = wb["DATA_INTERFACE"]
        assert ws.cell(row=4, column=1).value == "TRANSACTION_TYPE"
        assert ws.cell(row=5, column=1).value is None
        assert ws.cell(row=6, column=2).value is None
        wb.close()

    def test_untouched_parts_are_byte_identical(self, tmp_path):
        """Only the cleared data-sheet part may differ; everything else (styles,
        shared strings, the instructions sheet, VBA, ActiveX) is copied verbatim."""
        src = tmp_path / "Template.xlsm"
        dst = tmp_path / "Template_blank.xlsm"
        _make_template_with_extras(src)

        clear_workbook(str(src), str(dst))

        with zipfile.ZipFile(src) as zs, zipfile.ZipFile(dst) as zd:
            data_part = _sheet_part(zs, "DATA_INTERFACE")
            assert sorted(zs.namelist()) == sorted(zd.namelist())
            for name in zs.namelist():
                if name != data_part:
                    assert zs.read(name) == zd.read(name), name

    def test_hidden_sheets_not_cleared(self, tmp_path):
        """Hidden sheets hold the LOVs behind named ranges/dropdowns — never sample data."""
        src = tmp_path / "Template.xlsm"
        dst = tmp_path / "Template_blank.xlsm"
        _make_template_with_extras(src)

        clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        lov = wb["LOV"]
        assert lov.cell(row=2, column=1).value == "Afghanistan"
        assert lov.cell(row=3, column=2).value == "AL"
        wb.close()

    def test_cleared_cell_keeps_style_index_and_drops_type(self, tmp_path):
        src = tmp_path / "Template.xlsm"
        dst = tmp_path / "Template_blank.xlsm"
        _make_fbdi_workbook(src, header_row=4, headers=UPPER_SNAKE_HEADERS,
                            data_rows=[SAMPLE_DATA_ROW])
        with zipfile.ZipFile(src) as zf:
            part = _sheet_part(zf, "DATA_INTERFACE")
        # Give B5 an explicit style index the way Oracle templates do.
        _rewrite_parts(src, {part: lambda b: re.sub(rb'<c r="B5"( s="\d+")?', b'<c r="B5" s="7"', b, count=1)})

        clear_workbook(str(src), str(dst))

        with zipfile.ZipFile(dst) as zf:
            xml = zf.read(part).decode("utf-8")
        assert '<c r="B5" s="7"/>' in xml

    def test_calc_chain_dropped_when_formula_cleared(self, tmp_path):
        """A formula below the header gets cleared; a stale calcChain.xml would then
        point at a non-formula cell and Excel would prompt to repair the file."""
        src = tmp_path / "Template.xlsm"
        dst = tmp_path / "Template_blank.xlsm"
        _make_fbdi_workbook(src, header_row=4, headers=UPPER_SNAKE_HEADERS,
                            data_rows=[SAMPLE_DATA_ROW])
        wb = load_workbook(src)
        wb["DATA_INTERFACE"]["G5"] = "=1+1"
        wb.save(src)
        wb.close()
        _rewrite_parts(
            src,
            {
                "[Content_Types].xml": lambda b: b.replace(b"</Types>", CALC_CHAIN_CT + b"</Types>"),
                "xl/_rels/workbook.xml.rels": lambda b: b.replace(
                    b"</Relationships>", CALC_CHAIN_REL + b"</Relationships>"),
            },
            add={"xl/calcChain.xml": b'<calcChain xmlns="http://schemas.openxmlformats.org/'
                                     b'spreadsheetml/2006/main"><c r="G5" i="2"/></calcChain>'},
        )

        clear_workbook(str(src), str(dst))

        with zipfile.ZipFile(dst) as zf:
            assert "xl/calcChain.xml" not in zf.namelist()
            assert b"calcChain" not in zf.read("xl/_rels/workbook.xml.rels")
            assert b"calcChain" not in zf.read("[Content_Types].xml")
        wb = load_workbook(dst)  # still a valid package
        assert wb["DATA_INTERFACE"]["G5"].value is None
        wb.close()

    def test_calc_chain_kept_when_no_formula_cleared(self, tmp_path):
        src = tmp_path / "Template.xlsm"
        dst = tmp_path / "Template_blank.xlsm"
        _make_fbdi_workbook(src, header_row=4, headers=UPPER_SNAKE_HEADERS,
                            data_rows=[SAMPLE_DATA_ROW])
        _rewrite_parts(
            src,
            {
                "[Content_Types].xml": lambda b: b.replace(b"</Types>", CALC_CHAIN_CT + b"</Types>"),
                "xl/_rels/workbook.xml.rels": lambda b: b.replace(
                    b"</Relationships>", CALC_CHAIN_REL + b"</Relationships>"),
            },
            add={"xl/calcChain.xml": b'<calcChain xmlns="http://schemas.openxmlformats.org/'
                                     b'spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>'},
        )

        clear_workbook(str(src), str(dst))

        with zipfile.ZipFile(dst) as zf:
            assert "xl/calcChain.xml" in zf.namelist()

    def test_settings_sheet_in_skip_list_not_cleared(self, tmp_path):
        """ProjectImportTemplate::Options is a Yes/No settings form, not a load sheet."""
        src = tmp_path / "ProjectImportTemplate.xlsm"
        dst = tmp_path / "blank" / "ProjectImportTemplate.xlsm"
        dst.parent.mkdir()
        wb = Workbook()
        wb.active.title = "Instructions and CSV Generation"
        ws = wb.create_sheet("Options")
        ws["B2"] = "Project Data Import Options"
        ws["B7"], ws["C7"], ws["D7"] = "Project Financial Management", "Select if used", "Yes"
        ws["B9"], ws["C9"], ws["D9"] = "Project Execution Management", "Select if used", "Yes"
        wb.save(src)
        wb.close()

        clear_workbook(str(src), str(dst))

        wb = load_workbook(dst, read_only=True)
        assert wb["Options"]["D9"].value == "Yes"
        wb.close()
