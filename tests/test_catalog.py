"""Tests for fbdi.catalog — master catalog generation."""

import pytest
from pathlib import Path
from openpyxl import Workbook, load_workbook

from fbdi.catalog import (
    CatalogRow,
    IssueRow,
    DriftRow,
    extract_tab_rows,
)


def _make_thin_tab(ws, labels: list[str], header_row: int = 4):
    """Build a thin-tab workbook: just a title/legend and a label row."""
    ws.cell(row=2, column=1, value="Some Import")
    ws.cell(row=3, column=1, value="* Required")
    for col_idx, label in enumerate(labels, start=1):
        ws.cell(row=header_row, column=col_idx, value=label)


class TestExtractTabRowsThin:
    def test_thin_tab_labels_only(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "XCC_BUDGET_INTERFACE"
        _make_thin_tab(ws, [
            "*Source Budget Type",
            "*Source Budget Name",
            "Line Number",
            "Amount",
        ])
        rows, issues = extract_tab_rows(
            ws, file_stem="BudgetImportTemplate", release="26B"
        )

        assert issues == []
        assert len(rows) == 4
        # Position 1: required (had asterisk), normalized label
        assert rows[0].position == 1
        assert rows[0].column_label == "Source Budget Type"
        assert rows[0].column_technical == ""
        assert rows[0].data_type == ""
        assert rows[0].length is None
        assert rows[0].scale is None
        assert rows[0].data_type_raw == ""
        assert rows[0].required is True
        # Position 2: required
        assert rows[1].column_label == "Source Budget Name"
        assert rows[1].required is True
        # Position 3: not required (no asterisk)
        assert rows[2].column_label == "Line Number"
        assert rows[2].required is False
        # Position 4: not required
        assert rows[3].column_label == "Amount"
        assert rows[3].required is False

    def test_thin_tab_sets_release_and_file_and_tab(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "MY_TAB"
        _make_thin_tab(ws, ["*Field A", "Field B"])
        rows, _ = extract_tab_rows(ws, file_stem="MyTemplate", release="26A")
        assert rows[0].release == "26A"
        assert rows[0].file_name == "MyTemplate"
        assert rows[0].tab_name == "MY_TAB"

    def test_thin_tab_no_header_emits_issue(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "EmptyTab"
        # Only a title — no detectable header row
        ws.cell(row=1, column=1, value="Just a title")

        rows, issues = extract_tab_rows(
            ws, file_stem="Tpl", release="26A"
        )

        assert rows == []
        assert len(issues) == 1
        assert issues[0].issue_type == "NO_HEADER"
        assert issues[0].tab == "EmptyTab"
        assert issues[0].release == "26A"
        assert issues[0].file == "Tpl"


def _make_rich_tab(
    ws,
    labels: list[str],
    descriptions: list[str] | None = None,
    data_types: list[str] | None = None,
    required_flags: list[str] | None = None,
    technicals: list[str] | None = None,
    header_row: int = 5,
    table_name: str = "RCS_ATTACHMENTS_INT",
):
    """Build a rich-tab workbook with metadata rows + technical header."""
    def _row(row_idx, label_a, values):
        ws.cell(row=row_idx, column=1, value=label_a)
        for col_idx, v in enumerate(values, start=2):
            ws.cell(row=row_idx, column=col_idx, value=v)
    # Header row: "Column name of the Table X" in col A, then tech names col B..
    _row(header_row, f"Column name of the Table {table_name}", technicals or labels)
    # Name row above, then Description, Data Type, Required
    if header_row >= 2:
        _row(header_row - 1, "Required or Optional", required_flags or ["Optional"] * len(labels))
    if header_row >= 3:
        _row(header_row - 2, "Data Type", data_types or [""] * len(labels))
    if header_row >= 4:
        _row(header_row - 3, "Description", descriptions or [""] * len(labels))
    if header_row >= 5:
        _row(header_row - 4, "Name", labels)


class TestExtractTabRowsRich:
    def test_rich_tab_all_metadata_rows(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Attachment Details"
        _make_rich_tab(
            ws,
            labels=["Attachment Type", "Attachment Name", "Document ID"],
            data_types=["VARCHAR2(5 CHAR)", "VARCHAR2(2048 CHAR)", "NUMBER(18)"],
            required_flags=["Required", "Required", "Optional"],
            technicals=["ATTACHMENT_TYPE", "ATTACHMENT_NAME", "DOCUMENT_ID"],
            table_name="RCS_ATTACHMENTS_INT",
            header_row=5,
        )

        rows, issues = extract_tab_rows(ws, file_stem="AttachmentsImportTemplate", release="26B")

        assert issues == []
        assert len(rows) == 3
        r0 = rows[0]
        assert r0.position == 1
        assert r0.column_label == "Attachment Type"
        assert r0.column_technical == "ATTACHMENT_TYPE"
        assert r0.data_type == "VARCHAR2"
        assert r0.length == 5
        assert r0.scale is None
        assert r0.data_type_raw == "VARCHAR2(5 CHAR)"
        assert r0.required is True
        assert rows[1].length == 2048
        assert rows[2].data_type == "NUMBER"
        assert rows[2].length == 18
        assert rows[2].required is False

    def test_rich_tab_with_bom_on_required_row(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "TabWithBOM"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["VARCHAR2(80)"],
            technicals=["COL_A"],
        )
        # Overwrite the 'Required or Optional' col-A label with BOM-prefixed variant
        # In _make_rich_tab, required is at header_row - 1 = row 4
        ws.cell(row=4, column=1, value="\ufeffRequired or Optional")
        ws.cell(row=4, column=2, value="Required")

        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        assert rows[0].required is True

    def test_rich_tab_case_insensitive_col_a_match(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "MixedCase"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["VARCHAR2(80)"],
            technicals=["COL_A"],
        )
        # Lowercase the 'Data Type' label
        ws.cell(row=3, column=1, value="DATA type")  # row 3 = Data Type row
        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        assert rows[0].data_type == "VARCHAR2"
        assert rows[0].length == 80

    def test_rich_tab_missing_data_type_row(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "NoDataType"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["VARCHAR2(80)"],
            technicals=["COL_A"],
        )
        # Clear the Data Type row's column A label so it's not recognized
        ws.cell(row=3, column=1, value="Reserved for Future Use")

        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        # Type fields blank; other fields still populate
        assert rows[0].column_label == "Col A"
        assert rows[0].column_technical == "COL_A"
        assert rows[0].data_type == ""
        assert rows[0].length is None
        assert rows[0].data_type_raw == ""

    def test_rich_tab_unparseable_type_emits_warning_issue(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "WeirdType"
        _make_rich_tab(
            ws, labels=["Col A"],
            data_types=["???junk???"],
            technicals=["COL_A"],
        )
        rows, issues = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        # Row still emitted, raw preserved, parsed fields blank
        assert rows[0].data_type == ""
        assert rows[0].length is None
        assert rows[0].data_type_raw == "???junk???"
        # One warning issue for that raw string
        warnings = [i for i in issues if i.issue_type == "TYPE_PARSE_WARNING"]
        assert len(warnings) == 1
        assert warnings[0].detail == "???junk???"

    def test_rich_tab_asterisk_in_label_stripped(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "StarLabel"
        _make_rich_tab(
            ws, labels=["*Required Label"],
            data_types=["VARCHAR2(80)"],
            technicals=["REQ_LABEL"],
        )
        rows, _ = extract_tab_rows(ws, file_stem="Tpl", release="26B")
        # Asterisk stripped by normalize_label; required comes from R4 row anyway
        assert rows[0].column_label == "Required Label"


from fbdi.catalog import extract_file


class TestExtractFile:
    def test_extract_file_multiple_tabs(self, tmp_path):
        path = tmp_path / "MultiTab.xlsm"
        wb = Workbook()
        wb.remove(wb.active)

        # Thin tab
        ws1 = wb.create_sheet("THIN_TAB")
        _make_thin_tab(ws1, ["*Field A", "Field B"])

        # Rich tab
        ws2 = wb.create_sheet("RICH_TAB")
        _make_rich_tab(
            ws2,
            labels=["Col A", "Col B"],
            data_types=["VARCHAR2(50)", "NUMBER(10)"],
            required_flags=["Required", "Optional"],
            technicals=["COL_A", "COL_B"],
        )
        wb.save(path)

        rows, issues = extract_file(path, release="26B")
        tabs = {r.tab_name for r in rows}
        assert tabs == {"THIN_TAB", "RICH_TAB"}
        # Thin tab contributes 2 rows, rich tab contributes 2 rows
        assert len(rows) == 4
        assert issues == []

    def test_extract_file_skips_instruction_tabs(self, tmp_path):
        from fbdi.config import SKIP_TABS
        path = tmp_path / "WithInstructions.xlsm"
        wb = Workbook()
        wb.remove(wb.active)
        for name in list(SKIP_TABS)[:2]:
            ws = wb.create_sheet(name)
            ws.cell(row=1, column=1, value="Instruction content")
        data_ws = wb.create_sheet("DATA_TAB")
        _make_thin_tab(data_ws, ["*Field One", "Field Two"])
        wb.save(path)

        rows, issues = extract_file(path, release="26B")
        tabs = {r.tab_name for r in rows}
        assert tabs == {"DATA_TAB"}
        assert issues == []

    def test_extract_file_load_error_yields_issue(self, tmp_path):
        path = tmp_path / "Corrupt.xlsm"
        path.write_bytes(b"not a real xlsx file")

        rows, issues = extract_file(path, release="26B")
        assert rows == []
        assert len(issues) == 1
        assert issues[0].issue_type == "FILE_ERROR"
        assert issues[0].file == "Corrupt"
        assert issues[0].tab == ""
        assert issues[0].release == "26B"
