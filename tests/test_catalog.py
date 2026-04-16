"""Tests for fbdi.catalog — master catalog generation."""

import pytest
from pathlib import Path
from openpyxl import Workbook, load_workbook

from fbdi.catalog import (
    CatalogRow,
    IssueRow,
    DriftRow,
    extract_tab_rows,
    _write_master_workbook,
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


from fbdi.catalog import _compute_drift


def _row(**kwargs) -> CatalogRow:
    defaults = dict(
        release="26A", file_name="F", tab_name="T", position=1,
        column_label="L", column_technical="T", data_type="VARCHAR2",
        length=50, scale=None, data_type_raw="VARCHAR2(50)", required=False,
    )
    defaults.update(kwargs)
    return CatalogRow(**defaults)


class TestComputeDrift:
    def test_unchanged_rows_not_in_drift(self):
        old = [_row(release="26A")]
        new = [_row(release="26B")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift == []

    def test_added_column(self):
        old = []
        new = [_row(release="26B")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].change_type == "ADDED"
        assert drift[0].col_label_old == ""
        assert drift[0].col_label_new == "L"

    def test_removed_column(self):
        old = [_row(release="26A")]
        new = []
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].change_type == "REMOVED"
        assert drift[0].col_label_new == ""

    def test_renamed_label_only(self):
        old = [_row(release="26A", column_label="Old Name")]
        new = [_row(release="26B", column_label="New Name")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].change_type == "RENAMED"

    def test_renamed_technical_only(self):
        old = [_row(release="26A", column_technical="OLD_NAME")]
        new = [_row(release="26B", column_technical="NEW_NAME")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "RENAMED"

    def test_type_changed_only(self):
        old = [_row(release="26A", data_type="VARCHAR2")]
        new = [_row(release="26B", data_type="NUMBER")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "TYPE_CHANGED"

    def test_length_changed_only(self):
        old = [_row(release="26A", length=50)]
        new = [_row(release="26B", length=100)]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "LENGTH_CHANGED"

    def test_scale_changed_classified_as_length(self):
        old = [_row(release="26A", data_type="NUMBER", length=18, scale=None, data_type_raw="NUMBER(18)")]
        new = [_row(release="26B", data_type="NUMBER", length=18, scale=4, data_type_raw="NUMBER(18,4)")]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        # data_type same, length same, only scale differs — classified as LENGTH_CHANGED
        assert drift[0].change_type == "LENGTH_CHANGED"

    def test_required_changed_only(self):
        old = [_row(release="26A", required=False)]
        new = [_row(release="26B", required=True)]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "REQUIRED_CHANGED"

    def test_multi_change(self):
        old = [_row(release="26A", data_type="VARCHAR2", length=50, required=False)]
        new = [_row(release="26B", data_type="NUMBER", length=18, required=True)]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert drift[0].change_type == "MULTI"

    def test_aligns_by_file_tab_position(self):
        old = [
            _row(release="26A", file_name="F1", tab_name="T1", position=1),
            _row(release="26A", file_name="F2", tab_name="T2", position=1),
        ]
        new = [
            _row(release="26B", file_name="F1", tab_name="T1", position=1, column_label="NEW"),
            _row(release="26B", file_name="F2", tab_name="T2", position=1),
        ]
        drift = _compute_drift(old, new, release_old="26A", release_new="26B")
        assert len(drift) == 1
        assert drift[0].file == "F1"


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


class TestWriteMasterWorkbook:
    def test_writes_release_tab_with_correct_headers(self, tmp_path):
        out = tmp_path / "Master.xlsx"
        rows_by_release = {
            "26A": [_row(release="26A", file_name="F", tab_name="T", position=1)],
        }
        _write_master_workbook(
            out,
            rows_by_release=rows_by_release,
            issues=[],
            drift=[],
            release_old=None,
            release_new="26A",
        )
        assert out.exists()
        wb = load_workbook(out)
        assert "26A" in wb.sheetnames
        assert "Issues" in wb.sheetnames
        assert "Drift" in wb.sheetnames
        ws = wb["26A"]
        headers = [c.value for c in ws[1]]
        assert headers == [
            "release", "file_name", "tab_name", "position",
            "column_label", "column_technical",
            "data_type", "length", "scale", "data_type_raw",
            "required",
        ]
        # Data row
        row2 = [c.value for c in ws[2]]
        assert row2[0] == "26A"
        assert row2[1] == "F"

    def test_writes_issues_tab(self, tmp_path):
        out = tmp_path / "Master.xlsx"
        issues = [IssueRow("26B", "F", "T", "FILE_ERROR", "boom")]
        _write_master_workbook(
            out, rows_by_release={}, issues=issues, drift=[],
            release_old=None, release_new=None,
        )
        wb = load_workbook(out)
        ws = wb["Issues"]
        headers = [c.value for c in ws[1]]
        assert headers == ["release", "file", "tab", "issue_type", "detail"]
        assert [c.value for c in ws[2]] == ["26B", "F", "T", "FILE_ERROR", "boom"]

    def test_writes_drift_tab(self, tmp_path):
        out = tmp_path / "Master.xlsx"
        drift = [DriftRow(
            file="F", tab="T", position=1,
            col_label_old="A", col_label_new="B",
            col_technical_old="A1", col_technical_new="B1",
            data_type_old="VARCHAR2", data_type_new="VARCHAR2",
            length_old="50", length_new="100",
            required_old="FALSE", required_new="FALSE",
            change_type="LENGTH_CHANGED",
        )]
        _write_master_workbook(
            out, rows_by_release={}, issues=[], drift=drift,
            release_old="26A", release_new="26B",
        )
        wb = load_workbook(out)
        ws = wb["Drift"]
        headers = [c.value for c in ws[1]]
        assert "col_label_26A" in headers
        assert "col_label_26B" in headers
        assert "change_type" in headers

    def test_idempotent_content(self, tmp_path):
        out1 = tmp_path / "M1.xlsx"
        out2 = tmp_path / "M2.xlsx"
        rows_by_release = {"26A": [_row(release="26A")]}
        for out in (out1, out2):
            _write_master_workbook(
                out, rows_by_release=rows_by_release, issues=[], drift=[],
                release_old=None, release_new="26A",
            )
        wb1 = load_workbook(out1)
        wb2 = load_workbook(out2)
        assert wb1.sheetnames == wb2.sheetnames
        for sn in wb1.sheetnames:
            r1 = [[c.value for c in row] for row in wb1[sn].iter_rows()]
            r2 = [[c.value for c in row] for row in wb2[sn].iter_rows()]
            assert r1 == r2, f"Tab {sn} content differs"

    def test_preserves_existing_release_tabs(self, tmp_path):
        """Writing release X shouldn't wipe release Y if Y was already present in the file."""
        out = tmp_path / "Master.xlsx"
        # First run: writes 26A
        _write_master_workbook(
            out, rows_by_release={"26A": [_row(release="26A")]},
            issues=[], drift=[],
            release_old=None, release_new="26A",
        )
        # Second run: writes 26B but must preserve 26A
        # (caller has loaded 26A rows from existing workbook and passes both)
        _write_master_workbook(
            out, rows_by_release={
                "26A": [_row(release="26A")],
                "26B": [_row(release="26B")],
            },
            issues=[], drift=[],
            release_old="26A", release_new="26B",
        )
        wb = load_workbook(out)
        assert "26A" in wb.sheetnames
        assert "26B" in wb.sheetnames


from fbdi.catalog import generate_catalog


def _make_rich_xlsm(path: Path, tab_name: str, labels, types, techs, required):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title=tab_name)
    _make_rich_tab(
        ws, labels=labels, data_types=types,
        required_flags=required, technicals=techs,
    )
    wb.save(path)


class TestGenerateCatalog:
    def test_end_to_end_single_release(self, tmp_path):
        release_dir = tmp_path / "baselines" / "TESTA" / "originals"
        release_dir.mkdir(parents=True)
        _make_rich_xlsm(
            release_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A", "Col B"],
            types=["VARCHAR2(50)", "NUMBER(18)"],
            techs=["COL_A", "COL_B"],
            required=["Required", "Optional"],
        )
        master = tmp_path / "Catalog.xlsx"
        generate_catalog(
            release="TESTA",
            baselines_dir=release_dir,
            master_path=master,
            timeout=60,
        )
        assert master.exists()
        wb = load_workbook(master)
        assert "TESTA" in wb.sheetnames
        assert "Issues" in wb.sheetnames
        assert "Drift" in wb.sheetnames
        ws = wb["TESTA"]
        data = [[c.value for c in row] for row in ws.iter_rows(min_row=2)]
        assert len(data) == 2
        drift_ws = wb["Drift"]
        drift_rows = [
            r for r in drift_ws.iter_rows(min_row=2)
            if any(c.value is not None for c in r)
        ]
        assert drift_rows == []

    def test_end_to_end_two_releases_drift_classifications(self, tmp_path):
        testa_dir = tmp_path / "baselines" / "TESTA" / "originals"
        testa_dir.mkdir(parents=True)
        _make_rich_xlsm(
            testa_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A", "Col B", "Col C"],
            types=["VARCHAR2(50)", "NUMBER(18)", "DATE"],
            techs=["COL_A", "COL_B", "COL_C"],
            required=["Required", "Optional", "Optional"],
        )
        testb_dir = tmp_path / "baselines" / "TESTB" / "originals"
        testb_dir.mkdir(parents=True)
        _make_rich_xlsm(
            testb_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A", "Col B", "Col C", "Col D"],
            types=["VARCHAR2(50)", "NUMBER(32)", "DATE", "VARCHAR2(10)"],
            techs=["COL_A_RENAMED", "COL_B", "COL_C", "COL_D"],
            required=["Required", "Optional", "Required", "Optional"],
        )
        master = tmp_path / "Catalog.xlsx"
        generate_catalog(
            release="TESTA", baselines_dir=testa_dir,
            master_path=master, timeout=60,
        )
        generate_catalog(
            release="TESTB", baselines_dir=testb_dir,
            master_path=master, timeout=60,
        )
        wb = load_workbook(master)
        assert "TESTA" in wb.sheetnames
        assert "TESTB" in wb.sheetnames
        drift_ws = wb["Drift"]
        drift = [[c.value for c in row] for row in drift_ws.iter_rows(min_row=2)]
        change_types = {r[-1] for r in drift}
        assert "RENAMED" in change_types
        assert "LENGTH_CHANGED" in change_types
        assert "REQUIRED_CHANGED" in change_types
        assert "ADDED" in change_types

    def test_end_to_end_file_error_in_issues(self, tmp_path):
        release_dir = tmp_path / "baselines" / "TESTA" / "originals"
        release_dir.mkdir(parents=True)
        (release_dir / "Broken.xlsm").write_bytes(b"not a real xlsx file")
        master = tmp_path / "Catalog.xlsx"
        generate_catalog(
            release="TESTA", baselines_dir=release_dir,
            master_path=master, timeout=60,
        )
        wb = load_workbook(master)
        ws = wb["Issues"]
        issue_rows = [[c.value for c in row] for row in ws.iter_rows(min_row=2)]
        assert any(r[3] == "FILE_ERROR" and r[1] == "Broken" for r in issue_rows)

    def test_end_to_end_idempotent(self, tmp_path):
        release_dir = tmp_path / "baselines" / "TESTA" / "originals"
        release_dir.mkdir(parents=True)
        _make_rich_xlsm(
            release_dir / "Fake.xlsm",
            tab_name="MY_TAB",
            labels=["Col A"],
            types=["VARCHAR2(50)"],
            techs=["COL_A"],
            required=["Required"],
        )
        master = tmp_path / "Catalog.xlsx"
        generate_catalog(release="TESTA", baselines_dir=release_dir,
                         master_path=master, timeout=60)
        wb1 = load_workbook(master)
        snap1 = {sn: [[c.value for c in row] for row in wb1[sn].iter_rows()]
                 for sn in wb1.sheetnames}
        generate_catalog(release="TESTA", baselines_dir=release_dir,
                         master_path=master, timeout=60)
        wb2 = load_workbook(master)
        snap2 = {sn: [[c.value for c in row] for row in wb2[sn].iter_rows()]
                 for sn in wb2.sheetnames}
        assert snap1 == snap2
