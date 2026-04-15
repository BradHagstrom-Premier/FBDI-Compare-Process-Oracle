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
