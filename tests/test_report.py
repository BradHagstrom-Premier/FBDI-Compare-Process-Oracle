"""Tests for fbdi.report — Applaud-free release-change view-model + loaders."""

import json
from pathlib import Path

import jinja2
from openpyxl import Workbook

from fbdi.report import (
    ChangeRow,
    FileSection,
    ReportContext,
    build_report_context,
    load_catalog_release,
    load_file_modules,
    _oracle_type_str,
)
from fbdi.align import AlignedField


def _render_report(ctx, print_mode=False):
    template_dir = Path(__file__).parent.parent / "fbdi" / "templates"
    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(template_dir),
        autoescape=jinja2.select_autoescape(["html", "j2"]),
        trim_blocks=True, lstrip_blocks=True,
    )
    return env.get_template("report.html.j2").render(ctx=ctx, print_mode=print_mode)


def _aligned(position, label, technical, data_type=None, length=None, required=None, scale=None):
    return AlignedField(position=position, label=label, technical=technical,
                        data_type=data_type, length=length, scale=scale, required=required)


class TestScope:
    def test_all_changed_tabs_included_without_mapping(self):
        catalog_old = {("F1", "T"): [_aligned(1, "A", "A_F")]}
        catalog_new = {("F1", "T"): [_aligned(1, "A", "A_F"), _aligned(2, "B", "B_F")]}
        ctx = build_report_context(catalog_old, catalog_new, {"F1": "Financials"}, "26A", "26B")
        assert len(ctx.file_sections) == 1
        assert ctx.file_sections[0].module == "Financials"

    def test_unchanged_tab_excluded(self):
        rows = [_aligned(1, "A", "A_F")]
        ctx = build_report_context({("F", "T"): rows}, {("F", "T"): rows}, {}, "26A", "26B")
        assert ctx.file_sections == []

    def test_unknown_file_groups_unclassified(self):
        catalog_old = {("Mystery", "T"): []}
        catalog_new = {("Mystery", "T"): [_aligned(1, "A", "A_F")]}
        ctx = build_report_context(catalog_old, catalog_new, {}, "26A", "26B")
        assert ctx.file_sections[0].module == "Unclassified"


class TestFieldName:
    def test_uses_technical_when_present(self):
        ctx = build_report_context(
            {("F", "T"): []},
            {("F", "T"): [_aligned(1, "Some Label", "FIELD_NAME", "VARCHAR2", 30, True)]},
            {}, "26A", "26B",
        )
        added = ctx.file_sections[0].changes_by_type["ADDED"]
        assert added[0].field_name == "FIELD_NAME"

    def test_falls_back_to_label_when_technical_none(self):
        ctx = build_report_context(
            {("F", "T"): []},
            {("F", "T"): [_aligned(1, "Landed Cost Enabled", None, None, None, True)]},
            {}, "26A", "26B",
        )
        added = ctx.file_sections[0].changes_by_type["ADDED"]
        assert added[0].field_name == "Landed Cost Enabled"


class TestTotals:
    def test_totals_sum_adds_across_sections(self):
        catalog_old = {("F1", "T1"): [], ("F2", "T2"): []}
        catalog_new = {
            ("F1", "T1"): [_aligned(1, "A", "A_F"), _aligned(2, "B", "B_F")],
            ("F2", "T2"): [_aligned(1, "C", "C_F")],
        }
        ctx = build_report_context(catalog_old, catalog_new, {}, "26A", "26B")
        assert ctx.totals.tabs == 2
        assert ctx.totals.add == 3


class TestOracleTypeStr:
    def test_char_unit_preserved(self):
        f = AlignedField(position=1, label="L", technical="T", data_type="VARCHAR2",
                         length=30, scale=None, required=None, data_type_raw="VARCHAR2(30 CHAR)")
        assert _oracle_type_str(f) == "VARCHAR2(30 CHAR)"

    def test_reconstructed_when_raw_absent(self):
        f = AlignedField(position=1, label="L", technical="T", data_type="VARCHAR2",
                         length=30, scale=None, required=None)
        assert _oracle_type_str(f) == "VARCHAR2(30)"


class TestLoaders:
    def test_load_catalog_release_groups_by_file_and_tab(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "26B"
        ws.append(["release", "file_name", "tab_name", "position", "column_label",
                   "column_technical", "data_type", "length", "scale", "data_type_raw", "required"])
        ws.append(["26B", "F1", "T1", 1, "Lab", "TECH", "VARCHAR2", 30, None, "VARCHAR2(30)", "TRUE"])
        ws.append(["26B", "F1", "T1", 2, "Lab2", "TECH2", "NUMBER", 18, None, "NUMBER(18)", "FALSE"])
        path = tmp_path / "cat.xlsx"
        wb.save(path)
        result = load_catalog_release(path, "26B")
        assert len(result[("F1", "T1")]) == 2
        assert result[("F1", "T1")][0].required is True

    def test_load_file_modules_reads_json(self, tmp_path, monkeypatch):
        (tmp_path / "baselines" / "26b").mkdir(parents=True)
        (tmp_path / "baselines" / "26b" / "file_modules.json").write_text(
            json.dumps({"AutoInvoiceImportTemplate.xlsm": "Financials"}))
        monkeypatch.chdir(tmp_path)
        assert load_file_modules("26B") == {"AutoInvoiceImportTemplate.xlsm": "Financials"}

    def test_load_file_modules_missing_returns_empty(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        assert load_file_modules("99Z") == {}


class TestRender:
    def test_renders_title_and_field_without_applaud(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Label", "MYFIELD", "VARCHAR2", 30, required=None)]}
        ctx = build_report_context(catalog_old, catalog_new, {"F": "Financials"}, "26A", "26B")
        html = _render_report(ctx)
        assert "FBDI Release Change Report" in html
        assert "MYFIELD" in html
        assert "applaud" not in html.lower()
        assert 'class="nil">&middot;</span>' in html  # required=None renders a muted nil dot

    def test_empty_context_renders_no_changes_message(self):
        ctx = build_report_context({}, {}, {}, "26A", "26B")
        html = _render_report(ctx)
        assert "No field-level changes detected" in html

    def test_print_mode_drops_toolbar_and_script(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Label", "MYFIELD", "VARCHAR2", 30)]}
        ctx = build_report_context(catalog_old, catalog_new, {"F": "Financials"}, "26A", "26B")
        screen = _render_report(ctx, print_mode=False)
        assert 'class="toolbar"' in screen and "<script>" in screen
        printed = _render_report(ctx, print_mode=True)
        assert 'class="toolbar"' not in printed and "<script>" not in printed
        assert "MYFIELD" in printed  # content is fully present in the print render
        # Each template's impact strip rides in the (print) navy chip so "am I
        # impacted, and how much?" is answerable without reading the field rows.
        assert 'class="file__mini"' in printed  # strip markup emitted
        assert '<span class="m-add">+1</span>' in printed  # ...with the add count
        assert ".file__mini{ display:flex; }" in printed  # ...and un-hidden in print


class TestGenerateReportModuleGrouping:
    """Regression: catalog file_name is extension-less but file_modules.json keys
    carry .xlsm — generate_report must reconcile them so files aren't all
    grouped under 'Unclassified'."""

    def _write_catalog(self, path):
        wb = Workbook()
        wb.remove(wb.active)
        header = ["release", "file_name", "tab_name", "position", "column_label",
                  "column_technical", "data_type", "length", "scale", "data_type_raw", "required"]
        # 26A: one field; 26B: two fields (so the tab shows an ADDED change).
        ws_a = wb.create_sheet("26A")
        ws_a.append(header)
        ws_a.append(["26A", "AutoInvoiceImportTemplate", "RA", 1, "A", "A_F", "VARCHAR2", 30, None, "VARCHAR2(30)", "TRUE"])
        ws_b = wb.create_sheet("26B")
        ws_b.append(header)
        ws_b.append(["26B", "AutoInvoiceImportTemplate", "RA", 1, "A", "A_F", "VARCHAR2", 30, None, "VARCHAR2(30)", "TRUE"])
        ws_b.append(["26B", "AutoInvoiceImportTemplate", "RA", 2, "B", "B_F", "VARCHAR2", 30, None, "VARCHAR2(30)", "FALSE"])
        wb.save(path)
        wb.close()

    def test_module_resolved_from_extensioned_json_keys(self, tmp_path, monkeypatch):
        from fbdi.report import generate_report
        catalog = tmp_path / "cat.xlsx"
        self._write_catalog(catalog)
        (tmp_path / "baselines" / "26b").mkdir(parents=True)
        (tmp_path / "baselines" / "26b" / "file_modules.json").write_text(
            json.dumps({"AutoInvoiceImportTemplate.xlsm": "Financials"}))
        monkeypatch.chdir(tmp_path)

        html_path, pdf_path = generate_report(catalog, "26A", "26B", tmp_path)
        assert pdf_path is None
        html = html_path.read_text(encoding="utf-8")
        assert "<h2>Financials</h2>" in html
        assert "Unclassified" not in html


class TestDashboardAggregates:
    def test_module_totals_roll_up_per_module(self):
        catalog_old = {("F1", "T"): [], ("F2", "T"): []}
        catalog_new = {
            ("F1", "T"): [_aligned(1, "A", "A_F"), _aligned(2, "B", "B_F")],  # +2, Financials
            ("F2", "T"): [_aligned(1, "C", "C_F")],                           # +1, Supply Chain
        }
        ctx = build_report_context(
            catalog_old, catalog_new,
            {"F1": "Financials", "F2": "Supply Chain"}, "26A", "26B",
        )
        mods = {mt.module: mt for mt in ctx.module_totals}
        assert mods["Financials"].add == 2
        assert mods["Financials"].files == 1 and mods["Financials"].tabs == 1
        assert mods["Supply Chain"].add == 1
        # The dashboard must reconcile with the cover totals — never disagree.
        assert sum(mt.add for mt in ctx.module_totals) == ctx.totals.add
        assert sum(mt.shift for mt in ctx.module_totals) == ctx.totals.shift

    def test_top_files_sorted_by_high_signal_desc(self):
        catalog_old = {("Big", "T"): [], ("Small", "T"): []}
        catalog_new = {
            ("Big", "T"): [_aligned(i, f"L{i}", f"F{i}") for i in range(1, 6)],  # +5
            ("Small", "T"): [_aligned(1, "L", "F")],                             # +1
        }
        ctx = build_report_context(catalog_old, catalog_new, {}, "26A", "26B")
        assert [f.file for f in ctx.top_files] == ["Big", "Small"]
        assert ctx.top_files[0].high_signal == 5

    def test_dashboard_rendered_in_both_surfaces(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Label", "MYFIELD", "VARCHAR2", 30)]}
        ctx = build_report_context(catalog_old, catalog_new, {"F": "Financials"}, "26A", "26B")
        assert "Summary by module" in _render_report(ctx, print_mode=False)
        assert "Summary by module" in _render_report(ctx, print_mode=True)


class TestShiftOnlyRollup:
    """Condensed PDF: tabs whose only change is a position shift collapse into a
    single 'Position shifts' rollup in print mode, while the interactive HTML
    keeps them as full, browsable file cards (two surfaces, two truths)."""

    def _row(self, **kw):
        base = dict(change_type="ADDED", field_name="F", label="L", oracle_type_str="",
                    old_position=None, new_position=None, required=None, axes=(), sub_kinds=())
        base.update(kw)
        return ChangeRow(**base)

    def _ctx(self):
        hi = FileSection(
            file="HiFile", tab="T1", module="SCM",
            changes_by_type={"ADDED": [self._row(field_name="NEW_F", new_position=3)]},
        )
        sh = FileSection(
            file="ShiftFile", tab="T2", module="SCM",
            changes_by_type={"SHIFTED": [self._row(
                change_type="SHIFTED", field_name="MOVED", old_position=2,
                new_position=5, axes=("position",))]},
            shift_summary="1 field shifted from positions 2-2 to 5-5.",
            is_shift_only=True,
        )
        return ReportContext(old_release="26A", new_release="26B",
                             generated_date="2026-01-01", file_sections=[hi, sh])

    def test_print_collapses_shift_only_into_rollup(self):
        printed = _render_report(self._ctx(), print_mode=True)
        assert "Position shifts" in printed
        # Shift-only tab is NOT rendered as a detail card, but IS in the rollup.
        assert 'data-file="ShiftFile"' not in printed
        assert "ShiftFile" in printed
        # High-signal tab is still rendered in full.
        assert 'data-file="HiFile"' in printed

    def test_screen_keeps_shift_only_as_full_card(self):
        screen = _render_report(self._ctx(), print_mode=False)
        assert 'data-file="ShiftFile"' in screen   # normal, browsable card
        assert "Position shifts" not in screen      # no print-only rollup on screen


class TestSelfContained:
    """The generated HTML must be one portable file: Definian tokens + Aptos
    fonts inlined, and zero external CSS/font/JS references (so it renders the
    same in a browser, in email, and through weasyprint)."""

    def _generate(self, tmp_path, monkeypatch):
        wb = Workbook()
        wb.remove(wb.active)
        header = ["release", "file_name", "tab_name", "position", "column_label",
                  "column_technical", "data_type", "length", "scale", "data_type_raw", "required"]
        ws_a = wb.create_sheet("26A")
        ws_a.append(header)
        ws_a.append(["26A", "AutoInvoiceImportTemplate", "RA", 1, "A", "A_F", "VARCHAR2", 30, None, "VARCHAR2(30)", "TRUE"])
        ws_b = wb.create_sheet("26B")
        ws_b.append(header)
        ws_b.append(["26B", "AutoInvoiceImportTemplate", "RA", 1, "A", "A_F", "VARCHAR2", 30, None, "VARCHAR2(30)", "TRUE"])
        ws_b.append(["26B", "AutoInvoiceImportTemplate", "RA", 2, "B", "B_F", "VARCHAR2", 30, None, "VARCHAR2(30)", "FALSE"])
        catalog = tmp_path / "cat.xlsx"
        wb.save(catalog)
        wb.close()
        monkeypatch.chdir(tmp_path)
        from fbdi.report import generate_report
        html_path, _ = generate_report(catalog, "26A", "26B", tmp_path)
        return html_path.read_text(encoding="utf-8")

    def test_fonts_and_tokens_inlined(self, tmp_path, monkeypatch):
        html = self._generate(tmp_path, monkeypatch)
        assert "<style>" in html
        assert "@font-face" in html
        assert "data:font/ttf;base64," in html   # Aptos embedded, not linked
        assert "--def-blue" in html               # Definian tokens present

    def test_no_external_references(self, tmp_path, monkeypatch):
        # Check for real external-reference syntax only. A blanket "http" scan
        # would false-positive on the base64 font blobs (case-insensitive
        # substrings), so we look for the exact patterns that fetch a resource;
        # none of "(", quote chars, or "<" appear inside base64.
        html = self._generate(tmp_path, monkeypatch).lower()
        assert "<link" not in html
        assert 'src="http' not in html
        assert "src='http" not in html
        assert "url(http" not in html
        assert "@import url(" not in html
