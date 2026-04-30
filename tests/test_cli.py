"""Tests for fbdi.cli — CLI helpers."""

import pytest
from pathlib import Path
from fbdi.cli import _resolve_dir


class TestResolveDir:
    def test_existing_directory_passes_through(self, tmp_path):
        """A path that is already a directory is returned unchanged."""
        assert _resolve_dir(tmp_path) == tmp_path

    def test_release_label_resolves_to_originals(self, tmp_path, monkeypatch):
        """A non-directory path like '26A' resolves to baselines/26A/originals/."""
        baselines = tmp_path / "baselines" / "26A" / "originals"
        baselines.mkdir(parents=True)
        monkeypatch.chdir(tmp_path)
        result = _resolve_dir(Path("26A"))
        assert result == Path("baselines") / "26A" / "originals"
        assert result.is_dir()

    def test_nonexistent_path_passes_through(self, tmp_path, monkeypatch):
        """A path that doesn't exist and has no baselines match passes through."""
        monkeypatch.chdir(tmp_path)
        result = _resolve_dir(Path("nonexistent"))
        # Should return the original path for downstream error handling
        assert result == Path("nonexistent")

    def test_explicit_originals_path_passes_through(self, tmp_path):
        """An explicit path to originals/ is returned unchanged."""
        originals = tmp_path / "baselines" / "26A" / "originals"
        originals.mkdir(parents=True)
        assert _resolve_dir(originals) == originals


class TestCatalogCLI:
    def test_catalog_cli_requires_release(self, tmp_path, capsys):
        from fbdi.cli import main
        with pytest.raises(SystemExit):
            main(["catalog"])

    def test_catalog_cli_missing_baselines_errors(self, tmp_path, capsys):
        from fbdi.cli import main
        with pytest.raises(SystemExit):
            main([
                "catalog", "--release", "99Z",
                "--baselines-dir", str(tmp_path / "does-not-exist"),
                "--master", str(tmp_path / "M.xlsx"),
            ])
        captured = capsys.readouterr()
        assert "not found" in captured.out.lower()

    def test_catalog_cli_end_to_end(self, tmp_path):
        """Build a tiny release dir + run catalog CLI + verify file written."""
        from openpyxl import Workbook
        from fbdi.cli import main

        baselines = tmp_path / "baselines" / "TESTZ" / "originals"
        baselines.mkdir(parents=True)
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet("MY_TAB")
        # Thin tab — need MIN_CELLS=2 for header detection
        ws.cell(row=4, column=1, value="*Only Field")
        ws.cell(row=4, column=2, value="Second Field")
        wb.save(baselines / "Tpl.xlsm")

        master = tmp_path / "Catalog.xlsx"
        main([
            "catalog", "--release", "TESTZ",
            "--baselines-dir", str(baselines),
            "--master", str(master),
            "--timeout", "30",
        ])
        assert master.exists()


def test_populate_module_subcommand_invocation(tmp_path, monkeypatch, capsys):
    """`python -m fbdi populate-module` invokes populate_module_column with the
    right args and prints the summary."""
    import fbdi.cli as cli_mod
    from openpyxl import Workbook

    # Build minimal artifacts in tmp_path
    (tmp_path / "baselines" / "26a").mkdir(parents=True)
    (tmp_path / "baselines" / "26b").mkdir(parents=True)
    (tmp_path / "baselines" / "26a" / "file_modules.json").write_text(
        '{"AutoInvoiceImportTemplate.xlsm": "Financials"}'
    )
    (tmp_path / "baselines" / "26b" / "file_modules.json").write_text(
        '{"AutoInvoiceImportTemplate.xlsm": "Financials"}'
    )

    mapping_path = tmp_path / "mapping.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "FBDI Mapping"
    headers = ["FBDI Template", "FBDI Tab", "Applaud Table", "Prefix",
               "Status", "Module", "In Base System?"]
    for c_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c_idx, value=h)
    ws.cell(row=2, column=1, value="AutoInvoiceImportTemplate")
    ws.cell(row=2, column=2, value="RA_TAB")
    wb.save(mapping_path)
    wb.close()

    monkeypatch.chdir(tmp_path)
    cli_mod.main(["populate-module", "--new", "26b", "--old", "26a",
                  "--mapping", str(mapping_path)])

    out = capsys.readouterr().out
    assert "populated" in out.lower()
    assert "1" in out  # one row populated
