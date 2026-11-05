"""openpyxl's read-time UserWarnings (e.g. "Data Validation extension is not
supported and will be removed") are noise for this read-only engine — nothing
here re-saves a workbook through openpyxl. They must not escape compare,
catalog or diagnose.

openpyxl emits them from its own modules; in read_only mode they fire lazily
while sheets are iterated, not just inside load_workbook. The tests simulate
both by emitting an openpyxl-attributed warning from load_workbook AND from
header detection (which runs mid-iteration).
"""

import warnings
from pathlib import Path

import pytest
from openpyxl import Workbook

from fbdi.utils import quiet_openpyxl

MSG = "Data Validation extension is not supported and will be removed"


def _openpyxl_warn():
    warnings.warn_explicit(MSG, UserWarning, "openpyxl/worksheet/_reader.py", 1,
                           module="openpyxl.worksheet._reader")


def _make_book(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Instructions"
    ws["A1"] = "Instructions"
    data = wb.create_sheet("DATA_TAB")
    for c, v in enumerate(["BATCH_ID", "ITEM_NUMBER", "ORG_CODE"], start=1):
        data.cell(row=4, column=c, value=v)
    wb.save(path)


def _patch_noisy(monkeypatch, module):
    real_load = module.load_workbook
    real_detect = module.detect_header_row

    def noisy_load(*a, **kw):
        _openpyxl_warn()
        return real_load(*a, **kw)

    def noisy_detect(*a, **kw):
        _openpyxl_warn()
        return real_detect(*a, **kw)

    monkeypatch.setattr(module, "load_workbook", noisy_load)
    monkeypatch.setattr(module, "detect_header_row", noisy_detect)


def _openpyxl_warnings(record):
    return [w for w in record if MSG in str(w.message)]


class TestQuietOpenpyxl:
    def test_suppresses_openpyxl_userwarning(self):
        with warnings.catch_warnings(record=True) as rec:
            warnings.simplefilter("always")
            with quiet_openpyxl():
                _openpyxl_warn()
        assert _openpyxl_warnings(rec) == []

    def test_leaves_other_warnings_alone(self):
        with warnings.catch_warnings(record=True) as rec:
            warnings.simplefilter("always")
            with quiet_openpyxl():
                warnings.warn("not from openpyxl", UserWarning)
        assert [str(w.message) for w in rec] == ["not from openpyxl"]

    def test_usable_as_decorator(self):
        @quiet_openpyxl()
        def f():
            _openpyxl_warn()
            return 1

        with warnings.catch_warnings(record=True) as rec:
            warnings.simplefilter("always")
            assert f() == 1
        assert _openpyxl_warnings(rec) == []


class TestCallSitesAreQuiet:
    def test_compare_fbdi_pair(self, tmp_path, monkeypatch):
        import fbdi.compare as compare
        old, new = tmp_path / "old.xlsx", tmp_path / "new.xlsx"
        _make_book(old)
        _make_book(new)
        _patch_noisy(monkeypatch, compare)
        with warnings.catch_warnings(record=True) as rec:
            warnings.simplefilter("always")
            compare.compare_fbdi_pair(old, new)
        assert _openpyxl_warnings(rec) == []

    def test_catalog_extract_file(self, tmp_path, monkeypatch):
        import fbdi.catalog as catalog
        p = tmp_path / "book.xlsx"
        _make_book(p)
        _patch_noisy(monkeypatch, catalog)
        with warnings.catch_warnings(record=True) as rec:
            warnings.simplefilter("always")
            rows, _ = catalog.extract_file(p, "26D")
        assert rows
        assert _openpyxl_warnings(rec) == []

    def test_diagnose_file(self, tmp_path, monkeypatch):
        import fbdi.diagnose as diagnose
        p = tmp_path / "book.xlsx"
        _make_book(p)
        _patch_noisy(monkeypatch, diagnose)
        with warnings.catch_warnings(record=True) as rec:
            warnings.simplefilter("always")
            diagnose.diagnose_file(p)
        assert _openpyxl_warnings(rec) == []
