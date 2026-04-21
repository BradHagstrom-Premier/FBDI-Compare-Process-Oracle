import json
import tempfile
from pathlib import Path
from openpyxl import Workbook

from fbdi.audit import (
    SnapshotField, SnapshotKeySeq, SnapshotTable, ApplaudSnapshot,
    Candidate, EvidenceBundle, PriorRow, AuditRow,
    load_snapshot, load_catalog, load_prior_mapping, CatalogIndex,
)

def test_data_classes_importable():
    field = SnapshotField(
        name="TA4INVOICE_ID",
        bare_name="INVOICE_ID",
        is_legacy_tracking=False,
        data_type="N",
        length=15,
    )
    assert field.bare_name == "INVOICE_ID"
    assert not field.is_legacy_tracking

def test_legacy_tracking_field():
    field = SnapshotField(
        name="@TA4SITE",
        bare_name="SITE",
        is_legacy_tracking=True,
        data_type="X",
        length=10,
    )
    assert field.is_legacy_tracking
    assert field.bare_name == "SITE"


def _make_snapshot_dict(tables=None, missing=None) -> dict:
    return {
        "mdb_path": "C:/test/AP0STE.mdb",
        "extracted_at": "2026-04-21T12:00:00Z",
        "extractor_version": "1",
        "tables": tables or [],
        "missing_tables": missing or [],
    }


def test_load_snapshot_basic(tmp_path):
    snap_data = _make_snapshot_dict(tables=[{
        "name": "T_RA_INTERFACE_LINES_ALL",
        "prefix": "TA4",
        "description": "T_RA_INTERFACE_LINES_ALL (TA4)",
        "type": "1",
        "key_sequences": [{"seq": "1", "keys": ["TA4INVOICE_ID"]}],
        "fields": [
            {"name": "TA4INVOICE_ID", "bare_name": "INVOICE_ID",
             "is_legacy_tracking": False, "data_type": "N", "length": 15},
            {"name": "@TA4SITE", "bare_name": "SITE",
             "is_legacy_tracking": True, "data_type": "X", "length": 10},
        ],
    }])
    snap_file = tmp_path / "applaud_snapshot.json"
    snap_file.write_text(json.dumps(snap_data))
    snap = load_snapshot(snap_file)
    assert isinstance(snap, ApplaudSnapshot)
    assert len(snap.tables) == 1
    t = snap.tables[0]
    assert t.name == "T_RA_INTERFACE_LINES_ALL"
    assert t.prefix == "TA4"
    assert len(t.fields) == 2
    biz = t.business_fields()
    assert len(biz) == 1
    assert biz[0].bare_name == "INVOICE_ID"


def test_load_snapshot_missing_file(tmp_path):
    import pytest
    with pytest.raises(FileNotFoundError):
        load_snapshot(tmp_path / "missing.json")


def _make_catalog_xlsx(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "26B"
    ws.append(["release", "file_name", "tab_name", "position",
                "column_label", "column_technical",
                "data_type", "length", "scale", "data_type_raw", "required"])
    ws.append(["26B", "AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL",
                1, "Invoice ID", "INVOICE_ID", "N", 15, None, "NUMBER(15)", "TRUE"])
    ws.append(["26B", "AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL",
                2, "Batch Source Name", "BATCH_SOURCE_NAME", "X", 50, None, "VARCHAR2(50)", "FALSE"])
    path = tmp_path / "FBDI_Master_Catalog.xlsx"
    wb.save(path)
    return path


def test_load_catalog_basic(tmp_path):
    path = _make_catalog_xlsx(tmp_path)
    index = load_catalog(path, release="26B")
    key = ("AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL")
    assert key in index
    assert "INVOICE_ID" in index[key]
    assert "BATCH_SOURCE_NAME" in index[key]


def test_load_catalog_missing_release(tmp_path):
    import pytest
    path = _make_catalog_xlsx(tmp_path)
    with pytest.raises(ValueError, match="25D"):
        load_catalog(path, release="25D")


def _make_prior_mapping_xlsx(tmp_path: Path) -> Path:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "FBDI Mapping"
    ws2 = wb.create_sheet("Applaud Tables")
    ws2.append(["#", "applaud_table", "status", "prefix",
                 "fbdi_template_mappings", "module", "notes"])
    ws2.append([1, "T_RA_INTERFACE_LINES_ALL", "YES", "TA4",
                 "AutoInvoiceImportTemplate / RA_INTERFACE_LINES_ALL", "Financials", ""])
    ws2.append([2, "T_GHOST_TABLE", "UNMAPPED", "", "", "HR", "no match found"])
    path = tmp_path / "fbdi_applaud_mapping.xlsx"
    wb.save(path)
    return path


def test_load_prior_mapping_basic(tmp_path):
    path = _make_prior_mapping_xlsx(tmp_path)
    mapping = load_prior_mapping(path)
    assert "T_RA_INTERFACE_LINES_ALL" in mapping
    row = mapping["T_RA_INTERFACE_LINES_ALL"]
    assert row.prior_status == "YES"
    assert row.prefix == "TA4"
    assert "AutoInvoiceImportTemplate" in row.mapping_text
    assert "T_GHOST_TABLE" in mapping
    assert mapping["T_GHOST_TABLE"].prior_status == "UNMAPPED"
