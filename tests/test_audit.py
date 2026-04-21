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


from fbdi.audit import extract_prefix, derive_bare_name

def test_extract_prefix_standard():
    assert extract_prefix("T_RA_INTERFACE_LINES_ALL (TA4)") == "TA4"

def test_extract_prefix_alphanumeric():
    assert extract_prefix("T_EGP_COMPONENTS_INTERFACE (T91)") == "T91"

def test_extract_prefix_no_parens():
    assert extract_prefix("T_GHOST_TABLE") is None

def test_derive_bare_name_regular():
    bare, is_legacy = derive_bare_name("TA4INVOICE_ID", "TA4")
    assert bare == "INVOICE_ID"
    assert not is_legacy

def test_derive_bare_name_at_prefix():
    bare, is_legacy = derive_bare_name("@TA4SITE", "TA4")
    assert bare == "SITE"
    assert is_legacy

def test_derive_bare_name_no_prefix_match():
    bare, is_legacy = derive_bare_name("SOMETHING_ELSE", "TA4")
    assert bare == "SOMETHING_ELSE"
    assert not is_legacy


from fbdi.audit import (
    compute_name_alignment, compute_key_coverage,
    compute_column_overlap, check_prefix_conformance,
    SnapshotField,
)

# --- name_alignment ---

def test_name_alignment_exact():
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "RA_INTERFACE_LINES_ALL") == "EXACT"

def test_name_alignment_partial_strip_all():
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "RA_INTERFACE_LINES") == "PARTIAL"

def test_name_alignment_partial_strip_interface():
    assert compute_name_alignment("T_RCV_HEADERS_INTERFACE", "RCV_HEADERS") == "PARTIAL"

def test_name_alignment_none():
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "TOTALLY_DIFFERENT_TAB") == "NONE"

def test_name_alignment_case_insensitive():
    assert compute_name_alignment("T_RA_INTERFACE_LINES_ALL", "ra_interface_lines_all") == "EXACT"

# --- key_coverage ---

def test_key_coverage_full():
    keys = {"INVOICE_ID", "LINE_NUMBER"}
    fbdi_cols = {"INVOICE_ID", "LINE_NUMBER", "BATCH_SOURCE_NAME"}
    assert compute_key_coverage(keys, fbdi_cols) == 1.0

def test_key_coverage_partial():
    keys = {"INVOICE_ID", "LINE_NUMBER"}
    fbdi_cols = {"INVOICE_ID", "BATCH_SOURCE_NAME"}
    assert compute_key_coverage(keys, fbdi_cols) == 0.5

def test_key_coverage_empty_keys():
    assert compute_key_coverage(set(), {"INVOICE_ID"}) == 0.0

# --- column_overlap ---

def _biz_field(bare_name: str) -> SnapshotField:
    return SnapshotField(name=f"TA4{bare_name}", bare_name=bare_name,
                         is_legacy_tracking=False, data_type="X", length=30)

def _legacy_field(bare_name: str) -> SnapshotField:
    return SnapshotField(name=f"@TA4{bare_name}", bare_name=bare_name,
                         is_legacy_tracking=True, data_type="X", length=10)

def test_column_overlap_excludes_legacy():
    fields = [_biz_field("INVOICE_ID"), _biz_field("LINE_NUM"), _legacy_field("SITE")]
    fbdi_cols = {"INVOICE_ID", "LINE_NUM"}
    assert compute_column_overlap(fields, fbdi_cols) == 1.0

def test_column_overlap_partial():
    fields = [_biz_field("INVOICE_ID"), _biz_field("LINE_NUM"), _biz_field("BATCH_NAME")]
    fbdi_cols = {"INVOICE_ID", "LINE_NUM"}
    assert abs(compute_column_overlap(fields, fbdi_cols) - 2/3) < 0.001

def test_column_overlap_all_legacy():
    fields = [_legacy_field("SITE"), _legacy_field("LEGACY_HEADER")]
    fbdi_cols = {"INVOICE_ID"}
    assert compute_column_overlap(fields, fbdi_cols) == 0.0

def test_column_overlap_case_insensitive():
    fields = [_biz_field("INVOICE_ID")]
    fbdi_cols = {"invoice_id"}
    assert compute_column_overlap(fields, fbdi_cols) == 1.0

# --- prefix conformance ---

def test_prefix_conformance_true():
    assert check_prefix_conformance("T_RA_INTERFACE_LINES_ALL", "TA4", "RA_INTERFACE_LINES_ALL") is True

def test_prefix_conformance_false():
    assert check_prefix_conformance("T_RA_INTERFACE_LINES_ALL", "TA4", "RA_INTERFACE_LINES") is False


# --- Pass 1 candidate index ---

from fbdi.audit import build_candidate_index, ApplaudSnapshot, SnapshotTable, SnapshotField, SnapshotKeySeq

def _make_snapshot_table(
    name: str, prefix: str, fields: list[SnapshotField], key_fields: list[str]
) -> SnapshotTable:
    return SnapshotTable(
        name=name, prefix=prefix,
        description=f"{name} ({prefix})", type="1",
        key_sequences=[SnapshotKeySeq(seq="1", keys=key_fields)],
        fields=fields,
    )

def _make_snap(*tables: SnapshotTable) -> ApplaudSnapshot:
    return ApplaudSnapshot(
        mdb_path="", extracted_at="", extractor_version="1",
        tables=list(tables), missing_tables=[],
    )


def test_pass1_exact_name_match_kept():
    fields = [SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)]
    table = _make_snapshot_table("T_RA_INTERFACE_LINES_ALL", "TA4", fields, ["TA4INVOICE_ID"])
    catalog = {("AutoInvoiceImportTemplate", "RA_INTERFACE_LINES_ALL"): {"INVOICE_ID"}}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    assert "T_RA_INTERFACE_LINES_ALL" in idx
    candidates = idx["T_RA_INTERFACE_LINES_ALL"]
    assert len(candidates) == 1
    c = candidates[0]
    assert c.fbdi_file == "AutoInvoiceImportTemplate"
    assert c.fbdi_tab == "RA_INTERFACE_LINES_ALL"
    assert c.name_alignment == "EXACT"


def test_pass1_below_threshold_dropped():
    fields = [SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)]
    table = _make_snapshot_table("T_RA_INTERFACE_LINES_ALL", "TA4", fields, ["TA4INVOICE_ID"])
    catalog = {("SomeTemplate", "UNRELATED_TAB"): {"UNRELATED_COL"}}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    assert idx.get("T_RA_INTERFACE_LINES_ALL", []) == []


def test_pass1_high_column_overlap_kept():
    fields = [SnapshotField(f"TA4F{i}", f"F{i}", False, "X", 10) for i in range(5)]
    table = _make_snapshot_table("T_SOME_TABLE", "TA4", fields, [])
    fbdi_cols = {f"F{i}" for i in range(4)}
    catalog = {("SomeTemplate", "UNRELATED_TAB"): fbdi_cols}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    candidates = idx.get("T_SOME_TABLE", [])
    assert len(candidates) == 1
    assert candidates[0].column_overlap >= 0.3


def test_pass1_legacy_fields_excluded_from_overlap():
    biz = SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)
    leg = SnapshotField("@TA4SITE", "SITE", True, "X", 10)
    table = _make_snapshot_table("T_RA", "TA4", [biz, leg], [])
    catalog = {("AnyTemplate", "RA"): {"INVOICE_ID"}}
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    candidates = idx.get("T_RA", [])
    assert candidates  # kept because name PARTIAL match or high overlap
    assert candidates[0].column_overlap == 1.0  # 1/1 biz field matched


def test_pass1_sorted_strongest_first():
    fields = [SnapshotField("TA4INVOICE_ID", "INVOICE_ID", False, "N", 15)]
    table = _make_snapshot_table("T_RA_INTERFACE_LINES_ALL", "TA4", fields, ["TA4INVOICE_ID"])
    catalog = {
        ("TemplateA", "RA_INTERFACE_LINES_ALL"): {"INVOICE_ID"},   # EXACT + key 1.0
        ("TemplateB", "RA_INTERFACE_LINES"): {"INVOICE_ID"},       # PARTIAL + key 1.0
    }
    snap = _make_snap(table)
    idx = build_candidate_index(snap, catalog)
    candidates = idx["T_RA_INTERFACE_LINES_ALL"]
    assert len(candidates) == 2
    assert candidates[0].name_alignment == "EXACT"
