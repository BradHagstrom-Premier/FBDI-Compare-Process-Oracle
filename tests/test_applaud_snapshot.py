import pytest

from fbdi.applaud_snapshot import (
    DataColumn, FileField, SnapshotTable, ApplaudSnapshot,
    assert_complete, build_file_fields, build_table, is_audit_field,
    SnapshotIncompleteError,
)


def test_snapshot_roundtrips_through_json(tmp_path):
    snap = ApplaudSnapshot(
        system="ORACLE_MASTER",
        mdb_path="X:/AP0STE.mdb",
        extracted_at="2026-06-02T00:00:00+00:00",
        extractor_version="1",
        tables={
            "T_BANKS_BRANCHES": SnapshotTable(
                name="T_BANKS_BRANCHES", prefix="T32", prefix_fallback=False,
                description="T_BANKS_BRANCHES (T32)", key_seqs=[["T32COUNTRY"]],
                columns=[DataColumn(ddid="T32BANK_NAME", bare="BANK_NAME",
                                    data_type="X", size=100, dec_places=None,
                                    odbc_name="BANK_NAME", row=2)],
            )
        },
        imports={"I_T_BANKS_BRANCHES": [FileField(row=1, ddid="T32COUNTRY",
                  bare="COUNTRY", pic="X(60)", input_type="C", column_header=None)]},
        exports={"T_BANKS_BRANCHES": [FileField(row=1, ddid="T32COUNTRY",
                  bare="COUNTRY", pic="X(60)", input_type=None, column_header="")]},
        applications={"I_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES",
                  "description": "", "steps": [{"order": 1, "func_type": "IF",
                  "func_name": "I_T_BANKS_BRANCHES"}]}},
    )
    path = tmp_path / "snap.json"
    snap.write(path)
    loaded = ApplaudSnapshot.load(path)
    assert loaded == snap
    assert loaded.tables["T_BANKS_BRANCHES"].columns[0].bare == "BANK_NAME"


def test_assert_complete_raises_on_truncation():
    rows = [{"Row": i} for i in range(100)]
    with pytest.raises(SnapshotIncompleteError) as exc:
        assert_complete("ImportDetail", "I_X", rows, expected_count=137)
    assert "I_X" in str(exc.value) and "100" in str(exc.value) and "137" in str(exc.value)


def test_assert_complete_passes_when_counts_match():
    rows = [{"Row": i} for i in range(23)]
    assert_complete("ImportDetail", "I_T_BANKS_BRANCHES", rows, expected_count=23)


def test_is_audit_field_detects_at_prefix():
    assert is_audit_field("@T32LEGACY_HEADER1") is True
    assert is_audit_field("T32BANK_NAME") is False


def test_build_file_fields_strips_prefix_orders_and_drops_audit_fields():
    raw = [
        {"Row": 2, "DDID": "T32BANK_NAME", "Pic": "X(100)", "InputType": "C"},
        {"Row": 1, "DDID": "T32COUNTRY", "Pic": "X(60)", "InputType": "C"},
        {"Row": 3, "DDID": "@T32DO_NOT_LOAD", "Pic": "X(1)", "InputType": "C"},  # audit field
    ]
    fields = build_file_fields(raw, prefix="T32", kind="IF")
    assert [f.bare for f in fields] == ["COUNTRY", "BANK_NAME"]   # @ field dropped
    assert fields[0].row == 1 and fields[0].input_type == "C"
    assert fields[0].column_header is None


def test_build_table_joins_datadictionary_type_and_drops_audit_fields():
    # DatabaseDetail carries blank type (real-data shape); DataDictionary has the type.
    raw_cols = [
        {"Row": 1, "DDID": "T32COUNTRY", "DataType": "", "Size": 0,
         "DecPlaces": 0, "ODBCName": ""},
        {"Row": 2, "DDID": "@T32DO_NOT_LOAD", "DataType": "", "Size": 0,
         "DecPlaces": 0, "ODBCName": ""},                          # audit field
    ]
    dd_by_ddid = {"T32COUNTRY": {"DataType": "X", "Size": 60, "DecPlaces": 0}}
    table = build_table("T_BANKS_BRANCHES", prefix="T32", prefix_fallback=False,
                        description="T_BANKS_BRANCHES (T32)", key_seqs=[["T32COUNTRY"]],
                        raw_columns=raw_cols, dd_by_ddid=dd_by_ddid)
    assert [c.bare for c in table.columns] == ["COUNTRY"]          # @ field dropped
    # type/size come from DataDictionary, NOT the blank DatabaseDetail row
    assert table.columns[0].data_type == "X" and table.columns[0].size == 60


def test_build_table_raises_when_business_ddid_missing_from_datadictionary():
    # A non-audit column with no DataDictionary entry means an incomplete DD slice;
    # build_table must fail loud, not emit an empty-typed DataColumn.
    # NB: T32MISSING carries the table prefix, so it is a genuine data element —
    # its absence from the DD slice signals truncation and must still fail loud.
    raw_cols = [{"Row": 1, "DDID": "T32MISSING", "DataType": "", "Size": 0,
                 "DecPlaces": 0, "ODBCName": ""}]
    with pytest.raises(SnapshotIncompleteError) as exc:
        build_table("T_BANKS_BRANCHES", prefix="T32", prefix_fallback=False,
                    description="T_BANKS_BRANCHES (T32)", key_seqs=[["T32COUNTRY"]],
                    raw_columns=raw_cols, dd_by_ddid={})
    assert "T32MISSING" in str(exc.value)


def test_build_table_excludes_non_prefix_phantom_column():
    # A column whose DDID lacks the table prefix (e.g. Applaud's X_PHANTOM system
    # field, "Phantom Run?") is not one of the table's data elements — within the
    # T_* family every real data element shares the table's TableId prefix. So it
    # is excluded (not failed-loud), while a prefix-matching DDID missing from the
    # DD slice still raises (truncation guard, asserted in the test above).
    raw_cols = [
        {"Row": 1, "DDID": "T91ASSEMBLY_ITEM_NUMBER", "DataType": "", "Size": 0,
         "DecPlaces": 0, "ODBCName": ""},
        {"Row": 126, "DDID": "X_PHANTOM", "DataType": "", "Size": 0,
         "DecPlaces": 0, "ODBCName": ""},   # phantom: no T91 prefix, not in DD slice
    ]
    dd_by_ddid = {"T91ASSEMBLY_ITEM_NUMBER": {"DataType": "X", "Size": 40, "DecPlaces": 0}}
    table = build_table("T_EGP_COMPONENTS_INTERFACE", prefix="T91", prefix_fallback=False,
                        description="T_EGP_COMPONENTS_INTF (T91)", key_seqs=[],
                        raw_columns=raw_cols, dd_by_ddid=dd_by_ddid)
    assert [c.bare for c in table.columns] == ["ASSEMBLY_ITEM_NUMBER"]   # X_PHANTOM excluded
