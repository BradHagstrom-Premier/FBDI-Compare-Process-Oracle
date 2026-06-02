import logging

from fbdi.applaud_appmap import (
    derive_prefix, AppMapRow, derive_appmap,
    write_appmap_workbook, load_appmap_workbook, merge_appmap,
)


def test_derive_prefix_from_parenthetical():
    prefix, fallback = derive_prefix("T_BANKS_BRANCHES (T32)", ["T32COUNTRY", "T32BANK_NAME"])
    assert prefix == "T32" and fallback is False


def test_derive_prefix_falls_back_to_lcp_and_logs(caplog):
    with caplog.at_level(logging.WARNING):
        prefix, fallback = derive_prefix("O_BANKS", ["O33BANK_NAME", "O33BRANCH_NUMBER"])
    assert prefix == "O33" and fallback is True
    assert any("fallback" in r.message.lower() for r in caplog.records)


def test_derive_prefix_none_when_no_columns_and_no_parenthetical():
    prefix, fallback = derive_prefix("WEIRD_TABLE", [])
    assert prefix is None and fallback is True


def test_derive_prefix_fallback_ignores_audit_fields():
    # @-audit fields must not skew the longest-common-prefix derivation.
    prefix, fallback = derive_prefix(
        "O_BANKS", ["O33BANK_NAME", "@O33LEGACY_FIELD1", "O33BRANCH_NUMBER"])
    assert prefix == "O33" and fallback is True


def test_derive_appmap_resolves_if_and_ef_in_order():
    applications = {
        "I_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES", "description": "",
            "steps": [{"order": 1, "func_type": "IF", "func_name": "I_T_BANKS_BRANCHES"}]},
        "X_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES", "description": "FBDI Fields",
            "steps": [{"order": 1, "func_type": "EF", "func_name": "T_BANKS_BRANCHES"},
                      {"order": 2, "func_type": "EF", "func_name": "X_T_BANKS_BRANCHES_VAL"}]},
        "CQ_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES", "description": "",
            "steps": [{"order": 1, "func_type": "CS", "func_name": "CS_REQ"}]},
        "X_T_OTHER": {"dbid": "T_OTHER", "description": "",
            "steps": [{"order": 1, "func_type": "EF", "func_name": "T_OTHER"}]},
    }
    rows = derive_appmap(applications, {"T_BANKS_BRANCHES"})
    assert len(rows) == 1
    row = rows[0]
    assert row.target_table == "T_BANKS_BRANCHES"
    assert row.import_files == ["I_T_BANKS_BRANCHES"]
    assert row.export_files == ["T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES_VAL"]
    assert set(row.source_applications) == {"I_T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES"}
    assert row.origin == "derived"


def test_derive_appmap_table_with_no_apps_yields_empty_row():
    rows = derive_appmap({}, {"T_LONELY"})
    assert rows[0].target_table == "T_LONELY"
    assert rows[0].import_files == [] and rows[0].export_files == []


def test_appmap_workbook_roundtrip(tmp_path):
    rows = [AppMapRow("T_BANKS_BRANCHES", ["I_T_BANKS_BRANCHES"],
                      ["T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES_VAL"],
                      ["I_T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES"], "derived")]
    path = tmp_path / "appmap.xlsx"
    write_appmap_workbook(rows, path)
    loaded = load_appmap_workbook(path)
    assert loaded["T_BANKS_BRANCHES"].import_files == ["I_T_BANKS_BRANCHES"]
    assert loaded["T_BANKS_BRANCHES"].export_files == ["T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES_VAL"]


def test_merge_keeps_confirmed_and_adds_new_derived():
    confirmed = {"T_A": AppMapRow("T_A", ["I_HAND_EDITED"], [], ["X"], "confirmed")}
    derived = [
        AppMapRow("T_A", ["I_T_A_AUTO"], ["E_T_A"], ["X"], "derived"),  # must NOT override confirmed
        AppMapRow("T_B", ["I_T_B"], [], ["X"], "derived"),              # new -> added
    ]
    merged = merge_appmap(derived, confirmed)
    by = {r.target_table: r for r in merged}
    assert by["T_A"].import_files == ["I_HAND_EDITED"] and by["T_A"].origin == "confirmed"
    assert by["T_B"].import_files == ["I_T_B"] and by["T_B"].origin == "derived"
