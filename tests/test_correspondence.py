from fbdi.correspondence import (
    ABBREVIATIONS, APPLAUD_NAME_CAP, MAX_SUFFIX_SLACK,
    expand_abbreviations, normalize_name, truncation_window, _split_trailing_digits,
)


def test_squash_collapses_all_underscores():
    # Audit §2.3: underscore position carries no information -> full squash both sides.
    assert normalize_name("REMIT_ADVICEDELIVERY_METHOD") == "REMITADVICEDELIVERYMETHOD"
    assert normalize_name("REMIT_ADVICEDELIVERYMETHOD") == "REMITADVICEDELIVERYMETHOD"


def test_normalize_strips_star_and_uppercases():
    assert normalize_name("Supplier_Name*") == "SUPPLIERNAME"


def test_strip_bool_suffix_full_forms():
    # Full _FLAG/_FLG/_F are stripped; truncated forms are left to names_correspond (Task 3).
    assert normalize_name("ALWAYS_TAKE_DISCOUNT_FLAG") == normalize_name("ALWAYS_TAKE_DISCOUNT")


def test_expand_abbreviations_is_token_wise_and_bidirectional_safe():
    # BU -> BUSINESSUNIT expansion on the Oracle side.
    assert expand_abbreviations("PROCUREMENT_BU") == "PROCUREMENT_BUSINESSUNIT"
    # A token not in the table passes through unchanged (idempotent).
    assert expand_abbreviations("ITEM_NUMBER") == "ITEM_NUMBER"
    # Already-expanded input is stable.
    assert expand_abbreviations("PROCUREMENT_BUSINESSUNIT") == "PROCUREMENT_BUSINESSUNIT"


def test_abbreviation_table_has_data_grounded_seed():
    # Audit §8.3 seed entries.
    for k in ("BU", "BUS", "DISC", "NUM", "DESCR", "DESC", "AMT", "INV", "COMP", "REFER"):
        assert k in ABBREVIATIONS


def test_truncation_window_is_30_minus_prefix():
    assert truncation_window("T09") == APPLAUD_NAME_CAP - 3
    assert truncation_window("TA1") == 27


def test_split_trailing_digits():
    assert _split_trailing_digits("TIMESTAMP10") == ("TIMESTAMP", "10")
    assert _split_trailing_digits("VENDOR_NAME") == ("VENDOR_NAME", "")


from fbdi.align import AlignedField
from fbdi.applaud_snapshot import DataColumn
from fbdi.correspondence import (
    FieldCorrespondence, names_correspond, score_candidate, TIER_BANDS,
    derive_table_correspondences,
)


def _col(ddid, bare, dt="X", size=100, dec=None, row=1):
    return DataColumn(ddid=ddid, bare=bare, data_type=dt, size=size,
                      dec_places=dec, odbc_name=None, row=row)


def _of(technical, dt="VARCHAR2", length=100, scale=None, position=1):
    return AlignedField(position=position, label=None, technical=technical,
                        data_type=dt, length=length, scale=scale, required=None)


# --- names_correspond ---

def test_right_truncation_within_window():
    # CONSUMPTION_ADVICE_LINE_NUMBER -> Applaud lost the final R (truncated NUMBER).
    win = 27
    assert names_correspond("CONSUMPTIONADVICELINENUMBER",
                            "CONSUMPTIONADVICELINENUMBE", applaud_bare_len=26, window=win)


def test_appended_then_truncated_suffix():
    # PROCUREMENT_BU -> expand -> PROCUREMENTBUSINESSUNIT; Applaud appended NAME, truncated to NAM.
    assert names_correspond("PROCUREMENTBUSINESSUNIT",
                            "PROCUREMENTBUSINESSUNITNAM", applaud_bare_len=27, window=27)


def test_digit_run_truncation():
    # Audit §1.1 named test case. Digits (10) must be equal; stems differ by the dropped P.
    assert names_correspond("GLOBALATTRIBUTETIMESTAMP10",
                            "GLOBALATTRIBUTETIMESTAM10", applaud_bare_len=25, window=27)


def test_truncated_bool_suffix_fla():
    # Pass-1 audit LOW #2: ALLOW_SUBSTITUTE_RECEIPTS -> Applaud appended FLAG, truncated to FLA.
    # Matches via the append path (append-delta 3 <= MAX_SUFFIX_SLACK).
    assert names_correspond("ALLOWSUBSTITUTERECEIPTS",
                            "ALLOWSUBSTITUTERECEIPTSFLA", applaud_bare_len=26, window=27)


def test_digit_run_must_be_equal():
    # TIMESTAMP10 must NOT match TIMESTAMP1 just because stems share a prefix.
    assert not names_correspond("GLOBALATTRIBUTETIMESTAMP10",
                                "GLOBALATTRIBUTETIMESTAMP1", applaud_bare_len=25, window=27)


def test_coincidental_short_prefix_does_not_match():
    # 'BANK' is a prefix of 'BANKACCOUNTNUMBER' but the delta is far past MAX_SUFFIX_SLACK
    # and Applaud was not truncated at the cap -> reject.
    assert not names_correspond("BANKACCOUNTNUMBER", "BANK", applaud_bare_len=4, window=27)


# --- type veto ---

def test_char_vs_numeric_vetoes_candidate():
    # Name matches but Oracle char vs Applaud numeric -> no candidate emitted.
    oracle = {"AMOUNT": _of("AMOUNT", dt="VARCHAR2", length=50)}
    cols = [_col("T01AMOUNT", "AMOUNT", dt="N", size=18, dec=2)]
    out = derive_table_correspondences("T_X", "T01", oracle, cols, decided=set())
    assert out == []


def test_u_column_not_vetoed():
    # Audit §1.2: U is char; an Oracle char field still matches a U column on a name divergence.
    oracle = {"VENDOR_NAME_NEW": _of("VENDOR_NAME_NEW", dt="VARCHAR2", length=100)}
    cols = [_col("T07VENDOR_NAMENEW", "VENDOR_NAMENEW", dt="U", size=100)]
    out = derive_table_correspondences("T_POZ", "T07", oracle, cols, decided=set())
    assert len(out) == 1 and out[0].applaud_bare == "VENDOR_NAMENEW"


def test_date_vs_char_does_not_veto():
    # Audit §1.2: Applaud stores TIMESTAMP as X (char); Oracle TIMESTAMP -> 'date'. No veto.
    oracle = {"GLOBAL_ATTRIBUTE_TIMESTAMP10": _of("GLOBAL_ATTRIBUTE_TIMESTAMP10",
                                                  dt="TIMESTAMP", length=None)}
    cols = [_col("T09GLOBAL_ATTRIBUTE_TIMESTAM10", "GLOBAL_ATTRIBUTE_TIMESTAM10",
                 dt="X", size=150)]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols, decided=set())
    assert len(out) == 1


# --- exact pre-pass + exclusions ---

def test_exact_matches_are_not_proposed():
    oracle = {"ITEM_NUMBER": _of("ITEM_NUMBER")}
    cols = [_col("T01ITEM_NUMBER", "ITEM_NUMBER")]
    assert derive_table_correspondences("T_X", "T01", oracle, cols, decided=set()) == []


def test_derivation_excludes_audit_and_nonprefix():
    # Audit §1.3/§1.4: @-fields and non-prefix working columns never enter the candidate pool.
    oracle = {"PROCUREMENT_BU": _of("PROCUREMENT_BU")}
    cols = [
        _col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM", size=25),
        _col("@T09LEGACY_AUDIT", "@T09LEGACY_AUDIT"),   # @-field (defensive)
        _col("X_PHANTOM", "X_PHANTOM"),                 # non-prefix working column
    ]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols, decided=set())
    bares = {c.applaud_bare for c in out}
    assert "PROCUREMENT_BUSINESSUNITNAM" in bares
    assert all(not b.startswith("@") and b != "X_PHANTOM" for b in bares)


def test_decided_pairs_are_skipped():
    oracle = {"PROCUREMENT_BU": _of("PROCUREMENT_BU")}
    cols = [_col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM", size=25)]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols,
                                       decided={("T_POZ", "PROCUREMENT_BU")})
    assert out == []


def test_bijection_one_to_one_per_table():
    # Pass-1 audit LOW #1: use two Oracle keys that BOTH genuinely match the one column,
    # so the greedy bijection is actually exercised (the prior PROCUREMENT_BUS case had a
    # >MAX_SUFFIX_SLACK delta and matched nothing, so the test passed even with bijection
    # deleted). The one column normalizes to PROCUREMENTBUSINESSUNITNAM:
    #   PROCUREMENT_BUSINESSUNIT_NAM -> normalize -> PROCUREMENTBUSINESSUNITNAM == column
    #       -> name=1.00 (exact after squash); listed FIRST so its position score is also top.
    #   PROCUREMENT_BU               -> expand -> PROCUREMENTBUSINESSUNIT (append NAM, delta 3)
    #       -> name=0.80 (truncation). Strictly lower combined score -> loses the column.
    oracle = {"PROCUREMENT_BUSINESSUNIT_NAM": _of("PROCUREMENT_BUSINESSUNIT_NAM"),
              "PROCUREMENT_BU": _of("PROCUREMENT_BU")}
    cols = [_col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM", size=25)]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols, decided=set())
    assert len(out) == 1                                          # the one column assigned once
    assert out[0].oracle_key == "PROCUREMENT_BUSINESSUNIT_NAM"    # the higher-scoring key wins
    assert all(c.oracle_key != "PROCUREMENT_BU" for c in out)     # the loser is absent


def test_score_uses_match_key_not_technical():
    # Pass-1 audit §2.1: when of.technical is None the match key comes from the label;
    # the name score must be computed from that key, not re-derived from of.technical
    # (which would score name=0.00 and mis-tier a correct correspondence into WEAK).
    of = AlignedField(position=1, label="Procurement BU", technical=None,
                      data_type="VARCHAR2", length=100, scale=None, required=None)
    score, signals = score_candidate("PROCUREMENT_BU", of,
                                     _col("T09PROCUREMENT_BUSINESSUNITNAM",
                                          "PROCUREMENT_BUSINESSUNITNAM", size=25),
                                     window=27, position_score=1.0)
    assert score > 0.0
    assert "name=0.00" not in signals


def test_tiers_are_ordered_high_probable_weak():
    names = [b for b, _ in TIER_BANDS]
    assert names == ["HIGH", "PROBABLE", "WEAK"]
    # New signature (§2.1): score_candidate(oracle_key, of, col, window, position_score).
    pc = score_candidate("PROCUREMENT_BU", _of("PROCUREMENT_BU"),
                         _col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM"),
                         window=27, position_score=1.0)
    assert pc[0] > 0.0 and pc[1]  # (score, signals) — signals non-empty


# ---------------------------------------------------------------------------
# Task 4: fieldmap workbook I/O + merge_fieldmap / merge_decisions
# ---------------------------------------------------------------------------

from fbdi.correspondence import (
    write_fieldmap_workbook, load_fieldmap_workbook, merge_fieldmap, merge_decisions,
)


def _fc(table, okey, bare, origin="derived", conf="HIGH"):
    return FieldCorrespondence(applaud_table=table, oracle_key=okey, applaud_bare=bare,
                               applaud_ddid=table[:3] + bare, confidence=conf, origin=origin)


def test_fieldmap_workbook_roundtrip(tmp_path):
    rows = [_fc("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM",
                origin="confirmed", conf="HIGH")]
    path = tmp_path / "fieldmap.xlsx"
    write_fieldmap_workbook(rows, path)
    loaded = load_fieldmap_workbook(path)
    assert loaded["T_POZ"][0].oracle_key == "PROCUREMENT_BU"
    assert loaded["T_POZ"][0].applaud_bare == "PROCUREMENT_BUSINESSUNITNAM"
    assert loaded["T_POZ"][0].origin == "confirmed"


def test_merge_confirmed_wins_over_rederive():
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "HAND_PICKED_BARE",
                               origin="confirmed")]}
    derived = [_fc("T_POZ", "PROCUREMENT_BU", "AUTO_BARE", origin="derived"),  # must NOT win
               _fc("T_POZ", "NEW_KEY", "NEW_BARE", origin="derived")]          # undecided -> added
    merged = merge_fieldmap(derived, committed)
    by = {(fc.oracle_key): fc for fc in merged["T_POZ"]}
    assert by["PROCUREMENT_BU"].applaud_bare == "HAND_PICKED_BARE"
    assert by["PROCUREMENT_BU"].origin == "confirmed"
    assert by["NEW_KEY"].origin == "derived"


def test_merge_rejected_also_wins_and_suppresses_reproposal():
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "", origin="rejected")]}
    derived = [_fc("T_POZ", "PROCUREMENT_BU", "AUTO_BARE", origin="derived")]
    merged = merge_fieldmap(derived, committed)
    assert merged["T_POZ"][0].origin == "rejected"


def test_rederive_idempotence_across_releases():
    # 26B decision survives a fresh 26B->next derive that re-proposes the same key.
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "CONFIRMED_BARE",
                               origin="confirmed")]}
    rederived = [_fc("T_POZ", "PROCUREMENT_BU", "AUTO_BARE", origin="derived")]
    once = merge_fieldmap(rederived, committed)
    twice = merge_fieldmap(rederived, once)
    assert twice["T_POZ"][0].applaud_bare == "CONFIRMED_BARE"
    assert len(twice["T_POZ"]) == 1


def test_confirm_overrides_prior_decision():
    # Pass-1 audit §1.1 (BLOCKER): at confirm time, a NEW human decision must win so a
    # wrong confirmation can be revised through the tooling.
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "OLD_BARE", origin="confirmed")]}
    new = [_fc("T_POZ", "PROCUREMENT_BU", "NEW_BARE", origin="confirmed")]
    merged = merge_decisions(new, committed)
    assert merged["T_POZ"][0].applaud_bare == "NEW_BARE"


def test_merge_decisions_carries_forward_untouched_committed():
    # A decision for one key must not disturb an unrelated committed row.
    committed = {"T_POZ": [_fc("T_POZ", "KEEP_KEY", "KEEP_BARE", origin="confirmed")]}
    new = [_fc("T_POZ", "NEW_KEY", "NEW_BARE", origin="rejected")]
    merged = merge_decisions(new, committed)
    by = {fc.oracle_key: fc for fc in merged["T_POZ"]}
    assert by["KEEP_KEY"].applaud_bare == "KEEP_BARE"
    assert by["NEW_KEY"].origin == "rejected"


def test_load_fieldmap_drops_stray_derived_rows(tmp_path):
    # Precedence invariant (§1.1): a hand-edited 'derived' row must not survive load and
    # silently block a future decision. Only confirmed/rejected persist.
    rows = [_fc("T_POZ", "GOOD_KEY", "GOOD_BARE", origin="confirmed"),
            _fc("T_POZ", "STRAY_KEY", "STRAY_BARE", origin="derived")]
    path = tmp_path / "fieldmap.xlsx"
    write_fieldmap_workbook(rows, path)
    loaded = load_fieldmap_workbook(path)
    keys = {fc.oracle_key for fc in loaded.get("T_POZ", [])}
    assert keys == {"GOOD_KEY"}


# ---------------------------------------------------------------------------
# Task 5: review workbook emit + load + apply_review_decisions
# ---------------------------------------------------------------------------

import pytest
from fbdi.correspondence import (
    ReviewRow, write_review_workbook, load_review_workbook,
    apply_review_decisions, InvalidCorrectedBareError,
)


def _review(table, okey, cand, confirm="", corrected=""):
    return ReviewRow(applaud_table=table, oracle_key=okey, oracle_type="char 100",
                     candidate_bare=cand, applaud_ddid=table[:3] + cand,
                     applaud_type="char 25", confidence="HIGH", score=0.88,
                     signals="name=0.80", alternatives="", confirm=confirm,
                     corrected_bare=corrected)


def test_review_workbook_roundtrip(tmp_path):
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM")]
    path = tmp_path / "review.xlsx"
    write_review_workbook(rows, path, exact_counts={"T_POZ": (212, 226)})
    loaded = load_review_workbook(path)
    assert loaded[0].oracle_key == "PROCUREMENT_BU"
    assert loaded[0].candidate_bare == "PROCUREMENT_BUSINESSUNITNAM"


def test_apply_confirm_yes_becomes_confirmed():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM", confirm="Y")]
    valid = {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}}
    out = apply_review_decisions(rows, valid)
    assert out[0].origin == "confirmed"
    assert out[0].applaud_bare == "PROCUREMENT_BUSINESSUNITNAM"


def test_apply_confirm_no_becomes_rejected():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM", confirm="N")]
    out = apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})
    assert out[0].origin == "rejected"


def test_apply_corrected_bare_overrides_candidate():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "WRONG_GUESS", corrected="REAL_BARE")]
    out = apply_review_decisions(rows, {"T_POZ": {"REAL_BARE", "WRONG_GUESS"}})
    assert out[0].origin == "confirmed" and out[0].applaud_bare == "REAL_BARE"


def test_apply_corrected_bare_not_in_table_fails_loud():
    # Audit §4.1: a typo'd Corrected Bare must abort the merge, not commit a dead alias.
    rows = [_review("T_POZ", "PROCUREMENT_BU", "WRONG_GUESS", corrected="TYPOO_BARE")]
    with pytest.raises(InvalidCorrectedBareError):
        apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})


def test_apply_confirm_yes_with_unknown_candidate_fails_loud():
    # Audit §4.1 (Y path): a reviewer who edits the candidate cell and marks Y must not
    # silently commit an alias to a non-existent column — same fail-loud guard as Corrected Bare.
    rows = [_review("T_POZ", "PROCUREMENT_BU", "EDITED_TO_NONEXISTENT", confirm="Y")]
    with pytest.raises(InvalidCorrectedBareError):
        apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})


def test_apply_skips_undecided_rows():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM")]  # no Y/N
    out = apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})
    assert out == []


# ---------------------------------------------------------------------------
# Code-review fixes (Fix 1, Fix 2, Fix 3)
# ---------------------------------------------------------------------------

def test_apply_warns_on_unrecognized_confirm_value(caplog):
    import logging
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM", confirm="YES")]
    with caplog.at_level(logging.WARNING, logger="fbdi.correspondence"):
        out = apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})
    assert out == []                                  # still skipped, but...
    assert any("PROCUREMENT_BU" in r.message and "YES" in r.message
               for r in caplog.records)               # ...loudly


def test_apply_corrected_bare_stores_canonical_casing_and_empty_ddid():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "WRONG_GUESS", corrected="real_bare")]
    out = apply_review_decisions(rows, {"T_POZ": {"REAL_BARE", "WRONG_GUESS"}})
    assert out[0].applaud_bare == "REAL_BARE"   # canonical from the table, not as-typed
    assert out[0].applaud_ddid == ""            # candidate's DDID must not leak in


def test_load_fieldmap_warns_on_duplicate_key_last_wins(tmp_path, caplog):
    import logging
    rows = [_fc("T_POZ", "DUP_KEY", "FIRST_BARE", origin="confirmed"),
            _fc("T_POZ", "DUP_KEY", "SECOND_BARE", origin="confirmed")]
    path = tmp_path / "fieldmap.xlsx"
    write_fieldmap_workbook(rows, path)
    with caplog.at_level(logging.WARNING, logger="fbdi.correspondence"):
        loaded = load_fieldmap_workbook(path)
    assert [fc.applaud_bare for fc in loaded["T_POZ"]] == ["SECOND_BARE"]
    assert any("DUP_KEY" in r.message for r in caplog.records)


def test_load_fieldmap_warns_on_stray_derived_rows(tmp_path, caplog):
    import logging
    rows = [_fc("T_POZ", "GOOD_KEY", "GOOD_BARE", origin="confirmed"),
            _fc("T_POZ", "STRAY_KEY", "STRAY_BARE", origin="derived")]
    path = tmp_path / "fieldmap.xlsx"
    write_fieldmap_workbook(rows, path)
    with caplog.at_level(logging.WARNING, logger="fbdi.correspondence"):
        loaded = load_fieldmap_workbook(path)
    keys = {fc.oracle_key for fc in loaded.get("T_POZ", [])}
    assert keys == {"GOOD_KEY"}
    assert any("STRAY_KEY" in r.message for r in caplog.records)


# ---------------------------------------------------------------------------
# Task 6: build_alias resolver + confidence gate
# ---------------------------------------------------------------------------

from fbdi.correspondence import build_alias


def test_build_alias_confirmed_only_by_default():
    rows = [_fc("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM",
                origin="confirmed"),
            _fc("T_POZ", "OTHER_KEY", "OTHER_BARE", origin="derived", conf="HIGH")]
    alias = build_alias(rows, accept_confidence="confirmed")
    assert alias == {"PROCUREMENT_BUSINESSUNITNAM": "PROCUREMENT_BU"}


def test_build_alias_admits_derived_at_or_above_tier():
    rows = [_fc("T_POZ", "K1", "BARE_HIGH", origin="derived", conf="HIGH"),
            _fc("T_POZ", "K2", "BARE_WEAK", origin="derived", conf="WEAK")]
    alias = build_alias(rows, accept_confidence="HIGH")
    assert alias == {"BARE_HIGH": "K1"}   # WEAK excluded


def test_build_alias_never_aliases_rejected():
    rows = [_fc("T_POZ", "K1", "BARE", origin="rejected", conf="HIGH")]
    assert build_alias(rows, accept_confidence="WEAK") == {}


def test_build_alias_tier_gate_extends_confirmed_not_replaces():
    # A tier gate ADDS derived rows on top of confirmed rows — confirmed always aliases.
    rows = [_fc("T_POZ", "K_CONF", "BARE_CONF", origin="confirmed", conf="WEAK"),
            _fc("T_POZ", "K_DER", "BARE_DER", origin="derived", conf="HIGH")]
    alias = build_alias(rows, accept_confidence="HIGH")
    assert alias == {"BARE_CONF": "K_CONF", "BARE_DER": "K_DER"}


# ---------------------------------------------------------------------------
# Review note #1: Y-confirmed rows fold provenance into notes
# ---------------------------------------------------------------------------

from fbdi.correspondence import assemble_derivation_inputs


def _snapshot_with_procurement_sites():
    from fbdi.applaud_snapshot import ApplaudSnapshot, SnapshotTable, DataColumn
    col = DataColumn("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM",
                     "X", 25, None, None, 1)
    t = SnapshotTable("T_POZ_SUPPLIER_SITES_INT", "T09", False, "(T09)", [], [col])
    return ApplaudSnapshot("ORACLE_MASTER", "x", "2026-06-11", "t",
                           tables={"T_POZ_SUPPLIER_SITES_INT": t})


def test_assemble_derivation_inputs_groups_by_table():
    snap = _snapshot_with_procurement_sites()
    catalog = {("PO_TPL", "Suppliers"): [_of("PROCUREMENT_BU")]}
    mapping = {("PO_TPL", "Suppliers"): {"applaud_table": "T_POZ_SUPPLIER_SITES_INT"}}
    inputs = assemble_derivation_inputs(snap, catalog, mapping)
    prefix, oracle_by_key, cols = inputs["T_POZ_SUPPLIER_SITES_INT"]
    assert prefix == "T09"
    assert "PROCUREMENT_BU" in oracle_by_key
    assert cols and cols[0].bare == "PROCUREMENT_BUSINESSUNITNAM"


def test_assemble_merges_multiple_tabs_to_one_table():
    # Pass-1 audit §2.2: two (template, tab) rows mapping to ONE Applaud table must MERGE
    # their Oracle keys, not overwrite — both tabs' divergent fields enter the pool.
    snap = _snapshot_with_procurement_sites()
    catalog = {("PO_TPL", "Suppliers"): [_of("PROCUREMENT_BU")],
               ("PO_TPL", "Addresses"): [_of("REMIT_ADVICE_DELIVERY")]}
    mapping = {("PO_TPL", "Suppliers"): {"applaud_table": "T_POZ_SUPPLIER_SITES_INT"},
               ("PO_TPL", "Addresses"): {"applaud_table": "T_POZ_SUPPLIER_SITES_INT"}}
    inputs = assemble_derivation_inputs(snap, catalog, mapping)
    _prefix, oracle_by_key, _cols = inputs["T_POZ_SUPPLIER_SITES_INT"]
    assert "PROCUREMENT_BU" in oracle_by_key      # from the Suppliers tab
    assert "REMIT_ADVICE_DELIVERY" in oracle_by_key  # from the Addresses tab — NOT dropped


def test_apply_confirm_yes_folds_provenance_into_notes(tmp_path):
    # Carried review note #1: Y-confirmed FieldCorrespondence must carry provenance in
    # .notes so write_fieldmap_workbook can persist it (Notes column). The score/signals
    # live in memory only; the workbook has no Score/Signals columns.
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM", confirm="Y")]
    valid = {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}}
    out = apply_review_decisions(rows, valid)
    assert "HIGH" in out[0].notes
    assert "name=0.80" in out[0].notes

    # Round-trip persistence: write -> reload -> notes survives.
    write_fieldmap_workbook(out, tmp_path / "fm.xlsx")
    reloaded = load_fieldmap_workbook(tmp_path / "fm.xlsx")
    reloaded_row = reloaded["T_POZ"][0]
    assert "HIGH" in reloaded_row.notes
    assert "name=0.80" in reloaded_row.notes
