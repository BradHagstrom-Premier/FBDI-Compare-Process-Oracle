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
