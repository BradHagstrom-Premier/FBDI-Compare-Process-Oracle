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
