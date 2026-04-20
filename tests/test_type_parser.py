"""Tests for fbdi.type_parser."""

from fbdi.type_parser import parse_data_type, ParsedType


class TestParseDataType:
    def test_varchar2_with_char_suffix(self):
        result = parse_data_type("VARCHAR2(5 CHAR)")
        assert result == ParsedType("VARCHAR2", 5, None, False)

    def test_varchar2_large_with_char_suffix(self):
        result = parse_data_type("VARCHAR2(2048 CHAR)")
        assert result == ParsedType("VARCHAR2", 2048, None, False)

    def test_varchar2_without_char_suffix(self):
        result = parse_data_type("VARCHAR2(80)")
        assert result == ParsedType("VARCHAR2", 80, None, False)

    def test_lowercase_varchar2_normalizes(self):
        result = parse_data_type("Varchar2(250)")
        assert result == ParsedType("VARCHAR2", 250, None, False)

    def test_number_precision_only(self):
        result = parse_data_type("NUMBER(18)")
        assert result == ParsedType("NUMBER", 18, None, False)

    def test_number_with_scale(self):
        result = parse_data_type("NUMBER(18,4)")
        assert result == ParsedType("NUMBER", 18, 4, False)

    def test_number_with_scale_and_spaces(self):
        result = parse_data_type("NUMBER(18, 4)")
        assert result == ParsedType("NUMBER", 18, 4, False)

    def test_date_no_parens(self):
        result = parse_data_type("DATE")
        assert result == ParsedType("DATE", None, None, False)

    def test_clob_no_parens(self):
        result = parse_data_type("CLOB")
        assert result == ParsedType("CLOB", None, None, False)

    def test_blob_no_parens(self):
        result = parse_data_type("BLOB")
        assert result == ParsedType("BLOB", None, None, False)

    def test_varchar2_with_byte_suffix(self):
        # Some templates use BYTE instead of CHAR
        result = parse_data_type("VARCHAR2(100 BYTE)")
        assert result == ParsedType("VARCHAR2", 100, None, False)

    def test_empty_string_no_warning(self):
        # Empty input is a legitimate blank, not a parse failure
        result = parse_data_type("")
        assert result == ParsedType("", None, None, False)

    def test_none_no_warning(self):
        result = parse_data_type(None)
        assert result == ParsedType("", None, None, False)

    def test_whitespace_only_no_warning(self):
        result = parse_data_type("   ")
        assert result == ParsedType("", None, None, False)

    def test_garbage_string_sets_warning(self):
        result = parse_data_type("???weird junk???")
        assert result.parse_warning is True
        assert result.data_type == ""
        assert result.length is None
        assert result.scale is None

    def test_extra_text_sets_warning(self):
        result = parse_data_type("VARCHAR2(50) NOT NULL DEFAULT 'x'")
        assert result.parse_warning is True


class TestTrailingPeriod:
    def test_varchar2_with_trailing_period(self):
        result = parse_data_type("VARCHAR2(1 CHAR).")
        assert result == ParsedType("VARCHAR2", 1, None, False)

    def test_number_with_scale_and_trailing_period(self):
        result = parse_data_type("NUMBER(18,4).")
        assert result == ParsedType("NUMBER", 18, 4, False)

    def test_date_with_trailing_period(self):
        result = parse_data_type("DATE.")
        assert result == ParsedType("DATE", None, None, False)


class TestTemporalFormatMask:
    def test_date_upper_slash_format(self):
        result = parse_data_type("DATE(YYYY/MM/DD)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_upper_space_then_format(self):
        result = parse_data_type("DATE (YYYY/MM/DD)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_mixed_case_format(self):
        result = parse_data_type("Date(YYYY/MM/DD)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_lower_format(self):
        result = parse_data_type("Date(yyyy/mm/dd)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_datetime_format(self):
        result = parse_data_type("Date(yyyy/mm/dd hh24:mm)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_stray_leading_paren(self):
        # 40-row Oracle typo: 'Date((yyyy/mm/dd hh24:mm)'
        result = parse_data_type("Date((yyyy/mm/dd hh24:mm)")
        assert result == ParsedType("DATE", None, None, False)

    def test_timestamp_time_only(self):
        result = parse_data_type("TimeStamp(hh24:mm)")
        assert result == ParsedType("TIMESTAMP", None, None, False)

    def test_timestamp_full_datetime(self):
        result = parse_data_type("TimeStamp(yyyy/mm/dd hh24:mm:ss)")
        assert result == ParsedType("TIMESTAMP", None, None, False)


class TestStillWarnsOnTrulyMalformed:
    # These 11-or-so Oracle strings are genuinely broken. The parser should
    # NOT silently swallow them — Brad wants to see real data-quality issues.
    def test_unbalanced_leading_paren_warns(self):
        result = parse_data_type("(VARCHAR2(150)")
        assert result.parse_warning is True

    def test_missing_closing_paren_warns(self):
        result = parse_data_type("varchar2(4")
        assert result.parse_warning is True

    def test_non_digit_length_warns(self):
        result = parse_data_type("VARCHAR2(18R)")
        assert result.parse_warning is True

    def test_sentence_warns(self):
        result = parse_data_type("For desc Asset it is mandatory for create")
        assert result.parse_warning is True

    def test_label_without_paren_warns(self):
        # Single-word tokens that are not recognized types (VARCHAR2 etc.)
        # with no parens cannot be distinguished from type names like DATE
        # by the strict regex, so they match as a bare type name. That's
        # the correct behavior — the only way to know "Item Number" is
        # wrong is that it has a space and no parens. Current strict regex
        # requires [A-Za-z0-9]* after the leading letter, so "Item Number"
        # fails due to the space and sets the warning flag.
        result = parse_data_type("Item Number")
        assert result.parse_warning is True
