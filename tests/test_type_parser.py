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
