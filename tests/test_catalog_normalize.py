"""Tests for fbdi.catalog_normalize."""

from fbdi.catalog_normalize import normalize_label


class TestNormalizeLabel:
    def test_strips_leading_asterisk(self):
        assert normalize_label("*Source Budget Type") == "Source Budget Type"

    def test_strips_punctuation_keeps_alphanumeric_and_underscore(self):
        assert normalize_label("$Weird, Chars!") == "Weird Chars"

    def test_preserves_underscore(self):
        assert normalize_label("COLUMN_NAME") == "COLUMN_NAME"

    def test_preserves_mixed_snake_case(self):
        assert normalize_label("my_col_name") == "my_col_name"

    def test_collapses_runs_of_whitespace(self):
        assert normalize_label("  *Foo  Bar  ") == "Foo Bar"

    def test_empty_string(self):
        assert normalize_label("") == ""

    def test_none_returns_empty(self):
        assert normalize_label(None) == ""

    def test_only_punctuation_returns_empty(self):
        assert normalize_label("!!!") == ""

    def test_digits_preserved(self):
        assert normalize_label("Column123") == "Column123"

    def test_unicode_alphanumerics_pass_through(self):
        # Python's str.isalnum() returns True for Unicode letters
        assert normalize_label("Café Name") == "Café Name"

    def test_collapses_tabs_and_newlines(self):
        assert normalize_label("Foo\tBar\nBaz") == "Foo Bar Baz"
