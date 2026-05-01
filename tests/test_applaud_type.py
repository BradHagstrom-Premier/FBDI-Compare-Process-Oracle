"""Tests for fbdi.applaud_type — Oracle → Applaud type translator."""

from fbdi.applaud_type import applaud_type_for
from fbdi.type_parser import ParsedType


def _pt(data_type, length=None, scale=None, parse_warning=False):
    return ParsedType(data_type=data_type, length=length, scale=scale, parse_warning=parse_warning)


class TestApplaudTypeFor:
    def test_varchar2_with_length(self):
        assert applaud_type_for(_pt("VARCHAR2", length=30)) == "char 30"

    def test_varchar2_without_length(self):
        # Should never happen in real data, but be deterministic anyway
        assert applaud_type_for(_pt("VARCHAR2")) == "char"

    def test_number_with_precision_and_scale(self):
        assert applaud_type_for(_pt("NUMBER", length=18, scale=4)) == "numeric 18,4"

    def test_number_with_precision_only(self):
        assert applaud_type_for(_pt("NUMBER", length=18)) == "numeric 18"

    def test_number_with_precision_and_zero_scale(self):
        # scale=0 is distinct from scale=None: integer columns declared NUMBER(p,0)
        # must not fall through to the precision-only branch.
        assert applaud_type_for(_pt("NUMBER", length=18, scale=0)) == "numeric 18,0"

    def test_number_no_precision(self):
        # Plain NUMBER → "numeric" with no defaults invented
        assert applaud_type_for(_pt("NUMBER")) == "numeric"

    def test_date(self):
        assert applaud_type_for(_pt("DATE")) == "date"

    def test_timestamp(self):
        assert applaud_type_for(_pt("TIMESTAMP")) == "date"

    def test_clob_passthrough(self):
        assert applaud_type_for(_pt("CLOB")) == "clob"

    def test_blob_passthrough(self):
        assert applaud_type_for(_pt("BLOB")) == "blob"

    def test_unknown_type_passes_through_lowercase(self):
        assert applaud_type_for(_pt("XMLTYPE")) == "xmltype"

    def test_empty_returns_empty(self):
        # Blank input (thin tab with no type info) returns empty string
        assert applaud_type_for(_pt("")) == ""

    def test_parse_warning_returns_empty(self):
        # Don't fabricate a type for unparseable input
        assert applaud_type_for(_pt("VARCHAR2", parse_warning=True)) == ""
