"""Tests for fbdi.align — LCS-style alignment of FBDI tab rows across releases."""

from fbdi.align import AlignedField, Change, align_tabs


def _row(position, label, technical=None, data_type=None, length=None, required=None):
    """Build an AlignedField input row."""
    return AlignedField(
        position=position, label=label, technical=technical,
        data_type=data_type, length=length, required=required,
    )


class TestEmpty:
    def test_both_empty(self):
        result = align_tabs([], [])
        assert result == []

    def test_old_empty_new_one_field(self):
        new = [_row(1, "Item Name", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs([], new)
        assert len(result) == 1
        assert result[0].change_type == "ADDED"
        assert result[0].new_position == 1
        assert result[0].old_position is None

    def test_new_empty_old_one_field(self):
        old = [_row(1, "Item Name", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs(old, [])
        assert len(result) == 1
        assert result[0].change_type == "REMOVED"
        assert result[0].old_position == 1
        assert result[0].new_position is None
