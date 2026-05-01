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


class TestIdentityMatching:
    def test_perfect_match_no_changes_emits_nothing(self):
        rows = [_row(1, "Item Name", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs(rows, rows)
        assert result == []

    def test_match_by_technical_name_when_present(self):
        old = [_row(1, "Old Label", "ITEM_NAME", "VARCHAR2", 30, True)]
        new = [_row(1, "New Label", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs(old, new)
        # Same technical name → matched (so RENAMED, not ADD+REMOVE)
        assert len(result) == 1
        assert result[0].change_type == "RENAMED"

    def test_match_by_label_when_technical_is_none(self):
        # Thin tabs have no technical name; fall back to label
        old = [_row(1, "Item Name", None, None, None, True)]
        new = [_row(1, "Item Name", None, None, None, True)]
        result = align_tabs(old, new)
        assert result == []

    def test_no_match_when_both_label_and_technical_differ(self):
        old = [_row(1, "Old", "OLD_FIELD", "VARCHAR2", 30, True)]
        new = [_row(1, "New", "NEW_FIELD", "VARCHAR2", 30, True)]
        result = align_tabs(old, new)
        # Disjoint identity — REMOVE the old + ADD the new
        types = sorted(c.change_type for c in result)
        assert types == ["ADDED", "REMOVED"]
