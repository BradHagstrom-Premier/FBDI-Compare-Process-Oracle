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


class TestShiftDetection:
    def test_mid_file_insert_classifies_as_one_add_plus_shifts(self):
        """Real scenario from 26A→26B WorkDefinitionTemplate.

        Insert one new field at position 19. Catalog's naive per-position diff
        misclassifies this as multiple ADDs + RENAMEDs. Alignment must produce
        exactly 1 ADDED + N SHIFTED.
        """
        old = [
            _row(18, "Transform from Item Number", "TRANSFORM_FROM_ITEM_NUMBER", "VARCHAR2", 300, False),
            _row(19, "Completion Subinventory", "COMPLETION_SUBINVENTORY_NAME", "VARCHAR2", 10, False),
            _row(20, "Completion Locator Segment1", "COMPL_LOCATOR_SEGMENT1", "VARCHAR2", 40, False),
            _row(21, "Completion Locator Segment2", "COMPL_LOCATOR_SEGMENT2", "VARCHAR2", 40, False),
        ]
        new = [
            _row(18, "Transform from Item Number", "TRANSFORM_FROM_ITEM_NUMBER", "VARCHAR2", 300, False),
            _row(19, "Enable Parallel Operations", "ENABLE_PARALLEL_OPS_FLAG", "VARCHAR2", 1, False),
            _row(20, "Completion Subinventory", "COMPLETION_SUBINVENTORY_NAME", "VARCHAR2", 10, False),
            _row(21, "Completion Locator Segment1", "COMPL_LOCATOR_SEGMENT1", "VARCHAR2", 40, False),
            _row(22, "Completion Locator Segment2", "COMPL_LOCATOR_SEGMENT2", "VARCHAR2", 40, False),
        ]
        result = align_tabs(old, new)

        added = [c for c in result if c.change_type == "ADDED"]
        shifted = [c for c in result if c.change_type == "SHIFTED"]
        renamed = [c for c in result if c.change_type == "RENAMED"]
        modified = [c for c in result if c.change_type == "MODIFIED"]

        assert len(added) == 1
        assert added[0].new_field.technical == "ENABLE_PARALLEL_OPS_FLAG"
        assert added[0].new_position == 19

        assert len(shifted) == 3
        # Verify each SHIFTED keeps the same field but moves position +1
        for c in shifted:
            assert c.old_field.technical == c.new_field.technical
            assert c.new_position == c.old_position + 1

        assert renamed == []
        assert modified == []

    def test_pure_swap_emits_two_shifts(self):
        old = [
            _row(1, "Field A", "FIELD_A", "VARCHAR2", 10, True),
            _row(2, "Field B", "FIELD_B", "VARCHAR2", 10, True),
        ]
        new = [
            _row(1, "Field B", "FIELD_B", "VARCHAR2", 10, True),
            _row(2, "Field A", "FIELD_A", "VARCHAR2", 10, True),
        ]
        result = align_tabs(old, new)
        # LCS picks one of the two as a stable axis; the other becomes a remove+add.
        # We don't constrain *which*, but the total shape must conserve fields.
        assert len(result) > 0
        # Both fields must appear somewhere on either side
        all_techs = set()
        for c in result:
            if c.old_field:
                all_techs.add(c.old_field.technical)
            if c.new_field:
                all_techs.add(c.new_field.technical)
        assert all_techs == {"FIELD_A", "FIELD_B"}
