"""Tests for fbdi.report — view-model construction and scope filtering."""

import pytest
from pathlib import Path

from fbdi.report import (
    FileSection,
    ReportContext,
    PendingBaseEntry,
    build_report_context,
)
from fbdi.align import AlignedField, Change


# ---- helpers ----

def _aligned(position, label, technical, data_type=None, length=None, required=None, scale=None):
    return AlignedField(position=position, label=label, technical=technical,
                       data_type=data_type, length=length, scale=scale, required=required)


def _mapping(template, tab, applaud_table="T_X", prefix="TX1",
             module="Financials", in_base=None):
    """Build one mapping dict entry."""
    return {
        (template, tab): {
            "applaud_table": applaud_table,
            "prefix": prefix,
            "module": module,
            "in_base": in_base,
        }
    }


# ---- tests ----

class TestScopeFiltering:
    def test_unmapped_file_is_silently_excluded(self):
        catalog_old = {("UnmappedFile", "TabA"): [_aligned(1, "A", "A")]}
        catalog_new = {("UnmappedFile", "TabA"): [_aligned(1, "A", "A"),
                                                  _aligned(2, "B", "B")]}
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping={}, old_release="26A", new_release="26B",
        )
        assert ctx.file_sections == []
        assert ctx.pending_base == []

    def test_mapped_in_base_routes_to_main_body(self):
        catalog_old = {("MappedFile", "TabA"): [_aligned(1, "Old", "OLD_F")]}
        catalog_new = {("MappedFile", "TabA"): [_aligned(1, "New", "NEW_F")]}
        mapping = _mapping("MappedFile", "TabA", in_base=None)
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        assert len(ctx.file_sections) == 1
        assert ctx.pending_base == []

    def test_pending_base_routes_to_separate_section(self):
        catalog_old = {("PendingFile", "TabA"): [_aligned(1, "Old", "OLD_F")]}
        catalog_new = {("PendingFile", "TabA"): [_aligned(1, "Old", "OLD_F"),
                                                 _aligned(2, "New", "NEW_F")]}
        mapping = _mapping("PendingFile", "TabA",
                          in_base="Needs to be created in base system")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        assert ctx.file_sections == []
        assert len(ctx.pending_base) == 1
        assert ctx.pending_base[0].file == "PendingFile"
        assert ctx.pending_base[0].tab == "TabA"
        assert ctx.pending_base[0].change_count == 1


class TestApplaudFieldNameConstruction:
    def test_uses_technical_when_present(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Some Label", "FIELD_NAME",
                                              "VARCHAR2", 30, True)]}
        mapping = _mapping("F", "T", prefix="TX1")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        section = ctx.file_sections[0]
        added = section.changes_by_type["ADDED"]
        assert added[0].applaud_field_name == "TX1FIELD_NAME"

    def test_falls_back_to_normalized_label_when_technical_is_none(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Some Label!", None,
                                              None, None, True)]}
        mapping = _mapping("F", "T", prefix="TX1")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        section = ctx.file_sections[0]
        added = section.changes_by_type["ADDED"]
        # normalize_label strips '!' and joins → "Some Label"
        assert added[0].applaud_field_name == "TX1Some Label"

    def test_thirty_char_warning_set_when_name_exceeds_limit(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "X",
                                              "COPY_LOTS_AND_SERIAL_NUMBERS_FROM_PARENT_TXN",
                                              "VARCHAR2", 1, True)]}
        mapping = _mapping("F", "T", prefix="TH8")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        added = ctx.file_sections[0].changes_by_type["ADDED"]
        assert added[0].name_exceeds_30
        assert added[0].name_length > 30


class TestModuleRollup:
    def test_module_counts_aggregate_across_files(self):
        catalog_old = {
            ("F1", "T1"): [_aligned(1, "X", "X")],
            ("F2", "T1"): [_aligned(1, "X", "X")],
        }
        catalog_new = {
            ("F1", "T1"): [_aligned(1, "X", "X"), _aligned(2, "Y", "Y")],
            ("F2", "T1"): [_aligned(1, "X", "X")],  # no changes
        }
        mapping = {**_mapping("F1", "T1", module="Financials"),
                   **_mapping("F2", "T1", module="Financials")}
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        assert "Financials" in ctx.module_rollup
        rollup = ctx.module_rollup["Financials"]
        assert rollup["tabs"] == 1   # only F1/T1 has changes
        assert rollup["added"] == 1
