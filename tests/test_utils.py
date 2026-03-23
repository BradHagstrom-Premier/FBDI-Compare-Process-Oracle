"""Tests for fbdi.utils — column letter generation and file matching."""

import pytest
from pathlib import Path
from fbdi.utils import col_index_to_letter, match_fbdi_files


class TestColIndexToLetter:
    def test_single_letters(self):
        assert col_index_to_letter(1) == "A"
        assert col_index_to_letter(2) == "B"
        assert col_index_to_letter(26) == "Z"

    def test_double_letters(self):
        assert col_index_to_letter(27) == "AA"
        assert col_index_to_letter(28) == "AB"
        assert col_index_to_letter(52) == "AZ"
        assert col_index_to_letter(53) == "BA"
        assert col_index_to_letter(702) == "ZZ"

    def test_triple_letters(self):
        assert col_index_to_letter(703) == "AAA"

    def test_edge_cases(self):
        assert col_index_to_letter(0) == ""
        assert col_index_to_letter(-1) == ""


class TestMatchFbdiFiles:
    def test_basic_matching(self, tmp_path):
        old_dir = tmp_path / "old"
        new_dir = tmp_path / "new"
        old_dir.mkdir()
        new_dir.mkdir()

        (old_dir / "TemplateA.xlsm").touch()
        (old_dir / "TemplateB.xlsm").touch()
        (new_dir / "TemplateA.xlsm").touch()
        (new_dir / "TemplateB.xlsm").touch()

        matched, old_only, new_only = match_fbdi_files(old_dir, new_dir)
        assert len(matched) == 2
        assert len(old_only) == 0
        assert len(new_only) == 0

    def test_old_only_and_new_only(self, tmp_path):
        old_dir = tmp_path / "old"
        new_dir = tmp_path / "new"
        old_dir.mkdir()
        new_dir.mkdir()

        (old_dir / "OnlyInOld.xlsm").touch()
        (old_dir / "Shared.xlsm").touch()
        (new_dir / "Shared.xlsm").touch()
        (new_dir / "OnlyInNew.xlsm").touch()

        matched, old_only, new_only = match_fbdi_files(old_dir, new_dir)
        assert len(matched) == 1
        assert len(old_only) == 1
        assert len(new_only) == 1
        assert old_only[0].stem == "OnlyInOld"
        assert new_only[0].stem == "OnlyInNew"

    def test_case_insensitive_matching(self, tmp_path):
        old_dir = tmp_path / "old"
        new_dir = tmp_path / "new"
        old_dir.mkdir()
        new_dir.mkdir()

        (old_dir / "templateA.xlsm").touch()
        (new_dir / "TemplateA.xlsm").touch()

        matched, old_only, new_only = match_fbdi_files(old_dir, new_dir)
        assert len(matched) == 1
        assert len(old_only) == 0
        assert len(new_only) == 0

    def test_mixed_extensions(self, tmp_path):
        old_dir = tmp_path / "old"
        new_dir = tmp_path / "new"
        old_dir.mkdir()
        new_dir.mkdir()

        (old_dir / "Template.xlsm").touch()
        (new_dir / "Template.xlsx").touch()

        matched, old_only, new_only = match_fbdi_files(old_dir, new_dir)
        assert len(matched) == 1
        assert len(old_only) == 0
        assert len(new_only) == 0

    def test_sorted_output(self, tmp_path):
        old_dir = tmp_path / "old"
        new_dir = tmp_path / "new"
        old_dir.mkdir()
        new_dir.mkdir()

        (old_dir / "Zebra.xlsm").touch()
        (old_dir / "Alpha.xlsm").touch()
        (new_dir / "Zebra.xlsm").touch()
        (new_dir / "Alpha.xlsm").touch()

        matched, _, _ = match_fbdi_files(old_dir, new_dir)
        stems = [p[0].stem for p in matched]
        assert stems == sorted(stems, key=str.lower)
