"""Tests for the fbdi-compare-release skill's bundled scripts."""

import sys
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parent.parent / ".claude" / "skills" / "fbdi-compare-release"
# Make `from scripts import <module>` resolve the skill's bundled scripts.
sys.path.insert(0, str(SKILL_ROOT))


def test_skill_folder_exists():
    assert SKILL_ROOT.is_dir(), f"expected skill folder at {SKILL_ROOT}"


def test_skill_md_has_frontmatter():
    skill_md = SKILL_ROOT / "SKILL.md"
    assert skill_md.is_file()
    text = skill_md.read_text(encoding="utf-8")
    assert text.startswith("---\n")
    assert "\nname: fbdi-compare-release\n" in text
    assert "\ndescription:" in text


def test_scripts_dir_is_python_package():
    scripts_dir = SKILL_ROOT / "scripts"
    assert scripts_dir.is_dir()
    assert (scripts_dir / "__init__.py").is_file()
