"""Tests for the fbdi-compare-release skill's bundled scripts."""

import json
import subprocess
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


from scripts import check_env  # noqa: E402 — importable when cwd is repo root


def _run_check_env(tmp_path):
    """Invoke check_env.py as a subprocess with a cwd of tmp_path and return (exit_code, stdout_json)."""
    cmd = [sys.executable, str(SKILL_ROOT / "scripts" / "check_env.py")]
    proc = subprocess.run(cmd, cwd=tmp_path, capture_output=True, text=True)
    return proc.returncode, json.loads(proc.stdout)


def test_check_env_exposes_main():
    assert hasattr(check_env, "main")


def test_check_env_python_version_check_passes_on_314():
    result = check_env.check_python_version(current=(3, 14, 3))
    assert result["ok"] is True
    assert "3.14" in result["detail"]


def test_check_env_python_version_check_fails_on_old():
    result = check_env.check_python_version(current=(3, 11, 0))
    assert result["ok"] is False
    assert "3.14" in result["detail"]


def test_check_env_deps_check_detects_missing():
    result = check_env.check_deps(required=["definitely_not_a_real_package_xyz"])
    assert result["ok"] is False
    assert "definitely_not_a_real_package_xyz" in result["detail"]


def test_check_env_deps_check_passes_on_stdlib():
    # json is stdlib — always importable
    result = check_env.check_deps(required=["json"])
    assert result["ok"] is True


def test_check_env_baselines_dir_creates_if_missing(tmp_path):
    result = check_env.check_baselines_dir(root=tmp_path)
    assert result["ok"] is True
    assert (tmp_path / "baselines").is_dir()


def test_check_env_produces_structured_json(tmp_path):
    """check_env.py always emits a parseable payload with the documented
    shape. Exit code is not asserted here because a dev machine may legitimately
    be missing Chrome (→ exit 1) or deps (→ exit 2); those paths have their
    own unit tests via the helper functions."""
    (tmp_path / "baselines").mkdir()
    (tmp_path / "baseline_files.txt").write_text("stub\n")
    _, payload = _run_check_env(tmp_path)
    assert "checks" in payload
    assert "missing_deps" in payload
    assert "fatal" in payload


def test_check_env_json_output_parseable(tmp_path):
    exit_code, payload = _run_check_env(tmp_path)
    assert isinstance(payload, dict)
    assert isinstance(payload["checks"], list)


from scripts import verify_download  # noqa: E402


INVENTORY_FIXTURE = """\
FBDI Baseline File Inventory
Generated: 2026-04-23
============================

26A has 3 files. 26B has 4 files.

============================
26A ORIGINALS (3 files)
============================
AccountCombinationsImportTemplate.xlsm
BudgetImportTemplate.xlsm
RapidImplementationForCashManagement.xlsm

============================
26B ORIGINALS (4 files)
============================
AccountCombinationsImportTemplate.xlsm
BudgetImportTemplate.xlsm
ItemImportReferenceOrgTemplate.xlsm
RapidImplementationForCashManagement.xlsm

============================
DIFFERENCES
============================
Only in 26B: ItemImportReferenceOrgTemplate.xlsm
"""


def test_parse_inventory_extracts_both_sections():
    inventory = verify_download.parse_inventory(INVENTORY_FIXTURE)
    assert set(inventory.keys()) == {"26A", "26B"}
    assert inventory["26A"] == [
        "AccountCombinationsImportTemplate.xlsm",
        "BudgetImportTemplate.xlsm",
        "RapidImplementationForCashManagement.xlsm",
    ]
    assert len(inventory["26B"]) == 4


def test_parse_inventory_ignores_differences_footer():
    inventory = verify_download.parse_inventory(INVENTORY_FIXTURE)
    # The "DIFFERENCES" block is after the last ORIGINALS header;
    # its content must NOT leak into 26B.
    assert "Only in 26B: ItemImportReferenceOrgTemplate.xlsm" not in inventory["26B"]


def test_parse_inventory_empty_text():
    assert verify_download.parse_inventory("") == {}


def test_diff_clean_case():
    inventory = {"26A": ["A.xlsm", "B.xlsm"]}
    result = verify_download.diff_against_inventory(
        release="26A",
        downloaded_names=["A.xlsm", "B.xlsm"],
        inventory=inventory,
        manual_files=[],
    )
    assert result["missing"] == []
    assert result["extras"] == []


def test_diff_detects_missing_and_extras():
    inventory = {"26A": ["A.xlsm", "B.xlsm", "C.xlsm"]}
    result = verify_download.diff_against_inventory(
        release="26A",
        downloaded_names=["A.xlsm", "B.xlsm", "D.xlsm"],
        inventory=inventory,
        manual_files=[],
    )
    assert result["missing"] == ["C.xlsm"]
    assert result["extras"] == ["D.xlsm"]


def test_diff_excludes_manual_files_from_missing():
    inventory = {"26A": ["A.xlsm", "RapidImplementationForCashManagement.xlsm"]}
    result = verify_download.diff_against_inventory(
        release="26A",
        downloaded_names=["A.xlsm"],
        inventory=inventory,
        manual_files=["RapidImplementationForCashManagement.xlsm"],
    )
    assert result["missing"] == []  # manual file excluded


def test_diff_is_locale_agnostic():
    """Guard against non-`LC_ALL=C` environments where Mac default sort
    misorders mixed-case filenames. Set operations are locale-independent —
    we verify the diff is identical regardless of filename case ordering."""
    inventory = {"26A": ["AccountCombinationsImportTemplate.xlsm", "zxCustomTemplate.xlsm"]}
    # Downloaded in a different case-order
    result = verify_download.diff_against_inventory(
        release="26A",
        downloaded_names=["zxCustomTemplate.xlsm", "AccountCombinationsImportTemplate.xlsm"],
        inventory=inventory,
        manual_files=[],
    )
    assert result["missing"] == []
    assert result["extras"] == []
