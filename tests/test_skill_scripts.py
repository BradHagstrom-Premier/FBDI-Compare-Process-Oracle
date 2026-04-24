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
