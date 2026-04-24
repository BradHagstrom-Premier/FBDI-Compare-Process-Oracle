# FBDI Compare-Release Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a project-level Claude Code skill, `fbdi-compare-release`, that walks a coworker from "Oracle released 26C" through the 8-stage pipeline (env check → version resolve → download + verify → clear → compare → catalog → summary → post-run verify) to a finished `Comparison_Report_<OLD>_<NEW>.xlsx` plus an updated `FBDI_Master_Catalog.xlsx`, with human-in-the-loop prompts at six decision points and no re-implementation of comparison logic.

**Architecture:** `SKILL.md` is the orchestrator; Claude executes an ordered workflow and delegates narrow deterministic work to four bundled Python scripts (`check_env.py`, `verify_download.py`, `summarize_report.py`, `verify_run.py`). Scripts print JSON to stdout + set exit codes so SKILL.md can branch cleanly; nothing imports `fbdi/*` internals (scripts only call the existing `python -m fbdi ...` CLI and read output xlsx files). References (`troubleshooting.md`, `release-version-format.md`) are read-on-demand. Everything lives under `.claude/skills/fbdi-compare-release/` and is committed with the repo.

**Tech Stack:** Python 3.14+, openpyxl (read comparison + catalog output), argparse + json (script I/O), pytest (139 → ~179 tests), existing `fbdi/` package and `tools/download_and_clear.py` (unchanged).

---

## File Structure

| Action | Path | Responsibility |
|---|---|---|
| Create | `.claude/skills/fbdi-compare-release/SKILL.md` | 8-stage orchestrator workflow, HITL prompts, invocation triggers |
| Create | `.claude/skills/fbdi-compare-release/scripts/check_env.py` | Stage 1 — Python/deps/Chrome/OS/baselines preflight; JSON + exit codes |
| Create | `.claude/skills/fbdi-compare-release/scripts/verify_download.py` | Stage 3 — diff downloads vs `baseline_files.txt`; first-run bootstrap; inventory-write |
| Create | `.claude/skills/fbdi-compare-release/scripts/summarize_report.py` | Stage 7 — parse Comparison_Report, count changes, top-5 files, Stage 4 timeouts passthrough |
| Create | `.claude/skills/fbdi-compare-release/scripts/verify_run.py` | Stage 8 — run diagnose, check catalog Issues regression thresholds |
| Create | `.claude/skills/fbdi-compare-release/references/troubleshooting.md` | Corrupt xlsm, Selenium failures, Oracle FSM walk-through, scraper silent-failure |
| Create | `.claude/skills/fbdi-compare-release/references/release-version-format.md` | Oracle quarterly naming (YY{A-D}); how to find latest release |
| Create | `tests/test_skill_scripts.py` | Unit tests for the four bundled scripts (~40 tests across Tasks 1–7) |

All four scripts follow the same contract: stdlib + openpyxl only; argparse entrypoint in `main()`; `print(json.dumps(result))` to stdout; exit code communicates branch. Scripts are importable for unit tests (`from scripts.verify_download import parse_inventory, diff_against_inventory, ...`).

**`fbdi/` and `tools/` are not modified.** The skill is pure glue.

---

## Prerequisites (verify once before starting)

- [ ] Current branch has commit `9879aaf` (spec approval) as its head or ancestor:
  ```bash
  git log --oneline -1 docs/superpowers/specs/2026-04-23-fbdi-compare-release-skill-design.md
  ```
  Expected: shows `9879aaf` (or later).

- [ ] Ground-truth artifacts exist on disk for skill-eval #2:
  ```bash
  ls -la Comparison_Report_26A_26B.xlsx FBDI_Master_Catalog.xlsx baselines/26A/originals/ baselines/26B/originals/ | head -5
  ```
  Expected: all four exist; `baselines/26A/originals/` has 212 xlsm files; `baselines/26B/originals/` has 213.

- [ ] Full test suite green before changes:
  ```bash
  python -m pytest tests/ -q
  ```
  Expected: 139 passed.

If any prereq fails, stop and surface to the user.

---

## Task 1: Scaffold skill folder + empty test module

**Files:**
- Create: `.claude/skills/fbdi-compare-release/SKILL.md`
- Create: `.claude/skills/fbdi-compare-release/scripts/__init__.py` (empty — makes scripts importable in tests)
- Create: `tests/test_skill_scripts.py`

- [ ] **Step 1: Write the failing test**

The skill's `scripts/` folder lives at a non-standard path
(`.claude/skills/fbdi-compare-release/scripts/`). Inject it onto `sys.path`
at the top of the test file so `from scripts import check_env` (and later
scripts) resolve.

```python
# tests/test_skill_scripts.py
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
```

- [ ] **Step 2: Run test — verify it fails**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 3 failures (folder doesn't exist).

- [ ] **Step 3: Create the folder skeleton and minimal SKILL.md**

```bash
mkdir -p .claude/skills/fbdi-compare-release/scripts .claude/skills/fbdi-compare-release/references
```

Create `.claude/skills/fbdi-compare-release/scripts/__init__.py` (empty file).

Create `.claude/skills/fbdi-compare-release/SKILL.md` with **minimal** content (full orchestrator comes in Task 10):

```markdown
---
name: fbdi-compare-release
description: Use when Oracle ships a quarterly FBDI release and the user wants the full download → clear → compare → catalog pipeline run. Triggers on phrases like "Oracle released 26C", "compare 26A to 26B", "run the quarterly FBDI update", "update the FBDI Master Catalog", "new FBDI release dropped", "FBDI refresh for Q1". Does NOT trigger on unrelated Python/test-suite questions.
---

# FBDI Compare-Release

Placeholder — full workflow in Task 10.
```

- [ ] **Step 4: Run tests — verify they pass**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/fbdi-compare-release/ tests/test_skill_scripts.py
git commit -m "feat(skill): scaffold fbdi-compare-release skill folder"
```

---

## Task 2: check_env.py (Stage 1 preflight)

**Files:**
- Create: `.claude/skills/fbdi-compare-release/scripts/check_env.py`
- Test: `tests/test_skill_scripts.py`

**Exit codes:**
- `0` = all checks pass
- `1` = one or more fatal checks fail (Python too old, Chrome missing, etc.)
- `2` = deps missing only (skill.md can offer `pip install`)

**JSON stdout shape:**
```json
{
  "ok": false,
  "checks": [
    {"name": "python_version", "ok": true, "detail": "3.14.3"},
    {"name": "os", "ok": true, "detail": "Windows 11"},
    {"name": "deps", "ok": false, "detail": "missing: selenium, webdriver-manager"},
    {"name": "chrome", "ok": true, "detail": "C:\\Program Files\\Google\\Chrome\\..."},
    {"name": "baselines_dir", "ok": true, "detail": "baselines/ exists"},
    {"name": "baseline_files_txt", "ok": true, "detail": "present"}
  ],
  "missing_deps": ["selenium", "webdriver-manager"],
  "fatal": []
}
```

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_skill_scripts.py`:

```python
import json
import subprocess
import sys

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
```

- [ ] **Step 2: Run tests — verify they fail**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 8 new failures (`check_env` module doesn't exist).

- [ ] **Step 3: Implement `check_env.py`**

```python
# .claude/skills/fbdi-compare-release/scripts/check_env.py
"""Stage 1 preflight for fbdi-compare-release.

Checks: OS, Python >= 3.14, required deps importable, Chrome installed,
baselines/ exists, baseline_files.txt present.

JSON stdout. Exit codes: 0=ok, 1=fatal, 2=deps-only-missing.
"""

from __future__ import annotations

import importlib
import json
import os
import platform
import shutil
import sys
from pathlib import Path

REQUIRED_DEPS = ["openpyxl", "selenium", "webdriver_manager", "requests", "pytest"]
MIN_PYTHON = (3, 14)


def check_python_version(current=None) -> dict:
    current = current or sys.version_info[:3]
    ok = tuple(current[:2]) >= MIN_PYTHON
    detail = f"{current[0]}.{current[1]}.{current[2]} (need >= 3.14)"
    return {"name": "python_version", "ok": ok, "detail": detail}


def check_os() -> dict:
    system = platform.system()
    if system == "Windows":
        return {"name": "os", "ok": True, "detail": platform.platform()}
    return {
        "name": "os",
        "ok": True,  # non-fatal warning
        "detail": f"{system} (Windows is the supported platform; proceeding anyway)",
    }


def check_deps(required=None) -> dict:
    required = required or REQUIRED_DEPS
    missing = []
    for name in required:
        try:
            importlib.import_module(name)
        except ImportError:
            missing.append(name)
    if missing:
        return {
            "name": "deps",
            "ok": False,
            "detail": f"missing: {', '.join(missing)}",
            "missing": missing,
        }
    return {"name": "deps", "ok": True, "detail": "all required deps importable"}


def check_chrome() -> dict:
    # Windows default locations
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return {"name": "chrome", "ok": True, "detail": path}
    # Fall back to PATH lookup (covers Mac/Linux dev machines)
    for exe in ("chrome", "google-chrome", "chromium"):
        found = shutil.which(exe)
        if found:
            return {"name": "chrome", "ok": True, "detail": found}
    return {
        "name": "chrome",
        "ok": False,
        "detail": "Google Chrome not found. Install from https://www.google.com/chrome/",
    }


def check_baselines_dir(root: Path) -> dict:
    baselines = root / "baselines"
    if baselines.is_dir():
        return {"name": "baselines_dir", "ok": True, "detail": "baselines/ exists"}
    baselines.mkdir(parents=True, exist_ok=True)
    return {"name": "baselines_dir", "ok": True, "detail": "created baselines/"}


def check_baseline_files_txt(root: Path) -> dict:
    path = root / "baseline_files.txt"
    if path.is_file():
        return {"name": "baseline_files_txt", "ok": True, "detail": "present"}
    return {
        "name": "baseline_files_txt",
        "ok": False,  # non-fatal — caller decides
        "detail": "baseline_files.txt not found; download verification will be limited",
    }


def main(argv=None) -> int:
    root = Path.cwd()
    checks = [
        check_python_version(),
        check_os(),
        check_deps(),
        check_chrome(),
        check_baselines_dir(root),
        check_baseline_files_txt(root),
    ]

    fatal = []
    missing_deps: list[str] = []
    for c in checks:
        if not c["ok"]:
            if c["name"] == "deps":
                missing_deps = c.get("missing", [])
            elif c["name"] == "baseline_files_txt":
                pass  # non-fatal warning
            else:
                fatal.append(c["name"])

    payload = {
        "ok": not fatal and not missing_deps,
        "checks": checks,
        "missing_deps": missing_deps,
        "fatal": fatal,
    }
    print(json.dumps(payload, indent=2))

    if fatal:
        return 1
    if missing_deps:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Run tests — verify they pass**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: all tests pass (8 new).

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/fbdi-compare-release/scripts/check_env.py tests/test_skill_scripts.py
git commit -m "feat(skill): add Stage 1 preflight check_env.py"
```

---

## Task 3: verify_download.py — inventory parser + basic diff

**Files:**
- Create: `.claude/skills/fbdi-compare-release/scripts/verify_download.py` (stub + parser + diff)
- Test: `tests/test_skill_scripts.py`

**`baseline_files.txt` format** (inferred from the committed file):
- Free text preamble.
- One or more release sections, each delimited by:
  ```
  ============================
  26A ORIGINALS (212 files)
  ============================
  <one filename per line, sorted>
  ```
- A trailing `DIFFERENCES` block after the last section header.
- Blank lines and comment lines (anything not ending in `.xlsm`) are ignored inside sections.

**Parser contract:**
- `parse_inventory(text) -> dict[str, list[str]]` — keys are release labels (e.g., `"26A"`), values are sorted filename lists for that release.
- `diff_against_inventory(release, downloaded_names, inventory, manual_files)` returns `{"missing": [...], "extras": [...]}` with manual files excluded from `missing`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_skill_scripts.py`:

```python
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
```

- [ ] **Step 2: Run tests — verify they fail**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 7 failures (`verify_download` doesn't exist yet).

- [ ] **Step 3: Implement the parser + diff**

```python
# .claude/skills/fbdi-compare-release/scripts/verify_download.py
"""Stage 3 download verification for fbdi-compare-release.

Diffs baselines/<ver>/originals/ against the <ver> section of
baseline_files.txt. Handles first-run bootstrap (no <ver> section yet) and
commits an updated inventory on demand.

Exit codes:
    0 = clean (missing == 0, extras == 0)
    1 = missing > 0  (triggers retry / §5 #5 prompt)
    2 = extras only  (triggers §5 #6 prompt)
    3 = first-run bootstrap required (no <ver> section in inventory)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

MANUAL_FILES = ["RapidImplementationForCashManagement.xlsm"]
FIRST_RUN_DELTA_THRESHOLD = 0.15  # 15%, per spec §5 #6

_SECTION_RE = re.compile(r"^(\d{2}[A-D])\s+ORIGINALS\s*\(\d+\s+files?\)\s*$", re.IGNORECASE)


def parse_inventory(text: str) -> dict[str, list[str]]:
    """Parse baseline_files.txt into {release: [filenames...]}.

    Recognizes sections of the form:
        ============================
        26A ORIGINALS (212 files)
        ============================
        <filename>.xlsm
        ...

    Lines not ending in .xlsm are ignored inside sections. Sections end at
    the next `===` delimiter or EOF. A 'DIFFERENCES' section header is not
    an ORIGINALS section and its content is discarded.
    """
    result: dict[str, list[str]] = {}
    current_release: str | None = None
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _SECTION_RE.match(line)
        if m:
            current_release = m.group(1).upper()
            result.setdefault(current_release, [])
            i += 1
            continue
        if line.startswith("==="):
            # Delimiter line — doesn't change state on its own; next non-delim
            # line decides. Sections are terminated by the next SECTION_RE match
            # or a non-.xlsm header block.
            i += 1
            continue
        if current_release is not None and line.lower().endswith(".xlsm"):
            result[current_release].append(line)
        elif current_release is not None and line and not line.lower().endswith(".xlsm"):
            # A non-blank non-.xlsm line inside a section could be a new
            # free-text block (e.g. "DIFFERENCES"). End the current section.
            if line.upper() == "DIFFERENCES" or re.search(r"[A-Za-z]", line) and ":" in line:
                current_release = None
        i += 1
    # Sort each section for deterministic diffs
    for k in result:
        result[k] = sorted(result[k])
    return result


def diff_against_inventory(
    release: str,
    downloaded_names: list[str],
    inventory: dict[str, list[str]],
    manual_files: list[str],
) -> dict:
    """Return {"missing": [...], "extras": [...]}.

    missing = inventory[release] − downloaded − manual_files
    extras  = downloaded − inventory[release]
    """
    expected = set(inventory.get(release.upper(), []))
    actual = set(downloaded_names)
    manual = set(manual_files)

    missing = sorted((expected - actual) - manual)
    extras = sorted(actual - expected)
    return {"missing": missing, "extras": extras}


def list_downloaded(originals_dir: Path) -> list[str]:
    if not originals_dir.is_dir():
        return []
    return sorted(
        p.name for p in originals_dir.iterdir()
        if p.suffix.lower() == ".xlsm" and not p.name.startswith("~$")
    )


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Stage 3 download verification")
    parser.add_argument("--release", required=True, help="Release label, e.g. 26B")
    parser.add_argument(
        "--inventory", type=Path, default=Path("baseline_files.txt"),
        help="Path to baseline_files.txt (default: ./baseline_files.txt)",
    )
    parser.add_argument(
        "--originals", type=Path, default=None,
        help="Path to baselines/<release>/originals/ (default: derived from --release)",
    )
    args = parser.parse_args(argv)

    release = args.release.upper()
    originals = args.originals or (Path("baselines") / release / "originals")
    downloaded = list_downloaded(originals)
    inventory_text = args.inventory.read_text(encoding="utf-8") if args.inventory.is_file() else ""
    inventory = parse_inventory(inventory_text)

    # First-run: no section for this release
    if release not in inventory:
        payload = {
            "release": release,
            "first_run": True,
            "downloaded_count": len(downloaded),
            "downloaded": downloaded,
        }
        print(json.dumps(payload, indent=2))
        return 3

    diff = diff_against_inventory(release, downloaded, inventory, MANUAL_FILES)
    payload = {
        "release": release,
        "first_run": False,
        "downloaded_count": len(downloaded),
        "expected_count": len(inventory[release]),
        "missing": diff["missing"],
        "extras": diff["extras"],
    }
    print(json.dumps(payload, indent=2))

    if diff["missing"]:
        return 1
    if diff["extras"]:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Run tests — verify they pass**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: all new tests pass.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/fbdi-compare-release/scripts/verify_download.py tests/test_skill_scripts.py
git commit -m "feat(skill): add verify_download.py inventory parser + diff"
```

---

## Task 4: verify_download.py — module grouping + first-run delta guard

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/scripts/verify_download.py`
- Test: `tests/test_skill_scripts.py`

Adds two capabilities:
1. `group_missing_by_module(missing) -> dict[str, list[str]]` — rough grouping of missing filenames by Oracle docs module URL, so §5 #5 can surface "which module pages failed to expand".
2. `compute_first_run_delta(downloaded_count, inventory) -> dict` — for the first-run bootstrap case, return `{prior_release, prior_count, delta_pct, over_threshold}` using the most recent prior release in the inventory.

Module grouping is heuristic (filename prefix → module). Based on the four `MODULE_URL_TEMPLATES` in `tools/download_and_clear.py`:
- `project-management` → `Import*`, `Project*`, `Resource*`, `Idea*`, `Lease*`, `Revenue*`
- `financials` → `Payables*`, `Receivables*`, `FixedAsset*`, `Lease*`, `Cash*`, `General*`, `Journal*`, `Account*`, `ChartOf*`, `Daily*`, `AutoInvoice*`, `Cross*`, `Intercompany*`, `Gl*`, `Netting*`, `Tax*`, `Budget*`, `Attachment*`, `Xla*`, `Zx*`
- `procurement` → `PO*`, `Requisition*`, `Supplier*`, `Change*`, `Poi*`, `PONN*`, `Sch*`
- `supply-chain-and-manufacturing` → `Scp*`, `Work*`, `Cse*`, `Maintenance*`, `Mnt*`, `Inventory*`, `Item*`, `Order*`, `Egp*`, `Sus*`, `Vcs*`, `Ship*`, `Source*`, `Production*`, `Perform*`, `Process*`, `Cycle*`, `Dos*`, `Fiscal*`, `Inbound*`, `Interface*`, `Receiv*`, `Requirement*`, `Standard*`, `Upload*`, `IdeaImport*`, `Cost*`, `Discount*`, `Price*`, `Iby*`

Leftovers map to `"other"`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_skill_scripts.py`:

```python
def test_group_missing_by_module_basic():
    groups = verify_download.group_missing_by_module([
        "POBlanketPurchaseAgreementImportTemplate.xlsm",
        "SupplierImportTemplate.xlsm",
        "FixedAssetMassAdditionsImportTemplate.xlsm",
        "ScpItemCostImportTemplate.xlsm",
        "ImportAwards.xlsm",
        "WeirdUnknownFileXYZ.xlsm",
    ])
    assert "POBlanketPurchaseAgreementImportTemplate.xlsm" in groups["procurement"]
    assert "SupplierImportTemplate.xlsm" in groups["procurement"]
    assert "FixedAssetMassAdditionsImportTemplate.xlsm" in groups["financials"]
    assert "ScpItemCostImportTemplate.xlsm" in groups["supply-chain-and-manufacturing"]
    assert "ImportAwards.xlsm" in groups["project-management"]
    assert "WeirdUnknownFileXYZ.xlsm" in groups["other"]


def test_group_missing_by_module_empty():
    assert verify_download.group_missing_by_module([]) == {}


def test_group_missing_by_module_longest_prefix_wins():
    """Regression guard: project-management's 'Import' prefix must not
    swallow financials-specific 'ImportStandaloneFiscal*' filenames. The
    longer prefix should win across module boundaries."""
    groups = verify_download.group_missing_by_module([
        "ImportStandaloneFiscalDocumentTemplate.xlsm",
        "ImportAwards.xlsm",
    ])
    assert "ImportStandaloneFiscalDocumentTemplate.xlsm" in groups["financials"]
    assert "ImportAwards.xlsm" in groups["project-management"]


def test_compute_first_run_delta_within_threshold():
    inventory = {"26A": ["a.xlsm"] * 212, "26B": ["a.xlsm"] * 213}
    result = verify_download.compute_first_run_delta(
        downloaded_count=215,
        inventory=inventory,
    )
    assert result["prior_release"] == "26B"
    assert result["prior_count"] == 213
    assert abs(result["delta_pct"] - (2 / 213)) < 1e-6
    assert result["over_threshold"] is False


def test_compute_first_run_delta_over_threshold():
    inventory = {"26B": ["a.xlsm"] * 213}
    result = verify_download.compute_first_run_delta(
        downloaded_count=107,  # -49.8% — matches the 2026-04-23 module-silent-failure case
        inventory=inventory,
    )
    assert result["prior_release"] == "26B"
    assert result["over_threshold"] is True


def test_compute_first_run_delta_no_prior():
    # Empty inventory = no prior to compare against — non-fatal
    result = verify_download.compute_first_run_delta(downloaded_count=100, inventory={})
    assert result["prior_release"] is None
    assert result["over_threshold"] is False


def test_most_recent_release_sorts_ascii():
    inventory = {"25D": [], "26A": [], "26B": []}
    assert verify_download.most_recent_release(inventory) == "26B"


def test_most_recent_release_empty():
    assert verify_download.most_recent_release({}) is None
```

- [ ] **Step 2: Run tests — verify they fail**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 8 new failures.

- [ ] **Step 3: Add module grouping + delta-guard helpers**

In `verify_download.py`, add near the top after constants:

```python
MODULE_PREFIXES = {
    "project-management": [
        "Import", "Project", "Resource", "Idea", "Lease", "Revenue",
        "FinancialProject", "ExpenseLease",
    ],
    "financials": [
        "Payables", "Receivables", "FixedAsset", "Cash", "General", "Journal",
        "Account", "ChartOf", "Daily", "AutoInvoice", "Cross", "Intercompany",
        "Gl", "Netting", "Tax", "Budget", "Attachment", "Xla", "ZX_",
        "Configurator", "Create", "IbyLegacy", "FiscalDocument",
        "ImportStandaloneFiscal", "InboundFiscal", "UploadCredit", "UploadCustomers",
    ],
    "procurement": [
        "PO", "Requisition", "Supplier", "ChangeOrder", "Poi", "PONN",
        "Sch", "ImportDocumentActions",
    ],
    "supply-chain-and-manufacturing": [
        "Scp", "Work", "Cse", "Maintenance", "Mnt", "Inventory", "Item",
        "Order", "Egp", "Sus", "Vcs", "Ship", "Source", "Production",
        "Perform", "Process", "CycleCount", "Dos", "InterfacedPick",
        "Receiving", "Requirement", "StandardCost", "CostLists",
        "DiscountList", "PriceList", "CustomerImport",
    ],
}


def _module_for_filename(name: str) -> str:
    """Match filename to Oracle module using the longest prefix across all
    modules. Longest-first ordering matters because several modules share a
    common short prefix (e.g. project-management's "Import" would otherwise
    swallow financials-specific "ImportStandaloneFiscal*" files).
    """
    candidates: list[tuple[int, str, str]] = [
        (len(prefix), module, prefix)
        for module, prefixes in MODULE_PREFIXES.items()
        for prefix in prefixes
        if name.startswith(prefix)
    ]
    if not candidates:
        return "other"
    # Longest prefix wins; tie-break by module order is unimportant
    # since same-length collisions across modules don't occur in this set.
    candidates.sort(key=lambda t: -t[0])
    return candidates[0][1]


def group_missing_by_module(missing: list[str]) -> dict[str, list[str]]:
    """Group missing filenames by best-guess Oracle docs module.

    Heuristic prefix-based match. Returns {module: sorted_names}.
    Empty input returns {}.
    """
    if not missing:
        return {}
    groups: dict[str, list[str]] = {}
    for name in missing:
        module = _module_for_filename(name)
        groups.setdefault(module, []).append(name)
    return {k: sorted(v) for k, v in groups.items()}


def most_recent_release(inventory: dict[str, list[str]]) -> str | None:
    """Return the ASCII-max release key from the inventory, or None."""
    if not inventory:
        return None
    return max(inventory.keys())


def compute_first_run_delta(
    downloaded_count: int,
    inventory: dict[str, list[str]],
) -> dict:
    """For the first-run bootstrap case, compare download count to the most
    recent prior release. Returns {prior_release, prior_count, delta_pct,
    over_threshold}. delta_pct is relative ((new-prior)/prior); always
    non-negative (we care about absolute deviation)."""
    prior = most_recent_release(inventory)
    if prior is None or not inventory[prior]:
        return {
            "prior_release": None,
            "prior_count": 0,
            "delta_pct": 0.0,
            "over_threshold": False,
        }
    prior_count = len(inventory[prior])
    delta = abs(downloaded_count - prior_count) / prior_count
    return {
        "prior_release": prior,
        "prior_count": prior_count,
        "delta_pct": delta,
        "over_threshold": delta > FIRST_RUN_DELTA_THRESHOLD,
    }
```

Then update `main()`'s first-run branch to include the delta guard:

```python
    # First-run: no section for this release
    if release not in inventory:
        delta = compute_first_run_delta(len(downloaded), inventory)
        payload = {
            "release": release,
            "first_run": True,
            "downloaded_count": len(downloaded),
            "downloaded": downloaded,
            **delta,
        }
        print(json.dumps(payload, indent=2))
        return 3
```

And update the missing-path branch to include module grouping:

```python
    payload = {
        "release": release,
        "first_run": False,
        "downloaded_count": len(downloaded),
        "expected_count": len(inventory[release]),
        "missing": diff["missing"],
        "extras": diff["extras"],
        "missing_by_module": group_missing_by_module(diff["missing"]),
    }
```

- [ ] **Step 4: Run tests — verify they pass**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: all tests pass (8 new + prior 18 = 26).

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/fbdi-compare-release/scripts/verify_download.py tests/test_skill_scripts.py
git commit -m "feat(skill): add module grouping + first-run delta guard"
```

---

## Task 5: verify_download.py — `--commit-inventory` writer

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/scripts/verify_download.py`
- Test: `tests/test_skill_scripts.py`

Adds `commit_inventory(inventory_text, release, filenames) -> str` that returns the **new full text** of `baseline_files.txt` with the given release's section inserted or replaced. The wrapping CLI gains `--commit-inventory` (writes the file in place).

**Section shape to emit:**
```
============================
<REL> ORIGINALS (<N> files)
============================
<filename>
<filename>
...

```
(trailing blank line preserved).

**Placement rules:**
- If `<REL>` already has a section: replace it in place.
- If not: append after the last existing `ORIGINALS` section, before any `DIFFERENCES` block.
- If no `ORIGINALS` section exists: append at end of file.

The DIFFERENCES block is rewritten from scratch at end of file — set-union of `Only in REL` for each pairwise adjacent release.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_skill_scripts.py`:

```python
def test_commit_inventory_replaces_existing_section():
    inventory_text = INVENTORY_FIXTURE
    new_26b = ["A.xlsm", "B.xlsm"]  # shrunk from 4 to 2
    result = verify_download.commit_inventory(
        inventory_text, release="26B", filenames=new_26b,
    )
    parsed = verify_download.parse_inventory(result)
    assert parsed["26B"] == ["A.xlsm", "B.xlsm"]
    assert parsed["26A"] == [
        "AccountCombinationsImportTemplate.xlsm",
        "BudgetImportTemplate.xlsm",
        "RapidImplementationForCashManagement.xlsm",
    ]
    assert "26B ORIGINALS (2 files)" in result


def test_commit_inventory_appends_new_section():
    inventory_text = INVENTORY_FIXTURE
    filenames_26c = ["NewFileA.xlsm", "NewFileB.xlsm", "NewFileC.xlsm"]
    result = verify_download.commit_inventory(
        inventory_text, release="26C", filenames=filenames_26c,
    )
    parsed = verify_download.parse_inventory(result)
    assert parsed["26C"] == sorted(filenames_26c)
    assert "26C ORIGINALS (3 files)" in result
    # 26B section must still be present and unchanged
    assert parsed["26B"] == sorted([
        "AccountCombinationsImportTemplate.xlsm",
        "BudgetImportTemplate.xlsm",
        "ItemImportReferenceOrgTemplate.xlsm",
        "RapidImplementationForCashManagement.xlsm",
    ])


def test_commit_inventory_sorts_filenames_ascii():
    result = verify_download.commit_inventory(
        INVENTORY_FIXTURE, release="26C",
        filenames=["Zebra.xlsm", "AAA.xlsm", "Middle.xlsm"],
    )
    idx_aaa = result.index("AAA.xlsm")
    idx_middle = result.index("Middle.xlsm")
    idx_zebra = result.index("Zebra.xlsm")
    assert idx_aaa < idx_middle < idx_zebra


def test_commit_inventory_cli_writes_file_in_place(tmp_path):
    inv_path = tmp_path / "baseline_files.txt"
    inv_path.write_text(INVENTORY_FIXTURE, encoding="utf-8")

    originals = tmp_path / "baselines" / "26C" / "originals"
    originals.mkdir(parents=True)
    for name in ("A.xlsm", "B.xlsm"):
        (originals / name).touch()

    exit_code = verify_download.main([
        "--release", "26C",
        "--inventory", str(inv_path),
        "--originals", str(originals),
        "--commit-inventory",
    ])
    assert exit_code == 0
    parsed = verify_download.parse_inventory(inv_path.read_text(encoding="utf-8"))
    assert parsed["26C"] == ["A.xlsm", "B.xlsm"]
```

- [ ] **Step 2: Run tests — verify they fail**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 4 new failures.

- [ ] **Step 3: Implement `commit_inventory`**

Add to `verify_download.py`:

```python
def _format_section(release: str, filenames: list[str]) -> str:
    sorted_names = sorted(filenames)
    banner = "=" * 28
    lines = [
        banner,
        f"{release} ORIGINALS ({len(sorted_names)} files)",
        banner,
        *sorted_names,
        "",
    ]
    return "\n".join(lines) + "\n"


def _strip_differences_block(text: str) -> str:
    """Remove any existing DIFFERENCES block (we'll regenerate it)."""
    return re.sub(
        r"={20,}\s*\nDIFFERENCES\s*\n={20,}\s*\n(?:.*\n?)*\Z",
        "",
        text,
        flags=re.IGNORECASE | re.MULTILINE,
    ).rstrip() + "\n"


def _render_differences(inventory: dict[str, list[str]]) -> str:
    """Regenerate the DIFFERENCES footer from the current inventory.

    For each pair of adjacent releases (by ASCII sort), emit
    'Only in <NEW>: <comma-list>'. Simple and mirrors the existing file.
    """
    releases = sorted(inventory.keys())
    if len(releases) < 2:
        return ""
    banner = "=" * 28
    lines = [banner, "DIFFERENCES", banner]
    for i in range(1, len(releases)):
        prev, cur = releases[i - 1], releases[i]
        only_in_cur = sorted(set(inventory[cur]) - set(inventory[prev]))
        only_in_prev = sorted(set(inventory[prev]) - set(inventory[cur]))
        if only_in_cur:
            lines.append(f"Only in {cur}: {', '.join(only_in_cur)}")
        if only_in_prev:
            lines.append(f"Only in {prev}: {', '.join(only_in_prev)}")
    return "\n".join(lines) + "\n"


def commit_inventory(
    inventory_text: str, release: str, filenames: list[str],
) -> str:
    """Return a new inventory text with release's section inserted or replaced,
    and the DIFFERENCES footer regenerated."""
    release = release.upper()
    inventory = parse_inventory(inventory_text)
    inventory[release] = sorted(filenames)

    # Strip old DIFFERENCES
    text = _strip_differences_block(inventory_text)

    # Replace existing section for `release` if present
    section_pattern = re.compile(
        r"={20,}\s*\n" + re.escape(release) + r"\s+ORIGINALS\s*\([^)]*\)\s*\n={20,}\s*\n"
        r"(?:[^\n]*\n)*?"
        r"(?=(?:={20,}\s*\n(?:\d{2}[A-D]\s+ORIGINALS|DIFFERENCES))|\Z)",
        re.IGNORECASE,
    )
    new_section = _format_section(release, inventory[release])
    if section_pattern.search(text):
        text = section_pattern.sub(new_section, text, count=1)
    else:
        # Append after last ORIGINALS section (or at end if none)
        text = text.rstrip() + "\n\n" + new_section

    # Re-append DIFFERENCES footer
    diffs = _render_differences(inventory)
    if diffs:
        text = text.rstrip() + "\n\n" + diffs
    return text
```

Replace the entire `main()` function in `verify_download.py` with the following. The `--commit-inventory` branch is placed **before** the first-run and diff branches so it short-circuits.

```python
def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Stage 3 download verification")
    parser.add_argument("--release", required=True, help="Release label, e.g. 26B")
    parser.add_argument(
        "--inventory", type=Path, default=Path("baseline_files.txt"),
        help="Path to baseline_files.txt (default: ./baseline_files.txt)",
    )
    parser.add_argument(
        "--originals", type=Path, default=None,
        help="Path to baselines/<release>/originals/ (default: derived from --release)",
    )
    parser.add_argument(
        "--commit-inventory", action="store_true",
        help="Rewrite inventory to match the downloaded files for --release",
    )
    args = parser.parse_args(argv)

    release = args.release.upper()
    originals = args.originals or (Path("baselines") / release / "originals")
    downloaded = list_downloaded(originals)
    inventory_text = args.inventory.read_text(encoding="utf-8") if args.inventory.is_file() else ""
    inventory = parse_inventory(inventory_text)

    # Short-circuit: --commit-inventory rewrites baseline_files.txt and exits 0.
    if args.commit_inventory:
        new_text = commit_inventory(inventory_text, release, downloaded)
        args.inventory.write_text(new_text, encoding="utf-8")
        payload = {
            "release": release,
            "committed": True,
            "count": len(downloaded),
            "inventory_path": str(args.inventory),
        }
        print(json.dumps(payload, indent=2))
        return 0

    # First-run: no section for this release
    if release not in inventory:
        delta = compute_first_run_delta(len(downloaded), inventory)
        payload = {
            "release": release,
            "first_run": True,
            "downloaded_count": len(downloaded),
            "downloaded": downloaded,
            **delta,
        }
        print(json.dumps(payload, indent=2))
        return 3

    diff = diff_against_inventory(release, downloaded, inventory, MANUAL_FILES)
    payload = {
        "release": release,
        "first_run": False,
        "downloaded_count": len(downloaded),
        "expected_count": len(inventory[release]),
        "missing": diff["missing"],
        "extras": diff["extras"],
        "missing_by_module": group_missing_by_module(diff["missing"]),
    }
    print(json.dumps(payload, indent=2))

    if diff["missing"]:
        return 1
    if diff["extras"]:
        return 2
    return 0
```

- [ ] **Step 4: Run tests — verify they pass**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Smoke-test against the real `baseline_files.txt`**

```bash
python -m scripts.verify_download --release 26A
```

Expected: exits 0 (clean), JSON shows `"downloaded_count": 212, "expected_count": 212, "missing": [], "extras": []`.

```bash
python -m scripts.verify_download --release 26B
```

Expected: exits 0, `downloaded_count: 213, expected_count: 213`.

Note: run from the skill directory so the `scripts` package resolves. Alternatively: `python .claude/skills/fbdi-compare-release/scripts/verify_download.py --release 26A`.

- [ ] **Step 6: Commit**

```bash
git add .claude/skills/fbdi-compare-release/scripts/verify_download.py tests/test_skill_scripts.py
git commit -m "feat(skill): add --commit-inventory + DIFFERENCES regen"
```

---

## Task 6: summarize_report.py (Stage 7)

**Files:**
- Create: `.claude/skills/fbdi-compare-release/scripts/summarize_report.py`
- Test: `tests/test_skill_scripts.py`

**What it prints:**
- Total change rows.
- Distinct files-with-changes count.
- Top-5 most-changed files with counts.
- Paths to the two output xlsx (`Comparison_Report_<OLD>_<NEW>.xlsx`, `FBDI_Master_Catalog.xlsx`).
- Any Stage 4 timeout filenames passed via `--timeouts` (so user knows which blanks need manual clearing).

**Comparison_Report shape** (confirmed from disk via `load_workbook`):
- Single `Sheet1`, header row: `('FBDI File', 'FBDI Tab', 'Column Letter', 'Column Number', 'Old FBDI Field Name', 'New FBDI Field Name', 'Difference?')`.
- Each data row = one field-level change; `FBDI File` column (col A) identifies which file the change belongs to.

JSON stdout shape:
```json
{
  "report_path": "Comparison_Report_26A_26B.xlsx",
  "catalog_path": "FBDI_Master_Catalog.xlsx",
  "total_changes": 706,
  "files_with_changes": 19,
  "top_files": [
    {"file": "ConfiguratorRedwoodRuleConversionTemplate", "changes": 112},
    ...
  ],
  "stage4_timeouts": ["PayablesCollectionDocuments.xlsm"]
}
```

Exit code: always 0 (summary is informational, never fatal).

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_skill_scripts.py`:

```python
from openpyxl import Workbook
from scripts import summarize_report  # noqa: E402


def _make_comparison_report(path, rows):
    """rows = list of (fbdi_file, fbdi_tab, col_letter, col_num, old, new, diff)."""
    wb = Workbook()
    ws = wb.active
    ws.append(["FBDI File", "FBDI Tab", "Column Letter", "Column Number",
               "Old FBDI Field Name", "New FBDI Field Name", "Difference?"])
    for row in rows:
        ws.append(list(row))
    wb.save(path)
    wb.close()


def test_summarize_counts_changes(tmp_path):
    path = tmp_path / "cmp.xlsx"
    _make_comparison_report(path, [
        ("FileA", "Tab1", "A", 1, "old1", "new1", "YES"),
        ("FileA", "Tab1", "B", 2, "old2", "new2", "YES"),
        ("FileB", "Tab1", "A", 1, "old3", "new3", "YES"),
    ])
    result = summarize_report.summarize(path)
    assert result["total_changes"] == 3
    assert result["files_with_changes"] == 2


def test_summarize_top_files_ordered(tmp_path):
    path = tmp_path / "cmp.xlsx"
    rows = (
        [("FileB", "T", "A", 1, "o", "n", "YES")] * 10
        + [("FileA", "T", "A", 1, "o", "n", "YES")] * 5
        + [("FileC", "T", "A", 1, "o", "n", "YES")] * 3
    )
    _make_comparison_report(path, rows)
    result = summarize_report.summarize(path)
    assert [t["file"] for t in result["top_files"]][:3] == ["FileB", "FileA", "FileC"]
    assert result["top_files"][0]["changes"] == 10


def test_summarize_top_files_capped_at_5(tmp_path):
    path = tmp_path / "cmp.xlsx"
    rows = [(f"File{i}", "T", "A", 1, "o", "n", "YES") for i in range(10)]
    _make_comparison_report(path, rows)
    result = summarize_report.summarize(path)
    assert len(result["top_files"]) <= 5


def test_summarize_empty_report(tmp_path):
    path = tmp_path / "cmp.xlsx"
    _make_comparison_report(path, [])
    result = summarize_report.summarize(path)
    assert result["total_changes"] == 0
    assert result["files_with_changes"] == 0
    assert result["top_files"] == []


def test_summarize_cli_passthrough_timeouts(tmp_path):
    path = tmp_path / "cmp.xlsx"
    _make_comparison_report(path, [])
    exit_code = summarize_report.main([
        "--report", str(path),
        "--catalog", "dummy.xlsx",
        "--timeouts", "Foo.xlsm,Bar.xlsm",
    ])
    assert exit_code == 0


def test_summarize_against_ground_truth():
    """Spec §8 eval #2 reference: 26A→26B run produced 706 changes in 19 files."""
    report = Path("Comparison_Report_26A_26B.xlsx")
    if not report.is_file():
        import pytest
        pytest.skip("ground-truth report not present")
    result = summarize_report.summarize(report)
    assert result["total_changes"] == 706
    assert result["files_with_changes"] == 19
```

- [ ] **Step 2: Run tests — verify they fail**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 6 new failures (the last is an auto-skip without the file; on Brad's machine it will run).

- [ ] **Step 3: Implement `summarize_report.py`**

```python
# .claude/skills/fbdi-compare-release/scripts/summarize_report.py
"""Stage 7 summary for fbdi-compare-release.

Reads Comparison_Report_<OLD>_<NEW>.xlsx and prints a JSON summary:
total changes, distinct files with changes, top-5 most-changed files.
Accepts Stage 4 timeout filenames via --timeouts for inclusion in the
summary (so the user knows which blanks files need manual clearing).
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook


def summarize(report_path: Path) -> dict:
    report_path = Path(report_path)
    wb = load_workbook(report_path, read_only=True, data_only=True)
    ws = wb.active
    counter: Counter[str] = Counter()
    total = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue  # skip header
        if row and row[0]:
            counter[row[0]] += 1
            total += 1
    wb.close()

    top = [{"file": name, "changes": n} for name, n in counter.most_common(5)]
    return {
        "total_changes": total,
        "files_with_changes": len(counter),
        "top_files": top,
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Stage 7 summary")
    parser.add_argument(
        "--report", type=Path, required=True,
        help="Path to Comparison_Report_<OLD>_<NEW>.xlsx",
    )
    parser.add_argument(
        "--catalog", type=Path, default=Path("FBDI_Master_Catalog.xlsx"),
        help="Path to FBDI_Master_Catalog.xlsx",
    )
    parser.add_argument(
        "--timeouts", type=str, default="",
        help="Comma-separated Stage 4 timeout filenames (optional)",
    )
    args = parser.parse_args(argv)

    summary = summarize(args.report)
    payload = {
        "report_path": str(args.report),
        "catalog_path": str(args.catalog),
        **summary,
        "stage4_timeouts": [t for t in args.timeouts.split(",") if t.strip()],
    }
    print(json.dumps(payload, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Run tests — verify they pass**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: all pass.

- [ ] **Step 5: Smoke-test against ground-truth**

```bash
python .claude/skills/fbdi-compare-release/scripts/summarize_report.py \
  --report Comparison_Report_26A_26B.xlsx
```

Expected: JSON with `total_changes: 706`, `files_with_changes: 19`.

- [ ] **Step 6: Commit**

```bash
git add .claude/skills/fbdi-compare-release/scripts/summarize_report.py tests/test_skill_scripts.py
git commit -m "feat(skill): add Stage 7 summarize_report.py"
```

---

## Task 7: verify_run.py (Stage 8)

**Files:**
- Create: `.claude/skills/fbdi-compare-release/scripts/verify_run.py`
- Test: `tests/test_skill_scripts.py`

**Two checks:**

1. **Diagnose regression** — run `python -m fbdi diagnose --release <NEW>` via subprocess, read the Diagnostic_Report_*.xlsx output. Flag if `NO_HEADER > 0` for the new release (regression — CLAUDE.md documents the "Phase 3 resolved NO_HEADER: 0" hazard).

2. **Catalog Issues regression** — read `FBDI_Master_Catalog.xlsx` Issues tab, filter rows where `release == <NEW>`, compare count to the prior release's count. Flag if `new > 2 * prior` OR `new - prior > 50` (spec §6).

Exit codes:
- `0` = no regressions
- `1` = regressions present (warnings only — does not block, but SKILL.md surfaces them in the final summary)

JSON stdout shape:
```json
{
  "release": "26B",
  "diagnose": {
    "no_header_count": 0,
    "file_error_count": 11,
    "regression": false
  },
  "catalog_issues": {
    "release_issue_count": 5,
    "prior_release": "26A",
    "prior_issue_count": 4,
    "regression": false,
    "threshold": {"multiplier": 2.0, "absolute": 50}
  },
  "overall_regression": false
}
```

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_skill_scripts.py`:

```python
from scripts import verify_run  # noqa: E402


def _make_catalog_with_issues(path, issues_by_release):
    """issues_by_release: {release: [(file, tab, issue_type, detail), ...]}"""
    wb = Workbook()
    # Remove default + add per-release tabs (any content, we only read Issues)
    wb.remove(wb.active)
    for release in issues_by_release:
        wb.create_sheet(release)
    issues_ws = wb.create_sheet("Issues")
    issues_ws.append(["release", "file", "tab", "issue_type", "detail"])
    for release, rows in issues_by_release.items():
        for row in rows:
            issues_ws.append([release, *row])
    wb.create_sheet("Drift")
    wb.save(path)
    wb.close()


def test_verify_run_catalog_check_no_regression(tmp_path):
    catalog = tmp_path / "cat.xlsx"
    _make_catalog_with_issues(catalog, {
        "26A": [("F", "T", "TYPE_PARSE_WARNING", "x")] * 4,
        "26B": [("F", "T", "TYPE_PARSE_WARNING", "x")] * 5,
    })
    result = verify_run.check_catalog_issues(catalog, release="26B")
    assert result["release_issue_count"] == 5
    assert result["prior_issue_count"] == 4
    assert result["regression"] is False


def test_verify_run_catalog_check_regression_2x(tmp_path):
    catalog = tmp_path / "cat.xlsx"
    _make_catalog_with_issues(catalog, {
        "26A": [("F", "T", "TYPE_PARSE_WARNING", "x")] * 4,
        "26B": [("F", "T", "TYPE_PARSE_WARNING", "x")] * 10,
    })
    result = verify_run.check_catalog_issues(catalog, release="26B")
    assert result["regression"] is True


def test_verify_run_catalog_check_regression_absolute(tmp_path):
    catalog = tmp_path / "cat.xlsx"
    _make_catalog_with_issues(catalog, {
        "26A": [("F", "T", "TYPE_PARSE_WARNING", "x")] * 100,
        "26B": [("F", "T", "TYPE_PARSE_WARNING", "x")] * 160,
    })
    result = verify_run.check_catalog_issues(catalog, release="26B")
    # delta = 60, >50 → regression even though <2x
    assert result["regression"] is True


def test_verify_run_catalog_check_no_prior(tmp_path):
    catalog = tmp_path / "cat.xlsx"
    _make_catalog_with_issues(catalog, {
        "26B": [("F", "T", "TYPE_PARSE_WARNING", "x")] * 5,
    })
    result = verify_run.check_catalog_issues(catalog, release="26B")
    assert result["prior_release"] is None
    assert result["regression"] is False
```

(The diagnose-check is exercised against the real ground-truth in the smoke step; mocking subprocess calls here would just duplicate `fbdi.diagnose`'s own test coverage.)

- [ ] **Step 2: Run tests — verify they fail**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: 4 new failures.

- [ ] **Step 3: Implement `verify_run.py`**

```python
# .claude/skills/fbdi-compare-release/scripts/verify_run.py
"""Stage 8 post-run verification for fbdi-compare-release.

- Runs `python -m fbdi diagnose --release <ver>` and parses the Diagnostic
  xlsx output for NO_HEADER regressions.
- Reads FBDI_Master_Catalog.xlsx Issues tab, filters by release, and flags
  catalog Issues-tab regression if:
      new_count > 2 * prior_count   OR   new_count - prior_count > 50

Never blocks. Exit 0 = clean, 1 = regression detected.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

ISSUE_MULTIPLIER_THRESHOLD = 2.0
ISSUE_ABSOLUTE_THRESHOLD = 50


def check_catalog_issues(catalog_path: Path, release: str) -> dict:
    """Read Issues tab, group by release, compare release against most-recent prior."""
    release = release.upper()
    wb = load_workbook(catalog_path, read_only=True, data_only=True)
    if "Issues" not in wb.sheetnames:
        wb.close()
        return {
            "release_issue_count": 0,
            "prior_release": None,
            "prior_issue_count": 0,
            "regression": False,
            "threshold": {"multiplier": ISSUE_MULTIPLIER_THRESHOLD,
                          "absolute": ISSUE_ABSOLUTE_THRESHOLD},
        }
    ws = wb["Issues"]
    counter: Counter[str] = Counter()
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not row or not row[0]:
            continue
        counter[str(row[0]).upper()] += 1
    wb.close()

    release_count = counter.get(release, 0)
    priors = sorted(r for r in counter if r < release)
    prior = priors[-1] if priors else None
    prior_count = counter.get(prior, 0) if prior else 0

    regression = False
    if prior and prior_count > 0:
        if release_count > ISSUE_MULTIPLIER_THRESHOLD * prior_count:
            regression = True
        if release_count - prior_count > ISSUE_ABSOLUTE_THRESHOLD:
            regression = True

    return {
        "release_issue_count": release_count,
        "prior_release": prior,
        "prior_issue_count": prior_count,
        "regression": regression,
        "threshold": {"multiplier": ISSUE_MULTIPLIER_THRESHOLD,
                      "absolute": ISSUE_ABSOLUTE_THRESHOLD},
    }


def run_diagnose(release: str, repo_root: Path) -> dict:
    """Invoke `python -m fbdi diagnose --release <release>` and parse its output xlsx."""
    diagnostic_path = repo_root / f"Diagnostic_Report_{release.upper()}.xlsx"
    if diagnostic_path.is_file():
        diagnostic_path.unlink()

    proc = subprocess.run(
        [sys.executable, "-m", "fbdi", "diagnose", "--release", release],
        cwd=repo_root,
        capture_output=True, text=True,
    )
    if proc.returncode != 0 or not diagnostic_path.is_file():
        return {
            "no_header_count": None,
            "file_error_count": None,
            "regression": False,
            "error": f"diagnose invocation failed: {proc.stderr[:500]}",
        }

    wb = load_workbook(diagnostic_path, read_only=True, data_only=True)
    ws = wb.active
    no_header = 0
    file_error = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not row:
            continue
        result = row[2]  # "Detection Result" column
        if result == "NO_HEADER":
            no_header += 1
        elif result == "FILE_ERROR":
            file_error += 1
    wb.close()

    return {
        "no_header_count": no_header,
        "file_error_count": file_error,
        "regression": no_header > 0,
        "diagnostic_path": str(diagnostic_path),
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Stage 8 post-run verification")
    parser.add_argument("--release", required=True)
    parser.add_argument(
        "--catalog", type=Path, default=Path("FBDI_Master_Catalog.xlsx"),
    )
    parser.add_argument(
        "--repo-root", type=Path, default=Path.cwd(),
        help="Repo root for fbdi diagnose invocation",
    )
    parser.add_argument(
        "--skip-diagnose", action="store_true",
        help="Skip the diagnose subprocess (for unit tests / quick runs)",
    )
    args = parser.parse_args(argv)

    release = args.release.upper()

    diag = {"skipped": True}
    if not args.skip_diagnose:
        diag = run_diagnose(release, args.repo_root)

    cat = check_catalog_issues(args.catalog, release)

    overall = bool(diag.get("regression")) or bool(cat.get("regression"))
    payload = {
        "release": release,
        "diagnose": diag,
        "catalog_issues": cat,
        "overall_regression": overall,
    }
    print(json.dumps(payload, indent=2))
    return 1 if overall else 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Run tests — verify they pass**

```bash
python -m pytest tests/test_skill_scripts.py -v
```

Expected: all pass.

- [ ] **Step 5: Smoke-test against ground-truth**

```bash
python .claude/skills/fbdi-compare-release/scripts/verify_run.py \
  --release 26B --skip-diagnose
```

Expected: JSON shows `release_issue_count: 5, prior_release: "26A", prior_issue_count: 4, regression: false`. (Counts derived from the 9-row committed `FBDI_Master_Catalog.xlsx` Issues tab: 5 × 26B + 4 × 26A.)

- [ ] **Step 6: Commit**

```bash
git add .claude/skills/fbdi-compare-release/scripts/verify_run.py tests/test_skill_scripts.py
git commit -m "feat(skill): add Stage 8 verify_run.py"
```

---

## Task 8: references/troubleshooting.md

**Files:**
- Create: `.claude/skills/fbdi-compare-release/references/troubleshooting.md`

Read-on-demand reference — loaded by Claude only when a matching failure arises. Covers the four hazards documented in CLAUDE.md + docs/scraper-gap-findings-2026-04-23.md.

- [ ] **Step 1: Write `troubleshooting.md`**

```markdown
# FBDI Compare-Release — Troubleshooting

Load this only when a stage fails or the user reports an anomaly. The
symptoms below are the ones documented in `CLAUDE.md` and
`docs/scraper-gap-findings-2026-04-23.md`; anything outside this list is
an unknown — escalate to Brad rather than guessing.

---

## Stage 3 — Download ran but produced zero files / short counts

**Symptom:** `scripts/verify_download.py` reports many missing files from a
single Oracle module URL (e.g., all procurement files absent), and
`download_and_clear.py` logs show `Navigating to ...` immediately followed
by `Completed: ...` with no "Downloading:" lines between them.

**Cause:** Silent module-page failure in Oracle's JET-rendered
navigation drawer. The `navigationDrawer` element loads, but section
expansion yields no children before the scraper harvests links. Observed
on 2026-04-23 (first run on Brad's Windows machine, 26B); resolved on
retry.

**Fix:** Retry `download_and_clear.py <ver>` once (per spec §5 #5, auto).
If the retry still shorts, `MODULE_URL_TEMPLATES` in
`tools/download_and_clear.py:44-49` may need updating — Oracle may have
restructured the docs URL pattern for that module. Do not attempt to
debug from within the skill; defer to Brad.

**Retry cap:** 3 total download attempts per run. After the 3rd
verification failure, option (a) "retry again" is no longer offered — only
abort or proceed-with-gaps remain.

---

## Stage 3 — `RapidImplementationForCashManagement.xlsm` is missing

**Cause:** This is an Oracle Fusion FSM (Functional Setup Manager)
template, not a standard FBDI template hosted on docs pages. The scraper
cannot fetch it — it must be obtained manually from Oracle Fusion.

**Walk-through (paste verbatim into the user prompt):**

1. Log into Oracle Fusion.
2. Navigate: **Setup and Maintenance** → click the **hamburger menu**
   (top-right) → **Search**.
3. Search for: `Create Banks, Branches, and Accounts in Spreadsheet`.
4. Click the task — the template downloads directly.
5. Place it in `baselines/<new_release>/originals/`.

**Fast alternative:** copy from the prior release's `originals/` folder.
Oracle rarely updates this template, and `sha256sum` confirms the 26A and
26B copies were bit-identical on 2026-04-23.

---

## Stage 4 — Per-file clear timed out

**Symptom:** `download_and_clear.py --clear-only` prints:
```
*** TIMED OUT (N files, >120s each) — clear these manually: ***
    PayablesCollectionDocuments.xlsm (9,xxxKB)
```

**Cause:** Large xlsm files (typically >8MB) can exceed the 120s per-file
subprocess timeout in `tools/download_and_clear.py:244-281`. Known
recurring offender on 2026-04-23: `PayablesCollectionDocuments.xlsm`
(~9MB, timed out in both 26A and 26B).

**Impact:** **Not a blocker for Stage 5.** The comparison engine reads
`originals/` (not `blanks/`). The timeout only affects whether the
`blanks/<ver>/` folder has a cleared copy of that file for downstream
client use.

**Fix (optional):** Clear the file manually — either in Excel (select all
data rows below the header, delete, save) or via the legacy VBA macro
`reference/Clear_FBDIs - 20210412.xlsm`.

**Flag in final summary:** SKILL.md Stage 7 surfaces the timeout filename
list in the terminal summary so the user knows what needs manual
clearing.

---

## Stage 5 — Per-pair compare failure

**Symptom:** Compare output includes a file pair where one side's xlsm
cannot be opened (corrupt xml, phantom `max_column=16384`, etc.).

**Cause:** Oracle occasionally ships xlsm files with corrupt metadata.
`compare.py` uses subprocess-per-pair isolation with a 120s timeout; a
failure here does not abort the run. Historically ~11 FILE_ERROR files
exist in 26B.

**Fix:** None available from within the skill. Failures are collected and
surfaced at end of Stage 5 per spec §5 #4. If >5 failures, the skill
pauses for user input (retry / skip / abort).

---

## Stage 6 — Catalog Issues tab jumped

**Symptom:** `verify_run.py` reports `catalog_issues.regression: true`
— the new release has either >2× the prior release's Issues count, or
>50 more absolute.

**Likely causes:**
- Oracle changed how data-type strings are encoded in a new release
  (e.g., new temporal format mask not covered by `fbdi/type_parser.py`).
- A large new tab was added with no header and produces `NO_HEADER` rows.

**Fix:** Inspect the Issues tab for the new release — if all new entries
are `TYPE_PARSE_WARNING` with a common pattern, extend
`fbdi/type_parser.py` (see the 2026-04-20 fix: temporal format masks,
trailing-period typos). This is `fbdi/`-package work, not skill work —
defer to Brad.

**Does not block:** the regression is flagged in the summary; the user
decides whether to defer.

---

## Selenium / Chrome

**`chromedriver` version mismatch:** `webdriver-manager` auto-installs a
matching chromedriver for the installed Chrome. If it fails, upgrade
Chrome to the latest stable and re-run.

**Chrome not found:** `check_env.py` reports `"chrome": {"ok": false}`.
Install Chrome from https://www.google.com/chrome/.
```

- [ ] **Step 2: Commit**

```bash
git add .claude/skills/fbdi-compare-release/references/troubleshooting.md
git commit -m "docs(skill): add troubleshooting reference"
```

---

## Task 9: references/release-version-format.md

**Files:**
- Create: `.claude/skills/fbdi-compare-release/references/release-version-format.md`

- [ ] **Step 1: Write `release-version-format.md`**

```markdown
# Oracle Quarterly Release Version Format

Oracle Cloud Fusion Applications ships quarterly releases. Each release
has a two-character label: `YYx` where `YY` is a two-digit year and `x` is
one of `A`, `B`, `C`, `D` (Q1–Q4 respectively).

## Examples

| Label | Period                |
|-------|-----------------------|
| 25A   | Feb 2025 – Apr 2025  |
| 25B   | May 2025 – Jul 2025  |
| 25C   | Aug 2025 – Oct 2025  |
| 25D   | Nov 2025 – Jan 2026  |
| 26A   | Feb 2026 – Apr 2026  |
| 26B   | May 2026 – Jul 2026  |

Quarterly cadence is stable — Oracle has not skipped or renamed these in
the years leading up to this skill's authorship (2026-04).

## Canonical form

In this repo, release labels are **uppercase** everywhere they are
user-visible (folders `baselines/26A/`, tab names in catalog workbook,
`baseline_files.txt` section headers, comparison report filenames). The
CLI accepts any case and upper-cases internally.

## How to find the latest Oracle release

1. Visit https://docs.oracle.com/en/cloud/saas/ — Oracle's Cloud SaaS
   landing page.
2. Pick a module (Financials, Procurement, Project Management, Supply
   Chain) — each lists "What's New" for the current release at the top.
3. The current-release URL pattern is:
   `https://docs.oracle.com/en/cloud/saas/<module>/<release_lowercase>/oe<code>/index.html`
   (e.g., `.../financials/26b/oefbf/index.html`).

## Expected release count (historical)

| Release | File count (originals) |
|---------|------------------------|
| 26A     | 212                    |
| 26B     | 213 (added `ItemImportReferenceOrgTemplate.xlsm`) |

Oracle rarely adds or removes more than ~5–10 templates in a quarterly
release — §5 #6's 15% delta guard in `verify_download.py` catches larger
swings on the first run of a new release.
```

- [ ] **Step 2: Commit**

```bash
git add .claude/skills/fbdi-compare-release/references/release-version-format.md
git commit -m "docs(skill): add release-version-format reference"
```

---

## Task 10: SKILL.md — full orchestrator workflow

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/SKILL.md`

Replaces the placeholder from Task 1 with the full 8-stage orchestrator. Target ~300 lines per spec §3.

- [ ] **Step 1: Replace `SKILL.md` with the full workflow**

```markdown
---
name: fbdi-compare-release
description: "Use when Oracle ships a quarterly FBDI release and the user wants the full download → clear → compare → catalog pipeline run end-to-end. Triggers on phrases like 'Oracle released 26C', 'compare 26A to 26B', 'run the quarterly FBDI update', 'update the FBDI Master Catalog for 26B', 'new FBDI release dropped', 'FBDI refresh for Q1'. Does NOT trigger on unrelated questions like 'what's the current Python version' or 'run the test suite'."
---

# FBDI Compare-Release Orchestrator

You are orchestrating an 8-stage pipeline that takes a coworker from
"Oracle released 26C" to a finished `Comparison_Report_<OLD>_<NEW>.xlsx`
plus an updated `FBDI_Master_Catalog.xlsx`. You do **not** re-implement
comparison logic — the existing `fbdi/` package and
`tools/download_and_clear.py` do the work. Your job is glue, checkpoints,
and human-in-the-loop prompts.

**Expected wall time:** 35–50 minutes for a full run (downloads dominate).

## Before you start

- [ ] Parse the user's invocation for explicit versions (`--old 26A --new 26B`).
      If absent, Stage 2 will auto-detect and confirm.
- [ ] **Tell the user, up front:** "This will take 35–50 minutes. Downloads
      dominate. If you plan to step away, disable Windows sleep/lock for
      the duration — the Selenium process runs foreground and will
      suspend with the OS."

## Stage 1 — Environment preflight

Run:

```
python .claude/skills/fbdi-compare-release/scripts/check_env.py
```

Interpret the exit code:

- `0` → proceed to Stage 2.
- `1` (fatal) → surface the `fatal` list. If `python_version` is fatal,
  quote the install hint for `pyenv-win`. If `chrome` is fatal, give the
  download URL `https://www.google.com/chrome/`. Do not proceed.
- `2` (deps missing) → ask the user: "Missing deps: <list>. Run
  `pip install -r requirements.txt`?" If yes, run it, then re-run
  `check_env.py`. If no, stop.

If `baseline_files_txt` is `ok: false`, warn the user but continue — Stage 3
will degrade gracefully.

## Stage 2 — Resolve OLD and NEW releases

If both were passed explicitly (`--old 26A --new 26B`), skip to the
HITL #3 version-mismatch check below.

Otherwise, auto-detect:
- List `baselines/*/` folders, filter to those matching `^\d{2}[A-D]$`,
  sort ASCII-descending. The most recent is the prospective `OLD`.
- Infer `NEW` as "the release the user just mentioned" from their prompt,
  or the next quarter after `OLD` if they didn't say.

**HITL #3 — version-mismatch sanity:** If the auto-detected `OLD` doesn't
match what the user said (e.g., user said 26C but newest baseline is 25D,
implying they want to skip 26A/26B), confirm:

> "You've asked for <OLD> → <NEW>, but the most recent release I have is
> <detected>. Skipping releases is unusual — are you sure, or did you
> mean <detected> → <NEW>?"

Wait for explicit confirmation.

**HITL #1 — prior-release missing:** If `baselines/<OLD>/originals/` is
empty or absent, ask:

> "I need `<OLD>` as the comparison baseline but don't see
> `baselines/<OLD>/`. Download it too, or point me at an existing copy?"

Wait for response. Download path runs Stage 3 twice (once for `<OLD>`,
once for `<NEW>`).

## Stage 3 — Download + verify

For each release that needs downloading:

**Step 3a — download:**

```
python tools/download_and_clear.py <ver> --skip-clear
```

Expected wall time: ~15–20 min. This **wipes `baselines/<ver>/originals/`
first** — do not rerun blindly, you'd lose an already-complete download.

**Step 3b — verify:**

```
python .claude/skills/fbdi-compare-release/scripts/verify_download.py --release <ver>
```

Interpret the exit code:

- `0` → clean, proceed to Stage 4.
- `1` → missing files. If this is the first verification attempt for
  `<ver>`, auto-retry the download **once** (re-run step 3a + 3b).
  Track retry count — the cap is 3 total download attempts per release
  (see HITL #5).
- `2` → extras only. Go to HITL #6.
- `3` → first-run bootstrap required (no `<ver>` section in
  `baseline_files.txt`). Go to First-run bootstrap below.

**HITL #5 — download still short after retry:** If exit code is still `1`
after one auto-retry, group missing filenames by module (the JSON payload
includes `missing_by_module`) and ask:

> "After one retry, `baselines/<ver>/originals/` is still missing N files
> vs `baseline_files.txt`. Affected Oracle modules:
>   - procurement: <file>, <file>, …
>   - financials: <file>, <file>, …
>
> This is beyond what I can fix automatically — likely the Oracle docs
> site changed structure or a module page is down. Options:
>   (a) Retry again (transient issues sometimes need more than one retry).
>   (b) Abort so `tools/download_and_clear.py` can be debugged directly. [default]
>   (c) Proceed with what's present and note the gap in the summary
>       (not recommended — compare output will be incomplete).
>
> Which?"

If user picks (a), re-run 3a + 3b. **Withdraw option (a) after the 3rd
total download attempt** — only (b) and (c) remain.

**HITL #6 — extras present:** Ask:

> "Downloaded N files not in `baseline_files.txt`'s inventory for
> `<ver>`: <list>. Options:
>   (a) Update `baseline_files.txt` to add these to the `<ver>` section. [default]
>   (b) Quarantine to `baselines/<ver>/_extras/` and keep them out of
>       compare — use only if you suspect the scraper pulled something
>       unexpected (rare).
>   (c) Show me the file(s) first so I can decide per-file.
>
> Which?"

On (a), run:
```
python .claude/skills/fbdi-compare-release/scripts/verify_download.py --release <ver> --commit-inventory
```
then re-run verification. It should now exit 0.

On (b), create `baselines/<ver>/_extras/`, move the extras there, then
re-run verification (should now exit 0).

On (c), `cat` the first few rows of each extra and present them for
decision.

**First-run bootstrap** (exit code 3): Read the JSON payload.

- If `over_threshold: true`, show the delta and ask:
  > "Bootstrapping `<ver>` inventory from <N> downloaded files. The most
  > recent prior release (`<prior>`) has <P> files — that's a <Δ%>
  > change. Oracle rarely adds or removes more than ~5–10 templates in a
  > quarterly release, so this is worth a second look. Proceed with
  > bootstrap, retry the download, or abort?"
- If `over_threshold: false`, proceed without prompting.

On "proceed", run `verify_download.py --release <ver> --commit-inventory`
to write the new section, then re-verify (should exit 0).

**HITL #2 — `RapidImplementationForCashManagement.xlsm` missing:** At any
point after download if the file isn't in `baselines/<ver>/originals/`,
ask:

> "`RapidImplementationForCashManagement.xlsm` isn't auto-downloadable.
> Options:
>   (a) Copy from `baselines/<prior>/originals/` — fast, safe since Oracle
>       rarely updates it. [default]
>   (b) I'll walk you through the Oracle Fusion FSM path (Setup and
>       Maintenance → hamburger menu → Search → 'Create Banks, Branches,
>       and Accounts in Spreadsheet').
>   (c) I already have it, let me drop it in `baselines/<new>/originals/`
>       now."

Do not proceed to Stage 4 until the file is present. On (a), `cp` from
the prior baseline. On (b), load `references/troubleshooting.md` for the
full walk-through. On (c), wait for the user, then re-check the path.

## Stage 4 — Smart-clear

```
python tools/download_and_clear.py <NEW> --clear-only
```

(Also `<OLD>` if it was freshly downloaded in Stage 3.)

**Capture the `*** TIMED OUT ...` block** from stdout if present — record
the filenames for Stage 7. Example stdout pattern:

```
  *** TIMED OUT (1 files, >120s each) — clear these manually: ***
      PayablesCollectionDocuments.xlsm (9,234KB)
```

Expected wall time: ~2–4 min. Per-file timeouts are **not a blocker** —
compare reads `originals/`, not `blanks/`.

## Stage 5 — Compare

```
python -m fbdi compare --old <OLD> --new <NEW> --output Comparison_Report_<OLD>_<NEW>.xlsx
```

Expected wall time: ~3–5 min for ~210 file pairs. Per-pair subprocess
isolation already handles corrupt-metadata files silently.

**HITL #4 — excessive compare failures:** If the run's stdout reports >5
per-pair failures (look for the "WARNING: N file(s) timed out" or similar
summary), pause:

> "Compare produced N per-pair failures (threshold: 5). Options:
>   (a) Retry compare (failures are sometimes transient).
>   (b) Skip — note the failures in the final summary and proceed.
>   (c) Abort.
>
> Which?"

Single-digit failures are expected and do not trigger this prompt.

## Stage 6 — Catalog update

```
python -m fbdi catalog --release <NEW>
```

Expected wall time: ~3–5 min.

## Stage 7 — Summary

```
python .claude/skills/fbdi-compare-release/scripts/summarize_report.py \
  --report Comparison_Report_<OLD>_<NEW>.xlsx \
  --catalog FBDI_Master_Catalog.xlsx \
  --timeouts "<Stage 4 timeout filenames, comma-separated>"
```

Render the JSON as a human-readable summary to the terminal:

```
FBDI Compare-Release — <OLD> → <NEW> complete.

Comparison Report: Comparison_Report_<OLD>_<NEW>.xlsx
  <total_changes> change rows across <files_with_changes> files.
  Top 5 most-changed:
    - <file>: <n> changes
    ...

Catalog:           FBDI_Master_Catalog.xlsx

Stage 4 timeouts (manual clear required in baselines/<NEW>/blanks/):
  - PayablesCollectionDocuments.xlsm
```

If the `stage4_timeouts` list is empty, omit that section.

## Stage 8 — Post-run verification

```
python .claude/skills/fbdi-compare-release/scripts/verify_run.py --release <NEW>
```

If `overall_regression: true`, append a warning block to the summary:

```
WARNING: post-run verification flagged potential regressions:
  - Diagnose: <N> NO_HEADER rows in <NEW> (expected 0 historically).
  - Catalog Issues: <N> for <NEW> vs <P> for <prior>; threshold is
    2× or +50 absolute.

These do not invalidate the comparison output but are worth a look.
Load references/troubleshooting.md for context on common causes.
```

Exit code 1 from `verify_run.py` does **not** make the skill fail — the
report and catalog are already produced. Just surface the warnings.

---

## Error handling — general rules

- **Never hand the coworker a raw Python traceback.** If a subprocess
  returns a non-zero exit code with a traceback in stderr, parse out the
  exception type + message, give a plain-English likely cause, and ask
  for permission to retry or abort.
- **If a Stage 3 download subprocess crashes mid-run** (connection drop,
  Selenium timeout), retry once automatically before prompting.
- **If the user Ctrl-C's at any stage**, do not auto-cleanup `baselines/`
  — the partial download may still be useful, and re-running the skill
  resumes from the next stage with outputs that exist.

## Resumability

Each stage is idempotent on output-existence terms:
- Stage 3: re-running **wipes** `originals/` first (destructive). The skill
  warns before retrying.
- Stages 4-6: re-running is safe; they overwrite their outputs.

If interrupted at Stage 5, re-invoking the skill skips 1-4 (env still
healthy, downloads still present, blanks still cleared) and resumes from
compare.
```

- [ ] **Step 2: Smoke-check — make sure SKILL.md parses and the frontmatter is well-formed**

```bash
python -c "
import re
txt = open('.claude/skills/fbdi-compare-release/SKILL.md', encoding='utf-8').read()
assert txt.startswith('---\n')
fm_end = txt.index('\n---\n', 4)
fm = txt[4:fm_end]
assert 'name: fbdi-compare-release' in fm
assert 'description:' in fm
print('Frontmatter OK.')
print('Body lines:', len(txt.splitlines()))
"
```

Expected: "Frontmatter OK.", body lines ~300 (±20% fine).

- [ ] **Step 3: Commit**

```bash
git add .claude/skills/fbdi-compare-release/SKILL.md
git commit -m "feat(skill): flesh out SKILL.md orchestrator workflow"
```

---

## Task 11: Layer-2 eval fixture (spec §8 layer 2)

**Files:**
- Create: `.claude/skills/fbdi-compare-release/evals/prompts.jsonl` (test inputs)
- Create: `.claude/skills/fbdi-compare-release/evals/README.md` (run instructions)

Per spec §8 these are realistic end-to-end prompts driven through `skill-creator`'s eval-viewer. This task sets up the fixtures; running them is a separate manual step (requires Claude-with-skill-access, not pytest).

- [ ] **Step 1: Write eval prompts fixture**

```
# .claude/skills/fbdi-compare-release/evals/prompts.jsonl
{"id": 1, "prompt": "Oracle released 26C, run the quarterly FBDI update", "expected_trigger": true, "notes": "auto-detect path; should offer to download 26C"}
{"id": 2, "prompt": "Compare 26A to 26B", "expected_trigger": true, "notes": "explicit --old/--new; ground-truth reference run"}
{"id": 3, "prompt": "Update the FBDI Master Catalog for 26B", "expected_trigger": true, "notes": "catalog-centric phrasing"}
{"id": 4, "prompt": "What's the current production Python version?", "expected_trigger": false, "notes": "mentions Python, unrelated to FBDI"}
{"id": 5, "prompt": "Run the test suite", "expected_trigger": false, "notes": "unrelated"}
{"id": 6, "prompt": "Open the catalog xlsx", "expected_trigger": false, "notes": "reading, not refreshing"}
{"id": 7, "prompt": "New FBDI release dropped", "expected_trigger": true, "notes": "paraphrase"}
{"id": 8, "prompt": "FBDI refresh for Q1", "expected_trigger": true, "notes": "quarter-coded paraphrase"}
```

- [ ] **Step 2: Write evals README**

```markdown
# Skill Evals — fbdi-compare-release

Per spec §8 Layer 2, these prompts exercise the skill's triggering and
end-to-end behavior.

## Running manually

Use `skill-creator`'s eval-viewer loop:

```bash
python ~/.claude/plugins/cache/claude-plugins-official/skill-creator/unknown/skills/skill-creator/scripts/run_eval.py \
  --skill-dir .claude/skills/fbdi-compare-release \
  --prompts .claude/skills/fbdi-compare-release/evals/prompts.jsonl \
  --output .claude/skills/fbdi-compare-release/evals/results/
```

Then open the generated HTML report:
```bash
python ~/.claude/plugins/cache/claude-plugins-official/skill-creator/unknown/skills/skill-creator/eval-viewer/generate_review.py \
  .claude/skills/fbdi-compare-release/evals/results/
```

## Ground-truth reference (eval #2)

Eval #2 ("Compare 26A to 26B") expects the skill to reproduce the
2026-04-23 end-to-end run:
- `Comparison_Report_26A_26B.xlsx`: 706 change rows, 19 files
- `FBDI_Master_Catalog.xlsx`: 9 Issues-tab rows, 748 Drift rows
- `baselines/26A/originals/`: 212 files
- `baselines/26B/originals/`: 213 files

These artifacts are already on disk and can be used as the oracle for
pass/fail scoring.

## Success criteria

- All `expected_trigger: true` prompts invoke the skill.
- All `expected_trigger: false` prompts do NOT invoke the skill.
- Eval #2 produces a `Comparison_Report_26A_26B.xlsx` byte-equivalent (or
  row-count-equivalent) to the committed ground truth.

If triggering misfires, move to Layer 3 (description optimization — see
Task 12).
```

- [ ] **Step 3: Commit**

```bash
git add .claude/skills/fbdi-compare-release/evals/
git commit -m "docs(skill): add Layer 2 eval fixtures"
```

---

## Task 12: Layer-3 description optimization — post-merge follow-up

**No files created in this task.** Document the follow-up so it's not lost.

- [ ] **Step 1: Append a follow-up note to NEXT_STEPS.md**

Add the following entry under the top-of-backlog section of `NEXT_STEPS.md`:

```markdown
### `fbdi-compare-release` skill — description optimization (post-merge)

After the skill has seen real use for ~2 weeks, run
`skill-creator`'s Layer 3 description-optimization loop on ~20
should-trigger + should-not-trigger queries to tune the `description`
frontmatter:

```bash
python ~/.claude/plugins/cache/claude-plugins-official/skill-creator/unknown/skills/skill-creator/scripts/run_loop.py \
  --skill .claude/skills/fbdi-compare-release \
  --positive-prompts ... \
  --negative-prompts ...
```

Trigger: observed false positives or false negatives in the wild.
```

Run:
```bash
git add NEXT_STEPS.md
git commit -m "docs(next-steps): add Layer 3 description-optimization follow-up"
```

---

## Task 13: Final end-to-end verification

**Files:** none modified.

Before declaring the skill complete, verify the full pipeline produces
outputs equivalent to the committed ground truth (eval #2 on real files).

- [ ] **Step 1: Run the full test suite**

```bash
python -m pytest tests/ -q
```

Expected: all pass (139 original + ~40 new = ~179 total).

- [ ] **Step 2: Invoke the skill end-to-end on 26A → 26B**

In Claude Code, issue:

> "Compare 26A to 26B"

Expected behavior:
- Skill triggers.
- Stage 1 passes (baselines already exist).
- Stage 2 resolves to `OLD=26A, NEW=26B` without prompting (no version
  mismatch; both baselines present).
- Stage 3 reports `verify_download` exit 0 for both (no download needed).
- Stage 4 clears — timeouts may include `PayablesCollectionDocuments.xlsm`;
  skill captures it.
- Stage 5 produces `Comparison_Report_26A_26B.xlsx` with 706 changes.
- Stage 6 produces `FBDI_Master_Catalog.xlsx` with 9 Issues, 748 Drift.
- Stage 7 summary shows 706 changes, 19 files, top-5 list.
- Stage 8 reports no regressions.

Any deviation is a plan bug — investigate before merging.

- [ ] **Step 3: Review the full skill folder for housekeeping**

```bash
find .claude/skills/fbdi-compare-release -type f | sort
```

Expected files:
```
.claude/skills/fbdi-compare-release/SKILL.md
.claude/skills/fbdi-compare-release/evals/README.md
.claude/skills/fbdi-compare-release/evals/prompts.jsonl
.claude/skills/fbdi-compare-release/references/release-version-format.md
.claude/skills/fbdi-compare-release/references/troubleshooting.md
.claude/skills/fbdi-compare-release/scripts/__init__.py
.claude/skills/fbdi-compare-release/scripts/check_env.py
.claude/skills/fbdi-compare-release/scripts/summarize_report.py
.claude/skills/fbdi-compare-release/scripts/verify_download.py
.claude/skills/fbdi-compare-release/scripts/verify_run.py
```

- [ ] **Step 4: Final commit (if any housekeeping changes)**

```bash
git status
# If clean, skip. Otherwise:
git add <whatever>
git commit -m "chore(skill): final housekeeping"
```

- [ ] **Step 5: Push to master (per user preference — see feedback memory `feedback_workflow.md`)**

```bash
git push origin master
```

---

## Self-Review Checklist (run before handing off)

**Spec coverage:**
- §1 Goal — Tasks 1–10 cover it. ✓
- §2 Scope (in) — 8 stages, env bootstrap, 6 HITL, summary. All in SKILL.md Task 10. ✓
- §2 Scope (out) — Applaud mapping, client report, Oracle URL changes, FSM auto-download, catalog schema: explicitly not touched. ✓
- §3 Architecture — folder layout, SKILL.md as orchestrator, bundled scripts narrow + stateless, references on-demand, project-level install. ✓
- §4 Pipeline — 8 stages, each mapped to a script or a command in SKILL.md. ✓
- §5 HITL #1–#6 — all six prompts in SKILL.md Task 10. ✓
- §6 Preflight — Task 2 check_env.py covers all 6 rows of the preflight table. ✓
- §6 Failure handling — each stage's failure path is addressed in SKILL.md + scripts. ✓
- §7 Invocation — slash command + natural-language triggers in SKILL.md description. ✓
- §8 Testing Layer 1 — Task 2–7 tests. ✓
- §8 Testing Layer 2 — Task 11 eval fixtures. ✓
- §8 Testing Layer 3 — Task 12 follow-up. ✓
- §9 Deps — no new deps; `requirements.txt` already exists. ✓
- §10 Resolved — retry cap = 3 (SKILL.md Stage 3), 15% first-run delta (Task 4), Stage 4 timeout handling (Task 6 + SKILL.md Stage 7). ✓

**Placeholder scan:** no "TBD" / "TODO" / "implement later" in any task. ✓

**Type consistency:** `check_env.py` exports `main()` + helper check functions; `verify_download.py` exports `parse_inventory`, `diff_against_inventory`, `list_downloaded`, `group_missing_by_module`, `compute_first_run_delta`, `most_recent_release`, `commit_inventory`, `main`. Tests import these names exactly. ✓

---

## Open decisions already resolved with user (2026-04-23)

1. Stage 4 → Stage 7 timeout hand-off: `--timeouts name1,name2` CLI flag on `summarize_report.py`, populated by SKILL.md from Stage 4 stdout capture.
2. `verify_download.py` exit codes: 0 clean / 1 missing / 2 extras / 3 first-run.
3. Inventory-write location: `verify_download.py --commit-inventory` (deterministic rewrite), not Claude's Edit tool.
4. `check_env.py` never prompts; exit code 2 signals "deps missing" and SKILL.md handles the `pip install` prompt.
5. Sleep-during-long-run: documentation-only — SKILL.md "Before you start" reminds user to disable Windows sleep. No code.
6. Layer 3 description optimization: post-merge follow-up (Task 12), not a blocker.
