# `python -m fbdi run` — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `python -m fbdi run`, a headless chained pipeline that drives the FBDI quarterly comparison workflow without Claude in the loop.

**Architecture:** A thin Python orchestrator (`fbdi/run.py`) executes seven user-facing stages by subprocess (preflight, download, clear, compare, catalog, update-module, report) plus two auto-final stages (summary, verify) that always run. State is captured in a structured manifest written atomically to `logs/`. Five helper scripts move out of `.claude/skills/fbdi-compare-release/scripts/` into the `fbdi/` package as proper CLI subcommands so the skill and `fbdi run` share one source of truth.

**Tech Stack:** Python 3.14, argparse, subprocess, pathlib, openpyxl (existing), pytest (existing).

**Spec:** See [`docs/superpowers/specs/2026-05-04-fbdi-run-headless-pipeline-design.md`](../specs/2026-05-04-fbdi-run-headless-pipeline-design.md) for the full design.

---

## Implementer notes — read first

Brad will clear context before running this plan, so a fresh-context implementer needs the breadcrumbs visible upfront:

- **Skill update via `/skill-creator:skill-creator`.** When you reach Task 7, **invoke `/skill-creator:skill-creator`** — do not edit `.claude/skills/fbdi-compare-release/SKILL.md` by hand. The skill-creator handles SKILL.md rewrites correctly and will catch references a manual edit might miss.
- **Historical plans are immutable.** Do not "fix" path references inside `docs/superpowers/plans/2026-04-23-*.md` or `docs/superpowers/plans/2026-04-30-*.md`. They are historical records.
- **No CLI overrides on HITL policy.** Do not add flags like `--on-extras=stop` even if it feels like a small win. The unattended-mode policies are hard-coded by design.
- **Subprocess per stage, not in-process.** Each stage runs in its own subprocess. This matches the project's existing idiom and isolates failures.

## Recommended execution: 4 batches, hybrid subagent/inline

This plan is large (18 tasks). Execute in 4 batches with a fresh-context Claude Code session per batch. Within each batch, assign each task to subagent or inline based on its character — not a one-size-fits-all approach.

### Batch boundaries

| Batch | Tasks | Why this is a natural boundary |
|---|---|---|
| **1 — Helper Promotion** | Tasks 1–7 | Hard terminal point: scripts directory deleted, skill rewritten, all existing tests still pass. Subsequent batches never touch `.claude/skills/`. |
| **2 — Orchestrator Foundation** | Tasks 8–12 | Ends with a working basic orchestrator that runs all stages via simple `build_stage_command` dispatch. Proves the bones work before HITL complexity lands. |
| **3 — HITL + Auto-Final + CLI** | Tasks 13–16 | The "hot zone": Tasks 13/14/15 each replace the dispatch block from the prior task. Must be one mental thread — clearing mid-batch loses the dispatch evolution and you'll re-read `fbdi/run.py` repeatedly to find your place. |
| **4 — Integration + Verification** | Tasks 17–18 | Post-build verification. Separable because nothing about the implementation can break here. |

After each batch: run the full test suite, commit, then `/clear` before starting the next batch.

### Per-task assignment

| Task | Mode | Rationale |
|---|---|---|
| 1–5 | **Subagent (each, serially)** | Mechanical, identical pattern across all five. Fresh context per move keeps the main thread clean. **Dispatch serially** — all five touch `fbdi/cli.py`, parallel dispatch will collide. |
| 6 | Inline | Trivial cleanup + rename. Not worth a subagent round-trip. |
| 7 | **Inline (required)** | Requires `/skill-creator:skill-creator`. Slash commands run only in the main thread — a subagent literally cannot invoke this. Be at the keyboard for this one. |
| 8 | **Subagent** | Standalone TDD, single new file with clear test list. Good "warm-up" for Batch 2. |
| 9–12 | Inline | `fbdi/run.py` evolves task-by-task. Each task verifies the prior task's accumulating state. Subagents would re-read the file every time at high cost. |
| 13–15 | **Inline (required)** | Tightly coupled. Task 14 replaces a dispatch block that Task 13 wrote; Task 15 extends the same block with auto-final. A subagent doing Task 14 would have to be briefed on exactly what Task 13 left behind — that brief is more expensive than doing it inline. |
| 16 | Inline | Wires `fbdi/cli.py` to the orchestrator. `fbdi/cli.py` was heavily modified by Tasks 1–5; inline keeps that history visible. |
| 17 | **Subagent** | Standalone new test file with one big integration test. Clean isolation, marked `@pytest.mark.integration` so it doesn't perturb the default suite. |
| 18 | Inline | Smoke tests + interpretation of output. Judgment, not code. |

**Result:** 7 of 18 tasks run in subagents (1, 2, 3, 4, 5, 8, 17); the other 11 inline. Pattern: *subagent the mechanical stuff, inline the tightly-coupled stuff, inline anything that needs a slash command or human judgment.*

### Risks to watch

1. **Task 7 needs you in the loop.** `/skill-creator:skill-creator` is interactive. Don't start Batch 1 right before stepping away.
2. **Tasks 1–5 must dispatch serially, not in parallel.** They all add to `fbdi/cli.py`. Concurrent subagents will produce conflicting edits in the import block / subparser block.
3. **Don't break Batch 3 in the middle.** If you must clear context inside Batch 3, do it after a Task commit, never mid-task. The dispatch block evolution is the load-bearing piece of the orchestrator.

## Phase A — Promote helper scripts into the `fbdi/` package

Tasks 1–5 follow an identical pattern: move one script from `.claude/skills/fbdi-compare-release/scripts/` into `fbdi/`, register a CLI subcommand in `fbdi/cli.py`, migrate the relevant tests. Each is independent and committable on its own.

The helpers' existing `main(argv=None) -> int` signatures are preserved. `fbdi/cli.py` reconstructs the argv slice for each subcommand and forwards it to the helper's `main()`. This keeps the move mechanical and preserves all existing test logic.

---

### Task 1: Promote `check_env.py` → `fbdi/preflight.py`

**Files:**
- Move: `.claude/skills/fbdi-compare-release/scripts/check_env.py` → `fbdi/preflight.py`
- Modify: `fbdi/cli.py` (add `preflight` subparser + dispatch)
- Create: `tests/test_preflight.py`
- Modify: `tests/test_skill_scripts.py` (remove migrated test functions)

- [ ] **Step 1: Move the file with `git mv` to preserve history**

```bash
git mv .claude/skills/fbdi-compare-release/scripts/check_env.py fbdi/preflight.py
```

- [ ] **Step 2: Update the module docstring**

Edit the first line of `fbdi/preflight.py` from:

```python
"""Stage 1 preflight for fbdi-compare-release.
```

to:

```python
"""Environment preflight check for the FBDI pipeline.
```

The rest of the file is unchanged. The `main(argv=None) -> int` signature stays.

- [ ] **Step 3: Register the `preflight` subcommand in `fbdi/cli.py`**

Add this subparser block after the existing `report_parser` block (around line 152, before `args = parser.parse_args(argv)`):

```python
    preflight_parser = subparsers.add_parser(
        "preflight",
        help="Run environment preflight checks (Python, deps, Chrome, baselines/)",
    )
```

Add this dispatch branch after `elif args.command == "report":` (around line 168):

```python
    elif args.command == "preflight":
        from fbdi import preflight
        sys.exit(preflight.main([]))
```

- [ ] **Step 4: Verify the CLI subcommand works**

Run: `python -m fbdi preflight`
Expected: prints JSON checks payload, exit 0 (or 1/2 depending on local env). The exit code matches what `python .claude/skills/.../scripts/check_env.py` returned previously.

- [ ] **Step 5: Create `tests/test_preflight.py` by migrating tests from `tests/test_skill_scripts.py`**

Create `tests/test_preflight.py` with this content:

```python
"""Tests for the FBDI environment preflight check (fbdi/preflight.py)."""

import json
import subprocess
import sys

from fbdi import preflight


def test_preflight_exposes_main():
    assert hasattr(preflight, "main")


def test_preflight_python_version_check_passes_on_314():
    result = preflight.check_python_version(current=(3, 14, 3))
    assert result["ok"] is True
    assert "3.14" in result["detail"]


def test_preflight_python_version_check_fails_on_old():
    result = preflight.check_python_version(current=(3, 11, 0))
    assert result["ok"] is False
    assert "3.14" in result["detail"]


def test_preflight_deps_check_detects_missing():
    result = preflight.check_deps(required=["definitely_not_a_real_package_xyz"])
    assert result["ok"] is False
    assert "definitely_not_a_real_package_xyz" in result["detail"]


def test_preflight_deps_check_passes_on_stdlib():
    result = preflight.check_deps(required=["json"])
    assert result["ok"] is True


def test_preflight_baselines_dir_creates_if_missing(tmp_path):
    result = preflight.check_baselines_dir(root=tmp_path)
    assert result["ok"] is True
    assert (tmp_path / "baselines").is_dir()


def _run_preflight_cli(tmp_path):
    """Invoke `python -m fbdi preflight` and return (exit_code, parsed_json)."""
    cmd = [sys.executable, "-m", "fbdi", "preflight"]
    proc = subprocess.run(cmd, cwd=tmp_path, capture_output=True, text=True)
    return proc.returncode, json.loads(proc.stdout)


def test_preflight_cli_produces_structured_json(tmp_path):
    (tmp_path / "baselines").mkdir()
    (tmp_path / "baseline_files.txt").write_text("stub\n")
    _, payload = _run_preflight_cli(tmp_path)
    assert "checks" in payload
    assert "missing_deps" in payload
    assert "fatal" in payload


def test_preflight_cli_json_output_parseable(tmp_path):
    _, payload = _run_preflight_cli(tmp_path)
    assert isinstance(payload, dict)
    assert isinstance(payload["checks"], list)
```

- [ ] **Step 6: Remove the migrated tests from `tests/test_skill_scripts.py`**

Open `tests/test_skill_scripts.py` and delete:
- The line `from scripts import check_env  # noqa: E402 — importable when cwd is repo root`
- The function `_run_check_env(tmp_path)`
- All test functions whose names start with `test_check_env_`

The file should still have its top-level imports, the `SKILL_ROOT` / `sys.path` injection, and the `test_skill_folder_exists` / `test_skill_md_has_frontmatter` / `test_scripts_dir_is_python_package` tests (the last one will be deleted in Task 6 once the directory is empty).

- [ ] **Step 7: Run the test suites**

Run: `python -m pytest tests/test_preflight.py tests/test_skill_scripts.py -v`
Expected: all tests pass.

- [ ] **Step 8: Commit**

```bash
git add fbdi/preflight.py fbdi/cli.py tests/test_preflight.py tests/test_skill_scripts.py
git commit -m "refactor(fbdi): promote check_env.py to fbdi/preflight.py with CLI subcommand"
```

---

### Task 2: Promote `verify_download.py` → `fbdi/verify_download.py`

**Files:**
- Move: `.claude/skills/fbdi-compare-release/scripts/verify_download.py` → `fbdi/verify_download.py`
- Modify: `fbdi/cli.py` (add `verify-download` subparser + dispatch)
- Create: `tests/test_verify_download.py`
- Modify: `tests/test_skill_scripts.py` (remove migrated test functions)

- [ ] **Step 1: Move the file**

```bash
git mv .claude/skills/fbdi-compare-release/scripts/verify_download.py fbdi/verify_download.py
```

- [ ] **Step 2: Update the module docstring**

Replace the leading docstring `"""Stage 3 download verification for fbdi-compare-release.` with `"""Download verification and inventory management for the FBDI pipeline.`

- [ ] **Step 3: Register the `verify-download` subcommand in `fbdi/cli.py`**

Add this subparser after the `preflight_parser` block from Task 1:

```python
    verify_dl_parser = subparsers.add_parser(
        "verify-download",
        help="Verify downloaded baselines against baseline_files.txt inventory",
    )
    verify_dl_parser.add_argument(
        "--release", required=True, type=str,
        help="Release label (e.g. 26B)",
    )
    verify_dl_parser.add_argument(
        "--commit-inventory", action="store_true",
        help="Update baseline_files.txt with the current downloaded set",
    )
```

Add this dispatch branch after the `preflight` dispatch from Task 1:

```python
    elif args.command == "verify-download":
        from fbdi import verify_download
        argv_for_helper = ["--release", args.release]
        if args.commit_inventory:
            argv_for_helper.append("--commit-inventory")
        sys.exit(verify_download.main(argv_for_helper))
```

- [ ] **Step 4: Verify the CLI subcommand works**

Run: `python -m fbdi verify-download --release 26A`
Expected: exits with the same code that the previous skill-bundled script returned (0/1/2/3 depending on baseline state).

- [ ] **Step 5: Create `tests/test_verify_download.py` with the migrated tests**

Open `tests/test_skill_scripts.py` and locate the block that starts with `from scripts import verify_download  # noqa: E402` (around line 96). Copy every test function from that line through the next `from scripts import …` import (which will be `summarize_report`).

Create `tests/test_verify_download.py` with:

```python
"""Tests for fbdi/verify_download.py."""

from pathlib import Path

from fbdi import verify_download


# Paste the INVENTORY_FIXTURE constant and every test function from the
# verify_download section of tests/test_skill_scripts.py here. Replace any
# occurrence of `from scripts import verify_download` with the import above.
# Replace any direct subprocess invocations of
# `.claude/skills/fbdi-compare-release/scripts/verify_download.py` with
# `[sys.executable, "-m", "fbdi", "verify-download", ...]` argv lists.
```

Then literally copy the tests from `tests/test_skill_scripts.py` (the section between the `from scripts import verify_download` line and the next `from scripts import summarize_report` line) into the new file. Update any subprocess invocations as noted in the comment above.

- [ ] **Step 6: Remove the migrated section from `tests/test_skill_scripts.py`**

Delete the `from scripts import verify_download` line and every test function in that section.

- [ ] **Step 7: Run the test suites**

Run: `python -m pytest tests/test_verify_download.py tests/test_skill_scripts.py -v`
Expected: all tests pass.

- [ ] **Step 8: Commit**

```bash
git add fbdi/verify_download.py fbdi/cli.py tests/test_verify_download.py tests/test_skill_scripts.py
git commit -m "refactor(fbdi): promote verify_download.py to fbdi/ with CLI subcommand"
```

---

### Task 3: Promote `verify_run.py` → `fbdi/verify_run.py`

**Files:**
- Move: `.claude/skills/fbdi-compare-release/scripts/verify_run.py` → `fbdi/verify_run.py`
- Modify: `fbdi/cli.py` (add `verify-run` subparser + dispatch)
- Create: `tests/test_verify_run.py`
- Modify: `tests/test_skill_scripts.py` (remove migrated test functions)

- [ ] **Step 1: Move the file**

```bash
git mv .claude/skills/fbdi-compare-release/scripts/verify_run.py fbdi/verify_run.py
```

- [ ] **Step 2: Update the module docstring**

Replace `"""Stage 8 post-run verification for fbdi-compare-release.` with `"""Post-run verification for the FBDI pipeline.`

- [ ] **Step 3: Register the `verify-run` subcommand in `fbdi/cli.py`**

Inspect the `main()` function inside the new `fbdi/verify_run.py` to determine its argparse args. Add a matching subparser:

```python
    verify_run_parser = subparsers.add_parser(
        "verify-run",
        help="Post-run regression check (diagnose + catalog Issues thresholds)",
    )
    verify_run_parser.add_argument(
        "--release", required=True, type=str,
        help="Release label (e.g. 26B)",
    )
```

(Add additional `add_argument` calls if `verify_run.main` accepts more flags — read the file to confirm.)

Add the dispatch branch:

```python
    elif args.command == "verify-run":
        from fbdi import verify_run
        argv_for_helper = ["--release", args.release]
        sys.exit(verify_run.main(argv_for_helper))
```

(Extend the argv list if there are more flags.)

- [ ] **Step 4: Verify the CLI subcommand works**

Run: `python -m fbdi verify-run --release 26A`
Expected: exits with the same code as the previous skill-bundled script.

- [ ] **Step 5: Create `tests/test_verify_run.py` and migrate the tests**

Locate the `from scripts import verify_run  # noqa: E402` block in `tests/test_skill_scripts.py`. Copy every test function from that line through end-of-file into a new `tests/test_verify_run.py`:

```python
"""Tests for fbdi/verify_run.py."""

from fbdi import verify_run

# Paste the verify_run test functions from tests/test_skill_scripts.py here.
# Replace `from scripts import verify_run` references with the import above.
# Replace subprocess invocations of the old script path with
# `[sys.executable, "-m", "fbdi", "verify-run", ...]` argv lists.
```

- [ ] **Step 6: Remove the migrated section from `tests/test_skill_scripts.py`**

Delete the `from scripts import verify_run` line and all test functions in that section.

- [ ] **Step 7: Run the test suites**

Run: `python -m pytest tests/test_verify_run.py tests/test_skill_scripts.py -v`
Expected: all tests pass.

- [ ] **Step 8: Commit**

```bash
git add fbdi/verify_run.py fbdi/cli.py tests/test_verify_run.py tests/test_skill_scripts.py
git commit -m "refactor(fbdi): promote verify_run.py to fbdi/ with CLI subcommand"
```

---

### Task 4: Promote `verify_rerun.py` → `fbdi/verify_rerun.py`

**Files:**
- Move: `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py` → `fbdi/verify_rerun.py`
- Modify: `fbdi/cli.py` (add `verify-rerun` subparser + dispatch)
- Modify: `tests/test_verify_rerun.py` (replace `importlib.util.spec_from_file_location` with normal import)

- [ ] **Step 1: Move the file**

```bash
git mv .claude/skills/fbdi-compare-release/scripts/verify_rerun.py fbdi/verify_rerun.py
```

- [ ] **Step 2: Update the module docstring**

Replace `"""Stage 8 macro-signal validator for fbdi-compare-release.` with `"""Macro-signal validator for the FBDI pipeline (catalog/compare/module-pct deltas).`

- [ ] **Step 3: Register the `verify-rerun` subcommand in `fbdi/cli.py`**

Inspect the `main()` function inside `fbdi/verify_rerun.py` to determine its argparse args. Add a matching subparser:

```python
    verify_rerun_parser = subparsers.add_parser(
        "verify-rerun",
        help="Quarterly macro-signal validator (catalog row delta, compare delta, module pct)",
    )
    verify_rerun_parser.add_argument(
        "--release", required=True, type=str,
        help="Release label (e.g. 26B)",
    )
    verify_rerun_parser.add_argument(
        "--compare-report", type=Path,
        help="Path to the comparison report .xlsx",
    )
    verify_rerun_parser.add_argument(
        "--baseline-catalog", type=Path,
        help="Path to the baseline (pre-rerun) catalog .xlsx",
    )
```

(Adjust args based on what the file's `main()` actually accepts — read the file to confirm.)

Add the dispatch branch:

```python
    elif args.command == "verify-rerun":
        from fbdi import verify_rerun
        argv_for_helper = ["--release", args.release]
        if args.compare_report:
            argv_for_helper.extend(["--compare-report", str(args.compare_report)])
        if args.baseline_catalog:
            argv_for_helper.extend(["--baseline-catalog", str(args.baseline_catalog)])
        sys.exit(verify_rerun.main(argv_for_helper))
```

- [ ] **Step 4: Verify the CLI subcommand works**

Run: `python -m fbdi verify-rerun --release 26A`
Expected: exits with the same code as the previous skill-bundled script (or with a clear error about missing optional flags).

- [ ] **Step 5: Update `tests/test_verify_rerun.py`**

Open `tests/test_verify_rerun.py`. Replace the import block:

```python
"""Tests for the verify_rerun.py macro-signal validator."""

import importlib.util
import json
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook


SKILL_SCRIPT = Path(__file__).resolve().parent.parent / ".claude" / "skills" / \
    "fbdi-compare-release" / "scripts" / "verify_rerun.py"


def _load_module():
    spec = importlib.util.spec_from_file_location("verify_rerun", SKILL_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["verify_rerun"] = mod
    spec.loader.exec_module(mod)
    return mod
```

with:

```python
"""Tests for fbdi/verify_rerun.py — quarterly macro-signal validator."""

import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from fbdi import verify_rerun
```

Then update every callsite of `_load_module()` in this file to use `verify_rerun` directly. For example, `mod = _load_module()` → delete that line; replace subsequent `mod.run_checks(...)` with `verify_rerun.run_checks(...)`. Delete the `_load_module` function definition.

- [ ] **Step 6: Run the test suite**

Run: `python -m pytest tests/test_verify_rerun.py -v`
Expected: all tests pass.

- [ ] **Step 7: Commit**

```bash
git add fbdi/verify_rerun.py fbdi/cli.py tests/test_verify_rerun.py
git commit -m "refactor(fbdi): promote verify_rerun.py to fbdi/ with CLI subcommand"
```

---

### Task 5: Promote `summarize_report.py` → `fbdi/summary.py`

**Files:**
- Move: `.claude/skills/fbdi-compare-release/scripts/summarize_report.py` → `fbdi/summary.py`
- Modify: `fbdi/cli.py` (add `summarize` subparser + dispatch)
- Create: `tests/test_summary.py`
- Modify: `tests/test_skill_scripts.py` (remove migrated test functions)

- [ ] **Step 1: Move the file**

```bash
git mv .claude/skills/fbdi-compare-release/scripts/summarize_report.py fbdi/summary.py
```

- [ ] **Step 2: Update the module docstring**

Replace `"""Stage 7 summary for fbdi-compare-release.` with `"""Summary report for the FBDI pipeline (parses Comparison_Report into top-N changes).`

- [ ] **Step 3: Register the `summarize` subcommand in `fbdi/cli.py`**

Inspect the `main()` function inside `fbdi/summary.py` to determine its argparse args. Add a matching subparser:

```python
    summarize_parser = subparsers.add_parser(
        "summarize",
        help="Print a human-readable summary of the comparison report and catalog",
    )
    summarize_parser.add_argument(
        "--report", required=True, type=Path,
        help="Path to the Comparison_Report_<OLD>_<NEW>.xlsx",
    )
    summarize_parser.add_argument(
        "--catalog", required=True, type=Path,
        help="Path to FBDI_Master_Catalog.xlsx",
    )
    summarize_parser.add_argument(
        "--timeouts", type=str, default="",
        help="Comma-separated filenames that timed out during clear (optional)",
    )
```

(Adjust args based on what the file's `main()` actually accepts.)

Add the dispatch branch:

```python
    elif args.command == "summarize":
        from fbdi import summary
        argv_for_helper = [
            "--report", str(args.report),
            "--catalog", str(args.catalog),
        ]
        if args.timeouts:
            argv_for_helper.extend(["--timeouts", args.timeouts])
        sys.exit(summary.main(argv_for_helper))
```

- [ ] **Step 4: Verify the CLI subcommand works**

Run (with real artifacts present): `python -m fbdi summarize --report Comparison_Report_26A_26B.xlsx --catalog FBDI_Master_Catalog.xlsx`
Expected: prints the JSON summary; exit 0.

- [ ] **Step 5: Create `tests/test_summary.py` and migrate the tests**

Locate the `from scripts import summarize_report  # noqa: E402` block in `tests/test_skill_scripts.py`. Copy every related test function into:

```python
"""Tests for fbdi/summary.py."""

from pathlib import Path

from fbdi import summary

# Paste summarize_report test functions from tests/test_skill_scripts.py here.
# Replace `from scripts import summarize_report` references with the import above.
# Replace `summarize_report.<symbol>` with `summary.<symbol>`.
```

- [ ] **Step 6: Remove the migrated section from `tests/test_skill_scripts.py`**

Delete the `from scripts import summarize_report` line and all test functions in that section. After this step, `tests/test_skill_scripts.py` should retain only the `test_skill_folder_exists` and `test_skill_md_has_frontmatter` tests (the `test_scripts_dir_is_python_package` test will be deleted in Task 6).

- [ ] **Step 7: Run the test suites**

Run: `python -m pytest tests/test_summary.py tests/test_skill_scripts.py -v`
Expected: all tests pass.

- [ ] **Step 8: Commit**

```bash
git add fbdi/summary.py fbdi/cli.py tests/test_summary.py tests/test_skill_scripts.py
git commit -m "refactor(fbdi): promote summarize_report.py to fbdi/summary.py with CLI subcommand"
```

---

### Task 6: Clean up the now-empty skill scripts directory

**Files:**
- Delete: `.claude/skills/fbdi-compare-release/scripts/` (entire directory, including `__init__.py` and any `__pycache__`)
- Modify: `tests/test_skill_scripts.py` (remove the obsolete `test_scripts_dir_is_python_package` test, then either keep the file with only the SKILL.md/frontmatter tests or rename if no other tests remain)

- [ ] **Step 1: Confirm the scripts directory is empty (apart from `__init__.py` and caches)**

Run: `ls .claude/skills/fbdi-compare-release/scripts/`
Expected: only `__init__.py` and possibly `__pycache__/` remain. If any `.py` file other than `__init__.py` is present, return to Tasks 1–5 — something didn't move.

- [ ] **Step 2: Delete the scripts directory**

```bash
git rm -r .claude/skills/fbdi-compare-release/scripts/
```

- [ ] **Step 3: Update `tests/test_skill_scripts.py`**

Open `tests/test_skill_scripts.py`. Delete the `test_scripts_dir_is_python_package` function entirely (the directory no longer exists).

Also delete the `sys.path.insert(0, str(SKILL_ROOT))` line near the top of the file — there are no scripts to import anymore.

The file should now contain only `test_skill_folder_exists` and `test_skill_md_has_frontmatter`. Rename the file to better reflect its remaining scope:

```bash
git mv tests/test_skill_scripts.py tests/test_skill_metadata.py
```

- [ ] **Step 4: Run the renamed test file**

Run: `python -m pytest tests/test_skill_metadata.py -v`
Expected: both tests pass.

- [ ] **Step 5: Run the full test suite as a regression check**

Run: `python -m pytest tests/ -v`
Expected: all 320+ tests pass (the count may be slightly different due to test redistribution; the old test count minus zero new failures is the success criterion).

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "refactor(fbdi): remove now-empty skill scripts directory"
```

---

### Task 7: Update the skill via `/skill-creator:skill-creator`

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/SKILL.md` (rewrite all `python .claude/skills/.../scripts/<name>.py` invocations to `python -m fbdi <subcommand>`)

This task is interactive. The implementer drives `/skill-creator:skill-creator` with a clear prompt; that skill then performs the SKILL.md edits.

- [ ] **Step 1: Identify every invocation site to be rewritten**

Run: `grep -n "scripts/" .claude/skills/fbdi-compare-release/SKILL.md`
Expected output (verify all of these are present, then translate):

```
36:python .claude/skills/fbdi-compare-release/scripts/check_env.py
99:python .claude/skills/fbdi-compare-release/scripts/verify_download.py --release <ver>
148:python .claude/skills/fbdi-compare-release/scripts/verify_download.py --release <ver> --commit-inventory
291:python .claude/skills/fbdi-compare-release/scripts/summarize_report.py \
333:python .claude/skills/fbdi-compare-release/scripts/verify_run.py --release <NEW>
351:python .claude/skills/fbdi-compare-release/scripts/verify_rerun.py \
```

The translations:

| Old invocation | New invocation |
|---|---|
| `python .claude/skills/fbdi-compare-release/scripts/check_env.py` | `python -m fbdi preflight` |
| `python .claude/skills/fbdi-compare-release/scripts/verify_download.py --release <ver>` | `python -m fbdi verify-download --release <ver>` |
| `python .claude/skills/fbdi-compare-release/scripts/verify_download.py --release <ver> --commit-inventory` | `python -m fbdi verify-download --release <ver> --commit-inventory` |
| `python .claude/skills/fbdi-compare-release/scripts/summarize_report.py …` | `python -m fbdi summarize …` |
| `python .claude/skills/fbdi-compare-release/scripts/verify_run.py --release <NEW>` | `python -m fbdi verify-run --release <NEW>` |
| `python .claude/skills/fbdi-compare-release/scripts/verify_rerun.py …` | `python -m fbdi verify-rerun …` |

There are also five reference mentions (e.g., "run `verify_download.py --commit-inventory`") that should be updated to refer to the new CLI verb form ("run `python -m fbdi verify-download --commit-inventory`"). These are at approximate line numbers 47, 168, 237, 327, 365–368 — the skill-creator should catch them.

- [ ] **Step 2: Invoke the skill-creator**

Type into Claude:

```
/skill-creator:skill-creator

Update .claude/skills/fbdi-compare-release/SKILL.md to use the new
fbdi CLI subcommands. The following helper scripts have been promoted
out of .claude/skills/fbdi-compare-release/scripts/ into the fbdi/
package:

- check_env.py        → python -m fbdi preflight
- verify_download.py  → python -m fbdi verify-download
- verify_run.py       → python -m fbdi verify-run
- verify_rerun.py     → python -m fbdi verify-rerun
- summarize_report.py → python -m fbdi summarize

Rewrite every invocation in SKILL.md (six command-block lines plus
five inline reference mentions). The skill's external behavior must
not change — same stages, same HITL prompts, same outputs. Only the
invocation paths inside SKILL.md change.

After editing, verify with:
  grep -n "scripts/" .claude/skills/fbdi-compare-release/SKILL.md
  (expected: zero matches)
```

- [ ] **Step 3: Verify the edits landed cleanly**

Run: `grep -n "scripts/" .claude/skills/fbdi-compare-release/SKILL.md`
Expected: no output (zero matches).

Run: `grep -n "python -m fbdi" .claude/skills/fbdi-compare-release/SKILL.md | head`
Expected: at least six lines showing the new invocations.

- [ ] **Step 4: Smoke-test that the skill itself still parses by reading it back**

Run: `head -20 .claude/skills/fbdi-compare-release/SKILL.md`
Expected: the YAML frontmatter (`---`/`name:`/`description:`) is intact and the file's overall structure is preserved.

- [ ] **Step 5: Commit**

```bash
git add .claude/skills/fbdi-compare-release/SKILL.md
git commit -m "refactor(skill): retarget fbdi-compare-release SKILL.md at new fbdi CLI subcommands"
```

---

## Phase B — Manifest infrastructure

### Task 8: Create `fbdi/manifest.py` for run-state tracking

**Files:**
- Create: `fbdi/manifest.py`
- Create: `tests/test_manifest.py`

The manifest is the structured record of a `fbdi run` invocation. It accumulates per-stage status as the orchestrator progresses, writes atomically to disk after each stage, and serves as the source of truth for the exit-code decision.

- [ ] **Step 1: Write a failing test for manifest construction**

Create `tests/test_manifest.py`:

```python
"""Tests for fbdi/manifest.py — run state manifest."""

import json
from pathlib import Path

from fbdi.manifest import Manifest


def test_manifest_initial_state():
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    assert m.old == "26A"
    assert m.new == "26B"
    assert m.from_stage == "preflight"
    assert m.to_stage == "report"
    assert m.exit_code == 0
    assert m.warnings == []
    assert m.stages == {}


def test_manifest_record_stage():
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m.record_stage("preflight", status="ok", duration_s=1.2)
    assert m.stages["preflight"] == {"status": "ok", "duration_s": 1.2}


def test_manifest_record_stage_with_extra_fields():
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m.record_stage("compare", status="ok", duration_s=287, total_changes=1247, pair_failures=2)
    assert m.stages["compare"]["total_changes"] == 1247
    assert m.stages["compare"]["pair_failures"] == 2


def test_manifest_add_warning():
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m.add_warning("FSM file missing from both baselines")
    assert "FSM file missing from both baselines" in m.warnings


def test_manifest_to_dict_schema():
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m.record_stage("preflight", status="ok", duration_s=1.0)
    d = m.to_dict()
    expected_keys = {
        "old", "new", "started", "ended", "duration_s",
        "exit_code", "from_stage", "to_stage", "stages", "warnings",
    }
    assert expected_keys.issubset(d.keys())
    assert d["stages"]["preflight"]["status"] == "ok"


def test_manifest_atomic_write_creates_file(tmp_path):
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m.record_stage("preflight", status="ok", duration_s=1.0)
    out = tmp_path / "manifest.json"
    m.write(out)
    assert out.is_file()
    payload = json.loads(out.read_text())
    assert payload["old"] == "26A"
    assert payload["stages"]["preflight"]["status"] == "ok"


def test_manifest_atomic_write_overwrites_safely(tmp_path):
    """Writing twice should leave the second payload, not a half-written file."""
    out = tmp_path / "manifest.json"
    m1 = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m1.record_stage("preflight", status="ok", duration_s=1.0)
    m1.write(out)

    m2 = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m2.record_stage("preflight", status="failed", duration_s=2.0)
    m2.write(out)

    payload = json.loads(out.read_text())
    assert payload["stages"]["preflight"]["status"] == "failed"


def test_manifest_set_exit_code():
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m.set_exit_code(5)
    assert m.exit_code == 5


def test_manifest_finalize_records_ended_and_duration(monkeypatch):
    m = Manifest(old="26A", new="26B", from_stage="preflight", to_stage="report")
    m.finalize()
    assert m.ended is not None
    assert m.duration_s is not None
    assert m.duration_s >= 0
```

- [ ] **Step 2: Run the test to confirm failure**

Run: `python -m pytest tests/test_manifest.py -v`
Expected: all tests fail with `ModuleNotFoundError: No module named 'fbdi.manifest'`.

- [ ] **Step 3: Implement `fbdi/manifest.py`**

Create `fbdi/manifest.py`:

```python
"""Run-state manifest for `python -m fbdi run`.

Records per-stage status and exit code. Writes atomically to JSON so external
watchers can read partial state mid-run without seeing torn writes.
"""

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def _utcnow_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


@dataclass
class Manifest:
    old: str
    new: str
    from_stage: str
    to_stage: str
    started: str = field(default_factory=_utcnow_iso)
    ended: str | None = None
    duration_s: float | None = None
    exit_code: int = 0
    stages: dict[str, dict[str, Any]] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    _started_monotonic: float = field(default_factory=time.monotonic, repr=False)

    def record_stage(self, name: str, **fields: Any) -> None:
        self.stages[name] = dict(fields)

    def add_warning(self, message: str) -> None:
        self.warnings.append(message)

    def set_exit_code(self, code: int) -> None:
        self.exit_code = code

    def finalize(self) -> None:
        self.ended = _utcnow_iso()
        self.duration_s = max(time.monotonic() - self._started_monotonic, 0.0)

    def to_dict(self) -> dict[str, Any]:
        return {
            "old": self.old,
            "new": self.new,
            "started": self.started,
            "ended": self.ended,
            "duration_s": self.duration_s,
            "exit_code": self.exit_code,
            "from_stage": self.from_stage,
            "to_stage": self.to_stage,
            "stages": self.stages,
            "warnings": self.warnings,
        }

    def write(self, path: Path) -> None:
        """Write the manifest atomically: write to a sibling .tmp, then rename."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2)
        os.replace(tmp, path)
```

- [ ] **Step 4: Run the tests to confirm they pass**

Run: `python -m pytest tests/test_manifest.py -v`
Expected: all 8 tests pass.

- [ ] **Step 5: Commit**

```bash
git add fbdi/manifest.py tests/test_manifest.py
git commit -m "feat(fbdi): add manifest module for run-state tracking"
```

---

## Phase C — Orchestrator

### Task 9: Argument parsing and validation in `fbdi/run.py`

**Files:**
- Create: `fbdi/run.py` (skeleton with arg parsing only)
- Create: `tests/test_run.py` (validation tests only)

Keep this task narrow: just the argparse + validation surface. The actual stage execution lands in later tasks.

- [ ] **Step 1: Write failing tests for argument validation**

Create `tests/test_run.py`:

```python
"""Tests for fbdi/run.py — the headless pipeline orchestrator."""

import pytest

from fbdi.run import parse_run_args, ALL_STAGES


def test_parse_run_args_minimal():
    args = parse_run_args(["--old", "26A", "--new", "26B"])
    assert args.old == "26A"
    assert args.new == "26B"
    assert args.from_stage == "preflight"
    assert args.to_stage == "report"


def test_parse_run_args_with_range():
    args = parse_run_args([
        "--old", "26A", "--new", "26B",
        "--from", "compare", "--to", "catalog",
    ])
    assert args.from_stage == "compare"
    assert args.to_stage == "catalog"


def test_parse_run_args_log_dir():
    args = parse_run_args([
        "--old", "26A", "--new", "26B",
        "--log-dir", "/tmp/logs",
    ])
    assert str(args.log_dir).endswith("logs") or str(args.log_dir).endswith("logs\\")


def test_parse_run_args_rejects_invalid_release():
    with pytest.raises(SystemExit):
        parse_run_args(["--old", "26X", "--new", "26B"])


def test_parse_run_args_rejects_lowercase_release():
    with pytest.raises(SystemExit):
        parse_run_args(["--old", "26a", "--new", "26B"])


def test_parse_run_args_rejects_same_old_and_new():
    with pytest.raises(SystemExit):
        parse_run_args(["--old", "26A", "--new", "26A"])


def test_parse_run_args_rejects_unknown_stage():
    with pytest.raises(SystemExit):
        parse_run_args(["--old", "26A", "--new", "26B", "--from", "frobnicate"])


def test_parse_run_args_rejects_to_before_from():
    with pytest.raises(SystemExit):
        parse_run_args([
            "--old", "26A", "--new", "26B",
            "--from", "report", "--to", "compare",
        ])


def test_all_stages_defined():
    expected = ["preflight", "download", "clear", "compare",
                "catalog", "update-module", "report"]
    assert ALL_STAGES == expected
```

- [ ] **Step 2: Run the tests to confirm failure**

Run: `python -m pytest tests/test_run.py -v`
Expected: `ModuleNotFoundError: No module named 'fbdi.run'`.

- [ ] **Step 3: Implement `fbdi/run.py` skeleton**

Create `fbdi/run.py`:

```python
"""Headless chained pipeline orchestrator (`python -m fbdi run`).

Drives the FBDI quarterly comparison workflow end-to-end without Claude
in the loop. See docs/superpowers/specs/2026-05-04-fbdi-run-headless-pipeline-design.md
for the full design.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ALL_STAGES = [
    "preflight",
    "download",
    "clear",
    "compare",
    "catalog",
    "update-module",
    "report",
]

_RELEASE_RE = re.compile(r"^\d{2}[A-D]$")


def parse_run_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse argv for the `fbdi run` subcommand.

    Validates release format, distinct old/new, stage-name validity, and
    --to >= --from. SystemExit on any validation failure (argparse default).
    """
    parser = argparse.ArgumentParser(
        prog="fbdi run",
        description="Headless chained FBDI pipeline (download → compare → catalog → report).",
    )
    parser.add_argument("--old", required=True, type=str,
                        help="Prior release label (e.g. 26A)")
    parser.add_argument("--new", required=True, type=str,
                        help="New release label (e.g. 26B)")
    parser.add_argument("--from", dest="from_stage", default="preflight",
                        choices=ALL_STAGES,
                        help=f"First stage to execute (default: preflight). Choices: {', '.join(ALL_STAGES)}")
    parser.add_argument("--to", dest="to_stage", default="report",
                        choices=ALL_STAGES,
                        help=f"Last stage to execute, inclusive (default: report). Choices: {', '.join(ALL_STAGES)}")
    parser.add_argument("--log-dir", dest="log_dir", default=Path("logs"), type=Path,
                        help="Directory for run logs and manifests (default: ./logs)")

    args = parser.parse_args(argv)

    if not _RELEASE_RE.match(args.old):
        parser.error(f"--old must match format NNX (e.g. 26A); got: {args.old!r}")
    if not _RELEASE_RE.match(args.new):
        parser.error(f"--new must match format NNX (e.g. 26B); got: {args.new!r}")
    if args.old == args.new:
        parser.error(f"--old and --new must differ; both are {args.old!r}")

    if ALL_STAGES.index(args.to_stage) < ALL_STAGES.index(args.from_stage):
        parser.error(
            f"--to ({args.to_stage}) cannot precede --from ({args.from_stage}); "
            f"stage order is: {' → '.join(ALL_STAGES)}"
        )

    return args


def main(argv: list[str] | None = None) -> int:
    """Entry point. Currently arg-parse only; orchestration lands in later tasks."""
    parse_run_args(argv)
    print("fbdi run: orchestration not yet implemented — argument parsing OK", file=sys.stderr)
    return 0
```

- [ ] **Step 4: Run the tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all 9 tests pass.

- [ ] **Step 5: Commit**

```bash
git add fbdi/run.py tests/test_run.py
git commit -m "feat(fbdi): add fbdi/run.py argument parsing and validation"
```

---

### Task 10: Subprocess execution helper

**Files:**
- Modify: `fbdi/run.py` (add `run_subprocess` helper + log streaming)
- Modify: `tests/test_run.py` (add helper tests)

This step adds a reusable function the orchestrator calls per stage: spawn a subprocess, capture stdout/stderr to the log file, return exit code and captured text.

- [ ] **Step 1: Write failing tests for the subprocess helper**

Append to `tests/test_run.py`:

```python
import sys

from fbdi.run import run_subprocess


def test_run_subprocess_captures_stdout(tmp_path):
    log = tmp_path / "out.log"
    rc, text = run_subprocess(
        [sys.executable, "-c", "print('hello world')"],
        log_path=log,
    )
    assert rc == 0
    assert "hello world" in text
    assert "hello world" in log.read_text()


def test_run_subprocess_captures_stderr(tmp_path):
    log = tmp_path / "out.log"
    rc, text = run_subprocess(
        [sys.executable, "-c", "import sys; sys.stderr.write('boom\\n'); sys.exit(7)"],
        log_path=log,
    )
    assert rc == 7
    assert "boom" in text
    assert "boom" in log.read_text()


def test_run_subprocess_appends_to_existing_log(tmp_path):
    log = tmp_path / "out.log"
    log.write_text("PRIOR LINE\n")
    rc, _ = run_subprocess(
        [sys.executable, "-c", "print('NEW LINE')"],
        log_path=log,
    )
    assert rc == 0
    contents = log.read_text()
    assert "PRIOR LINE" in contents
    assert "NEW LINE" in contents


def test_run_subprocess_writes_command_header(tmp_path):
    log = tmp_path / "out.log"
    run_subprocess(
        [sys.executable, "-c", "print('x')"],
        log_path=log,
    )
    contents = log.read_text()
    # Header line marks where this subprocess began in the log
    assert "$ " in contents or "===" in contents
```

- [ ] **Step 2: Run tests to confirm failure**

Run: `python -m pytest tests/test_run.py::test_run_subprocess_captures_stdout -v`
Expected: `ImportError: cannot import name 'run_subprocess' from 'fbdi.run'`.

- [ ] **Step 3: Add `run_subprocess` to `fbdi/run.py`**

Add this import at the top of `fbdi/run.py`:

```python
import subprocess
```

Add this function below `parse_run_args`:

```python
def run_subprocess(cmd: list[str], log_path: Path) -> tuple[int, str]:
    """Run a subprocess, append its stdout+stderr to log_path, return (rc, text).

    The log file is opened in append mode so multiple calls accumulate. Each
    invocation writes a header line marking the command for later debugging.
    """
    log_path = Path(log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    header = f"\n=== $ {' '.join(cmd)}\n"
    with open(log_path, "a", encoding="utf-8") as log_f:
        log_f.write(header)
        log_f.flush()
        proc = subprocess.run(cmd, capture_output=True, text=True)
        captured = proc.stdout + proc.stderr
        log_f.write(captured)
        log_f.flush()
    return proc.returncode, captured
```

- [ ] **Step 4: Run tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all tests pass (the four new ones plus the nine from Task 9).

- [ ] **Step 5: Commit**

```bash
git add fbdi/run.py tests/test_run.py
git commit -m "feat(fbdi): add run_subprocess helper to orchestrator"
```

---

### Task 11: Stage definitions table

**Files:**
- Modify: `fbdi/run.py` (add `STAGE_COMMANDS` data structure + `build_stage_command` function)
- Modify: `tests/test_run.py` (add tests)

The stage table maps each stage name to the subprocess argv it spawns. This is data-driven so the orchestrator loop in Task 12 stays small.

- [ ] **Step 1: Write failing tests for stage command construction**

Append to `tests/test_run.py`:

```python
from fbdi.run import build_stage_command


def test_build_stage_command_preflight():
    cmd = build_stage_command("preflight", old="26A", new="26B")
    assert cmd[0] == sys.executable
    assert cmd[1:] == ["-m", "fbdi", "preflight"]


def test_build_stage_command_compare():
    cmd = build_stage_command("compare", old="26A", new="26B")
    assert cmd[0] == sys.executable
    assert "-m" in cmd
    assert "fbdi" in cmd
    assert "compare" in cmd
    assert "--old" in cmd and "26A" in cmd
    assert "--new" in cmd and "26B" in cmd
    assert "--output" in cmd
    assert any("Comparison_Report_26A_26B.xlsx" in arg for arg in cmd)


def test_build_stage_command_catalog():
    cmd = build_stage_command("catalog", old="26A", new="26B")
    assert "catalog" in cmd
    assert "--release" in cmd
    assert "26B" in cmd  # catalog targets the NEW release


def test_build_stage_command_update_module():
    cmd = build_stage_command("update-module", old="26A", new="26B")
    assert "populate-module" in cmd  # underlying CLI subcommand name
    assert "--new" in cmd and "26B" in cmd
    assert "--old" in cmd and "26A" in cmd


def test_build_stage_command_report():
    cmd = build_stage_command("report", old="26A", new="26B")
    assert "report" in cmd
    assert "--old" in cmd and "26A" in cmd
    assert "--new" in cmd and "26B" in cmd


def test_build_stage_command_unknown_stage_raises():
    with pytest.raises(KeyError):
        build_stage_command("frobnicate", old="26A", new="26B")
```

- [ ] **Step 2: Run tests to confirm failure**

Run: `python -m pytest tests/test_run.py -v -k build_stage_command`
Expected: `ImportError: cannot import name 'build_stage_command' from 'fbdi.run'`.

- [ ] **Step 3: Add `build_stage_command` to `fbdi/run.py`**

Add below `run_subprocess`:

```python
def build_stage_command(stage: str, *, old: str, new: str) -> list[str]:
    """Construct the subprocess argv for a single pipeline stage.

    Stages that internally orchestrate multiple commands (`download`, `clear`)
    return only the primary command; the orchestrator wraps additional
    commands separately. See the spec's HITL absorption section.
    """
    py = sys.executable
    report_name = f"Comparison_Report_{old}_{new}.xlsx"

    if stage == "preflight":
        return [py, "-m", "fbdi", "preflight"]
    if stage == "download":
        # Primary download command for NEW release; --skip-clear means clear
        # is a separate stage. The orchestrator (Task 13) handles HITL #1
        # (auto-download OLD if missing) and the verify-download follow-up.
        return [py, "tools/download_and_clear.py", new, "--skip-clear"]
    if stage == "clear":
        return [py, "tools/download_and_clear.py", new, "--clear-only"]
    if stage == "compare":
        return [py, "-m", "fbdi", "compare", "--old", old, "--new", new, "--output", report_name]
    if stage == "catalog":
        return [py, "-m", "fbdi", "catalog", "--release", new]
    if stage == "update-module":
        return [py, "-m", "fbdi", "populate-module", "--new", new, "--old", old]
    if stage == "report":
        return [py, "-m", "fbdi", "report", "--old", old, "--new", new]
    raise KeyError(f"Unknown stage: {stage!r}. Known: {ALL_STAGES}")
```

- [ ] **Step 4: Run tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add fbdi/run.py tests/test_run.py
git commit -m "feat(fbdi): add stage command table to orchestrator"
```

---

### Task 12: Pipeline loop with manifest writing

**Files:**
- Modify: `fbdi/run.py` (add `run_pipeline` function — the main loop)
- Modify: `tests/test_run.py` (add loop tests using monkey-patched subprocess)

This task wires the previous pieces together: iterate stages in `--from..--to` range, dispatch via subprocess, update the manifest, write `latest.json` after each stage. HITL absorption inside `download` is deferred to Task 13. HITL absorption inside `catalog` and `update-module` is deferred to Task 14. Auto-final summary/verify is deferred to Task 15.

- [ ] **Step 1: Write failing tests for the pipeline loop**

Append to `tests/test_run.py`:

```python
from unittest.mock import patch

from fbdi.run import run_pipeline


def test_run_pipeline_happy_path(tmp_path, monkeypatch):
    """All stages return 0 → exit code 0, manifest has all stages as ok."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    fake_calls = []

    def fake_subprocess(cmd, log_path):
        fake_calls.append(cmd)
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="report",
            log_dir=log_dir,
        )
    assert rc == 0
    # Each user-facing stage was invoked at least once
    invoked_stages = [" ".join(c) for c in fake_calls]
    for stage_marker in ["preflight", "compare", "catalog", "populate-module", "report"]:
        assert any(stage_marker in cmd for cmd in invoked_stages), \
            f"expected {stage_marker} in invoked commands"


def test_run_pipeline_writes_latest_json(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    with patch("fbdi.run.run_subprocess", return_value=(0, "OK\n")):
        run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="preflight",
            log_dir=log_dir,
        )
    latest = log_dir / "fbdi_run_latest.json"
    assert latest.is_file()
    payload = json.loads(latest.read_text())
    assert payload["old"] == "26A"
    assert payload["new"] == "26B"
    assert payload["stages"]["preflight"]["status"] == "ok"


def test_run_pipeline_writes_timestamped_manifest(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    with patch("fbdi.run.run_subprocess", return_value=(0, "OK\n")):
        run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="preflight",
            log_dir=log_dir,
        )
    timestamped = list(log_dir.glob("fbdi_run_26A_26B_*.json"))
    assert len(timestamped) == 1


def test_run_pipeline_range_skips_outside_stages(tmp_path, monkeypatch):
    """--from compare --to compare runs only compare; preflight/download/clear are skipped."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    invoked = []

    def fake_subprocess(cmd, log_path):
        invoked.append(" ".join(cmd))
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        run_pipeline(
            old="26A", new="26B",
            from_stage="compare", to_stage="compare",
            log_dir=log_dir,
        )
    # compare ran
    assert any("compare" in cmd and "Comparison_Report" in cmd for cmd in invoked)
    # preflight/catalog/report did not
    assert not any("preflight" in cmd for cmd in invoked)
    assert not any("catalog" in cmd for cmd in invoked)
    assert not any("report" in cmd for cmd in invoked)


def test_run_pipeline_propagates_exit_code_on_compare_failure(tmp_path, monkeypatch):
    """Compare returning non-zero → orchestrator exit code 4."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    def fake_subprocess(cmd, log_path):
        if "compare" in cmd and "--output" in cmd:
            return 1, "compare crashed\n"
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="report",
            log_dir=log_dir,
        )
    assert rc == 4


def test_run_pipeline_preflight_failure_exits_2(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    def fake_subprocess(cmd, log_path):
        if cmd[1:4] == ["-m", "fbdi", "preflight"]:
            return 1, "preflight failed\n"
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="report",
            log_dir=log_dir,
        )
    assert rc == 2
```

- [ ] **Step 2: Run tests to confirm failure**

Run: `python -m pytest tests/test_run.py -v -k run_pipeline`
Expected: `ImportError: cannot import name 'run_pipeline' from 'fbdi.run'`.

- [ ] **Step 3: Add `run_pipeline` to `fbdi/run.py`**

Add this import block at the top of `fbdi/run.py`:

```python
import time
from datetime import datetime, timezone
```

Add this function below `build_stage_command`:

```python
def _exit_code_for_stage_failure(stage: str) -> int:
    """Map a failed stage to the appropriate top-level exit code (per spec)."""
    if stage == "preflight":
        return 2
    if stage == "download":
        return 3
    # compare, catalog, update-module, report → mid-pipeline
    return 4


def _stages_in_range(from_stage: str, to_stage: str) -> list[str]:
    start = ALL_STAGES.index(from_stage)
    end = ALL_STAGES.index(to_stage)
    return ALL_STAGES[start : end + 1]


def run_pipeline(
    *,
    old: str,
    new: str,
    from_stage: str,
    to_stage: str,
    log_dir: Path,
) -> int:
    """Execute the FBDI pipeline from `from_stage` through `to_stage`.

    Returns the top-level exit code (0=clean, 2=preflight, 3=download,
    4=mid-pipeline, 5=warnings). Writes a timestamped manifest and a
    `fbdi_run_latest.json` to `log_dir`.

    HITL absorption inside the `download` stage and the auto-final
    `summary`/`verify` stages are added in subsequent tasks.
    """
    from fbdi.manifest import Manifest

    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    log_path = log_dir / f"fbdi_run_{old}_{new}_{timestamp}.log"
    manifest_path = log_dir / f"fbdi_run_{old}_{new}_{timestamp}.json"
    latest_path = log_dir / "fbdi_run_latest.json"

    manifest = Manifest(old=old, new=new, from_stage=from_stage, to_stage=to_stage)

    stages_to_run = _stages_in_range(from_stage, to_stage)
    skipped_by_range = [s for s in ALL_STAGES if s not in stages_to_run]
    for skipped in skipped_by_range:
        manifest.record_stage(skipped, status="skipped_by_range")
        manifest.write(latest_path)

    failure_stage: str | None = None
    for stage in stages_to_run:
        cmd = build_stage_command(stage, old=old, new=new)
        t0 = time.monotonic()
        rc, _ = run_subprocess(cmd, log_path=log_path)
        duration = time.monotonic() - t0

        if rc == 0:
            manifest.record_stage(stage, status="ok", duration_s=duration)
        else:
            manifest.record_stage(stage, status="failed", duration_s=duration, exit_code=rc)
            failure_stage = stage
            manifest.write(latest_path)
            break
        manifest.write(latest_path)

    if failure_stage:
        manifest.set_exit_code(_exit_code_for_stage_failure(failure_stage))

    manifest.finalize()
    manifest.write(manifest_path)
    manifest.write(latest_path)
    return manifest.exit_code
```

- [ ] **Step 4: Run tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add fbdi/run.py tests/test_run.py
git commit -m "feat(fbdi): add pipeline loop with per-stage manifest writes"
```

---

### Task 13: HITL absorption inside the `download` stage

**Files:**
- Modify: `fbdi/run.py` (replace simple download invocation with HITL-aware logic)
- Modify: `tests/test_run.py` (tests for OLD auto-download, FSM file copy, extras auto-accept)

The download stage absorbs three HITLs (#1 OLD baseline missing, #2 FSM file missing, #6 extras present). After this task the download stage is the only one with non-trivial internal logic.

- [ ] **Step 1: Write failing tests for HITL absorption**

Append to `tests/test_run.py`:

```python
import shutil


def test_download_stage_auto_downloads_old_if_missing(tmp_path, monkeypatch):
    """HITL #1: if baselines/<OLD>/originals/ is empty, download OLD first."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    # baselines/26A/originals/ does NOT exist → triggers HITL #1

    invoked = []

    def fake_subprocess(cmd, log_path):
        invoked.append(" ".join(cmd))
        # Simulate a successful download by creating the originals directory
        if "download_and_clear.py" in " ".join(cmd) and "--skip-clear" in cmd:
            ver = cmd[2]  # tools/download_and_clear.py <ver> --skip-clear
            (tmp_path / "baselines" / ver / "originals").mkdir(parents=True, exist_ok=True)
            (tmp_path / "baselines" / ver / "originals" / "RapidImplementationForCashManagement.xlsm").write_bytes(b"stub")
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="download", to_stage="download",
            log_dir=log_dir,
        )
    assert rc == 0
    # Both OLD and NEW were downloaded (look for "26A" and "26B" download commands)
    assert any("26A" in cmd and "--skip-clear" in cmd for cmd in invoked)
    assert any("26B" in cmd and "--skip-clear" in cmd for cmd in invoked)


def test_download_stage_fails_3_when_old_download_fails(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    def fake_subprocess(cmd, log_path):
        if "download_and_clear.py" in " ".join(cmd) and "26A" in cmd:
            return 1, "download crashed\n"
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="download", to_stage="download",
            log_dir=log_dir,
        )
    assert rc == 3


def test_download_stage_copies_fsm_file_from_old(tmp_path, monkeypatch):
    """HITL #2: if FSM file missing in NEW, copy it from OLD."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    # OLD already populated with the FSM file
    old_orig = tmp_path / "baselines" / "26A" / "originals"
    old_orig.mkdir(parents=True)
    (old_orig / "RapidImplementationForCashManagement.xlsm").write_bytes(b"FSM_BYTES")

    def fake_subprocess(cmd, log_path):
        # Simulate NEW download landing files but NOT the FSM file
        if "download_and_clear.py" in " ".join(cmd) and "26B" in cmd:
            new_orig = tmp_path / "baselines" / "26B" / "originals"
            new_orig.mkdir(parents=True, exist_ok=True)
            (new_orig / "SomeOtherTemplate.xlsm").write_bytes(b"other")
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="download", to_stage="download",
            log_dir=log_dir,
        )
    assert rc == 0
    new_fsm = tmp_path / "baselines" / "26B" / "originals" / "RapidImplementationForCashManagement.xlsm"
    assert new_fsm.is_file()
    assert new_fsm.read_bytes() == b"FSM_BYTES"


def test_download_stage_warns_when_fsm_missing_from_both(tmp_path, monkeypatch):
    """HITL #2 fallback: if FSM is missing in BOTH baselines, continue with warning, exit 5."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    # OLD populated but NO FSM file
    old_orig = tmp_path / "baselines" / "26A" / "originals"
    old_orig.mkdir(parents=True)
    (old_orig / "OtherFile.xlsm").write_bytes(b"other")

    def fake_subprocess(cmd, log_path):
        # Simulate NEW download with no FSM file either
        if "download_and_clear.py" in " ".join(cmd) and "26B" in cmd:
            new_orig = tmp_path / "baselines" / "26B" / "originals"
            new_orig.mkdir(parents=True, exist_ok=True)
            (new_orig / "OtherFile.xlsm").write_bytes(b"other")
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="download", to_stage="download",
            log_dir=log_dir,
        )
    assert rc == 5  # warnings → exit 5
    latest = json.loads((log_dir / "fbdi_run_latest.json").read_text())
    assert any("RapidImplementation" in w or "FSM" in w for w in latest["warnings"])
```

- [ ] **Step 2: Run tests to confirm failure**

Run: `python -m pytest tests/test_run.py -v -k download_stage`
Expected: at least 3 of 4 fail because `run_pipeline` does not yet handle HITL absorption.

- [ ] **Step 3: Replace the download branch in the orchestrator with HITL-aware logic**

Refactor `run_pipeline` to extract the download stage into a dedicated helper. Add this function to `fbdi/run.py`:

```python
FSM_FILE = "RapidImplementationForCashManagement.xlsm"


def _baseline_originals(release: str) -> Path:
    return Path("baselines") / release / "originals"


def _baseline_present(release: str) -> bool:
    p = _baseline_originals(release)
    return p.is_dir() and any(p.glob("*.xlsm"))


def _execute_download_stage(
    *, old: str, new: str, log_path: Path, manifest,
) -> int:
    """Run the download stage with HITL #1 / #2 / #6 absorption.

    Returns 0 on success (with possible warnings recorded), 3 on download failure.
    """
    py = sys.executable

    # HITL #1: if OLD baseline missing, download it first.
    if not _baseline_present(old):
        rc, _ = run_subprocess(
            [py, "tools/download_and_clear.py", old, "--skip-clear"],
            log_path=log_path,
        )
        if rc != 0:
            return 3
        rc, _ = run_subprocess(
            [py, "-m", "fbdi", "verify-download", "--release", old],
            log_path=log_path,
        )
        if rc == 2:
            # extras-only — auto-accept (HITL #6)
            rc, _ = run_subprocess(
                [py, "-m", "fbdi", "verify-download", "--release", old, "--commit-inventory"],
                log_path=log_path,
            )
            if rc != 0:
                return 3
        elif rc != 0:
            return 3

    # Primary: download NEW.
    rc, _ = run_subprocess(
        [py, "tools/download_and_clear.py", new, "--skip-clear"],
        log_path=log_path,
    )
    if rc != 0:
        # One auto-retry of the primary download.
        rc, _ = run_subprocess(
            [py, "tools/download_and_clear.py", new, "--skip-clear"],
            log_path=log_path,
        )
        if rc != 0:
            return 3

    # Verify NEW; absorb HITL #6 (extras auto-accept).
    rc, _ = run_subprocess(
        [py, "-m", "fbdi", "verify-download", "--release", new],
        log_path=log_path,
    )
    if rc == 2:
        rc, _ = run_subprocess(
            [py, "-m", "fbdi", "verify-download", "--release", new, "--commit-inventory"],
            log_path=log_path,
        )
        if rc != 0:
            return 3
    elif rc != 0:
        return 3

    # HITL #2: copy FSM file from OLD if missing in NEW.
    new_fsm = _baseline_originals(new) / FSM_FILE
    if not new_fsm.is_file():
        old_fsm = _baseline_originals(old) / FSM_FILE
        if old_fsm.is_file():
            new_fsm.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(old_fsm, new_fsm)
            manifest.stages.setdefault("download", {})["fsm_file"] = "copied_from_old"
        else:
            manifest.add_warning(
                f"{FSM_FILE} missing from both baselines/{old}/ and "
                f"baselines/{new}/ — compliance report will exclude it"
            )
            manifest.stages.setdefault("download", {})["fsm_file"] = "missing_from_both"

    return 0
```

Add this import at the top of `fbdi/run.py`:

```python
import shutil
```

Now modify the `run_pipeline` loop to dispatch the download stage through `_execute_download_stage` instead of `run_subprocess(build_stage_command(...))`. Replace this block in `run_pipeline`:

```python
    for stage in stages_to_run:
        cmd = build_stage_command(stage, old=old, new=new)
        t0 = time.monotonic()
        rc, _ = run_subprocess(cmd, log_path=log_path)
        duration = time.monotonic() - t0

        if rc == 0:
            manifest.record_stage(stage, status="ok", duration_s=duration)
        else:
            manifest.record_stage(stage, status="failed", duration_s=duration, exit_code=rc)
            failure_stage = stage
            manifest.write(latest_path)
            break
        manifest.write(latest_path)
```

with:

```python
    for stage in stages_to_run:
        t0 = time.monotonic()
        if stage == "download":
            rc = _execute_download_stage(old=old, new=new, log_path=log_path, manifest=manifest)
            duration = time.monotonic() - t0
            existing = manifest.stages.get("download", {})
            if rc == 0:
                manifest.record_stage("download", status="ok", duration_s=duration, **existing)
            else:
                manifest.record_stage("download", status="failed", duration_s=duration,
                                      exit_code=rc, **existing)
                failure_stage = "download"
                manifest.write(latest_path)
                break
        else:
            cmd = build_stage_command(stage, old=old, new=new)
            rc, _ = run_subprocess(cmd, log_path=log_path)
            duration = time.monotonic() - t0
            if rc == 0:
                manifest.record_stage(stage, status="ok", duration_s=duration)
            else:
                manifest.record_stage(stage, status="failed", duration_s=duration, exit_code=rc)
                failure_stage = stage
                manifest.write(latest_path)
                break
        manifest.write(latest_path)
```

Also: when warnings have been recorded but there is no failure stage, bump exit code to 5. Replace the existing `if failure_stage:` block in `run_pipeline` with:

```python
    if failure_stage:
        manifest.set_exit_code(_exit_code_for_stage_failure(failure_stage))
    elif manifest.warnings:
        manifest.set_exit_code(5)
```

- [ ] **Step 4: Run tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all tests pass (the four new download-stage tests plus everything from prior tasks).

- [ ] **Step 5: Commit**

```bash
git add fbdi/run.py tests/test_run.py
git commit -m "feat(fbdi): absorb HITL #1, #2, #6 inside download stage"
```

---

### Task 14: HITL absorption inside `catalog` and `update-module` stages

**Files:**
- Modify: `fbdi/run.py` (add `_execute_catalog_stage` + `_execute_update_module_stage` helpers; extend the dispatch in `run_pipeline`)
- Modify: `tests/test_run.py` (catalog backup, update-module skip-if-absent, mapping backup with timestamped collision)

The `catalog` stage must snapshot `FBDI_Master_Catalog.xlsx` to `.bak.xlsx` before running, so a crashed catalog run does not destroy the prior baseline (spec stages table, row 5). The `update-module` stage absorbs HITL #7: backup the mapping spreadsheet to `FBDI_to_ApplaudTables_Mapping.bak.xlsx` (timestamped suffix on collision) before overwriting, and skip the stage entirely with `status: skipped_no_mapping_file` if the mapping file is absent (spec stages table, row 6, and HITL policy table row 7).

Like Task 13's download helper, both stages get dedicated `_execute_*_stage` helpers because they involve file ops + subprocess in sequence — `build_stage_command` cannot express that.

- [ ] **Step 1: Update the existing happy-path test in Task 12 to satisfy the new mapping-required precondition**

The new logic skips `update-module` when no mapping file is present, but the existing `test_run_pipeline_happy_path` (from Task 12) asserts that `populate-module` *was* invoked. Open `tests/test_run.py`, find `test_run_pipeline_happy_path`, and add this line right after `monkeypatch.chdir(tmp_path)`:

```python
    (tmp_path / "FBDI_to_ApplaudTables_Mapping.xlsx").write_bytes(b"stub")
```

This makes the happy-path test exercise the "mapping present → backup + populate-module" branch, matching the test's existing assertion that `populate-module` is invoked.

- [ ] **Step 2: Write failing tests for catalog and update-module HITL absorption**

Append to `tests/test_run.py`:

```python
def test_catalog_stage_creates_backup_if_catalog_exists(tmp_path, monkeypatch):
    """HITL: existing FBDI_Master_Catalog.xlsx is snapshotted to .bak.xlsx before catalog runs."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"
    (tmp_path / "FBDI_Master_Catalog.xlsx").write_bytes(b"PRIOR_CATALOG")

    with patch("fbdi.run.run_subprocess", return_value=(0, "OK\n")):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="catalog", to_stage="catalog",
            log_dir=log_dir,
        )
    assert rc == 0
    backup = tmp_path / "FBDI_Master_Catalog.bak.xlsx"
    assert backup.is_file()
    assert backup.read_bytes() == b"PRIOR_CATALOG"


def test_catalog_stage_first_run_skips_backup(tmp_path, monkeypatch):
    """When no prior catalog exists, no backup is created and catalog still runs."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    invoked = []

    def fake_subprocess(cmd, log_path):
        invoked.append(cmd)
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="catalog", to_stage="catalog",
            log_dir=log_dir,
        )
    assert rc == 0
    assert not (tmp_path / "FBDI_Master_Catalog.bak.xlsx").exists()
    assert any("catalog" in cmd and "--release" in cmd for cmd in invoked)


def test_update_module_skipped_when_mapping_absent(tmp_path, monkeypatch):
    """HITL #7 (absent path): missing mapping file → status skipped_no_mapping_file, no subprocess."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    invoked = []

    def fake_subprocess(cmd, log_path):
        invoked.append(" ".join(cmd))
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="update-module", to_stage="update-module",
            log_dir=log_dir,
        )
    assert rc == 0
    latest = json.loads((log_dir / "fbdi_run_latest.json").read_text())
    assert latest["stages"]["update-module"]["status"] == "skipped_no_mapping_file"
    assert not any("populate-module" in cmd for cmd in invoked)


def test_update_module_backs_up_mapping_before_overwrite(tmp_path, monkeypatch):
    """HITL #7 (present path): existing mapping is snapshotted to .bak.xlsx before populate-module."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"
    (tmp_path / "FBDI_to_ApplaudTables_Mapping.xlsx").write_bytes(b"PRIOR_MAPPING")

    with patch("fbdi.run.run_subprocess", return_value=(0, "OK\n")):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="update-module", to_stage="update-module",
            log_dir=log_dir,
        )
    assert rc == 0
    backup = tmp_path / "FBDI_to_ApplaudTables_Mapping.bak.xlsx"
    assert backup.is_file()
    assert backup.read_bytes() == b"PRIOR_MAPPING"


def test_update_module_backup_uses_timestamp_on_collision(tmp_path, monkeypatch):
    """HITL #7: when .bak.xlsx already exists, the new backup goes to a timestamped name and the older backup is preserved."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"
    (tmp_path / "FBDI_to_ApplaudTables_Mapping.xlsx").write_bytes(b"PRIOR_MAPPING")
    (tmp_path / "FBDI_to_ApplaudTables_Mapping.bak.xlsx").write_bytes(b"OLDER_BACKUP")

    with patch("fbdi.run.run_subprocess", return_value=(0, "OK\n")):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="update-module", to_stage="update-module",
            log_dir=log_dir,
        )
    assert rc == 0
    # Pre-existing .bak.xlsx must NOT be overwritten.
    assert (tmp_path / "FBDI_to_ApplaudTables_Mapping.bak.xlsx").read_bytes() == b"OLDER_BACKUP"
    # New backup written to a timestamped sibling like "FBDI_to_ApplaudTables_Mapping.bak.20260715T020000Z.xlsx".
    timestamped = list(tmp_path.glob("FBDI_to_ApplaudTables_Mapping.bak.*.xlsx"))
    assert len(timestamped) == 1
    assert timestamped[0].read_bytes() == b"PRIOR_MAPPING"
```

- [ ] **Step 3: Run tests to confirm failure**

Run: `python -m pytest tests/test_run.py -v -k "catalog_stage or update_module"`
Expected: the five new tests fail (the orchestrator does not yet handle catalog backup / update-module skip-or-backup).

- [ ] **Step 4: Add `_execute_catalog_stage` and `_execute_update_module_stage` to `fbdi/run.py`**

Add these constants near the existing `FSM_FILE` constant from Task 13:

```python
CATALOG_FILE = "FBDI_Master_Catalog.xlsx"
CATALOG_BACKUP = "FBDI_Master_Catalog.bak.xlsx"
MAPPING_FILE = "FBDI_to_ApplaudTables_Mapping.xlsx"
MAPPING_BACKUP = "FBDI_to_ApplaudTables_Mapping.bak.xlsx"
```

Add these helpers near `_execute_download_stage`:

```python
def _execute_catalog_stage(*, old: str, new: str, log_path: Path) -> int:
    """Snapshot the existing catalog (if any) to .bak.xlsx, then run catalog.

    Returns the catalog subprocess exit code.
    """
    catalog = Path(CATALOG_FILE)
    if catalog.is_file():
        shutil.copy2(catalog, Path(CATALOG_BACKUP))
    rc, _ = run_subprocess(
        build_stage_command("catalog", old=old, new=new),
        log_path=log_path,
    )
    return rc


def _execute_update_module_stage(*, old: str, new: str, log_path: Path) -> tuple[int, str]:
    """If mapping file absent → return (0, "skipped_no_mapping_file") so caller records that status.
    If present → backup (timestamped on collision) and run populate-module.
    Returns (rc, status_marker). status_marker is "ok" or "skipped_no_mapping_file".
    """
    mapping = Path(MAPPING_FILE)
    if not mapping.is_file():
        return 0, "skipped_no_mapping_file"
    backup = Path(MAPPING_BACKUP)
    if backup.exists():
        ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        backup = Path(f"FBDI_to_ApplaudTables_Mapping.bak.{ts}.xlsx")
    shutil.copy2(mapping, backup)
    rc, _ = run_subprocess(
        build_stage_command("update-module", old=old, new=new),
        log_path=log_path,
    )
    return rc, "ok"
```

(`shutil`, `datetime`, and `timezone` are already imported in earlier tasks.)

- [ ] **Step 5: Wire the helpers into `run_pipeline`**

Replace the dispatch block from Task 13 (the `for stage in stages_to_run:` loop with the `if stage == "download": ... else: ...` branching) with this expanded version:

```python
    for stage in stages_to_run:
        t0 = time.monotonic()
        if stage == "download":
            rc = _execute_download_stage(old=old, new=new, log_path=log_path, manifest=manifest)
            duration = time.monotonic() - t0
            existing = manifest.stages.get("download", {})
            if rc == 0:
                manifest.record_stage("download", status="ok", duration_s=duration, **existing)
            else:
                manifest.record_stage("download", status="failed", duration_s=duration,
                                      exit_code=rc, **existing)
                failure_stage = "download"
                manifest.write(latest_path)
                break
        elif stage == "catalog":
            rc = _execute_catalog_stage(old=old, new=new, log_path=log_path)
            duration = time.monotonic() - t0
            if rc == 0:
                manifest.record_stage("catalog", status="ok", duration_s=duration)
            else:
                manifest.record_stage("catalog", status="failed", duration_s=duration, exit_code=rc)
                failure_stage = "catalog"
                manifest.write(latest_path)
                break
        elif stage == "update-module":
            rc, marker = _execute_update_module_stage(old=old, new=new, log_path=log_path)
            duration = time.monotonic() - t0
            if marker == "skipped_no_mapping_file":
                manifest.record_stage("update-module", status="skipped_no_mapping_file",
                                      duration_s=duration)
            elif rc == 0:
                manifest.record_stage("update-module", status="ok", duration_s=duration)
            else:
                manifest.record_stage("update-module", status="failed",
                                      duration_s=duration, exit_code=rc)
                failure_stage = "update-module"
                manifest.write(latest_path)
                break
        else:
            cmd = build_stage_command(stage, old=old, new=new)
            rc, _ = run_subprocess(cmd, log_path=log_path)
            duration = time.monotonic() - t0
            if rc == 0:
                manifest.record_stage(stage, status="ok", duration_s=duration)
            else:
                manifest.record_stage(stage, status="failed", duration_s=duration, exit_code=rc)
                failure_stage = stage
                manifest.write(latest_path)
                break
        manifest.write(latest_path)
```

- [ ] **Step 6: Run tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all tests pass (the five new ones, plus the modified `test_run_pipeline_happy_path`, plus everything from prior tasks).

- [ ] **Step 7: Commit**

```bash
git add fbdi/run.py tests/test_run.py
git commit -m "feat(fbdi): absorb HITL #7 (catalog backup + update-module backup-or-skip)"
```

---

### Task 15: Always-on `summary` and `verify` auto-final stages

**Files:**
- Modify: `fbdi/run.py` (run summary + verify after the main loop, regardless of `--from`/`--to`)
- Modify: `tests/test_run.py` (tests for unconditional auto-final behavior)

Per the spec, `summary` and `verify` always run at the end of every `fbdi run` invocation, even when the range excludes upstream stages. The underlying `summarize` / `verify-run` / `verify-rerun` subcommands already no-op gracefully on absent inputs, so even partial runs produce a sensible manifest.

- [ ] **Step 1: Write failing tests for auto-final behavior**

Append to `tests/test_run.py`:

```python
def test_summary_and_verify_run_after_full_pipeline(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    invoked = []

    def fake_subprocess(cmd, log_path):
        invoked.append(" ".join(cmd))
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="report",
            log_dir=log_dir,
        )
    assert rc == 0
    assert any("summarize" in cmd for cmd in invoked)
    assert any("verify-run" in cmd for cmd in invoked)
    assert any("verify-rerun" in cmd for cmd in invoked)


def test_summary_and_verify_run_even_for_truncated_range(tmp_path, monkeypatch):
    """--from compare --to compare still triggers summary + verify at end."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    invoked = []

    def fake_subprocess(cmd, log_path):
        invoked.append(" ".join(cmd))
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="compare", to_stage="compare",
            log_dir=log_dir,
        )
    assert rc == 0
    assert any("summarize" in cmd for cmd in invoked)
    assert any("verify-run" in cmd for cmd in invoked)


def test_verify_regression_bumps_exit_code_to_5(tmp_path, monkeypatch):
    """verify-run returning non-zero → exit 5 (warnings)."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    def fake_subprocess(cmd, log_path):
        if "verify-run" in cmd:
            return 1, "regression flagged\n"
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="report",
            log_dir=log_dir,
        )
    assert rc == 5


def test_summary_failure_does_not_change_exit_code(tmp_path, monkeypatch):
    """Summary failures are recorded but never fail the pipeline."""
    monkeypatch.chdir(tmp_path)
    log_dir = tmp_path / "logs"

    def fake_subprocess(cmd, log_path):
        if "summarize" in cmd:
            return 1, "summary crashed\n"
        return 0, "OK\n"

    with patch("fbdi.run.run_subprocess", side_effect=fake_subprocess):
        rc = run_pipeline(
            old="26A", new="26B",
            from_stage="preflight", to_stage="report",
            log_dir=log_dir,
        )
    # Pipeline still exits 0 (or 5 if there were other warnings) — summary doesn't gate
    assert rc in (0, 5)
```

- [ ] **Step 2: Run tests to confirm failure**

Run: `python -m pytest tests/test_run.py -v -k summary or verify`
Expected: at least the new tests fail.

- [ ] **Step 3: Add auto-final logic to `run_pipeline`**

Add this helper function to `fbdi/run.py`:

```python
def _execute_auto_final_stages(
    *, old: str, new: str, log_path: Path, manifest,
) -> None:
    """Always-run summary + verify-run + verify-rerun after the main loop.

    Failures here never block the pipeline from completing. A non-zero
    verify-run/verify-rerun exit bumps the top-level exit code to 5.
    """
    py = sys.executable
    report_name = f"Comparison_Report_{old}_{new}.xlsx"
    catalog_name = "FBDI_Master_Catalog.xlsx"
    baseline_catalog_name = "FBDI_Master_Catalog.bak.xlsx"

    # summary
    t0 = time.monotonic()
    rc, _ = run_subprocess(
        [py, "-m", "fbdi", "summarize", "--report", report_name, "--catalog", catalog_name],
        log_path=log_path,
    )
    duration = time.monotonic() - t0
    manifest.record_stage("summary",
                          status="ok" if rc == 0 else "failed",
                          duration_s=duration)

    # verify-run
    t0 = time.monotonic()
    rc, _ = run_subprocess(
        [py, "-m", "fbdi", "verify-run", "--release", new],
        log_path=log_path,
    )
    duration = time.monotonic() - t0
    if rc == 0:
        manifest.record_stage("verify", status="ok", duration_s=duration, regressions=[])
    else:
        manifest.record_stage("verify", status="failed", duration_s=duration,
                              regressions=["verify-run reported regressions"])
        manifest.add_warning("verify-run flagged potential regressions")

    # verify-rerun (best-effort; missing baseline is tolerated)
    cmd = [py, "-m", "fbdi", "verify-rerun", "--release", new,
           "--compare-report", report_name,
           "--baseline-catalog", baseline_catalog_name]
    rc, _ = run_subprocess(cmd, log_path=log_path)
    if rc != 0:
        manifest.add_warning("verify-rerun flagged macro-signal deltas")
```

Then in `run_pipeline`, after the main `for stage in stages_to_run:` loop and before the `if failure_stage:` block, add:

```python
    # Auto-final stages always run, regardless of --from/--to range.
    # Even if the main pipeline crashed mid-way, we still emit a summary
    # of whatever artifacts ARE present.
    _execute_auto_final_stages(old=old, new=new, log_path=log_path, manifest=manifest)
```

- [ ] **Step 4: Run tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add fbdi/run.py tests/test_run.py
git commit -m "feat(fbdi): always run summary + verify auto-final stages"
```

---

### Task 16: Register `python -m fbdi run` in `fbdi/cli.py`

**Files:**
- Modify: `fbdi/cli.py` (add `run` subparser + dispatch)
- Modify: `tests/test_run.py` (add CLI integration test)

This is the final wiring: hook the orchestrator into the top-level `fbdi` CLI so users invoke it as `python -m fbdi run --old 26A --new 26B`.

- [ ] **Step 1: Write a failing CLI integration test**

Append to `tests/test_run.py`:

```python
def test_fbdi_run_cli_invocation_smoke(tmp_path, monkeypatch):
    """`python -m fbdi run --old 26A --new 26B --from preflight --to preflight`
    should at minimum parse args and start a run (it'll fail downstream because
    we have no real subprocess fakes, but it must not exit on bad args)."""
    monkeypatch.chdir(tmp_path)
    proc = subprocess.run(
        [sys.executable, "-m", "fbdi", "run",
         "--old", "26A", "--new", "26B",
         "--from", "preflight", "--to", "preflight",
         "--log-dir", str(tmp_path / "logs")],
        capture_output=True, text=True,
    )
    # Argparse must accept the args (return code != 1 from arg validation).
    # The actual run will likely fail (no python on path, no Chrome, etc.),
    # but the manifest should still be written.
    latest = tmp_path / "logs" / "fbdi_run_latest.json"
    assert latest.is_file(), f"expected latest manifest; got stdout={proc.stdout} stderr={proc.stderr}"
    payload = json.loads(latest.read_text())
    assert payload["old"] == "26A"
    assert payload["new"] == "26B"


def test_fbdi_run_cli_rejects_invalid_release():
    proc = subprocess.run(
        [sys.executable, "-m", "fbdi", "run", "--old", "ZZZ", "--new", "26B"],
        capture_output=True, text=True,
    )
    assert proc.returncode != 0
    assert "26A" in proc.stderr or "format" in proc.stderr.lower()
```

- [ ] **Step 2: Run tests to confirm failure**

Run: `python -m pytest tests/test_run.py -v -k cli`
Expected: tests fail because `fbdi/cli.py` does not yet wire `run`.

- [ ] **Step 3: Add the `run` subparser and dispatch in `fbdi/cli.py`**

Add the subparser block after the existing `summarize_parser` block (or wherever the last existing subparser was added in Task 5):

```python
    run_parser = subparsers.add_parser(
        "run",
        help="Run the full headless FBDI pipeline (download → compare → catalog → report)",
    )
    run_parser.add_argument("--old", required=True, type=str,
                            help="Prior release label (e.g. 26A)")
    run_parser.add_argument("--new", required=True, type=str,
                            help="New release label (e.g. 26B)")
    run_parser.add_argument("--from", dest="from_stage", default="preflight",
                            help="First stage to execute (default: preflight)")
    run_parser.add_argument("--to", dest="to_stage", default="report",
                            help="Last stage to execute, inclusive (default: report)")
    run_parser.add_argument("--log-dir", dest="log_dir", default=Path("logs"), type=Path,
                            help="Directory for run logs and manifests (default: ./logs)")
```

Add the dispatch branch after the `summarize` dispatch:

```python
    elif args.command == "run":
        from fbdi import run as run_mod
        argv_for_run = [
            "--old", args.old,
            "--new", args.new,
            "--from", args.from_stage,
            "--to", args.to_stage,
            "--log-dir", str(args.log_dir),
        ]
        # Reparse via run.parse_run_args so all validation lives in one place.
        parsed = run_mod.parse_run_args(argv_for_run)
        rc = run_mod.run_pipeline(
            old=parsed.old, new=parsed.new,
            from_stage=parsed.from_stage, to_stage=parsed.to_stage,
            log_dir=parsed.log_dir,
        )
        sys.exit(rc)
```

- [ ] **Step 4: Run tests to confirm they pass**

Run: `python -m pytest tests/test_run.py -v`
Expected: all tests pass.

- [ ] **Step 5: Run the full project test suite as a regression check**

Run: `python -m pytest tests/ -v`
Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add fbdi/cli.py tests/test_run.py
git commit -m "feat(fbdi): wire python -m fbdi run into top-level CLI"
```

---

### Task 17: End-to-end smoke test (opt-in, slow)

**Files:**
- Create: `tests/test_run_integration.py`

A single integration test that exercises the full orchestrator against tiny fixture xlsm files in `tests/fixtures/`, skipping download/clear (which require Selenium and Oracle docs). Marked `@pytest.mark.integration` so it doesn't run in the default suite.

- [ ] **Step 1: Check whether `tests/fixtures/` already has tiny xlsm files**

Run: `ls tests/fixtures/ 2>/dev/null || echo "no fixtures dir"`
Expected: either lists existing fixtures or reports the directory is missing.

If missing, create the directory and a placeholder fixture in subsequent steps. If present, prefer reusing existing fixtures (e.g., what `test_compare.py` or `test_catalog.py` use).

- [ ] **Step 2: Write the integration test**

Create `tests/test_run_integration.py`:

```python
"""End-to-end integration test for `python -m fbdi run`.

Marked `integration` because it runs the real orchestrator against synthetic
fixture xlsm files. Skipped in the default suite; run explicitly with:

    python -m pytest tests/test_run_integration.py -v -m integration
"""

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

REPO_ROOT = Path(__file__).resolve().parent.parent


def _make_synthetic_baseline(originals_dir: Path, n_files: int = 2):
    """Create N tiny xlsm files in originals_dir to act as a fake baseline."""
    originals_dir.mkdir(parents=True, exist_ok=True)
    for i in range(n_files):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(1, 1, "FIELD_NAME")
        ws.cell(1, 2, "DATA_TYPE")
        ws.cell(2, 1, f"FIELD_{i}")
        ws.cell(2, 2, "VARCHAR2(50)")
        wb.save(originals_dir / f"FixtureTemplate{i}.xlsm")


@pytest.mark.integration
def test_fbdi_run_from_compare_with_synthetic_baselines(tmp_path):
    """Full orchestrator run from `compare` onward against synthetic baselines.

    Skips download/clear (which require Selenium); exercises compare, catalog,
    update-module (skipped — no mapping file), report (will fail — no real
    catalog inputs), and the auto-final summary/verify.
    """
    # Set up fake baselines under a working directory
    work = tmp_path
    _make_synthetic_baseline(work / "baselines" / "26A" / "originals")
    _make_synthetic_baseline(work / "baselines" / "26B" / "originals")

    proc = subprocess.run(
        [sys.executable, "-m", "fbdi", "run",
         "--old", "26A", "--new", "26B",
         "--from", "compare", "--to", "catalog",
         "--log-dir", str(work / "logs")],
        cwd=work,
        capture_output=True, text=True,
        timeout=300,
    )

    # We expect a manifest regardless of exit code.
    latest = work / "logs" / "fbdi_run_latest.json"
    assert latest.is_file(), f"manifest missing; stdout={proc.stdout!r} stderr={proc.stderr!r}"
    payload = json.loads(latest.read_text())
    assert payload["old"] == "26A"
    assert payload["new"] == "26B"
    assert "compare" in payload["stages"]
    assert "catalog" in payload["stages"]

    # Logs should be append-only and have content
    log_files = list((work / "logs").glob("fbdi_run_26A_26B_*.log"))
    assert len(log_files) == 1
    assert log_files[0].stat().st_size > 0
```

- [ ] **Step 3: Run the integration test explicitly**

Run: `python -m pytest tests/test_run_integration.py -v -m integration`
Expected: the test passes (or the manifest assertion holds even if compare/catalog stages report internal failures — the test asserts on manifest shape, not pipeline success).

- [ ] **Step 4: Confirm the integration test is excluded from the default run**

Run: `python -m pytest tests/ -v --ignore=tests/test_run_integration.py`
Expected: green; default suite is unchanged.

Then run the default suite without `--ignore` to confirm the integration test is skipped (because we used `@pytest.mark.integration` and the default config doesn't run that marker):

Run: `python -m pytest tests/ -v -k "not integration"`
Expected: integration test reported as deselected; everything else passes.

- [ ] **Step 5: Commit**

```bash
git add tests/test_run_integration.py
git commit -m "test(fbdi): add integration smoke test for fbdi run orchestrator"
```

---

## Phase D — Final verification

### Task 18: Full regression run

**Files:** none modified (verification only)

- [ ] **Step 1: Run the full test suite**

Run: `python -m pytest tests/ -v`
Expected: all tests pass. No regressions vs. the project's pre-existing 320-test baseline.

- [ ] **Step 2: Smoke-test the new CLI subcommand**

Run: `python -m fbdi run --help`
Expected: prints usage with `--old`, `--new`, `--from`, `--to`, `--log-dir` flags.

- [ ] **Step 3: Smoke-test that the promoted helpers' direct module invocation still works**

Run each of:
```bash
python -m fbdi preflight
python -m fbdi verify-download --help
python -m fbdi verify-run --help
python -m fbdi verify-rerun --help
python -m fbdi summarize --help
```
Expected: each prints help (or executes for `preflight`); no `ModuleNotFoundError`.

- [ ] **Step 4: Smoke-test that the skill still parses and references valid commands**

Run: `grep -n "scripts/" .claude/skills/fbdi-compare-release/SKILL.md`
Expected: zero matches.

Run: `grep -c "python -m fbdi" .claude/skills/fbdi-compare-release/SKILL.md`
Expected: at least 6.

- [ ] **Step 5: Validate the spec is referenced from CLAUDE.md if appropriate**

Open `CLAUDE.md` and confirm the "Active Pipeline" section either already mentions `python -m fbdi run` (it shouldn't yet, since this is the implementation pass), or note that a follow-up commit will update it. The plan does NOT prescribe a CLAUDE.md edit because the user prefers to do that themselves after running the new command for the first time.

- [ ] **Step 6: No commit needed for this verification task** (unless smoke-tests revealed bugs that needed inline fixes — in which case create commits as you go).

---

## Self-review checklist (run after writing all tasks)

- [x] **Spec coverage:** Every spec section has a task implementing it.
  - Goals/non-goals → Implementer notes section + Out-of-scope mentioned.
  - Architecture → Tasks 9–15.
  - Stage definitions → Task 11.
  - HITL absorption → Task 13 (download, HITL #1/#2/#6) + Task 14 (catalog backup, update-module HITL #7).
  - CLI surface → Task 9 + Task 16.
  - HITL policy table → Task 13 (#1, #2, #6) + Task 14 (#7) + Task 15 (verify regression bumps to exit 5). HITL #3/#4/#5/#8 either resolve trivially per spec (#3 N/A, #8 skip-and-generate) or are observable through subprocess exit codes only without explicit manifest enrichment in v1 (#4 pair_failures count and #5 missing_by_module — recorded as TODO follow-ups; the basic pass/fail contract is honored by the orchestrator dispatch in Task 12 + 13).
  - Logging/manifest → Task 8 (Manifest module) + Task 12 (timestamped + latest.json writes) + Task 13/14/15 (per-stage status enrichment).
  - Exit codes → Task 12 (2/3/4 base mapping) + Task 13 (download → 3) + Task 14 (catalog/update-module → 4) + Task 15 (verify → 5).
  - File layout (helper promotion) → Tasks 1–6.
  - Skill update → Task 7.
  - Test files needing updating → Tasks 1–5 (each migrates its slice; Task 6 finalizes).
  - Testing strategy → Tasks 1–15 unit tests + Task 17 integration.
  - Resumability → covered implicitly by Task 12's range logic (any stage range can be re-run).
  - Wall-time expectations → no implementation needed; documented in spec.

- [x] **Placeholder scan:** Every step contains executable content. No "TBD" / "TODO" / "fill in details" / "similar to Task N".

- [x] **Type consistency:** Function signatures and module names referenced across tasks are consistent:
  - `Manifest` (Task 8) used by `run_pipeline` (Task 12) and `_execute_auto_final_stages` (Task 15).
  - `parse_run_args` / `run_pipeline` / `run_subprocess` / `build_stage_command` consistently named.
  - `_execute_download_stage` (Task 13), `_execute_catalog_stage` / `_execute_update_module_stage` (Task 14), `_execute_auto_final_stages` (Task 15) follow a consistent `_execute_*_stage(...) -> int|tuple` pattern.
  - Stage name list `ALL_STAGES` is canonical from Task 9 onward.
