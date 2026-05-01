# CLAUDE.md — Oracle FBDI Pulldown

This file gives Claude Code persistent context for this project. Read it at the start of every session.

**For human readers:** [`docs/operator-guide.md`](docs/operator-guide.md) walks through running the quarterly refresh; [`docs/developer-guide.md`](docs/developer-guide.md) orients developers to the codebase.

---

## Project Purpose

This repo automates comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) between Oracle Cloud release versions. The goal is to identify field-level changes (added, removed, modified columns) across releases so that Brad and Dan can keep Definian's Oracle integrations current. The primary deliverable is a structured Excel comparison report.

---

## Quick Start

```bash
# Compare two releases (uses baselines/<ver>/originals/)
python -m fbdi compare --old 26A --new 26B --output Comparison_Report_26A_26B.xlsx

# Generate/update the FBDI master catalog for a release
python -m fbdi catalog --release 26B

# Diagnose header detection across releases
python -m fbdi diagnose --old baselines/26A/originals --new baselines/26B/originals --output Diagnostic_Report_26A_26B.xlsx

# Download + smart-clear FBDI templates for a new release
python tools/download_and_clear.py 26B                 # download + clear
python tools/download_and_clear.py 26B --clear-only    # re-clear only (skip download)

# Populate Module column in mapping spreadsheet (uses baselines/<ver>/file_modules.json)
python -m fbdi populate-module --new 26B --old 26A

# Tests
python -m pytest tests/            # full suite
python -m pytest tests/test_clear.py -v
```

**Requirements:** Python 3.14+. Install deps via `pip install -r requirements.txt` (openpyxl, selenium, webdriver-manager, requests, pytest).

---

## Active Pipeline (Built and Working)

- **`fbdi/` package** — Python comparison engine
  - `detect_header.py` — dynamically identifies the header row in each FBDI tab using content scoring (no hardcoded filename map). Uses `iter_rows` for streaming scans.
  - `compare.py` — diffs two releases tab-by-tab, field-by-field. Each pair runs in a fresh subprocess (via `_subprocess_util.run_worker`) with a 120s timeout to isolate openpyxl resource leaks.
  - `clear.py` — smart clearing of FBDI templates using `detect_header_row` (preserves headers at any row — 4, 5, 8, etc.)
  - `diagnose.py` — reports header-detection outcomes per tab (`DETECTED`, `NO_HEADER`, `SKIPPED_TAB`, `FILE_TOO_LARGE`, `FILE_ERROR`). Uses full (non-read_only) openpyxl mode.
  - `catalog.py` — generates `FBDI_Master_Catalog.xlsx` with per-release snapshots (file × tab × position × label × technical × type × length × scale × required) + `Issues` + `Drift` tabs. Shares `_subprocess_util.run_worker` with `compare.py`.
  - `type_parser.py` — parses Oracle data-type strings (`VARCHAR2(N CHAR)`, `NUMBER(p,s)`, `DATE`, `DATE(YYYY/MM/DD)`, `TimeStamp(hh24:mm)`, trailing-period variants) into structured fields. Emits `TYPE_PARSE_WARNING` only for genuinely malformed strings.
  - `_subprocess_util.py` — shared `run_worker(target, args, timeout)` helper used by `catalog.py` and `compare.py`. Drains the result queue *before* joining the child process — required to avoid a pipe-buffer deadlock on Windows when payloads exceed ~64 KB.
  - `catalog_normalize.py` — normalizes FBDI labels (strips non-alphanumeric/underscore/whitespace) for Applaud MDB compatibility.
  - `build_mapping.py` — builds the `fbdi_applaud_mapping.xlsx` workbook that maps FBDI tabs/fields to Applaud target tables for downstream integrations.
  - `audit.py` — FBDI ↔ Applaud mapping audit engine. Reads `baselines/applaud/applaud_snapshot.json` (gitignored), `FBDI_Master_Catalog.xlsx`, and the working `fbdi_applaud_mapping.xlsx`. Two-pass signal scoring + adjudication; emits `Claude_fbdi_applaud_mapping.xlsx` and a markdown audit report.
  - `populate_module.py` — surgical column-F updater for `FBDI_to_ApplaudTables_Mapping.xlsx`. Reads `baselines/<ver>/file_modules.json` (NEW wins, OLD fallback). Uses openpyxl full mode so formatting/formulas/freeze-panes are preserved.
  - `cli.py` / `__main__.py` — CLI entry point. `_resolve_dir()` makes `--old 26A` resolve to `baselines/26A/originals/`.
  - `config.py`, `utils.py` — shared configuration and helpers.
- **`tools/download_and_clear.py`** — standalone Selenium downloader + smart clearing entry point. Imports `fbdi.clear` but lives outside the `fbdi/` package so Selenium/webdriver dependencies stay out of the comparison engine.
- **`.claude/skills/fbdi-compare-release/`** — orchestrator skill that chains the full download → clear → compare → catalog → populate-module pipeline with human-in-the-loop checkpoints. Triggers on phrases like "Compare 26A to 26B" or "Oracle released 26C". Bundles Python helpers (`check_env.py`, `verify_download.py`, `summarize_report.py`, `verify_run.py`, `verify_rerun.py`) under `scripts/` and reference docs under `references/`. See the skill's `SKILL.md` for the 8-stage workflow (plus Stage 6.5 for Module column population).
- **`tests/`** — 281 unit tests, all passing (`python -m pytest tests/`)
- **Outputs:**
  - `Comparison_Report_<OLD>_<NEW>.xlsx` — 7-column diff for VBA validation (unchanged)
  - `FBDI_Master_Catalog.xlsx` — per-release snapshots + Issues + Drift tabs
- **Baseline layout** — `baselines/26A/originals/` (as-downloaded), `baselines/26A/blanks/` (smart-cleared copies for client use), and `baselines/26A/file_modules.json` (per-release `{filename: module}` map written by the downloader, consumed by `populate-module`)

---

## Current Frontier

- **FBDI → Applaud mapping** — `fbdi_applaud_mapping.xlsx` (built by `fbdi/build_mapping.py`) is partially populated; Brad fills in TBD rows manually. The Module column is now auto-populated via `python -m fbdi populate-module` from `file_modules.json` (100% as of 2026-05-01 rerun).
- **`report.py`** (not built) — Will reformat comparison output into the compliance change-tracking format used for client deliverables. Blocked on mapping completion.
- **`python -m fbdi run`** (not built) — Would chain download → compare → report in a single command.

---

## Key Architectural Decisions (Closed — Do Not Re-litigate)

| Decision | Choice |
|---|---|
| Baseline storage | Folder-based: `baselines/26A/originals/`, `baselines/26A/blanks/` — gitignored, not Git-tracked |
| Header detection | Dynamic content scoring per tab — no hardcoded filename-to-header map |
| Excel reading | `openpyxl` with `data_only=True` where formula evaluation is needed |
| Comparison output | 7-column `.xlsx` — columns A–G as specified |
| Column scan cap | Max 500 columns per tab (avoids phantom `max_column=16384` from corrupt xlsm metadata) |

---

## Known Hazards

- **`RapidImplementationForCashManagement.xlsm` is not auto-downloadable** — this is an Oracle Rapid Implementation (FSM) template, not a standard FBDI template. It is not hosted on Oracle docs pages so the Selenium downloader never finds it. Must be obtained manually from Oracle Fusion: Setup and Maintenance → hamburger menu (top-right) → Search → search "Create Banks, Branches, and Accounts in Spreadsheet" → click the task to download. Place in `baselines/<VER>/originals/` before running compare. The `download_and_clear.py` script will warn if it's missing after a download run. Once placed, the compare engine picks it up automatically.
- **Phantom columns (`max_column=16384`)** — some xlsm files report 16384 columns due to corrupt metadata. The engine caps column scanning at 500.
- **Corrupt XML in some xlsm files** — handled gracefully; engine catches `zipfile.BadZipFile` and logs the file as unreadable. Stage 8 `verify_run` flags a regression if the per-release FILE_ERROR count jumps vs. the prior release.
- **`Comparison_Report_25D_26A.xlsx` (VBA output)** — has a corrupt stylesheet. Cannot be loaded with standard `openpyxl.load_workbook`. Use `read_only=True` or `data_only=True` with exception handling if you need to read it.
- **Diagnose and build_mapping are still bounded by `MAX_FILE_SIZE_BYTES` (5MB)** — they load workbooks in full (non-read_only) mode for memory reasons. Comparison is unbounded and streams via `iter_rows`.
- **JET `<oj-tree-view>` race in `tools/download_and_clear.py`** — Oracle docs put the TOC inside `<oj-tree-view>` under `#navigationDrawer`. The drawer container appears in DOM before the tree-view's `<li role="treeitem">` children populate. Without a wait for at least one treeitem, `find_elements(...#navigationDrawer li)` returns empty and the URL is silently skipped (no error, no SKIP log — page just immediately "Completed" with zero downloads). Fixed in commit 82cd568; keep the wait when refactoring the scraper.
- **`RapidImplementationForCashManagement.xlsm` fallback when both baselines wiped** — skill HITL #2's default "copy from prior baseline" assumes a prior is present. If a rerun wipes both 26A and 26B at once, fall back to an external archive (e.g., `C:/Users/10193/Definian/<old release>_*_Compare/<old>_FBDI/Manual/RapidImplementationForCashManagement.xlsm`).

## Resolved Hazards (historical note)

- ~~6 files >5MB are currently skipped~~ — fixed by subprocess isolation + `iter_rows` optimization. Comparison now processes all 210 file pairs with no size limit.
- ~~~8 tabs with non-standard headers fail detection~~ — fixed in Phase 3. Diagnose reports `NO_HEADER: 0`.
- ~~Full comparison run is ~75 minutes~~ — much faster now due to `iter_rows` streaming (74s → 0.02s per tab on wide sheets).
- ~~`ChangeOrderImportTemplate` and `ItemImportTemplate` report bogus TIMEOUT in the catalog~~ — fixed by extracting `_subprocess_util.run_worker` with drain-before-join semantics. Catalog now fully ingests both files (~1,400 rows each per release).
- ~~463 TYPE_PARSE_WARNING rows in the catalog Issues tab~~ — collapsed to 9 (only the genuinely-broken Oracle strings) after `type_parser.py` was extended to accept temporal format masks (`DATE(YYYY/MM/DD)`, `TimeStamp(hh24:mm:ss)`) and the stray-trailing-period typo.
- ~~Financials URL silently scrapes 0 files (`<oj-tree-view>` race)~~ — fixed by waiting for `[role='treeitem']` to populate inside `#navigationDrawer` before iterating sections (commit 82cd568).

---

## Reference Files

Two read-only archives, distinct purposes:

- **`reference/`** — pre-Python pipeline artifacts (legacy VBA macros, Dan's original Selenium downloader). Do not modify.
- **`docs/archive/`** — historical narrative docs (audit notes, scraper gap findings) preserved via `git mv` so blame history survives. Do not modify.

`reference/` contents:

| File | What It Is |
|---|---|
| `fbdi_compare.xlsm` | Legacy VBA macro that did the comparison before the Python engine |
| `Clear_FBDIs - 20210412.xlsm` | Legacy VBA macro that cleared template files before re-download |
| `Oracle_26A_Comparison_Report.docx` | Sample output from the VBA macro for release 26A — used as a reference baseline during Phase 1 |
| `test.py` | Dan's original Selenium downloader — scrapes Oracle docs and downloads xlsm files |

---

## Testing

- `python -m pytest tests/` — run full suite (281 tests)
- `python -m pytest tests/test_clear.py -v` — run one module
- `tests/validate_against_vba.py` and `tests/vba_fieldrow_map.json` — ad-hoc validation against the legacy VBA macro's expected header rows (not pytest, kept for spot-checks against regressions)

Tests use `openpyxl.Workbook` to build synthetic FBDI files per test (see `_make_fbdi_workbook` / `_create_fbdi_workbook` helpers). There are no fixtures — fixture-like workbooks are built inline in each test so the expected layout is visible next to the assertions.

**Test-data gotcha:** `detect_header_row` scores rows by UPPER_SNAKE_CASE content. Synthetic sample data like `"CREATE"`, `"V1"`, `"DR_ECO_1"` will false-positive as headers. Use lowercase/mixed-case values in test data rows (e.g. `"Create Order"`, `"V1-org"`).

---

## Docs & Planning

Two patterns have been used in this repo — both are still valid:

- **Old:** `handoff_*.md` files (written in Claude Chat, executed by Claude Code) — now gitignored, kept in conversation or local scratch.
- **Current:** `docs/superpowers/specs/*.md` (design) and `docs/superpowers/plans/*.md` (implementation plans) — produced via the `superpowers:brainstorming` and `superpowers:writing-plans` skills and committed to the repo so the history is auditable.

Completed-project narrative docs (audit notes, one-off findings) live in `docs/archive/`. The two user-facing guides live at `docs/operator-guide.md` and `docs/developer-guide.md`.

---

## Plugins / Tooling

Project uses the `superpowers` skill family (brainstorming, writing-plans, executing-plans, systematic-debugging, verification-before-completion). CodeRabbit is wired up for PR review. See user-level `~/.claude/` config for the full plugin list — no project-specific plugin requirements.

## graphify

This project has a graphify knowledge graph at graphify-out/.

Rules:
- Before answering architecture or codebase questions, read graphify-out/GRAPH_REPORT.md for god nodes and community structure
- If graphify-out/wiki/index.md exists, navigate it instead of reading raw files
- After modifying code files in this session, run `python3 -c "from graphify.watch import _rebuild_code; from pathlib import Path; _rebuild_code(Path('.'))"` to keep the graph current
