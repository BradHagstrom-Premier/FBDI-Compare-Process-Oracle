# CLAUDE.md — Oracle FBDI Pulldown

This file gives Claude Code persistent context for this project. Read it at the start of every session.

---

## Project Purpose

This repo automates comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) between Oracle Cloud release versions. The goal is to identify field-level changes (added, removed, modified columns) across releases so that Brad and Dan can keep Definian's Oracle integrations current. The primary deliverable is a structured Excel comparison report.

---

## Quick Start

```bash
# Compare two releases (uses baselines/<ver>/originals/)
python -m fbdi compare --old 26A --new 26B --output Comparison_Report_26A_26B.xlsx

# Diagnose header detection across releases
python -m fbdi diagnose --old baselines/26A/originals --new baselines/26B/originals --output Diagnostic_Report_26A_26B.xlsx

# Download + smart-clear FBDI templates for a new release
python tools/download_and_clear.py 26B                 # download + clear
python tools/download_and_clear.py 26B --clear-only    # re-clear only (skip download)

# Tests
python -m pytest tests/            # full suite
python -m pytest tests/test_clear.py -v
```

**Requirements:** Python 3.14+, `openpyxl`, `selenium`, `webdriver-manager`, `requests`, `pytest`. No `requirements.txt` — dependencies are installed ad-hoc.

---

## Active Pipeline (Built and Working)

- **`fbdi/` package** — Python comparison engine
  - `detect_header.py` — dynamically identifies the header row in each FBDI tab using content scoring (no hardcoded filename map). Uses `iter_rows` for streaming scans.
  - `compare.py` — diffs two releases tab-by-tab, field-by-field. Each pair runs in a `multiprocessing.Process` with a 120s timeout (prevents openpyxl resource-leak hangs).
  - `clear.py` — smart clearing of FBDI templates using `detect_header_row` (preserves headers at any row — 4, 5, 8, etc.)
  - `diagnose.py` — reports header-detection outcomes per tab (`DETECTED`, `NO_HEADER`, `SKIPPED_TAB`, `FILE_TOO_LARGE`, `FILE_ERROR`). Uses full (non-read_only) openpyxl mode.
  - `build_mapping.py` — builds the `fbdi_applaud_mapping.xlsx` workbook that maps FBDI tabs/fields to Applaud target tables for downstream integrations.
  - `cli.py` / `__main__.py` — CLI entry point. `_resolve_dir()` makes `--old 26A` resolve to `baselines/26A/originals/`.
  - `config.py`, `utils.py` — shared configuration and helpers.
- **`tools/download_and_clear.py`** — standalone Selenium downloader + smart clearing entry point. Imports `fbdi.clear` but lives outside the `fbdi/` package so Selenium/webdriver dependencies stay out of the comparison engine.
- **`tests/`** — 54 unit tests, all passing (`python -m pytest tests/`)
- **Output** — `Comparison_Report_<OLD>_<NEW>.xlsx` — 7-column format (columns A–G): FBDI File, FBDI Tab, Column Letter, Column Number, Old FBDI Field Name, New FBDI Field Name, Difference?
- **Baseline layout** — `baselines/26A/originals/` (as-downloaded) and `baselines/26A/blanks/` (smart-cleared copies for client use)

---

## Current Frontier

- **FBDI → Applaud mapping** — `fbdi_applaud_mapping.xlsx` (built by `fbdi/build_mapping.py`) is partially populated. Brad is filling in TBD rows manually. See `Applaud Mapping Review To Do 04022026.md`.
- **`report.py`** (not built) — Will reformat comparison output into the Audrey change-tracking format used for client deliverables. Blocked on mapping completion.
- **`python -m fbdi run`** (not built) — Would chain download → compare → report in a single command.

See `NEXT_STEPS.md` for the prioritized backlog and historical phase-by-phase resolutions.

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

- **Phantom columns (`max_column=16384`)** — some xlsm files report 16384 columns due to corrupt metadata. The engine caps column scanning at 500.
- **Corrupt XML in some xlsm files** — handled gracefully; engine catches `zipfile.BadZipFile` and logs the file as unreadable. 26B has ~11 such files (diagnose reports FILE_ERROR).
- **`Comparison_Report_25D_26A.xlsx` (VBA output)** — has a corrupt stylesheet. Cannot be loaded with standard `openpyxl.load_workbook`. Use `read_only=True` or `data_only=True` with exception handling if you need to read it.
- **Diagnose and build_mapping are still bounded by `MAX_FILE_SIZE_BYTES` (5MB)** — they load workbooks in full (non-read_only) mode for memory reasons. Comparison is unbounded and streams via `iter_rows`.

## Resolved Hazards (historical note)

- ~~6 files >5MB are currently skipped~~ — fixed by subprocess isolation + `iter_rows` optimization. Comparison now processes all 210 file pairs with no size limit.
- ~~~8 tabs with non-standard headers fail detection~~ — fixed in Phase 3. Diagnose reports `NO_HEADER: 0`.
- ~~Full comparison run is ~75 minutes~~ — much faster now due to `iter_rows` streaming (74s → 0.02s per tab on wide sheets).

---

## Reference Files

`reference/` is a read-only archive. Do not modify these files. They exist for historical context only.

| File | What It Is |
|---|---|
| `fbdi_compare.xlsm` | Legacy VBA macro that did the comparison before the Python engine |
| `Clear_FBDIs - 20210412.xlsm` | Legacy VBA macro that cleared template files before re-download |
| `Oracle_26A_Comparison_Report.docx` | Sample output from the VBA macro for release 26A — used as a reference baseline during Phase 1 |
| `test.py` | Dan's original Selenium downloader — scrapes Oracle docs and downloads xlsm files |

---

## Testing

- `python -m pytest tests/` — run full suite (54 tests)
- `python -m pytest tests/test_clear.py -v` — run one module
- `tests/validate_against_vba.py` and `tests/vba_fieldrow_map.json` — ad-hoc validation against the legacy VBA macro's expected header rows (not pytest, kept for spot-checks against regressions)

Tests use `openpyxl.Workbook` to build synthetic FBDI files per test (see `_make_fbdi_workbook` / `_create_fbdi_workbook` helpers). There are no fixtures — fixture-like workbooks are built inline in each test so the expected layout is visible next to the assertions.

**Test-data gotcha:** `detect_header_row` scores rows by UPPER_SNAKE_CASE content. Synthetic sample data like `"CREATE"`, `"V1"`, `"DR_ECO_1"` will false-positive as headers. Use lowercase/mixed-case values in test data rows (e.g. `"Create Order"`, `"V1-org"`).

---

## Docs & Planning

Two patterns have been used in this repo — both are still valid:

- **Old:** `handoff_*.md` files (written in Claude Chat, executed by Claude Code) — now gitignored, kept in conversation or local scratch.
- **Current:** `docs/superpowers/specs/*.md` (design) and `docs/superpowers/plans/*.md` (implementation plans) — produced via the `superpowers:brainstorming` and `superpowers:writing-plans` skills and committed to the repo so the history is auditable.

`NEXT_STEPS.md` at the project root is the prioritized backlog / resolution log — read it for historical context on what was fixed when and why.

---

## Plugins / Tooling

Project uses the `superpowers` skill family (brainstorming, writing-plans, executing-plans, systematic-debugging, verification-before-completion). CodeRabbit is wired up for PR review. See user-level `~/.claude/` config for the full plugin list — no project-specific plugin requirements.

## graphify

This project has a graphify knowledge graph at graphify-out/.

Rules:
- Before answering architecture or codebase questions, read graphify-out/GRAPH_REPORT.md for god nodes and community structure
- If graphify-out/wiki/index.md exists, navigate it instead of reading raw files
- After modifying code files in this session, run `python3 -c "from graphify.watch import _rebuild_code; from pathlib import Path; _rebuild_code(Path('.'))"` to keep the graph current
