# CLAUDE.md — Oracle FBDI Pulldown

This file gives Claude Code persistent context for this project. Read it at the start of every session.

---

## Project Purpose

This repo automates comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) between Oracle Cloud release versions. The goal is to identify field-level changes (added, removed, modified columns) across releases so that Brad and Dan can keep Definian's Oracle integrations current. The primary deliverable is a structured Excel comparison report.

---

## Active Pipeline (Built and Working)

- **`fbdi/` package** — Python comparison engine
  - `detect_header.py` — dynamically identifies the header row in each FBDI tab using content scoring (no hardcoded filename map)
  - `compare.py` — diffs two releases tab-by-tab, field-by-field; uses subprocess isolation + per-file timeout
  - `clear.py` — smart clearing of FBDI templates using `detect_header_row` (preserves headers at any row)
  - `cli.py` / `__main__.py` — CLI entry point: `python -m fbdi compare --old 26A --new 26B` (release label or explicit path)
  - `config.py`, `utils.py` — shared configuration and helpers
- **`tools/download_and_clear.py`** — standalone Selenium downloader + smart clearing entry point. Not integrated into `fbdi` (keeps Selenium dependencies out of the comparison engine).
- **`tests/`** — 54 unit tests, all passing (`python -m pytest tests/`)
- **Output** — `Comparison_Report_<OLD>_<NEW>.xlsx` — 7-column format (columns A–G): Module, Template File, Tab, Field, Old Value, New Value, Change Type
- **Baseline layout** — `baselines/<ver>/originals/` (as-downloaded) and `baselines/<ver>/blanks/` (smart-cleared copies for client use)

---

## What Is Not Built Yet

- **Downloader integration** — `reference/test.py` (Dan's Selenium script) has not been ported into the `fbdi/` package. Baselines are still populated manually or by running the legacy script directly.
- **`report.py`** — Audrey report automation. Phase 2: reformats comparison output into the Audrey change-tracking format used for client deliverables.
- **`python -m fbdi run`** — Full pipeline command (Phase 3). Would chain download → compare → report in one call.

---

## Key Architectural Decisions (Closed — Do Not Re-litigate)

| Decision | Choice |
|---|---|
| Baseline storage | Folder-based: `baselines/25d/`, `baselines/26a/` — not Git-based |
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

## Workflow

Planning happens in Claude Chat (claude.ai). Implementation happens in Claude Code (this CLI). The bridge is a handoff `.md` file written in Claude Chat and executed here.

Pattern:
1. Brad and Claude Chat design the next phase → produce a `handoff_*.md`
2. Brad opens Claude Code → runs the handoff plan
3. Claude Code executes, commits, pushes

---

## Plugins Active in This Project

- **superpowers** — brainstorming, plan writing, plan execution, debugging, finishing branches, code review
- **context7** (MCP) — library documentation lookup
- **github** (MCP) — PR creation, branch management
- **coderabbit** — automated code review
- **commit-commands** — `/commit`, `/commit-push-pr`
- **claude-md-management** — CLAUDE.md auditing and improvement
- **feature-dev** — feature development with architecture focus
- **pr-review-toolkit** — PR review workflows

## graphify

This project has a graphify knowledge graph at graphify-out/.

Rules:
- Before answering architecture or codebase questions, read graphify-out/GRAPH_REPORT.md for god nodes and community structure
- If graphify-out/wiki/index.md exists, navigate it instead of reading raw files
- After modifying code files in this session, run `python3 -c "from graphify.watch import _rebuild_code; from pathlib import Path; _rebuild_code(Path('.'))"` to keep the graph current
