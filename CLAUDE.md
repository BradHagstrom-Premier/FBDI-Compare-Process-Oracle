# CLAUDE.md — Oracle FBDI Pulldown

This file gives Claude Code persistent context for this project. Read it at the start of every session.

---

## Project Purpose

This repo automates comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) between Oracle Cloud release versions. The goal is to identify field-level changes (added, removed, modified columns) across releases so that Brad and Dan can keep Definian's Oracle integrations current. The primary deliverable is a structured Excel comparison report.

---

## Active Pipeline (Built and Working)

- **`fbdi/` package** — Python comparison engine
  - `detect_header.py` — dynamically identifies the header row in each FBDI tab using content scoring (no hardcoded filename map)
  - `compare.py` — diffs two releases tab-by-tab, field-by-field
  - `cli.py` / `__main__.py` — CLI entry point: `python -m fbdi compare --old 25d --new 26a`
  - `config.py`, `utils.py` — shared configuration and helpers
- **`tests/`** — 33 unit tests, all passing (`python -m pytest tests/`)
- **Output** — `Comparison_Report_<OLD>_<NEW>.xlsx` — 7-column format (columns A–G): Module, Template File, Tab, Field, Old Value, New Value, Change Type

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

## Known Hazards (From Phase 1 Implementation)

- **6 files >5MB are currently skipped** — performance issue, not a bug. Large files time out during openpyxl load. Address in a future optimization pass.
- **~8 tabs with non-standard headers fail detection** — known edge case. Engine logs them and moves on. Do not assume detection is 100%.
- **Full comparison run is ~75 minutes** for 209 file pairs (25D vs 26A). This is expected given openpyxl's read speed on xlsm files.
- **Phantom columns (`max_column=16384`)** — some xlsm files report 16384 columns due to corrupt metadata. The engine caps column scanning at 500.
- **Corrupt XML in some xlsm files** — handled gracefully; engine catches `zipfile.BadZipFile` and logs the file as unreadable.
- **`Comparison_Report_25D_26A.xlsx` (VBA output)** — has a corrupt stylesheet. Cannot be loaded with standard `openpyxl.load_workbook`. Use `read_only=True` or `data_only=True` with exception handling if you need to read it.

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
