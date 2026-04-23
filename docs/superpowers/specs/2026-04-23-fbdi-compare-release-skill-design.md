# FBDI Compare-Release Skill — Design

**Date:** 2026-04-23
**Status:** Design approved; ready for implementation plan
**Trigger:** Brad wants a skill his coworkers can invoke when Oracle ships a quarterly FBDI release, to run the full download → clear → compare → catalog pipeline without needing to remember the individual commands.

---

## 1. Goal

Ship a project-level Claude Code skill, `fbdi-compare-release`, that takes a coworker from "Oracle released 26C" to a finished `Comparison_Report_<OLD>_<NEW>.xlsx` plus an updated `FBDI_Master_Catalog.xlsx`, handling:

- Environment bootstrap (Python 3.14+, deps, Chrome)
- Selenium download of the new release (and the prior release, if missing)
- The manual `RapidImplementationForCashManagement.xlsm` drop
- Smart-clear
- Compare
- Catalog refresh
- A short "what changed" summary in the terminal

The skill is the glue between coworkers and the existing `fbdi` package. It does not re-implement comparison logic.

---

## 2. Scope

### In scope
- End-to-end orchestration of the full pipeline — the five functional stages (download → manual-file drop → clear → compare → catalog refresh) plus three skill-added stages (environment check, summary, post-run verification). See §4 for the 8-stage breakdown.
- Environment bootstrap for coworkers whose Python/dep setup is inconsistent.
- Human-in-the-loop prompts at four decision points (see §5).
- A short change summary at the end of a run.

### Out of scope
- Applaud mapping refresh (`fbdi_applaud_mapping.xlsx`) — separate future skill.
- Client-deliverable report generation (`report.py`) — separate future skill.
- Oracle docs URL pattern changes — Dan/Brad update `tools/download_and_clear.py` directly.
- `RapidImplementationForCashManagement.xlsm` auto-download — technically infeasible (Oracle Fusion FSM only).
- Catalog schema changes — separate work.

---

## 3. Architecture

```
.claude/skills/fbdi-compare-release/
├── SKILL.md                       (workflow + decision points — target ~300 lines)
├── scripts/
│   ├── check_env.py               Stage 1 — Python/deps/OS/Chrome preflight
│   ├── summarize_report.py        Stage 7 — summary of Comparison_Report_*.xlsx
│   └── verify_run.py              Stage 8 — diagnose + Issues-tab sanity check
└── references/
    ├── troubleshooting.md         Corrupt xlsm, Selenium failures, Chrome drift, Oracle FSM walk-through
    └── release-version-format.md  Oracle quarterly naming (YY{A-D}), how to find latest
```

**Key properties:**

- **SKILL.md is the orchestrator.** Claude executes an ordered workflow; it does not delegate to a monolithic PowerShell wrapper. This keeps human-in-the-loop prompts and failure triage in natural language.
- **Bundled scripts are narrow and stateless.** They only call the existing `fbdi` CLI or read output xlsx files. They do not import `fbdi/*` internals — keeps the skill decoupled from library refactors.
- **References are read-on-demand.** `troubleshooting.md` and `release-version-format.md` load only when a matching situation arises, keeping SKILL.md context lean.
- **Nothing new lives in `fbdi/`.** The package stays a library; skill-specific glue stays inside the skill folder.
- **Project-level install** (`.claude/skills/...`) — committed to the repo. Coworkers get the skill on `git pull`.

---

## 4. Pipeline

```
STAGE 1 — Environment check
  scripts/check_env.py
  Verify: Windows (warn+continue on Mac), Python ≥3.14, deps, Chrome installed, baselines/ dir.
  Plain-English failure messages. Auto-offer `pip install -r requirements.txt` when deps missing.

STAGE 2 — Resolve versions
  If user passed --old/--new → use them.
  Else detect most recent release folder under baselines/ and confirm with user:
    "I'll compare <prior> → <new>. Correct?"
  If the prior release is missing, prompt to download both.

STAGE 3 — Download
  python tools/download_and_clear.py <ver> --skip-clear
  NOTE: download wipes originals/ first — re-running is destructive, not incremental.
  Expected wall time: ~15–20 min per release.
  After download, verify Rapid Impl file presence. If missing, prompt (see §5 #2).
  Sanity-check file-count delta vs prior release. If >15% or >20 files, warn + offer re-run.

STAGE 4 — Smart-clear
  python tools/download_and_clear.py <ver> --clear-only
  Populates baselines/<ver>/blanks/ from originals/, preserving header rows.
  Rapid Impl file flows through naturally.

STAGE 5 — Compare
  python -m fbdi compare --old <OLD> --new <NEW> --output Comparison_Report_<OLD>_<NEW>.xlsx
  Long-running. Per-pair subprocess isolation already handles most failures (see CLAUDE.md).
  Collect per-pair failures; surface at end rather than aborting.

STAGE 6 — Catalog update
  python -m fbdi catalog --release <NEW>
  Updates FBDI_Master_Catalog.xlsx with new release snapshot + Drift tab.

STAGE 7 — Summary
  scripts/summarize_report.py
  Prints: files processed, files with changes (adds/removes/modifies),
          top 5 most-changed files, paths to the two output xlsx.

STAGE 8 — Post-run verification
  scripts/verify_run.py
  Spot-checks: diagnose shows no new NO_HEADER regressions; catalog Issues tab
               row count is within expected bounds vs prior run.
```

**Resumability:** each stage is idempotent on output-existence terms. If the skill is interrupted at stage 5, rerunning the skill will skip stages 1–4 (downloads and clears already exist) and resume from compare. Stage 3's download wipes the originals folder, so a retry of stage 3 specifically is destructive — the skill surfaces this before retrying.

---

## 5. Human-in-the-loop decision points

Four places the skill stops and asks. Everywhere else it runs unattended.

1. **Prior-release missing (Stage 2).**
   "I need `<prior>` as the comparison baseline but don't see `baselines/<prior>/`. Download it too, or point me at an existing copy?"

2. **Rapid Implementation file missing (Stage 3).**
   "`RapidImplementationForCashManagement.xlsm` isn't auto-downloadable. Options:
   - (a) Copy from `baselines/<prior>/originals/` — fast, safe since Oracle rarely updates it. **[default]**
   - (b) I'll walk you through the Oracle Fusion FSM path (Setup and Maintenance → hamburger menu → Search → 'Create Banks, Branches, and Accounts in Spreadsheet').
   - (c) I already have it, let me drop it in `baselines/<new>/originals/` now."
   Don't proceed until the file is present.

3. **Version-mismatch sanity (Stage 2).**
   If autodetected prior ≠ what the user named, confirm the unusual pick before running (e.g., skipping a release).

4. **Excessive compare failures (end of Stage 5).**
   If more than ~5 per-pair failures, pause: retry / skip / abort? Single-digit failures are expected (corrupt xlsm metadata in some 26B templates per CLAUDE.md).

Dan's happy-path run hits zero of these prompts on a warm machine.

---

## 6. Environment bootstrap & failure handling

### Preflight (Stage 1)

| Check | Fail action |
|---|---|
| OS is Windows | If Mac/Linux: print warning, continue. If Windows: silent. |
| Python ≥ 3.14 | Print pyenv-win install command; exit. |
| `selenium`, `webdriver-manager`, `requests`, `openpyxl`, `pytest` importable | Offer `pip install -r requirements.txt`; run if user confirms. |
| Google Chrome installed | Print install URL; exit. |
| `baselines/` exists (or creatable) | Create it. |

### Failure handling by stage

- **Stage 3 download failures.** Timeout / connection / zero-file → retry once. File-count delta >15% or >20 files vs prior → warn + offer re-run.
- **Stage 3 manual file missing.** The §5 #2 prompt.
- **Stage 5 per-pair compare failures.** Already isolated by subprocess (CLAUDE.md §subprocess_util). Collect failures; surface at end. If >5 failures, §5 #4 prompt.
- **Stage 6 catalog regressions.** Issues-tab row count >2× prior OR >50 new rows → flag in summary. Does not block.
- **Any stage: Python traceback.** Print the exception, likely cause in plain English, and the remediation. Never hand coworkers a raw stack trace alone.

### What the skill will NOT try to fix
- Oracle docs site restructure (new URL patterns in `MODULE_URL_TEMPLATES`).
- `RapidImplementationForCashManagement.xlsm` auto-download.
- Catalog schema changes.

---

## 7. Invocation

**Slash command:** `/fbdi-compare-release` (with optional `--old 26A --new 26B`).

**Natural-language triggers** (description-driven). SKILL.md `description` field must be "pushy" enough to fire on paraphrases, per the skill-creator guide:
- "Oracle released 26C, run the quarterly FBDI update"
- "Compare 26A to 26B"
- "Update the FBDI Master Catalog for 26B"
- "New FBDI release dropped"
- "FBDI refresh for Q1"

**Should NOT fire on** (near-misses used as negative eval cases):
- "What's the current Python version?" (mentions Python, not FBDI)
- "Run the test suite" (unrelated)
- "Open the catalog xlsx" (reading, not refreshing)

---

## 8. Testing

### Layer 1 — Unit tests for bundled scripts

`tests/test_skill_scripts.py`:

- `check_env.py` — mocks Python version, missing deps, missing Chrome → each failure path exits correctly with readable message.
- `summarize_report.py` — feeds synthetic `Comparison_Report_*.xlsx` (inline `openpyxl.Workbook` per the repo's existing test pattern) → assert counts match.
- `verify_run.py` — synthetic catalog Issues tab → regression detection triggers at the configured threshold.

### Layer 2 — Skill eval cases (via skill-creator)

Realistic prompts, driven through the eval-viewer loop:

| # | Prompt | Expected |
|---|---|---|
| 1 | "Oracle released 26C, run the quarterly FBDI update" | Triggers skill; detects missing `baselines/26C/`; runs full pipeline |
| 2 | "Compare 26A to 26B" | Triggers skill with explicit versions; skips auto-detection |
| 3 | "Update the FBDI Master Catalog for 26B" | Triggers skill — validates catalog-centric phrasing |
| 4 | "What's the current production Python version?" | **Does NOT trigger** |
| 5 | "Run the test suite" | **Does NOT trigger** |

The 26A→26B end-to-end run done during design becomes the ground-truth reference for eval #2.

### Layer 3 — Description optimization

Per the skill-creator guide, after eval #1–5 pass, run the description-optimization loop (`scripts/run_loop.py`) against ~20 should-trigger + should-not-trigger queries to tune the `description` frontmatter for reliable triggering without false positives.

---

## 9. Dependencies

- Existing `fbdi/` package — unchanged.
- `tools/download_and_clear.py` — unchanged.
- New top-level `requirements.txt` — created during design (lists openpyxl, selenium, webdriver-manager, requests, pytest).
- `.python-version` — created during design (pins repo to 3.14.3 via pyenv).

No new runtime dependencies. No changes required to `fbdi/` or `tools/`.

---

## 10. Open questions

None at spec-approval time. All design questions resolved during brainstorming (see conversation log 2026-04-23).
