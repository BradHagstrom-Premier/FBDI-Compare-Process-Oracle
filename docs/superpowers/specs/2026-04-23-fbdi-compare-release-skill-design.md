# FBDI Compare-Release Skill — Design

**Date:** 2026-04-23
**Status:** Design approved
**Trigger:** Brad wants a skill his coworkers can invoke when Oracle ships a quarterly FBDI release, to run the full download → clear → compare → catalog pipeline without needing to remember the individual commands.

---

## 1. Goal

Ship a project-level Claude Code skill, `fbdi-compare-release`, that takes a coworker from "Oracle released 26C" to a finished `Comparison_Report_<OLD>_<NEW>.xlsx` plus an updated `FBDI_Master_Catalog.xlsx`, handling:

- Environment bootstrap (Python 3.14+, deps, Chrome)
- Selenium download of the new release (and the prior release, if missing)
- The manual `RapidImplementationForCashManagement.xlsm` drop or copy
- Smart-clear
- Compare
- Catalog refresh
- A short "what changed" summary in the terminal

The skill is the glue between coworkers and the existing `fbdi` package. It does not re-implement comparison logic.

---

## 2. Scope

### In scope
- End-to-end orchestration of the full pipeline — the five functional stages (download → manual-file drop or copy → clear → compare → catalog refresh) plus three skill-added stages (environment check, summary, post-run verification). See §4 for the 8-stage breakdown.
- Environment bootstrap for coworkers whose Python/dep setup is inconsistent.
- Human-in-the-loop prompts at six decision points (see §5).
- A short change summary at the end of a run.

### Out of scope
- Applaud mapping refresh (`fbdi_applaud_mapping.xlsx`) — separate future skill.
- Client-deliverable report generation (`report.py`) — separate future skill.
- Oracle docs URL pattern changes — Brad will need work to update `tools/download_and_clear.py` directly and we cannot predict Oracle future changes.
- `RapidImplementationForCashManagement.xlsm` auto-download — technically infeasible (Oracle Fusion FSM only).
- Catalog schema changes — separate work.

---

## 3. Architecture

```
.claude/skills/fbdi-compare-release/
├── SKILL.md                       (workflow + decision points — target ~300 lines)
├── scripts/
│   ├── check_env.py               Stage 1 — Python/deps/OS/Chrome preflight
│   ├── verify_download.py         Stage 3 — diff downloads vs baseline_files.txt; return missing/extras
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
- It is likely that the human running this skill will step away from the computer at some point and it will lock and/or sleep. This should not cause the process to fail.

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

STAGE 3 — Download + verify
  python tools/download_and_clear.py <ver> --skip-clear
  NOTE: download wipes originals/ first — re-running is destructive, not incremental.
  Expected wall time: ~15–20 min per release.

  After download, run scripts/verify_download.py <ver>:
    - Parses the <ver> section of baseline_files.txt into an expected set.
    - Lists actual filenames in baselines/<ver>/originals/.
    - Diffs with `LC_ALL=C sort | comm` (locale matters — default macOS sort
      misorders mixed-case filenames and silently hides real gaps).
    - Returns {missing: [...], extras: [...]}.
      Known manual files (e.g. RapidImplementationForCashManagement.xlsm) are
      excluded from "missing" — handled by §5 #2 instead.

  Verification outcomes:
    - missing == 0 and extras == 0 → proceed.
    - missing > 0 (first attempt) → retry the download once. Silent scraper
      failures on a module page (observed in 2026-04-23 Windows testing) are
      transient; a single retry resolved them. See docs/scraper-gap-findings-
      2026-04-23.md for the evidence.
    - missing > 0 (after retry) → surface to user grouped by Oracle module URL
      (per MODULE_URL_TEMPLATES). This indicates a genuine scraper/docs-site
      breakage that the skill cannot self-heal; see §5 #5 (capped at 3 total
      attempts before the user-initiated-retry option is withdrawn).
    - extras > 0 → these are files the scraper legitimately downloaded (i.e.
      Oracle served them) that aren't in baseline_files.txt's inventory for
      <ver>. Almost always means the inventory is stale, not that the file is
      suspect — Oracle's served set is ground truth. See §5 #6.

  First-run fallback (no baseline_files.txt entry for <ver>): the skill runs
  the same inventory-update flow as §5 #6, since "all downloaded files are
  extras relative to a nonexistent inventory" is the same problem. The §5 #6
  "First-run sanity check" guards against silent scraper failures on this path.

STAGE 4 — Smart-clear
  python tools/download_and_clear.py <ver> --clear-only
  Populates baselines/<ver>/blanks/ from originals/, preserving header rows.
  Rapid Impl file flows through naturally.

STAGE 5 — Compare
  python -m fbdi compare --old <OLD> --new <NEW> --output Comparison_Report_<OLD>_<NEW>.xlsx
  Expected wall time: ~3–5 min for ~210 file pairs (post-`iter_rows` optimization).
  Per-pair subprocess isolation already handles corrupt-metadata files silently
  (the historical "single-digit failures" hazard is rare post-fix).
  Collect per-pair failures; surface at end rather than aborting.

STAGE 6 — Catalog update
  python -m fbdi catalog --release <NEW>
  Updates FBDI_Master_Catalog.xlsx with new release snapshot + Drift tab.
  Expected wall time: ~3–5 min for ~210 files.

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

Six places the skill stops and asks. Everywhere else it runs unattended.

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

5. **Download still short after retry (Stage 3).**
   "After one retry, `baselines/<ver>/originals/` is still missing N files vs `baseline_files.txt`. Affected Oracle modules: [grouped list by URL]. This is beyond what I can fix automatically — likely the Oracle docs site changed structure or a module page is down. Options:
   - (a) Retry again (transient issues sometimes need more than one retry).
   - (b) Abort so `tools/download_and_clear.py` can be debugged directly.
   - (c) Proceed with what's present and note the gap in the summary (not recommended — compare output will be incomplete)."
   Default: (b).
   **Retry cap:** 3 total download attempts across the run (1 initial + 1 auto-retry + 1 user-initiated via (a)). After the 3rd attempt still fails verification, (a) is no longer offered — only (b) or (c) remain. Reason: 2026-04-23 evidence shows naive retries resolve the transient module-silent-failure class of bug, but once three independent attempts all come up short, further retries are unlikely to help and risk wasting another ~15–20 minutes per attempt.

6. **Extras present (Stage 3).**
   Oracle served files that aren't in `baseline_files.txt`'s inventory for `<ver>`. This is almost always stale inventory, not a bad download — Oracle's served set is ground truth.
   "Downloaded N files not in `baseline_files.txt`'s inventory for `<ver>`: [list]. Options:
   - (a) Update `baseline_files.txt` to add these to the `<ver>` section. **[default]**
   - (b) Quarantine to `baselines/<ver>/_extras/` and keep them out of compare — use only if you suspect the scraper pulled something unexpected (rare).
   - (c) Show me the file(s) first so I can decide per-file."
   On (a), the skill writes the inventory update in-place, re-runs the verification, and proceeds.

   **First-run sanity check.** When `<ver>` has no existing inventory section in `baseline_files.txt` (bootstrap case), the skill cannot diff against a prior `<ver>` inventory. Before committing the new section, it compares the downloaded file count to the most recent prior release's inventory count. If |new − prior| / prior > 15%, the skill surfaces the delta with the per-module breakdown and asks the user to confirm:
   "Bootstrapping `<ver>` inventory from <N> downloaded files. The most recent prior release (`<prior>`) has <P> files — that's a <Δ%> change. Oracle rarely adds or removes more than ~5–10 templates in a quarterly release, so this is worth a second look. Proceed with bootstrap, retry the download, or abort?"
   This is a cheap belt-and-suspenders guard against a silent scraper failure on the first run of a new release (where there's no prior `<ver>` inventory to catch it via §5 #5). Below 15% drift, the bootstrap proceeds without prompting.

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
| `baseline_files.txt` present | If missing: print warning; Stage 3 verification degrades to delta-vs-prior only. |
| `baseline_files.txt` contains an inventory for `<new>` (and `<old>` if being downloaded) | If missing: flag as first-run; Stage 3 invokes the §5 #6 "update inventory" flow post-download to bootstrap the section. |

### Failure handling by stage

- **Stage 3 download failures.** Timeout / connection / zero-file → retry once.
- **Stage 3 inventory verification.** Missing files after one retry → §5 #5 prompt. Extras → §5 #6 prompt (default: update inventory).
- **Stage 3 manual file missing.** The §5 #2 prompt.
- **Stage 3 first-run (no baseline inventory for `<ver>`).** Same flow as extras handling (§5 #6) — skill proposes a new inventory section for `<ver>` built from the downloaded filenames, runs the §5 #6 "First-run sanity check" (15% delta-vs-prior guard), and asks the user to confirm before committing to `baseline_files.txt`.
- **Stage 4 per-file clear timeouts.** `tools/download_and_clear.py` uses a 120s per-file subprocess timeout and already prints a `TIMED OUT — clear these manually` summary for oversized files. Known recurring offender: `PayablesCollectionDocuments.xlsm` (~9MB, observed timing out in both 26A and 26B on 2026-04-23). **Not a blocker for Stage 5** — compare reads originals, not blanks. The skill captures the timeout list in the final summary (Stage 7) so the user knows which blanks files need manual clearing (e.g. via Excel or `reference/Clear_FBDIs - 20210412.xlsm`) if they need complete blanks coverage for downstream client use. No HITL prompt needed.
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
- `verify_download.py` — parses a synthetic `baseline_files.txt` with multiple release sections, lists a synthetic `originals/` dir, returns correct `{missing, extras}`; exercises the locale-aware sort (regression guard for non-`LC_ALL=C` environments); exercises the known-manual-file exclusion; exercises the first-run case (no inventory for the release).
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
- `baseline_files.txt` — committed, hand-maintained authoritative inventory per release; consumed by `verify_download.py`. Brad updates it when a new release is accepted.
- New top-level `requirements.txt` — created during design (lists openpyxl, selenium, webdriver-manager, requests, pytest).
- `.python-version` — created during design (pins repo to 3.14.3 via pyenv).

No new runtime dependencies. No changes required to `fbdi/` or `tools/`.

---

## 10. Resolved questions

All design questions resolved during brainstorming (see conversation log 2026-04-23), post-design Windows verification (see `docs/scraper-gap-findings-2026-04-23.md`, "Update — 2026-04-23" section), and the 2026-04-23 design-approval pass.

- ~~First-run-of-a-new-release behavior.~~ Unified with extras handling (§5 #6): on first run of an unknown release, the skill proposes a bootstrap inventory from the download and asks the user to sign off before committing to `baseline_files.txt`.
- ~~Retry cadence for §5 #5.~~ Capped at 3 total download attempts (1 initial + 1 auto-retry + 1 user-initiated via (a)). See §5 #5 "Retry cap". Rationale: 2026-04-23 evidence shows naive retries resolve the transient module-silent-failure class of bug; once three independent attempts fall short, further retries are unlikely to help and each costs ~15–20 min.
- ~~First-run delta-vs-prior sanity check.~~ Added to §5 #6 "First-run sanity check": when bootstrapping inventory for a brand-new release, the skill warns the user if the downloaded file count deviates >15% from the most recent prior release. This catches silent scraper failures on the very first run of a new release (where §5 #5's per-release inventory diff can't help).
