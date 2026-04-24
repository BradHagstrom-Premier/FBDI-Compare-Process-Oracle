---
name: fbdi-compare-release
description: "Use when Oracle ships a quarterly FBDI release and the user wants the full download → clear → compare → catalog pipeline run end-to-end. Triggers on phrases like 'Oracle released 26C', 'compare 26A to 26B', 'run the quarterly FBDI update', 'update the FBDI Master Catalog for 26B', 'new FBDI release dropped', 'FBDI refresh for Q1'. Does NOT trigger on near-miss phrases like 'compare these two spreadsheets' or 'run the test suite'."
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

> **HITL numbering note:** `HITL #1`–`#6` below are stable IDs from the
> design spec, not sequential execution order — e.g., #3 appears before
> #1 in the flow because versions are resolved before baseline presence
> is checked.

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
  or the next quarter after `OLD` if they didn't say. Convention: after
  26D comes 27A; otherwise bump the letter (26A → 26B, etc.).

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

**Step 3c — FSM-file check:** After Step 3b returns exit 0 for `<NEW>`,
confirm `RapidImplementationForCashManagement.xlsm` exists in
`baselines/<NEW>/originals/`. It is not auto-downloadable (see
`references/troubleshooting.md`), so verify explicitly. If missing, go
to HITL #2.

**HITL #2 — `RapidImplementationForCashManagement.xlsm` missing:** When
Step 3c flags the file as absent, ask:

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

If Stage 4 captured no timeouts, omit the `--timeouts` flag entirely
rather than passing an empty string.

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
