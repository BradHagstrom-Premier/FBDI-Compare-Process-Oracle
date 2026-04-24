# Handoff Docs and Repo Cleanup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn this repo into a clean hand-off package — a cruft-free root, a tracked `docs/archive/` for narrative history, two self-serviceable guides (operator + developer), a refreshed README, and a CLAUDE.md that matches the final state.

**Architecture:** Sequential. First: tidy the filesystem (move/delete/relocate) so the guides describe the final state. Second: write the two new docs and archive the narrative history. Third: refresh README with pointers and humanize. Fourth: run the CLAUDE.md improver. Four commits pushed directly to master.

**Tech Stack:** git, Python 3.14, pytest, openpyxl, humanizer-skill, claude-md-management skill.

---

## File Structure

**Create (tracked):**
- `docs/operator-guide.md` — 7-section operator walkthrough (~2.5K words)
- `docs/developer-guide.md` — 9-section developer orientation (~2.5K words)
- `docs/archive/applaud-mapping-audit-notes.md` — via `git mv` from `docs/Applaud Mapping Audit.md`
- `docs/archive/scraper-gap-findings-2026-04-23.md` — via `git mv` from `docs/scraper-gap-findings-2026-04-23.md`
- `docs/archive/claude-fbdi-applaud-mapping-audit.md` — via `git mv` from root `Claude_fbdi_applaud_mapping_audit.md`

**Relocate (to gitignored paths):**
- `complete_mapping.py` → `Archive/complete_mapping.py` (untrack via `git rm --cached`)
- `format_workbooks.py` → `Archive/format_workbooks.py` (untrack via `git rm --cached`)
- `applaud_snapshot.json` → `baselines/applaud/applaud_snapshot.json` (untrack via `git rm --cached`)

**Delete (filesystem only — already untracked/gitignored):**
- `~$FBDI_Master_Catalog.xlsx` (Excel lock file)
- `_extract_cache/` (empty dir)
- `Comparison_Report.xlsx` (March 24 stale root-level report)
- `Diagnostic_Report_26B.xlsx` (stale root-level report)

**Modify (tracked):**
- `fbdi/audit.py:26` — `SNAPSHOT_PATH` constant
- `fbdi/audit.py:867` — audit report display string (for consistency)
- `README.md` — add guide pointers near top, refresh `Repo structure` tree, light humanizer pass
- `CLAUDE.md` — updated via `claude-md-management:claude-md-improver` skill

**Unchanged (explicitly out of scope):**
- `fbdi/` package logic (except the two lines in `audit.py`)
- `.claude/skills/fbdi-compare-release/` (SKILL.md, scripts/, references/)
- `docs/superpowers/specs/` and `plans/` (historical design artifacts)
- `reference/` (pre-Python VBA archive)
- `tests/` (no new tests, no test refactors — test_audit.py uses `tmp_path`, unaffected by the `audit.py` constant change)

---

## Task 1: Repo cleanup — Buckets 1, 2, 4 + audit.py path update

**Files:**
- Delete: `~$FBDI_Master_Catalog.xlsx`, `_extract_cache/`, `Comparison_Report.xlsx`, `Diagnostic_Report_26B.xlsx` (filesystem only)
- Relocate (untrack): `complete_mapping.py`, `format_workbooks.py`, `applaud_snapshot.json`
- Modify: `fbdi/audit.py:26` and `fbdi/audit.py:867`
- Verify: `tests/` (run full suite)

Bucket 3 (narrative docs → `docs/archive/`) is handled in Task 4 so it commits with the guides.

- [ ] **Step 1: Delete stale filesystem artifacts (Bucket 2)**

These are all untracked (gitignored or transient). `rm` removes them from the working tree without touching the git index.

```bash
rm -f "~\$FBDI_Master_Catalog.xlsx"
rm -rf _extract_cache
rm -f Comparison_Report.xlsx
rm -f Diagnostic_Report_26B.xlsx
```

Verify none remain:

```bash
ls -la | grep -E '^(~\$|_extract_cache|Comparison_Report\.xlsx$|Diagnostic_Report_26B\.xlsx$)'
```

Expected: no matches.

- [ ] **Step 2: Move orphan scripts to gitignored `Archive/` (Bucket 1)**

`Archive/` already exists and is gitignored (see `.gitignore` line 24). Both scripts are tracked — move on disk, then untrack.

```bash
mv complete_mapping.py Archive/
mv format_workbooks.py Archive/
git rm --cached complete_mapping.py format_workbooks.py
```

Verify:

```bash
ls Archive/ | grep -E '^(complete_mapping|format_workbooks)\.py$'
git status --short
```

Expected: both files listed in `Archive/`; `git status` shows `D  complete_mapping.py` and `D  format_workbooks.py` (deletions staged from the index).

- [ ] **Step 3: Relocate `applaud_snapshot.json` (Bucket 4)**

`baselines/` is gitignored (see `.gitignore` line 11). Move on disk, then untrack. This makes the file disappear from git history going forward while still being available at its logical location for `fbdi/audit.py`.

```bash
mkdir -p baselines/applaud
mv applaud_snapshot.json baselines/applaud/applaud_snapshot.json
git rm --cached applaud_snapshot.json
```

Verify:

```bash
ls -la baselines/applaud/applaud_snapshot.json
git status --short | grep applaud_snapshot
```

Expected: file is ~3.3 MB at new path; git status shows `D  applaud_snapshot.json`.

- [ ] **Step 4: Update `fbdi/audit.py` to read the snapshot from the new path**

Open `fbdi/audit.py` and update line 26. The repo root is `Path(__file__).parent.parent`, so the new path is `REPO_ROOT / "baselines" / "applaud" / "applaud_snapshot.json"`.

Replace this line:

```python
SNAPSHOT_PATH = REPO_ROOT / "applaud_snapshot.json"
```

with:

```python
SNAPSHOT_PATH = REPO_ROOT / "baselines" / "applaud" / "applaud_snapshot.json"
```

Also update the display string at line 867 so the generated audit report shows the new relative path. Replace:

```python
        f"**Snapshot:** applaud_snapshot.json @ {snapshot_meta.get('extracted_at', 'unknown')}",
```

with:

```python
        f"**Snapshot:** baselines/applaud/applaud_snapshot.json @ {snapshot_meta.get('extracted_at', 'unknown')}",
```

The docstring on line 4 (`Consumes applaud_snapshot.json + ...`) is fine as-is — bare filename in a module docstring is a common shorthand and doesn't need surgery.

Verify the edits landed:

```bash
grep -n 'applaud_snapshot.json' fbdi/audit.py
```

Expected: three lines — the module docstring (line 4), the new `SNAPSHOT_PATH` (line 26), and the new display string (line 867). Lines 26 and 867 now reference `baselines/applaud/`.

- [ ] **Step 5: Run the full test suite to verify the path change**

```bash
python -m pytest tests/
```

Expected: 241 passed (no failures, no errors). `tests/test_audit.py` uses `tmp_path / "applaud_snapshot.json"` at lines 59 and 805 — these are test-local paths unaffected by the production constant, so the change is benign.

If anything fails: stop, diagnose, fix — do not commit a broken state. The most likely failure mode is a stray hardcoded path reference somewhere else in `fbdi/` that wasn't caught by the grep earlier.

- [ ] **Step 6: Stage and commit (Commit 1 of 4)**

```bash
git add fbdi/audit.py
git status --short
```

Expected staged changes: `D complete_mapping.py`, `D format_workbooks.py`, `D applaud_snapshot.json`, `M fbdi/audit.py`.

```bash
git commit -m "$(cat <<'EOF'
chore(cleanup): archive orphan scripts, remove stale artifacts, relocate applaud_snapshot

- Move complete_mapping.py, format_workbooks.py to gitignored Archive/
- Relocate applaud_snapshot.json to baselines/applaud/ (gitignored logical home); update fbdi/audit.py SNAPSHOT_PATH accordingly
- Remove stale root-level artifacts: ~\$FBDI_Master_Catalog.xlsx, _extract_cache/, Comparison_Report.xlsx, Diagnostic_Report_26B.xlsx

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

Verify:

```bash
git log -1 --stat
```

Expected: commit title matches; four files deleted (three scripts/json + one audit.py modification). Do not push yet — push is Task 13.

---

## Task 2: Draft `docs/operator-guide.md`

**Files:**
- Create: `docs/operator-guide.md`
- Reference: `.claude/skills/fbdi-compare-release/SKILL.md` (for stage + HITL numbering)
- Reference: `fbdi/compare.py`, `fbdi/catalog.py`, `fbdi/clear.py` (for what stages do under the hood)
- Reference: `CLAUDE.md` (for the RapidImplementation hazard and known-issues list)

Target ~2.5K words total, second person, concrete. Tone: a clear-headed coworker explaining the quarterly refresh to the new hire sitting next to them. Normal voice — full humanizer pass comes later (Task 5).

Before writing a section, open `SKILL.md` and read the corresponding stage so the description matches what the skill actually does. The guide is for operators, not developers — no Python internals, no module names, no import paths. Translate to operator concepts: "the comparison step" not "`fbdi/compare.py`".

- [ ] **Step 1: Write the file skeleton and section 1 ("What this is")**

Create `docs/operator-guide.md` with this frontmatter and section headers. Section 1 is a single paragraph (~100 words) identifying the deliverable (quarterly field-level diff + per-release catalog), the consumers (Definian team members maintaining Oracle integrations), and when this runs (every Oracle quarterly release — 26A, 26B, 26C, 26D).

```markdown
# Operator Guide — Quarterly FBDI Refresh

> If you're looking to develop on this codebase rather than run the pipeline, start with [`developer-guide.md`](developer-guide.md).

## What this is and who it's for

<~100 words: Definian team member runs this once per Oracle release. Produces a field-level diff (`Comparison_Report_<OLD>_<NEW>.xlsx`) and a per-release snapshot (`FBDI_Master_Catalog.xlsx`). These feed the Applaud mapping and downstream compliance reporting. Oracle ships releases quarterly; this process is how we keep integrations current.>

## Before your first run
## The 8 stages, in order
## The 6 HITL checkpoints
## Reading the outputs
## When something goes sideways
## Next steps
```

Write section 1 in place of the `<…>` placeholder.

- [ ] **Step 2: Write section 2 — "Before your first run" (~200 words)**

Cover, in order:

1. Environment check: Python 3.14+, Google Chrome installed (Selenium dependency), Windows (supported platform). Point to `pip install -r requirements.txt`.
2. Baselines folder: `baselines/<release>/originals/` is where downloads land; gitignored; must exist before Stage 3. The skill creates it; for a manual CLI run you do it yourself.
3. Windows sleep warning: the download stage takes 15–20 min. If the laptop sleeps mid-run, Selenium loses its Chrome handle. Either disable sleep for the duration or run overnight with sleep disabled.
4. Time budget: ~35–50 minutes end-to-end (downloads dominate; comparison and catalog each take seconds after the iter_rows fix).
5. Two invocation paths: (A) Claude Code with the `/fbdi-compare-release` skill — recommended for first-timers because HITL checkpoints catch drift; (B) raw CLI with `python tools/download_and_clear.py` + `python -m fbdi compare` + `python -m fbdi catalog` — for experienced operators who know exactly what they want.

- [ ] **Step 3: Write section 3 — "The 8 stages, in order" (~1.5K words)**

Open `.claude/skills/fbdi-compare-release/SKILL.md` and copy the stage names/numbers from there (don't invent them). For each stage, write one subsection:

```markdown
### Stage N — <Stage Name>

**What it does (plain English):** <1–2 sentences>
**What you see on screen:** <1–2 sentences — e.g., stdout patterns, progress counters>
**Expected wall time:** <range>
**If it stalls:** <1 sentence — what to do>
```

Fill ALL 8 stages. The eight stages per SKILL.md are:
1. Environment preflight
2. Version resolve (old vs. new)
3. Download (Selenium)
4. Smart clear (preserves headers)
5. Compare (field-level diff)
6. Catalog (per-release snapshot)
7. Summary
8. Post-run verification (`verify_run.py`)

Re-read `SKILL.md` if unsure what any stage does. Do not paraphrase prompt text from the skill's HITL — that's section 4's job.

- [ ] **Step 4: Write section 4 — "The 6 HITL checkpoints" (~400 words)**

From `SKILL.md`, the six HITL checkpoints are numbered `HITL #1` through `HITL #6`. Use the same IDs. For each:

```markdown
### HITL #N — <checkpoint name>

**Trigger:** <when the pipeline hits this prompt>
**Options presented:** <bullet list of the choices>
**How to decide:** <1–2 sentences of operator-perspective guidance>
```

Cover all six. Read SKILL.md to get the exact trigger conditions — do not guess. The `RapidImplementationForCashManagement.xlsm` manual-fetch prompt is HITL #2 (from `README.md` line 54).

- [ ] **Step 5: Write section 5 — "Reading the outputs" (~200 words)**

Cover:

1. **`Comparison_Report_<OLD>_<NEW>.xlsx`** — seven columns. Name each: File, Tab, Position, Label, Technical, Change Type (Added/Removed/Modified), Details. Explain what a reader does with this (hands to Applaud mapping owner; triages changes against existing integrations).
2. **`FBDI_Master_Catalog.xlsx`** — three sheet types. (a) Per-release tabs (e.g., `26A`, `26B`) — one row per file × tab × column, with position, label, technical, type, length, scale, required. (b) `Issues` — malformed type strings and parse warnings (currently 9 rows — all genuinely broken Oracle strings per CLAUDE.md). (c) `Drift` — flags tabs that moved files or files that moved tabs across releases.

- [ ] **Step 6: Write section 6 — "When something goes sideways" (~150 words)**

Three concrete failure modes and how to respond:

1. **The FSM file is missing** — Stage 4 warns if `RapidImplementationForCashManagement.xlsm` isn't in `baselines/<ver>/originals/`. Manual fetch path: Oracle Fusion → Setup and Maintenance → hamburger menu → Search → "Create Banks, Branches, and Accounts in Spreadsheet". Download, drop into the folder, re-run the compare stage.
2. **Ctrl-C mid-run** — the skill is idempotent from Stage 3 onward. Re-invoking re-downloads any missing files but skips ones already on disk. Comparison and catalog rebuild cleanly from the xlsm files.
3. **For anything else** — point at `SKILL.md` "Error handling" section and at `CLAUDE.md` "Known hazards". If `FILE_ERROR` count jumps between releases, Stage 8 (`verify_run.py`) will surface it.

- [ ] **Step 7: Write section 7 — "Next steps" (~50 words)**

Short pointer: if you want to understand what happens under the hood, extend a stage, or fix a bug you hit, switch to [`developer-guide.md`](developer-guide.md).

- [ ] **Step 8: Sanity check the draft**

```bash
wc -w docs/operator-guide.md
```

Expected: 2,200–2,800 words. If under 2,000, the "8 stages" section is probably thin — go back to Step 3 and flesh out each stage. If over 3,000, trim section 3 (it's the biggest).

---

## Task 3: Draft `docs/developer-guide.md`

**Files:**
- Create: `docs/developer-guide.md`
- Reference: `CLAUDE.md` (for architecture decisions, hazards), `fbdi/` (for module responsibilities), `tests/` (for conventions), `.claude/skills/fbdi-compare-release/` (for skill shape)

Target ~2.5K words total, second person, specific. Tone: a dev who's spent three weeks with this codebase writing up what they'd want to know on day one. Normal voice — humanizer pass comes later (Task 6).

Before writing each section, open the referenced source files and read them. The guide must match what the code does, not what the reader expects.

- [ ] **Step 1: Write the file skeleton and section 1 — "Orientation" (~150 words)**

Create `docs/developer-guide.md`:

```markdown
# Developer Guide

> If you just want to run the quarterly refresh, start with [`operator-guide.md`](operator-guide.md).

## Orientation
## Local setup for development
## Codebase tour
## The `/fbdi-compare-release` skill
## Testing conventions
## How to add a new release handler
## Design docs and how we work
## Known hazards and gotchas
## Where to ask for help
```

Section 1 content (~150 words): State the problem (Oracle releases FBDI templates quarterly; Definian needs to know what changed field-by-field). Name what's shipped (comparison engine, catalog, Selenium downloader, smart-clear, Applaud mapping audit, 241 tests, orchestrator skill). Name what's on the frontier (FBDI → Applaud mapping manual review, future `report.py` for compliance reports, future `python -m fbdi run` one-shot command). Point reader at `CLAUDE.md` for the authoritative state.

- [ ] **Step 2: Write section 2 — "Local setup for development" (~200 words)**

Concrete, executable:

1. Clone + `cd`.
2. Python 3.14 via `pyenv-win` (`.python-version` pins this).
3. `pip install -r requirements.txt`.
4. `baselines/` is gitignored. Two ways to seed it: (a) run `python tools/download_and_clear.py 26B` to do a real download, or (b) copy `baselines/26A/` and `baselines/26B/` from a teammate's machine. Raw downloads live in `baselines/<ver>/originals/`; smart-cleared copies live in `baselines/<ver>/blanks/`.
5. `applaud_snapshot.json` for the audit tool lives at `baselines/applaud/applaud_snapshot.json` — also gitignored. Copy from a teammate or regenerate via the audit pipeline.
6. Smoke test: `python -m pytest tests/` — expect 241 passed.

- [ ] **Step 3: Write section 3 — "Codebase tour" (~600 words)**

Walk `fbdi/` module-by-module, explaining responsibility and how modules connect. Write for a reader, not a reference — prose with paragraph breaks, not a table. Cover (look at each file before writing):

- `fbdi/detect_header.py` — dynamic content scoring to find header row per tab. Why it exists: no hardcoded filename map (Oracle's header row drifts across releases).
- `fbdi/compare.py` — pair-wise diff orchestrator. Each (old, new) file pair runs in a fresh subprocess via `run_worker` with a 120s timeout. Streams via `iter_rows`.
- `fbdi/catalog.py` — per-release snapshot builder. Shares `run_worker` with compare.
- `fbdi/clear.py` — header-preserving smart clear. Uses `detect_header_row`, preserves row 4/5/8 headers intact.
- `fbdi/diagnose.py` — reports detection outcomes per tab (DETECTED / NO_HEADER / SKIPPED_TAB / FILE_TOO_LARGE / FILE_ERROR).
- `fbdi/type_parser.py` — parses Oracle type strings (`VARCHAR2(N CHAR)`, `NUMBER(p,s)`, temporal format masks). Emits `TYPE_PARSE_WARNING` only for genuinely broken strings.
- `fbdi/_subprocess_util.py` — shared `run_worker(target, args, timeout)`. Critically: drains the result queue before `join()` to avoid a Windows pipe-buffer deadlock when payloads exceed ~64 KB.
- `fbdi/audit.py` — Applaud mapping audit engine. Reads `applaud_snapshot.json`, `FBDI_Master_Catalog.xlsx`, and prior `fbdi_applaud_mapping.xlsx`. Two-pass (signal scoring → adjudication).
- `fbdi/build_mapping.py` — builds the working `fbdi_applaud_mapping.xlsx` scaffold.
- `fbdi/catalog_normalize.py` — label normalization for Applaud MDB compatibility.
- `fbdi/cli.py` and `fbdi/__main__.py` — CLI entry; `_resolve_dir()` maps `--old 26A` to `baselines/26A/originals/`.
- `tools/download_and_clear.py` — lives outside the package so Selenium deps stay out of the comparison engine. Chains download → smart-clear.

- [ ] **Step 4: Write section 4 — "The `/fbdi-compare-release` skill" (~250 words)**

Cover:

- **Purpose:** glue, not logic. The skill orchestrates the 8 stages with human-in-the-loop checkpoints. All domain logic lives in `fbdi/` and `tools/`.
- **Location:** `.claude/skills/fbdi-compare-release/`. Files: `SKILL.md` (Claude-facing workflow), `scripts/` (4 Python helpers), `references/` (markdown snippets Claude loads contextually), `evals/` (test harness).
- **The four bundled scripts and exit codes:** `check_env.py` (exit 0 = ok, 1 = missing deps), `verify_download.py` (exit 0 = full, 1 = partial, 2 = none), `summarize_report.py` (prints comparison stats), `verify_run.py` (exit 0 = no regression, 1 = regression flagged).
- **When to modify skill vs. CLI:** if the change is orchestration (a new HITL, a new stage, different stdout pattern), edit the skill. If the change is a new analysis or a different comparison algorithm, add it to `fbdi/` and expose via the CLI, then wire the skill to call it. Don't put domain logic in `SKILL.md`.

- [ ] **Step 5: Write section 5 — "Testing conventions" (~250 words)**

Cover:

- **Layout:** `tests/` flat, one `test_<module>.py` per `fbdi/` module. 241 tests currently.
- **Run:** `python -m pytest tests/` (full), `python -m pytest tests/test_clear.py -v` (single module).
- **No fixtures.** Each test inline-builds its synthetic FBDI workbook using `openpyxl.Workbook` helpers like `_make_fbdi_workbook` or `_create_fbdi_workbook`. The expected layout is visible next to the assertions. Do not introduce a shared `conftest.py` fixture — the team prefers the visible inline pattern.
- **UPPER_SNAKE_CASE gotcha:** `detect_header_row` scores rows by counting UPPER_SNAKE_CASE cells. Synthetic data like `"CREATE"`, `"V1"`, `"DR_ECO_1"` falsely trips header detection. Use lowercase/mixed-case in test data rows — e.g., `"Create Order"`, `"V1-org"`.
- **VBA spot-check:** `tests/validate_against_vba.py` and `tests/vba_fieldrow_map.json` are an ad-hoc comparison against the legacy VBA macro's expected header rows. Not pytest. Kept for regression spot-checks, not run by CI.

- [ ] **Step 6: Write section 6 — "How to add a new release handler" (~300 words)**

Anchor on a concrete scenario: "Oracle just shipped 27A." Walk through, in order:

1. Invoke `/fbdi-compare-release` with "Compare 26B to 27A" — the skill does the heavy lifting; read this section if you need to do it by hand instead.
2. Manual path: `python tools/download_and_clear.py 27A` — populates `baselines/27A/originals/` and `baselines/27A/blanks/`.
3. Handle the FSM file manually (`RapidImplementationForCashManagement.xlsm`) — see CLAUDE.md "Known hazards".
4. `python -m fbdi compare --old 26B --new 27A` — produces `Comparison_Report_26B_27A.xlsx`.
5. `python -m fbdi catalog --release 27A` — adds a `27A` tab to `FBDI_Master_Catalog.xlsx`.
6. `python -m fbdi diagnose --old baselines/26B/originals --new baselines/27A/originals` — verify detection health. A regression in `FILE_ERROR` count vs. prior release is the signal to investigate.
7. If Oracle changes the detection scoring threshold, edit `fbdi/detect_header.py` and add a test in `tests/test_detect_header.py` that pins the new behavior.
8. If a new tab layout breaks, the fix is in `detect_header.py` first, `compare.py` second. Never special-case a filename — that's against the architectural decision in CLAUDE.md.

- [ ] **Step 7: Write section 7 — "Design docs and how we work" (~200 words)**

Cover:

- **Specs and plans:** `docs/superpowers/specs/<date>-<feature>.md` for design, `docs/superpowers/plans/<date>-<feature>.md` for implementation plans. Specs come from the `superpowers:brainstorming` skill; plans from `superpowers:writing-plans`. Both live in git so the history is auditable. Handoff workflow: read the plan end-to-end, execute, push direct to master (no PR for solo work).
- **Resolved-hazards log:** CLAUDE.md has a "Resolved Hazards" section documenting bugs that used to exist (iter_rows perf fix, subprocess deadlock, type parser warnings). Read this before re-raising an "issue" — it's probably already been fixed.
- **CodeRabbit:** wired up for PR review. Not required for solo work on master.

- [ ] **Step 8: Write section 8 — "Known hazards and gotchas" (~300 words)**

Reformat the CLAUDE.md "Known hazards" list from a developer's perspective (what it means when you hit it, where to look in the code):

- **Phantom `max_column=16384`** — caused by corrupt xlsm metadata. Capped at 500 in `compare.py`. Don't lift the cap without understanding why it was added.
- **Corrupt XML in some xlsm files** — `zipfile.BadZipFile` caught in compare; logged as unreadable. `verify_run.py` flags a regression if FILE_ERROR count jumps.
- **5 MB cap in `diagnose` and `build_mapping`** — these load full (non-read_only) workbooks for memory reasons. `compare` is unbounded and streams via `iter_rows`. If you're modifying diagnose/build_mapping, keep the cap.
- **`Comparison_Report_25D_26A.xlsx` (the VBA-generated one)** — has a corrupt stylesheet. Cannot be loaded with vanilla `openpyxl.load_workbook`. Use `read_only=True` or `data_only=True` with exception handling. The Python-generated reports do not have this issue.
- **Test-data UPPER_SNAKE_CASE trap** — already covered in section 5 but worth restating here because it's a common foot-gun.
- **FSM file non-auto-downloadable** — covered in operator-guide.md and CLAUDE.md. Developers hit this if they're testing the download pipeline against a clean `baselines/` directory.

- [ ] **Step 9: Write section 9 — "Where to ask for help" (~50 words)**

Short: read `CLAUDE.md` and `SKILL.md` first. Check `docs/superpowers/specs/` for design context on a specific feature, and `docs/superpowers/plans/` for how it was built. Past audit notes live in `docs/archive/`. Don't treat "DM Brad" as the only answer — the repo should be self-serviceable.

- [ ] **Step 10: Sanity check the draft**

```bash
wc -w docs/developer-guide.md
```

Expected: 2,200–2,800 words. If thin in any section, the likely culprit is section 3 (codebase tour) or section 8 (hazards).

---

## Task 4: Archive narrative history (Bucket 3)

**Files:**
- Move: `docs/Applaud Mapping Audit.md` → `docs/archive/applaud-mapping-audit-notes.md`
- Move: `docs/scraper-gap-findings-2026-04-23.md` → `docs/archive/scraper-gap-findings-2026-04-23.md`
- Move: `Claude_fbdi_applaud_mapping_audit.md` → `docs/archive/claude-fbdi-applaud-mapping-audit.md`

All three are tracked in git — use `git mv` to preserve history.

- [ ] **Step 1: Create the archive directory**

```bash
mkdir -p docs/archive
```

- [ ] **Step 2: Move and rename the three narrative docs**

```bash
git mv "docs/Applaud Mapping Audit.md" docs/archive/applaud-mapping-audit-notes.md
git mv docs/scraper-gap-findings-2026-04-23.md docs/archive/scraper-gap-findings-2026-04-23.md
git mv Claude_fbdi_applaud_mapping_audit.md docs/archive/claude-fbdi-applaud-mapping-audit.md
```

- [ ] **Step 3: Verify the moves**

```bash
git status --short | grep -E 'docs/archive|Applaud|scraper|Claude_fbdi'
ls docs/archive/
```

Expected staged renames (git detects `R100` rename when content unchanged):

- `R  Claude_fbdi_applaud_mapping_audit.md -> docs/archive/claude-fbdi-applaud-mapping-audit.md`
- `R  docs/Applaud Mapping Audit.md -> docs/archive/applaud-mapping-audit-notes.md`
- `R  docs/scraper-gap-findings-2026-04-23.md -> docs/archive/scraper-gap-findings-2026-04-23.md`

`ls docs/archive/` shows all three files.

Do not commit yet — the docs commit (Task 7) bundles the archive moves with the two new guides.

---

## Task 5: Full humanizer pass on `docs/operator-guide.md`

**Files:**
- Modify: `docs/operator-guide.md`

Operator guide is full humanizer treatment per the spec. Run the humanizer on this file alone (not bundled with the developer guide) to keep voices distinct.

- [ ] **Step 1: Invoke the humanizer skill on the operator guide**

Use the `Skill` tool with `skill: "humanizer-skill:humanizer"` and pass the file path `docs/operator-guide.md` as the argument. The skill reads the file, removes AI tells (em-dash overuse, "serves as a testament", SaaS CTA scaffolding, rule-of-three, inflated significance, emoji bullet headers, "let's dive in" patterns, etc.), and writes back a humanized version.

- [ ] **Step 2: Diff-check the result**

```bash
git diff docs/operator-guide.md | head -200
```

Scan the diff. Expected: em-dashes replaced with commas/colons/parens where awkward, inflated phrases toned down, rule-of-three breakups, occasional sentence-length variance. Word count shouldn't change by more than ~10%.

```bash
wc -w docs/operator-guide.md
```

If word count dropped below ~2,100 or grew above ~2,900, the humanizer probably over-corrected. Spot-check — if content was dropped, restore from git and re-run with tighter scope.

- [ ] **Step 3: Read the humanized version top-to-bottom**

Open `docs/operator-guide.md` and read it straight through. Fix any sentences where the humanizer's swap read worse than the original. Watch for: stiff constructions where a dash was replaced with a comma, awkward paragraph breaks where the humanizer split a sentence.

---

## Task 6: Full humanizer pass on `docs/developer-guide.md`

**Files:**
- Modify: `docs/developer-guide.md`

Same process as Task 5, separate invocation.

- [ ] **Step 1: Invoke the humanizer skill on the developer guide**

Use the `Skill` tool with `skill: "humanizer-skill:humanizer"` and pass `docs/developer-guide.md`.

- [ ] **Step 2: Diff-check**

```bash
git diff docs/developer-guide.md | head -200
wc -w docs/developer-guide.md
```

Expect 2,100–2,900 words; confirm no structural damage.

- [ ] **Step 3: Read through and manually fix awkward swaps**

---

## Task 7: Commit docs — guides + archive (Commit 2 of 4)

**Files:**
- Staged new: `docs/operator-guide.md`, `docs/developer-guide.md`
- Staged renames: three narrative docs into `docs/archive/`

- [ ] **Step 1: Stage the new guides**

```bash
git add docs/operator-guide.md docs/developer-guide.md
```

The three renames from Task 4 are already staged.

- [ ] **Step 2: Review the full staged set**

```bash
git status --short
```

Expected:
- `A  docs/operator-guide.md`
- `A  docs/developer-guide.md`
- `R  Claude_fbdi_applaud_mapping_audit.md -> docs/archive/claude-fbdi-applaud-mapping-audit.md`
- `R  docs/Applaud Mapping Audit.md -> docs/archive/applaud-mapping-audit-notes.md`
- `R  docs/scraper-gap-findings-2026-04-23.md -> docs/archive/scraper-gap-findings-2026-04-23.md`

If anything else is staged, stop and investigate — clean this commit to these five files only.

- [ ] **Step 3: Commit**

```bash
git commit -m "$(cat <<'EOF'
docs: add operator and developer guides, archive narrative history

- Add docs/operator-guide.md — 8-stage walkthrough, 6 HITL checkpoints, output reading
- Add docs/developer-guide.md — codebase tour, testing conventions, how to add a release handler
- Move narrative history into docs/archive/: applaud mapping audit notes, scraper gap findings, and the Claude-generated audit doc

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 4: Verify**

```bash
git log -1 --stat
```

Expected: 2 additions + 3 renames.

---

## Task 8: Update `README.md` — guide pointers + repo tree refresh

**Files:**
- Modify: `README.md`

Two concrete changes:
1. Add two pointer lines near the top (after the one-paragraph intro, before `## Setup`).
2. Update the `## Repo structure` tree to reflect the cleaned state (removed files, `docs/archive/` added, `baselines/applaud/` noted).

- [ ] **Step 1: Add the guide pointers**

After the opening paragraph (line 3), before the `---` separator on line 5, insert these two pointer lines with a blank line above and below:

```markdown
> **Running it:** see [`docs/operator-guide.md`](docs/operator-guide.md).
> **Developing on it:** see [`docs/developer-guide.md`](docs/developer-guide.md).
```

- [ ] **Step 2: Refresh the `Repo structure` tree**

Locate the existing tree (lines 62–77). Replace with:

```
FBDI-Compare-Process-Oracle/
├── fbdi/                      # Python comparison/catalog/clear engine
├── tools/                     # Selenium downloader (download_and_clear.py)
├── tests/                     # 241 unit tests (pytest)
├── .claude/skills/            # Project-level Claude Code skills
│   └── fbdi-compare-release/  # Orchestrator for quarterly refreshes
├── docs/
│   ├── operator-guide.md      # End-to-end pipeline walkthrough
│   ├── developer-guide.md     # Codebase tour and extension guide
│   ├── archive/               # Historical narrative docs (audits, gap findings)
│   └── superpowers/           # Design specs and implementation plans
├── baselines/                 # GITIGNORED — downloaded xlsm per release + applaud_snapshot.json
├── reference/                 # Read-only archive of legacy VBA + scripts
├── baseline_files.txt         # Inventory of expected downloads per release
├── FBDI_Master_Catalog.xlsx   # Per-release snapshot catalog (git-tracked)
├── requirements.txt
├── CLAUDE.md                  # Persistent Claude Code context
└── README.md
```

Note the deletions relative to the old tree: no `Comparison_Report_26A_26B.xlsx` line (tracked but not worth listing in the tree), no `complete_mapping.py`, no `format_workbooks.py`, no `applaud_snapshot.json`.

- [ ] **Step 3: Sanity check**

```bash
git diff README.md
```

Expected: only the two pointer lines added and the tree block replaced. No other edits.

---

## Task 9: Light humanizer pass on `README.md`

**Files:**
- Modify: `README.md`

"Light" means: remove obvious AI tells (em-dash overuse, SaaS CTA patterns, "serves as", rule-of-three) without restructuring. README is already close to a human voice; don't over-correct.

- [ ] **Step 1: Invoke the humanizer skill on README.md**

Use the `Skill` tool with `skill: "humanizer-skill:humanizer"` and pass `README.md`. Tell the skill it's a light pass — preserve structure, just clean AI tells.

- [ ] **Step 2: Diff-check**

```bash
git diff README.md
```

If the skill restructured sections or removed content, restore and re-run with narrower scope.

- [ ] **Step 3: Read through**

Open README.md and read top-to-bottom. Confirm the two pointer lines from Task 8 survived the humanizer pass. If they read stilted after humanization, manually touch them up.

---

## Task 10: Commit README — pointers + humanizer (Commit 3 of 4)

**Files:**
- Staged: `README.md`

- [ ] **Step 1: Stage and verify**

```bash
git add README.md
git status --short
```

Expected: `M  README.md` and nothing else.

- [ ] **Step 2: Commit**

```bash
git commit -m "$(cat <<'EOF'
docs(readme): add guide pointers, update repo structure, light humanizer pass

- Link operator-guide.md and developer-guide.md near the top so new readers land on them immediately
- Refresh Repo structure tree to reflect the cleaned-up root (archived scripts, relocated applaud_snapshot, new docs/archive/ directory)
- Light humanizer pass: remove AI tells without restructuring

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 3: Verify**

```bash
git log -1 --stat
```

Expected: `README.md` modification only.

---

## Task 11: `CLAUDE.md` improver pass

**Files:**
- Modify: `CLAUDE.md`

The improver skill sees the post-cleanup state and tightens CLAUDE.md accordingly. It should prune references to deleted files (`complete_mapping.py`, `format_workbooks.py`, root-level `applaud_snapshot.json`), add pointers to the two new guides, and tighten the `reference/` (read-only VBA archive) vs. `docs/archive/` (narrative history) distinction.

- [ ] **Step 1: Invoke the claude-md-improver skill**

Use the `Skill` tool with `skill: "claude-md-management:claude-md-improver"`. The skill audits CLAUDE.md against templates, outputs a quality report, then makes targeted updates.

- [ ] **Step 2: Review the improver's changes**

```bash
git diff CLAUDE.md
```

Look for:
- Pointers to `docs/operator-guide.md` and `docs/developer-guide.md` added somewhere sensible (probably near the top, like README's treatment).
- Any references to deleted files removed.
- No invented content — if the improver added claims about the codebase that aren't true, undo those hunks manually.

If the improver rewrote sections drastically (e.g., condensed the "Active Pipeline" section into two lines), restore from `git checkout -- CLAUDE.md` and run it again with a narrower directive.

- [ ] **Step 3: Read CLAUDE.md top-to-bottom**

Every line should still be true. If the improver changed a version, a test count, or a performance number — verify and correct.

---

## Task 12: Commit CLAUDE.md (Commit 4 of 4)

**Files:**
- Staged: `CLAUDE.md`

- [ ] **Step 1: Stage and verify**

```bash
git add CLAUDE.md
git status --short
```

Expected: `M  CLAUDE.md` and nothing else.

- [ ] **Step 2: Commit**

```bash
git commit -m "$(cat <<'EOF'
docs(claude-md): refresh after handoff-docs cleanup

- Add pointers to docs/operator-guide.md and docs/developer-guide.md
- Remove references to files archived or relocated during cleanup
- Tighten reference/ vs. docs/archive/ distinction

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 3: Verify**

```bash
git log -1 --stat
```

Expected: `CLAUDE.md` modification only.

---

## Task 13: Final verification + push to master

**Files:**
- Verify only (no commits): `tests/`, `graphify-out/`

- [ ] **Step 1: Run the full test suite one more time**

```bash
python -m pytest tests/
```

Expected: 241 passed. If anything fails here, it means one of the docs/README/CLAUDE.md edits somehow broke an import — which would be surprising. Diagnose before pushing.

- [ ] **Step 2: Rebuild the local graph**

```bash
python -c "from graphify.watch import _rebuild_code; from pathlib import Path; _rebuild_code(Path('.'))"
```

`graphify-out/` is gitignored, so this produces no commit. It keeps the local knowledge graph current for future Claude sessions.

- [ ] **Step 3: Review the four-commit stack**

```bash
git log --oneline origin/master..HEAD
```

Expected four commits, in this order:

```
<hash> docs(claude-md): refresh after handoff-docs cleanup
<hash> docs(readme): add guide pointers, update repo structure, light humanizer pass
<hash> docs: add operator and developer guides, archive narrative history
<hash> chore(cleanup): archive orphan scripts, remove stale artifacts, relocate applaud_snapshot
```

If any commit title differs from the spec or any commit bundles the wrong files, fix before pushing. It is cheaper to `git reset --soft HEAD~N` and re-commit than to push a messy stack.

- [ ] **Step 4: Push to master**

```bash
git push origin master
```

Expected: 4 commits pushed, fast-forward. If remote has diverged (someone else pushed in the interim), stop — do not force-push master. Rebase locally and re-verify.

- [ ] **Step 5: Confirm success**

```bash
git log --oneline -5
git status
```

Expected: four new commits at the top; working tree clean; local is synced with `origin/master`.

---

## Success criteria (from the spec)

- A fork can be handed to a Definian coworker who runs `/fbdi-compare-release` end-to-end using only `docs/operator-guide.md`.
- A coworker can navigate the codebase and make a first change using only `docs/developer-guide.md`.
- Both new docs read as human-written (humanizer pass applied).
- README is the launchpad — guide pointers visible within the first 10 lines.
- CLAUDE.md accurately reflects the post-cleanup state.
- Root directory is clean: no orphan scripts, no Excel lock files, no stale reports.
- Test suite still passes (241 tests).
- Graph is current.
