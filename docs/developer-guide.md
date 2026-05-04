# Developer Guide

> If you just want to run the quarterly refresh, start with [`operator-guide.md`](operator-guide.md).

## Orientation

Oracle ships updated FBDI (File-Based Data Import) templates each quarter. Definian's integrations are built on top of those templates, and when Oracle renames, adds, or removes fields, the integrations break silently unless someone catches it. This repo automates the catching: given two releases (say 26A and 26B), it produces a field-level diff across every FBDI template so Brad and Dan know exactly what changed before a client goes live.

What's shipped and working: a comparison engine that diffs all template pairs, a per-release catalog that snapshots every field with its type and metadata, a Selenium-based downloader, a smart-clear tool that strips sample data from templates while preserving headers, an Applaud mapping audit that checks whether FBDI fields are covered by downstream Applaud target tables, a compliance report generator (HTML + PDF via `fbdi report`), 320 unit tests, and an orchestrator skill that chains all of it with human-in-the-loop checkpoints.

What's on the frontier: `python -m fbdi run` (not built) would chain download → compare → catalog → report in one shot. The FBDI-to-Applaud mapping (`FBDI_to_ApplaudTables_Mapping.xlsx`) is complete with no TBD rows as of 2026-05-04; new FBDI tabs added in future Oracle releases may introduce rows needing manual review.

The authoritative source of truth for current state is `CLAUDE.md` at the repo root.

## Local setup for development

1. Clone the repo and `cd` into it.

2. Install Python 3.14. The repo uses `pyenv-win` on Windows; `.python-version` pins the exact version. If `python --version` doesn't show 3.14.x, install pyenv-win and run `pyenv install` in the repo root.

3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
   Core deps are `openpyxl`, `selenium`, `webdriver-manager`, `requests`, and `pytest`.

4. Seed your `baselines/` directory. It's gitignored, so it won't be there after a fresh clone. Two ways:
   - Download fresh: `python tools/download_and_clear.py 26B`. This runs the Selenium scraper and populates `baselines/26B/originals/`, then smart-clears to `baselines/26B/blanks/`. Takes 15–20 minutes.
   - Copy from a teammate: grab `baselines/26A/` and `baselines/26B/` from someone who already has them. Raw Oracle downloads live under `originals/`; cleared copies under `blanks/`.

5. If you're working on the Applaud audit, you also need `applaud_snapshot.json`. It lives at `baselines/applaud/applaud_snapshot.json` and is also gitignored. Copy it from a teammate or ask Brad. It's a snapshot of the Applaud MDB schema.

6. Smoke test:
   ```
   python -m pytest tests/
   ```
   Expect 320 passed, no failures or errors.

## Codebase tour

Everything domain-specific lives under `fbdi/`. The `tools/` directory holds the Selenium downloader, which lives outside the package deliberately. Here's what each module does.

**`fbdi/detect_header.py`** solves a genuinely annoying Oracle problem: FBDI templates don't have a consistent header row. Some put the technical field names at row 4, some at row 5, some at row 8 or even row 11. The legacy VBA macro hardcoded a per-filename map; that breaks every release when Oracle reshuffles templates. `detect_header_row()` takes a worksheet and scans the first 20 rows, scoring each against two tiers. Tier 1 looks for a row where more than half the non-empty cells match `^[A-Z][A-Z0-9_]+$` (UPPER_SNAKE_CASE technical names like `INTERFACE_BATCH_CODE`). Tier 2 scores rows on a weighted combination of header-like characteristics (short strings, high fill ratio, all-string content) for templates that use mixed-case labels instead. It streams via `iter_rows` rather than random-access `ws.cell()` lookups; on sheets with 500+ columns, this was the difference between 74 seconds and under a second per tab.

**`fbdi/compare.py`** is the pairwise diff engine. Given two directories (one per release), it matches files by name, then for each matched pair spawns a fresh subprocess to diff the pair tab-by-tab. It detects the header row independently per file per tab (old and new can drift), reads the header values, and aligns by position to produce `ComparisonRow` objects: one row per column position, flagged YES/NO for difference. The subprocess isolation is essential. Openpyxl accumulates file handles and memory across sequential loads, and after roughly 50 loads in one process the behavior gets unpredictable. Each pair gets a 120s timeout. Results are written to a 7-column `Comparison_Report_<OLD>_<NEW>.xlsx`.

**`fbdi/catalog.py`** builds `FBDI_Master_Catalog.xlsx`. It uses the same subprocess-per-file isolation pattern as compare, but instead of diffing it extracts a structured snapshot of every field: position, human-readable label, technical name, data type, length, scale, and required flag. The workbook gets one tab per release (e.g., `26A`, `26B`) plus an `Issues` tab for anything that went wrong (parse failures, file errors, tabs with no detectable header) and a `Drift` tab that position-aligns the two most recent releases and classifies changes as ADDED, REMOVED, RENAMED, TYPE_CHANGED, LENGTH_CHANGED, REQUIRED_CHANGED, or MULTI. Re-running catalog for a single release regenerates only that release's tab. It reads and preserves the others from the existing workbook.

**`fbdi/clear.py`** produces client-ready blank templates. It opens each xlsm with `keep_vba=True`, calls `detect_header_row()` per sheet, and clears every row below the detected header. It also strips malformed VML XML references that cause openpyxl save errors on some Oracle files. The comparison engine reads from `originals/` and never touches this output.

**`fbdi/diagnose.py`** is a health-check tool. It runs header detection on every tab of every file and reports one of five outcomes per tab: `DETECTED` (found a header row), `NO_HEADER` (scored below the confidence threshold), `SKIPPED_TAB` (instructions or summary tab, deliberately ignored), `FILE_TOO_LARGE` (over 5 MB, not loaded in full mode), or `FILE_ERROR` (corrupt or unreadable file). It loads workbooks in full (non-read_only) mode, so it respects the 5 MB cap in `config.py`. The output is a `Diagnostic_Report_<label>.xlsx`. The thing to watch for across releases is the `FILE_ERROR` count. A jump usually means Oracle shipped a corrupt file that the comparison engine is silently skipping.

**`fbdi/type_parser.py`** parses Oracle's data type strings into structured fields. The inputs look like `VARCHAR2(5 CHAR)`, `NUMBER(18,4)`, `DATE`, `CLOB`, `DATE(YYYY/MM/DD)`, `TimeStamp(hh24:mm:ss)`. Two patterns cover everything Oracle actually ships: a strict shape pattern for standard SQL types with optional length/scale, and a temporal format-mask pattern for DATE and TIMESTAMP variants where the parenthesized content is a format string rather than a length. The parser also tolerates Oracle's occasional stray trailing period (e.g., `VARCHAR2(1 CHAR).`). `parse_warning=True` is emitted only for inputs that genuinely don't match either pattern; there are currently 9 of those across all releases.

**`fbdi/_subprocess_util.py`** is the shared worker harness. Both `compare.py` and `catalog.py` call `run_worker(target, args, timeout)` to run file processing in a fresh subprocess. The implementation detail that matters most: the queue is drained before `join()` is called. On Windows, `multiprocessing.Queue.put()` hands off to a background feeder thread that writes to an OS pipe with roughly a 64 KB buffer. If the parent calls `join()` before reading the queue, the feeder blocks because the pipe is full, the child can't exit, and `join()` times out. This was an actual historical bug that caused `ChangeOrderImportTemplate` and `ItemImportTemplate` to report bogus TIMEOUT in the catalog; both files produce payloads larger than 64 KB. Do not refactor `run_worker` without understanding this constraint.

**`fbdi/audit.py`** is the Applaud mapping audit engine. It reads `applaud_snapshot.json` (at `baselines/applaud/applaud_snapshot.json`), `FBDI_Master_Catalog.xlsx`, and the working `FBDI_to_ApplaudTables_Mapping.xlsx`. Two-pass adjudication: score signals like name similarity and prefix matching, then classify each FBDI tab as YES, NEEDS_REVIEW, or UNMAPPED. Outputs are `Claude_fbdi_applaud_mapping.xlsx` (three sheets) and a markdown audit report.

**`fbdi/build_mapping.py`** built the initial scaffold of `FBDI_to_ApplaudTables_Mapping.xlsx`. It's a one-shot utility that scanned the 25D and 26A baselines, enumerated tabs, merged 9 known hardcoded Applaud mappings, and wrote the starting point.

**`fbdi/catalog_normalize.py`** is a single function, `normalize_label()`. It strips characters Applaud doesn't handle cleanly (asterisks, punctuation, symbols) while preserving alphanumerics, underscores, and whitespace. Applied only to user-facing labels; technical names are left alone.

**`fbdi/cli.py` and `fbdi/__main__.py`** wire the package up as a CLI. The main convenience is `_resolve_dir()`: when you pass `--old 26A`, it resolves that to `baselines/26A/originals/` automatically. Three subcommands are exposed: `compare`, `catalog`, and `diagnose`.

**`tools/download_and_clear.py`** lives outside `fbdi/` on purpose. Selenium and webdriver-manager are heavy dependencies and don't belong in the comparison engine. This script chains download → smart-clear, or accepts `--clear-only` to re-clear without re-downloading.

## The `/fbdi-compare-release` skill

The skill's job is glue, not logic. All the domain work happens in `fbdi/` and `tools/`. The skill orchestrates the 8-stage pipeline (download, verify, smart-clear, compare, catalog, verify again) with human-in-the-loop checkpoints so a non-developer can run a quarterly refresh without understanding the internals.

The skill lives at `.claude/skills/fbdi-compare-release/`. The files:

- `SKILL.md` — the Claude-facing workflow document. This is what Claude reads when the skill triggers. It defines all 8 stages, the HITL prompts (numbered #1–#6, not sequential by execution order), and the exact commands to run.
- `scripts/` — four Python helpers that the skill calls out to.
- `references/` — markdown snippets Claude loads contextually (troubleshooting steps, known issues, etc.).
- `evals/` — test harness for the skill itself.

The four bundled scripts and their exit codes:

- `check_env.py` — preflights Python version, Chrome, and baseline file presence. Exit 0 = ok, 1 = fatal dep missing, 2 = pip deps missing.
- `verify_download.py` — checks the downloaded file count against `baseline_files.txt` inventory. Exit 0 = clean, 1 = missing files, 2 = extras only, 3 = first-run bootstrap needed (no entry for this release in the inventory).
- `summarize_report.py` — prints comparison statistics (changed/added/removed/total) from a completed report workbook.
- `verify_run.py` — post-run health check. Runs diagnose and checks the Issues tab for regressions vs. the prior release. Exit 0 = no regression, 1 = regression flagged.

When to modify the skill vs. the CLI: if the change is orchestration (a new HITL checkpoint, a new stage, different output to parse), edit `SKILL.md`. If the change is a new analysis, comparison algorithm, or data extraction, add it to `fbdi/` and expose it via the CLI, then have the skill call it. Domain logic does not belong in `SKILL.md`.

## Testing conventions

The test suite lives in `tests/`, flat structure, one file per `fbdi/` module. Currently 241 tests.

Run the full suite:
```
python -m pytest tests/
```

Run one module:
```
python -m pytest tests/test_clear.py -v
```

There are no pytest fixtures and no `conftest.py`. Each test builds its own synthetic FBDI workbook inline using `openpyxl.Workbook` helpers (`_make_fbdi_workbook`, `_create_fbdi_workbook`, or equivalent per-module variants). The expected layout sits right next to the assertions, which makes it easy to see what structure a test is exercising without jumping to a fixture file. Don't introduce shared fixtures; the team prefers the visible inline pattern.

The most common gotcha for anyone writing a new test: `detect_header_row()` scores rows by counting UPPER_SNAKE_CASE cells. If your test data rows contain values like `"CREATE"`, `"V1"`, or `"DR_ECO_1"`, the scorer will flag them as header rows and your test will fail in confusing ways. Use lowercase or mixed-case values in data rows: `"Create Order"`, `"V1-org"`, `"open"`. Keep UPPER_SNAKE_CASE only in the rows you intend to be headers.

There are also two ad-hoc spot-check files: `tests/validate_against_vba.py` and `tests/vba_fieldrow_map.json`. These compare the Python engine's header row detection against the legacy VBA macro's hardcoded row map. They're not pytest and not run in CI. They're there for regression spot-checks when you suspect a detection change broke something that was previously correct.

## How to add a new release handler

Say Oracle just shipped 27A. The fastest path is to trigger `/fbdi-compare-release` with "Compare 26B to 27A", and the skill handles everything. Read this section if you're doing it by hand instead, or if something in the skill's run goes sideways and you need to pick up mid-stream.

1. Download and clear the new release:
   ```
   python tools/download_and_clear.py 27A
   ```
   This populates `baselines/27A/originals/` and `baselines/27A/blanks/`. It takes 15–20 minutes. Don't re-run blindly on an existing directory; it wipes `originals/` first.

2. Handle `RapidImplementationForCashManagement.xlsm` manually. Oracle's Selenium scraper never finds this file because it's a Rapid Implementation (FSM) template, not a standard FBDI template. Get it from Oracle Fusion: Setup and Maintenance → hamburger menu → Search → "Create Banks, Branches, and Accounts in Spreadsheet" → download. Drop it into `baselines/27A/originals/`. If you skip this, the comparison runs fine but you'll have a gap in the coverage.

3. Run the comparison:
   ```
   python -m fbdi compare --old 26B --new 27A --output Comparison_Report_26B_27A.xlsx
   ```

4. Add the new release to the catalog:
   ```
   python -m fbdi catalog --release 27A
   ```
   This adds a `27A` tab to `FBDI_Master_Catalog.xlsx` and recomputes the Drift tab against 26B.

5. Verify detection health:
   ```
   python -m fbdi diagnose --old baselines/26B/originals --new baselines/27A/originals
   ```
   Check the `FILE_ERROR` count. If it jumped vs. the prior release, something Oracle shipped is corrupt or uses a format the engine doesn't handle. Investigate before handing off the report.

6. If Oracle changed the structure of a template's header rows and detection breaks, the fix lives in `fbdi/detect_header.py`. Adjust the scoring weights or thresholds, then add a test in `tests/test_detect_header.py` that pins the new behavior. Do not special-case a filename. That's the architectural decision the engine was built to avoid. The whole point of dynamic detection is that filenames shouldn't matter.

7. If a previously-working tab now produces `NO_HEADER`, `detect_header.py` is the first place to look. If it's failing on the comparison side (a header is detected but fields are wrong), check `compare.py`. If the catalog is missing fields, check `catalog.py`'s `_extract_rich` vs. `_extract_thin` dispatch. The distinction matters when Oracle's header row is UPPER_SNAKE_CASE vs. mixed-case labels.

## Design docs and how we work

Specs and plans follow a consistent pattern. Design docs live at `docs/superpowers/specs/<date>-<feature>.md`. Implementation plans live at `docs/superpowers/plans/<date>-<feature>.md`. Specs are produced with the `superpowers:brainstorming` skill; plans come from `superpowers:writing-plans`. Both are committed to git so there's an auditable trail of why decisions were made.

The handoff workflow: read the plan end-to-end before touching any code, execute it start to finish, push directly to master. No PR required for solo work; CodeRabbit is wired up for PR review but optional on master pushes.

`CLAUDE.md` has a "Resolved Hazards" section that documents bugs that used to exist: the iter_rows performance fix, the subprocess deadlock, the 463-to-9 type parser warning collapse. Before you raise something as a new issue, check that section first. If it's already there, the fix is already in the codebase and you're probably looking at a different root cause.

For cross-release historical context, the `reference/` directory at repo root is a read-only archive: the legacy VBA comparison macro, the legacy clearing macro, and a sample VBA output report. Don't modify anything there.

## Known hazards and gotchas

**Phantom `max_column=16384`.** Some xlsm files report 16,384 columns due to corrupt formatting metadata. Both `compare.py` and `detect_header.py` cap column scanning at 500. Don't lift that cap without understanding why it's there; scanning all 16,384 phantom columns on a wide sheet takes a very long time.

**Corrupt XML in some xlsm files.** `openpyxl` throws `zipfile.BadZipFile` on a handful of Oracle templates. The engine catches it and logs the file as a `FILE_ERROR`. The issue shows up in the catalog's Issues tab and in the diagnose report. `verify_run.py` flags a regression if the FILE_ERROR count jumps between releases. That's the signal to go investigate what Oracle shipped.

**5 MB cap in `diagnose` and `build_mapping`.** These two modules load workbooks in full (non-read_only) mode, which holds the whole file in memory. The cap is in `fbdi/config.py` as `MAX_FILE_SIZE_BYTES`. The comparison engine doesn't have this limit because it streams via `iter_rows` in read_only mode. If you're extending `diagnose` or `build_mapping`, keep the cap; don't silently remove it and wonder why memory spikes.

**`Comparison_Report_25D_26A.xlsx` (the VBA-generated one).** The legacy VBA macro produced a comparison report with a corrupt stylesheet. Standard `openpyxl.load_workbook()` will throw an exception on it. If you need to read this file programmatically, use `read_only=True` or `data_only=True` with exception handling wrapped around the load. Python-generated reports don't have this issue.

**Test-data UPPER_SNAKE_CASE trap.** Already covered in the testing section, but worth repeating because it bites people: any cell value matching `^[A-Z][A-Z0-9_]+$` looks like a header to the scorer. Strings like `"CREATE"`, `"NULL"`, `"V1"`, `"DR_ECO_1"` will all trigger false-positive header detection in tests. Use mixed-case or lowercase in data rows.

**`RapidImplementationForCashManagement.xlsm` is not auto-downloadable.** If you're testing the download pipeline against a fresh `baselines/` directory, this file will be missing after the Selenium run finishes. The downloader warns you, but it won't stop. The comparison and catalog engines pick it up automatically once you place it; they just skip it silently if it's absent.

## Where to ask for help

Read `CLAUDE.md` and the relevant skill's `SKILL.md` first. Most questions about current state are answered there. For design context on a specific feature, check `docs/superpowers/specs/`; the spec will explain the tradeoffs that led to the current implementation. For how a feature was built step by step, check `docs/superpowers/plans/`. Past audit notes and one-off findings are in `docs/archive/`. The repo is meant to be self-serviceable; "DM Brad" is a last resort, not a first step.
