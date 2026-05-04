# Developer Guide

> If you just want to run the quarterly refresh, start with [`operator-guide.md`](operator-guide.md).

## Orientation

Oracle ships updated FBDI (File-Based Data Import) templates each quarter. Definian's integrations sit on top of those templates, and when Oracle renames, adds, or removes a field, the integrations break silently unless someone catches it. This repo is the catching part: given two releases (say 26A and 26B), it produces a field-level diff across every FBDI template so the integration team knows exactly what changed before a client goes live.

Shipped and working today:

- A comparison engine that diffs all template pairs.
- A per-release catalog that snapshots every field with its type and metadata.
- A Selenium-based downloader.
- A smart-clear tool that strips sample data from templates while preserving headers.
- An Applaud mapping audit that checks whether FBDI fields are covered by downstream Applaud target tables.
- A compliance report generator that produces both HTML and PDF (`python -m fbdi report`).
- 320 unit tests.
- An orchestrator skill that chains the lot of it together with human-in-the-loop checkpoints.

On the frontier: `python -m fbdi run`, a headless chained pipeline (download, compare, catalog, populate-module, report) that runs without Claude in the loop. The implementation plan is at `docs/superpowers/plans/2026-05-04-fbdi-run-headless-pipeline.md`. The FBDI-to-Applaud mapping (`FBDI_to_ApplaudTables_Mapping.xlsx`) is complete with no TBD rows as of 2026-05-04; new tabs introduced in future Oracle releases may bring rows that need manual review.

`CLAUDE.md` at the repo root is the source of truth for current state. When this guide and CLAUDE.md disagree, trust CLAUDE.md.

## Local setup for development

1. Clone the repo and `cd` into it.

2. Install Python 3.14. The repo uses `pyenv-win` on Windows, and `.python-version` pins the exact patch version. If `python --version` doesn't show 3.14.x, install pyenv-win and run `pyenv install` in the repo root.

3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
   The core packages are `openpyxl`, `selenium`, `webdriver-manager`, `requests`, `pytest`, `jinja2`, and `weasyprint`. PDF rendering also requires MSYS2 mingw64 GTK on Windows; CLAUDE.md has the install steps under Known Hazards.

4. Seed your `baselines/` directory. It's gitignored, so it won't be there after a fresh clone. Two ways to get it:
   - Download fresh by running `python tools/download_and_clear.py 26B`. This runs the Selenium scraper, populates `baselines/26B/originals/`, then smart-clears the templates into `baselines/26B/blanks/`. Plan on 15–20 minutes.
   - Copy from a teammate. Grab `baselines/26A/` and `baselines/26B/` from someone who already has them. Raw Oracle downloads live under `originals/`, cleared copies under `blanks/`.

5. If you're working on the Applaud audit, you also need `applaud_snapshot.json`. It lives at `baselines/applaud/applaud_snapshot.json` and is gitignored. Copy it from a teammate. It's a snapshot of the Applaud MDB schema.

6. Smoke test:
   ```
   python -m pytest tests/
   ```
   Expect 320 passed, no failures or errors.

## Codebase tour

Everything domain-specific lives under `fbdi/`. The `tools/` directory holds the Selenium downloader, which sits outside the package on purpose so the comparison engine doesn't drag in a browser stack. Here's what each module does.

**`fbdi/detect_header.py`** solves the most aggravating Oracle quirk in the repo: FBDI templates don't have a consistent header row. Some put the technical field names at row 4, some at row 5, some at row 8, occasionally row 11. The legacy VBA macro hardcoded a per-filename map, which broke every release when Oracle reshuffled templates. `detect_header_row()` takes a worksheet and scans the first 20 rows, scoring each against two tiers. Tier 1 looks for a row where more than half the non-empty cells match `^[A-Z][A-Z0-9_]+$` (UPPER_SNAKE_CASE technical names like `INTERFACE_BATCH_CODE`). Tier 2 scores rows on a weighted combination of header-like characteristics (short strings, high fill ratio, all-string content) for the templates that use mixed-case labels instead of technical names. It streams via `iter_rows` rather than random-access `ws.cell()` lookups; on sheets with 500+ columns, that was the difference between 74 seconds and under a second per tab.

**`fbdi/compare.py`** is the pairwise diff engine. Given two directories, one per release, it matches files by name, then spawns a fresh subprocess for each matched pair to diff that pair tab by tab. It detects the header row independently per file per tab (old and new can drift), reads the header values, and aligns by position to produce `ComparisonRow` objects, one row per column position, flagged YES/NO for difference. The subprocess isolation is essential. Openpyxl accumulates file handles and memory across sequential loads, and after roughly 50 loads in one process the behavior gets unpredictable. Each pair gets a 120s timeout. Results are written to a 7-column `Comparison_Report_<OLD>_<NEW>.xlsx`.

**`fbdi/catalog.py`** builds `FBDI_Master_Catalog.xlsx`. It uses the same subprocess-per-file isolation pattern as compare, but instead of diffing it extracts a structured snapshot of every field: position, human-readable label, technical name, data type, length, scale, and required flag. The workbook has one tab per release (e.g., `26A`, `26B`), plus an `Issues` tab for anything that went wrong (parse failures, file errors, tabs with no detectable header), plus a `Drift` tab. The Drift tab uses alignment-driven classification (`align.align_tabs`) to label changes as ADDED, REMOVED, SHIFTED, RENAMED, MODIFIED, or MULTI. Re-running catalog for a single release regenerates only that release's tab. It reads and preserves the others from the existing workbook.

**`fbdi/clear.py`** produces client-ready blank templates. It opens each xlsm with `keep_vba=True`, calls `detect_header_row()` per sheet, and clears every row below the detected header. It also strips malformed VML XML references that cause openpyxl save errors on some Oracle files. The comparison engine reads from `originals/` and never touches this output.

**`fbdi/diagnose.py`** is a health-check tool. It runs header detection on every tab of every file and reports one of five outcomes per tab: `DETECTED` (found a header row), `NO_HEADER` (scored below the confidence threshold), `SKIPPED_TAB` (instructions or summary tab, deliberately ignored), `FILE_TOO_LARGE` (over 5 MB, not loaded in full mode), or `FILE_ERROR` (corrupt or unreadable file). It loads workbooks in full (non-read_only) mode, so it respects the 5 MB cap in `config.py`. The output is `Diagnostic_Report_<label>.xlsx`. The number to watch across releases is `FILE_ERROR`. A jump usually means Oracle shipped a corrupt file that the comparison engine is silently skipping.

**`fbdi/type_parser.py`** parses Oracle's data type strings into structured fields. The inputs look like `VARCHAR2(5 CHAR)`, `NUMBER(18,4)`, `DATE`, `CLOB`, `DATE(YYYY/MM/DD)`, `TimeStamp(hh24:mm:ss)`. Two patterns cover everything Oracle actually ships: a strict shape pattern for standard SQL types with optional length/scale, and a temporal format-mask pattern for DATE and TIMESTAMP variants where the parenthesized content is a format string rather than a length. The parser also tolerates Oracle's occasional stray trailing period (for example, `VARCHAR2(1 CHAR).`). `parse_warning=True` is emitted only for inputs that genuinely don't match either pattern; there are 9 of those across all releases at the moment.

**`fbdi/_subprocess_util.py`** is the shared worker harness. Both `compare.py` and `catalog.py` call `run_worker(target, args, timeout)` to run file processing in a fresh subprocess. The implementation detail that matters most: the queue is drained before `join()` is called. On Windows, `multiprocessing.Queue.put()` hands off to a background feeder thread that writes to an OS pipe with roughly a 64 KB buffer. If the parent calls `join()` before reading the queue, the feeder blocks because the pipe is full, the child can't exit, and `join()` times out. This was an actual historical bug that caused `ChangeOrderImportTemplate` and `ItemImportTemplate` to report bogus TIMEOUT in the catalog; both files produce payloads larger than 64 KB. Do not refactor `run_worker` without understanding this constraint.

**`fbdi/align.py`** is the alignment algorithm shared by `catalog.py` (the Drift writer) and `report.py`. `align_tabs(old_rows, new_rows) -> list[Change]` does an LCS-style alignment and classifies each row across three axes: label, metadata, and position. The result is a single `Change` per row, tagged SHIFTED, RENAMED, MODIFIED, ADDED, REMOVED, or MULTI. If you find yourself wanting to reimplement diff logic somewhere new, route it through `align.py` instead.

**`fbdi/audit.py`** is the Applaud mapping audit engine. It reads `applaud_snapshot.json` (at `baselines/applaud/applaud_snapshot.json`), `FBDI_Master_Catalog.xlsx`, and the working `FBDI_to_ApplaudTables_Mapping.xlsx`. Two-pass adjudication: score signals like name similarity and prefix matching, then classify each FBDI tab as YES, NEEDS_REVIEW, or UNMAPPED. Outputs are `Claude_fbdi_applaud_mapping.xlsx` (three sheets) and a markdown audit report.

**`fbdi/build_mapping.py`** built the initial scaffold of `FBDI_to_ApplaudTables_Mapping.xlsx`. It's a one-shot utility that scanned the 25D and 26A baselines, enumerated tabs, merged 9 hardcoded Applaud mappings, and wrote the starting point.

**`fbdi/populate_module.py`** is a surgical column-F updater for `FBDI_to_ApplaudTables_Mapping.xlsx`. It reads `baselines/<ver>/file_modules.json` (NEW wins, OLD as fallback) and writes the Module column. It uses openpyxl in full mode so formatting, formulas, and freeze-panes survive the rewrite.

**`fbdi/applaud_type.py`** translates Oracle types to Applaud types. `applaud_type_for(parsed_type) -> str`. `VARCHAR2(N)` becomes `char N`, `NUMBER(p,s)` becomes `numeric p,s`, `DATE` and `TIMESTAMP` both become `date`, and so on.

**`fbdi/report.py`** generates the compliance report. `generate_report(catalog_path, mapping_path, old_release, new_release, out_dir) -> (html_path, pdf_path)`. It filters to MAPPED in-scope tabs, routes pending-base tabs to a separate section, and renders both HTML and PDF from a single Jinja2 template (`fbdi/templates/report.html.j2`) via `weasyprint`. The PDF rendering needs MSYS2 mingw64 GTK on Windows; if you're on a fresh machine and `python -m fbdi report` blows up on Pango, that's the cause.

**`fbdi/catalog_normalize.py`** is a single function, `normalize_label()`. It strips characters Applaud doesn't handle cleanly (asterisks, punctuation, symbols) while preserving alphanumerics, underscores, and whitespace. It's applied only to user-facing labels; technical names are left alone.

**`fbdi/cli.py` and `fbdi/__main__.py`** wire the package up as a CLI. The main convenience is `_resolve_dir()`. When you pass `--old 26A`, it resolves that to `baselines/26A/originals/` automatically. Subcommands exposed today: `compare`, `catalog`, `diagnose`, `populate-module`, and `report`.

**`tools/download_and_clear.py`** lives outside `fbdi/` deliberately. Selenium and webdriver-manager are heavy dependencies and don't belong in the comparison engine. This script chains download then smart-clear, or accepts `--clear-only` to re-clear without re-downloading.

## The `/fbdi-compare-release` skill

The skill's job is glue, not logic. All the domain work happens in `fbdi/` and `tools/`. The skill orchestrates a 9-stage pipeline (plus an interim Stage 6.5 for mapping updates) with eight human-in-the-loop checkpoints so a non-developer can run a quarterly refresh without understanding the internals. The stages cover preflight, version resolve, download, smart-clear, compare, catalog, populate-module, summary, post-run verification, and the compliance report.

The skill lives at `.claude/skills/fbdi-compare-release/`. The files:

- `SKILL.md` is the Claude-facing workflow document. This is what Claude reads when the skill triggers. It defines all stages, the HITL prompts (numbered #1 through #8, not sequential by execution order), and the exact commands to run.
- `scripts/` holds the Python helpers that the skill calls out to.
- `references/` is markdown snippets Claude loads contextually (troubleshooting steps, known issues, and so on).
- `evals/` is the test harness for the skill itself.

The bundled scripts and their exit codes:

- `check_env.py` preflights Python version, Chrome, and baseline file presence. Exit 0 means ok, 1 means a fatal dep is missing, 2 means pip deps are missing.
- `verify_download.py` checks the downloaded file count against `baseline_files.txt`. Exit 0 means clean, 1 means files are missing, 2 means extras only, 3 means first-run bootstrap (no entry for this release in the inventory).
- `summarize_report.py` prints comparison statistics (changed, added, removed, total) from a completed report workbook.
- `verify_run.py` is a post-run health check. It runs diagnose and looks for regressions in the Issues tab versus the prior release. Exit 0 means no regression, 1 means a regression was flagged.
- `verify_rerun.py` is the macro-signal validator that watches for backslide between catalog runs.

When to modify the skill vs. the CLI: if the change is orchestration (a new HITL checkpoint, a new stage, different output to parse), edit `SKILL.md`. If the change is a new analysis, comparison algorithm, or data extraction, add it to `fbdi/` and expose it through the CLI, then have the skill call it. Domain logic does not belong in `SKILL.md`.

## Testing conventions

The test suite lives in `tests/` in a flat structure, one file per `fbdi/` module. 320 tests as of 2026-05-04.

Run the full suite:
```
python -m pytest tests/
```

Run one module:
```
python -m pytest tests/test_clear.py -v
```

There are no pytest fixtures and no `conftest.py`. Each test builds its own synthetic FBDI workbook inline using `openpyxl.Workbook` helpers (`_make_fbdi_workbook`, `_create_fbdi_workbook`, or equivalent per-module variants). The expected layout sits right next to the assertions, which makes it easy to see what structure a test is exercising without jumping out to a fixture file. Don't introduce shared fixtures. The team prefers the visible inline pattern.

The most common gotcha for anyone writing a new test: `detect_header_row()` scores rows by counting UPPER_SNAKE_CASE cells. If your test data rows contain values like `"CREATE"`, `"V1"`, or `"DR_ECO_1"`, the scorer will flag them as header rows and your test will fail in confusing ways. Use lowercase or mixed-case values in data rows: `"Create Order"`, `"V1-org"`, `"open"`. Keep UPPER_SNAKE_CASE only in the rows you intend to be headers.

Two ad-hoc spot-check files also live in `tests/`: `validate_against_vba.py` and `vba_fieldrow_map.json`. They compare the Python engine's header-row detection against the legacy VBA macro's hardcoded row map. They aren't pytest and aren't run in CI. They sit there for regression spot-checks when you suspect a detection change broke something that used to be correct.

## How to add a new release handler

Say Oracle just shipped 27A. The fastest path is to trigger `/fbdi-compare-release` with "Compare 26B to 27A" and let the skill handle everything. Read this section if you're doing it by hand, or if something in the skill's run goes sideways and you need to pick up mid-stream.

1. Download and clear the new release:
   ```
   python tools/download_and_clear.py 27A
   ```
   This populates `baselines/27A/originals/` and `baselines/27A/blanks/`. It takes 15–20 minutes. Don't re-run blindly on an existing directory. It wipes `originals/` first.

2. Handle `RapidImplementationForCashManagement.xlsm` manually. Oracle's Selenium scraper never finds this file because it's a Rapid Implementation (FSM) template, not a standard FBDI template. Get it from Oracle Fusion: Setup and Maintenance, hamburger menu, Search, "Create Banks, Branches, and Accounts in Spreadsheet", then download. Drop it into `baselines/27A/originals/`. If you skip this step, the comparison still runs fine, but you'll have a coverage gap.

3. Run the comparison:
   ```
   python -m fbdi compare --old 26B --new 27A --output Comparison_Report_26B_27A.xlsx
   ```

4. Add the new release to the catalog:
   ```
   python -m fbdi catalog --release 27A
   ```
   This adds a `27A` tab to `FBDI_Master_Catalog.xlsx` and recomputes the Drift tab against 26B.

5. Update the Module column in the mapping spreadsheet:
   ```
   python -m fbdi populate-module --new 27A --old 26B
   ```
   This reads `baselines/<ver>/file_modules.json` (NEW wins, OLD as fallback) and surgically updates column F of `FBDI_to_ApplaudTables_Mapping.xlsx`. Formatting and freeze-panes are preserved.

6. Generate the compliance report:
   ```
   python -m fbdi report --old 26B --new 27A
   ```
   You get an HTML and a PDF in the repo root.

7. Verify detection health:
   ```
   python -m fbdi diagnose --old baselines/26B/originals --new baselines/27A/originals
   ```
   Check the `FILE_ERROR` count. If it jumped vs. the prior release, something Oracle shipped is corrupt or uses a format the engine doesn't handle. Investigate before handing off the report.

8. If Oracle changed the structure of a template's header rows and detection breaks, the fix lives in `fbdi/detect_header.py`. Adjust the scoring weights or thresholds, then add a test in `tests/test_detect_header.py` that pins the new behavior. Do not special-case a filename. That's the architectural decision the engine was built to avoid. The whole point of dynamic detection is that filenames shouldn't matter.

9. If a previously-working tab now produces `NO_HEADER`, `detect_header.py` is the first place to look. If it's failing on the comparison side (a header is detected but fields are wrong), check `compare.py`. If the catalog is missing fields, check `catalog.py`'s `_extract_rich` vs. `_extract_thin` dispatch. The distinction matters when Oracle's header row is UPPER_SNAKE_CASE vs. mixed-case labels.

## Design docs and how we work

Specs and plans follow a consistent pattern. Design docs live at `docs/superpowers/specs/<date>-<feature>.md`. Implementation plans live at `docs/superpowers/plans/<date>-<feature>.md`. Specs are produced with the `superpowers:brainstorming` skill, plans come from `superpowers:writing-plans`. Both are committed to git so there's an auditable trail of why decisions were made.

Working style here: read the plan end-to-end before touching code, execute it start to finish, push directly to master. No PR required for solo work. CodeRabbit is wired up for PR review but it's optional on master pushes.

`CLAUDE.md` has a "Resolved Hazards" section that lists bugs that used to exist (the iter_rows performance fix, the subprocess pipe deadlock, the type-parser warning collapse from 463 to 9, the JET tree-view scraper race). Before you file something as a new issue, check that section first. If it's already there, the fix is already in the codebase and you're most likely looking at a different root cause.

For cross-release historical context, the `reference/` directory at repo root is a read-only archive (the legacy VBA comparison macro, the legacy clearing macro, a sample VBA output report, and Dan's original Selenium downloader). Don't modify anything in there.

## Known hazards and gotchas

**Phantom `max_column=16384`.** Some xlsm files report 16,384 columns because of corrupt formatting metadata. Both `compare.py` and `detect_header.py` cap column scanning at 500. Don't lift that cap without understanding why it's there. Scanning all 16,384 phantom columns on a wide sheet takes forever.

**Corrupt XML in some xlsm files.** `openpyxl` throws `zipfile.BadZipFile` on a handful of Oracle templates. The engine catches the exception and logs the file as a `FILE_ERROR`. The issue shows up in the catalog's Issues tab and in the diagnose report. `verify_run.py` flags a regression when the FILE_ERROR count jumps between releases. That's the signal to go investigate what Oracle shipped.

**5 MB cap in `diagnose` and `build_mapping`.** These two modules load workbooks in full (non-read_only) mode, which holds the whole file in memory. The cap is in `fbdi/config.py` as `MAX_FILE_SIZE_BYTES`. The comparison engine has no such limit because it streams via `iter_rows` in read_only mode. If you're extending `diagnose` or `build_mapping`, keep the cap. Don't silently remove it and then wonder why memory spikes.

**Legacy VBA-generated comparison reports.** The legacy VBA macro produced comparison reports with a corrupt stylesheet (this is what `Comparison_Report_25D_26A.xlsx` looks like, for example). Standard `openpyxl.load_workbook()` throws an exception on these. If you need to read one programmatically, use `read_only=True` or `data_only=True` with exception handling around the load. Python-generated reports don't have this issue.

**Test-data UPPER_SNAKE_CASE trap.** Covered in the testing section, but worth repeating because it bites people: any cell value matching `^[A-Z][A-Z0-9_]+$` looks like a header to the scorer. Strings like `"CREATE"`, `"NULL"`, `"V1"`, `"DR_ECO_1"` all trigger false-positive header detection in tests. Use mixed-case or lowercase in data rows.

**`RapidImplementationForCashManagement.xlsm` is not auto-downloadable.** If you're testing the download pipeline against a fresh `baselines/` directory, this file will be missing after the Selenium run finishes. The downloader warns about it but doesn't stop. The comparison and catalog engines pick it up automatically once you place it, and skip it silently if it's absent.

**JET `<oj-tree-view>` race in the scraper.** Oracle's docs put the table of contents inside an `<oj-tree-view>` under `#navigationDrawer`. The drawer container appears in the DOM before the tree-view's `<li role="treeitem">` children populate. Without a wait for at least one treeitem, the scraper finds an empty list and silently skips the URL. No error, no SKIP log. The page just immediately reports "Completed" with zero downloads. Fixed in commit 82cd568. Keep the wait when refactoring the scraper.

**PDF rendering needs MSYS2 mingw64 GTK on Windows.** `fbdi/report.py` uses weasyprint, which depends on libgobject, libpango, and libcairo. The standalone GtkD installer ships Pango 1.43 and breaks on weasyprint ≥53. MSYS2 mingw64 ships Pango 1.56+ and works. Install MSYS2 and run `pacman -S mingw-w64-x86_64-pango mingw-w64-x86_64-gtk3 mingw-w64-x86_64-pkg-config`. The probe order in `_GTK_WINDOWS_BIN_CANDIDATES` checks MSYS2 first; keep that ordering. Don't try to lower the `weasyprint>=62.0` pin to work around an older GTK. weasyprint <53 lacks flexbox and grid layout entirely, and the report cover collapses to white-on-white.

## Where to ask for help

Read `CLAUDE.md` and the relevant skill's `SKILL.md` first. Most questions about current state are answered there. For design context on a specific feature, check `docs/superpowers/specs/`. The spec will explain the tradeoffs that led to the current implementation. For how a feature was built step by step, look in `docs/superpowers/plans/`. Past audit notes and one-off findings are in `docs/archive/`. The repo is meant to be self-serviceable, so DMing the previous owner is a last resort, not a first step.
