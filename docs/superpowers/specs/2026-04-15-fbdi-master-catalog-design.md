# FBDI Master Catalog

**Date:** 2026-04-15
**Status:** Approved

---

## Problem

The existing comparison pipeline produces a position-aligned diff of header names between two releases (`Comparison_Report_26A_26B.xlsx`). It does not produce a standalone catalog of every FBDI file, tab, and column in each release with data type, length, and required information.

For the downstream Applaud MDB comparison (via `applaud-mcp`), we need a per-release snapshot that captures — in template order — the technical column name, user-friendly label, data type, length, and required flag for every column in every tab of every FBDI template Oracle publishes in that release. The snapshots must stay in sync across quarterly releases so drift is visible and Applaud integrations don't silently fall out of alignment.

## Goals

1. A single master workbook `FBDI_Master_Catalog.xlsx` containing one flat snapshot tab per Oracle release (`26A`, `26B`, …), an `Issues` tab consolidating coverage gaps, and a `Drift` tab diffing the two most-recent releases.
2. Honest representation of both rich templates (inline type/length metadata available) and thin templates (label-only) — the catalog reports what the templates contain without fabricating missing data.
3. Zero impact on the working comparison pipeline. The existing `Comparison_Report_*.xlsx` flow keeps generating exactly as today and remains the artifact validated against the manual VBA process.
4. Per-release regeneration is idempotent: `python -m fbdi catalog --release 26B` rewrites the `26B` tab and regenerates `Issues`/`Drift` deterministically without touching other release tabs.

## Non-Goals (v1)

1. Repairing the 6 FILE_ERROR templates with corrupt stylesheets (`ImportAssignmentLaborSchedules`, `ImportPayrollCosts`, `ProjectResourceAssignmentImportTemplate`, `ResourceBreakdownStructureImportTemplate`, `SupplierSiteImportTemplate`, `SusActivityImportTemplate`). These appear in `Issues` with no catalog rows. Separate workstream.
2. Capturing the `description` field from rich templates. Adds file bloat with limited value given that opening the template gives you the description.
3. Module/area classification (e.g., `Receivables`, `Payables`).
4. CSV companion exports alongside the master workbook for `applaud-mcp` ingestion. Can follow once `applaud-mcp`'s preferred ingestion format is known.
5. Auto-running catalog generation inside `python -m fbdi compare`. Keep the commands separate; chain later if that proves useful.
6. The Applaud MDB comparison itself. That's the downstream step this catalog feeds.

## Findings that grounded the design

Sampling 26B templates revealed the structural split that drives the design:

- **Rich tabs (349 of 664 data tabs in 26B)** — templates stack labeled metadata rows above example data. Each metadata row's column A tags the row's role. Observed values:

  | Column A value | Row role |
  |---|---|
  | `Name` | user-friendly labels |
  | `Description` | long descriptions |
  | `Data Type` | e.g., `VARCHAR2(5 CHAR)`, `NUMBER(18)`, `Varchar2(250)` |
  | `Required or Optional` (sometimes `\ufeffRequired or Optional`) | `Required` / `Optional` |
  | `Column name of the Table <NAME>` | UPPER_SNAKE_CASE technical column names |
  | `Reserved for Future Use` | skip |

  Header row position varies across rich templates (observed at R5 and R8). The current engine's `detect_header_row` correctly locates the technical-name row as "Tier 1".

- **Thin tabs (303 of 664)** — only a title, a `* Required` legend, and a single header row of user-friendly labels with `*` prefixes for required fields (e.g., `*Source Budget Type`). No inline data type, length, description, or technical name.

- **12 NO_HEADER tabs** in 26B when scanning with a conservative skip list. `Phase 3` cleanup reported 0 against the production `SKIP_TABS`; new tabs in future releases could regress.

- **6 FILE_ERROR files** in 26B fail `openpyxl.load_workbook` in both `read_only=True` and full modes with `could not read stylesheet from base`. Unchanged by the Phase 2/3 fixes.

- `AP5TBFLD.CSV` at the repo root is not a usable data-type source: its `DB_Name` values (prefixed `T_*`) do not match FBDI tab names. Out of scope.

## Design Decisions

### Master workbook layout

```
FBDI_Master_Catalog.xlsx
├── 26A        (per-release flat snapshot)
├── 26B        (per-release flat snapshot)
├── Issues     (consolidated coverage gaps across all releases)
└── Drift      (position-aligned diff between the two most-recent release tabs)
```

Per-release tab schema — one row per `(file, tab, position)`:

| column | notes |
|---|---|
| `release` | e.g., `26A`. Redundant with tab name, kept for flat-export and filtering. |
| `file_name` | FBDI `.xlsm` stem, e.g., `AttachmentsImportTemplate`. |
| `tab_name` | sheet name as it appears in the template. |
| `position` | 1-based column index in the template. |
| `column_label` | normalized user-friendly label. |
| `column_technical` | UPPER_SNAKE_CASE name. Blank on thin tabs. |
| `data_type` | normalized uppercase, e.g., `VARCHAR2`, `NUMBER`, `DATE`. Blank on thin tabs or parse failures. |
| `length` | integer max length / precision. Blank when absent. |
| `scale` | integer scale for `NUMBER(p,s)`. Blank otherwise. |
| `data_type_raw` | original string exactly as read (e.g., `Varchar2(250)`). Blank on thin tabs. |
| `required` | `TRUE` / `FALSE`. Universal: from R4 `Required or Optional` on rich tabs, from `*` prefix on thin tabs. |

`Issues` tab schema: `release | file | tab | issue_type | detail`.

`Drift` tab schema: `file | tab | position | col_label_<old> | col_label_<new> | col_technical_<old> | col_technical_<new> | data_type_<old> | data_type_<new> | length_<old> | length_<new> | required_<old> | required_<new> | change_type` where `<old>` and `<new>` are the two most-recent release names in sort order (`26A` < `26B` < `26C` < `27A`).

`change_type` classification, evaluated in this order (first match wins, except `MULTI`):

| `change_type` | Definition |
|---|---|
| `ADDED` | position exists in new release only (old tab has no row at this `(file, tab, position)`). |
| `REMOVED` | position exists in old release only. |
| `RENAMED` | both releases have the row; `column_label` OR `column_technical` differ; type/length/required unchanged. |
| `TYPE_CHANGED` | `data_type` differs; nothing else. |
| `LENGTH_CHANGED` | `length` OR `scale` differs; nothing else. |
| `REQUIRED_CHANGED` | `required` differs; nothing else. |
| `MULTI` | two or more of name/type/length/required differ in the same row. |

Only changed rows emitted. Rows unchanged between releases are absent from `Drift`.

### Label normalization

`normalize_label(s)` strips every character that is not alphanumeric, underscore, or whitespace, then collapses runs of whitespace, then trims. "Alphanumeric" here means Python's Unicode-aware `str.isalnum()` — non-ASCII letters (e.g., accented characters) pass through; only punctuation and symbols are stripped. Applied only to `column_label`. `column_technical` is already `UPPER_SNAKE_CASE` by construction and untouched.

Examples:
- `*Source Budget Type` → `Source Budget Type`
- `$Weird, Chars!` → `Weird Chars`
- `COLUMN_NAME` → `COLUMN_NAME` (underscore preserved)
- `  *Foo  Bar  ` → `Foo Bar`

### Type parsing

`fbdi/type_parser.py` exposes `parse_data_type(raw: str) -> ParsedType` returning `(data_type, length, scale, parse_warning)`. Handles the observed forms:

| Input | `data_type` | `length` | `scale` |
|---|---|---|---|
| `VARCHAR2(5 CHAR)` | `VARCHAR2` | `5` | blank |
| `VARCHAR2(2048 CHAR)` | `VARCHAR2` | `2048` | blank |
| `VARCHAR2(80)` | `VARCHAR2` | `80` | blank |
| `Varchar2(250)` | `VARCHAR2` | `250` | blank |
| `NUMBER(18)` | `NUMBER` | `18` | blank |
| `NUMBER(18,4)` | `NUMBER` | `18` | `4` |
| `DATE` | `DATE` | blank | blank |
| `CLOB` | `CLOB` | blank | blank |

Unrecognized strings return `parse_warning=True`; caller preserves the original in `data_type_raw` and leaves the parsed fields blank. Parse warnings surface in `Issues` with `issue_type=TYPE_PARSE_WARNING`.

### Metadata-row detection for rich tabs

`detect_header_row` already returns the UPPER_SNAKE_CASE row (Tier 1) for rich tabs. The catalog extends this by scanning column A of rows `1..header_row` and keyword-matching (case-insensitive, BOM-stripped, prefix-matched for `Column name of the Table`) to locate the label, data_type, and required rows. Absent role rows leave the corresponding field blank for the whole tab — rows still emit.

For Tier 2 (thin) tabs, the header row value becomes `column_label` (after normalization), `required` is inferred from the `*` prefix, and all other metadata fields are blank.

**Risk — unknown column-A variants in 26A or older releases.** Mitigation: during development, run the detector against 26A and log any unrecognized column-A strings on rich tabs (rows with ≥ `MIN_CELLS` non-empty cells whose column A isn't matched). Expand the keyword list from evidence, not speculation.

### Subprocess isolation & column cap

Each `.xlsm` is processed in its own `multiprocessing.Process` with a 120s timeout, mirroring `compare.py`'s pattern. This isolates `openpyxl` resource accumulation and caps any single file's impact on the batch. Column scanning is capped at 500 to avoid phantom `max_column=16384` behavior (same cap as `compare.py`).

The shared logic is `detect_header.detect_header_row` and `utils.match_fbdi_files`. No refactor to extract further shared code in v1 — duplication is small, the compare pipeline is the VBA-validated reference, and keeping it pure is more valuable than DRY.

### Idempotent per-release regeneration

`python -m fbdi catalog --release 26B`:

1. Resolve `baselines/26B/originals/` (via the same `_resolve_dir` pattern as `compare`).
2. Load the existing master workbook, or create a new one.
3. Spawn a subprocess per file; collect rows and issues.
4. Drop the existing `26B` tab (if any); write the new `26B` tab.
5. Rebuild the `Issues` tab: preserve issues from untouched release tabs, replace all `release=26B` issues with the new set.
6. Rebuild the `Drift` tab: identify the two most-recent release tabs by sort order; compute position-aligned diff; emit only changed rows.
7. Save atomically: write to `<master>.tmp`, then rename.

Other release tabs (e.g., `26A`) are not touched. Re-running the command is safe and produces content-identical output: row-by-row equality of every tab's data, though `.xlsx` serialization may differ at the byte level due to openpyxl's internal timestamps and ordering.

### Master workbook location

Repo root: `FBDI_Master_Catalog.xlsx`. Gitignored, consistent with `Comparison_Report_*.xlsx` and `Diagnostic_Report_*.xlsx`. The catalog is a build output, not source.

## Architecture

```
fbdi/
├── catalog.py                 NEW — generation, subprocess worker, writer
├── type_parser.py             NEW — type string parsing
├── catalog_normalize.py       NEW — label normalization helper
├── cli.py                     MODIFIED — add `catalog` subcommand
├── __main__.py                unchanged (routes through cli.py)
├── compare.py                 unchanged
├── clear.py                   unchanged
├── diagnose.py                unchanged
├── detect_header.py           unchanged (catalog imports from here)
├── config.py                  unchanged (catalog uses existing SKIP_TABS)
└── utils.py                   unchanged (catalog imports match_fbdi_files pattern)

tests/
├── test_catalog.py            NEW
├── test_type_parser.py        NEW
└── test_catalog_normalize.py  NEW
```

### Key functions in `fbdi/catalog.py`

- `extract_tab_rows(ws, file_stem, tab_name, release) -> tuple[list[CatalogRow], list[IssueRow]]` — per-tab extraction. Uses `detect_header_row`; for Tier 1, scans column A of rows above the header to find metadata rows; for Tier 2, uses the header row as labels only.
- `extract_file(path, release) -> tuple[list[CatalogRow], list[IssueRow]]` — opens workbook, iterates tabs, delegates to `extract_tab_rows`. Converts load failures into `FILE_ERROR` issues.
- `_catalog_worker(path, release, queue)` — subprocess entry; mirrors `_compare_worker`.
- `generate_catalog(release, baselines_dir, master_path, timeout=120)` — orchestrates subprocess fan-out, aggregates results, rewrites release tab, rebuilds `Issues` and `Drift`, saves.
- `_compute_drift(workbook, release_old, release_new) -> list[DriftRow]` — position-aligned diff between two release tabs, emits only changed rows with classified `change_type`.

## Error handling

Every exceptional path that prevents a tab from contributing rows records an entry in `Issues`. The catalog does not silently drop data.

| `issue_type` | Trigger | `detail` format |
|---|---|---|
| `FILE_ERROR` | `load_workbook` raises | `<ExceptionClass>: <message>` |
| `TIMEOUT` | subprocess exceeds 120s | `exceeded 120s` |
| `SUBPROCESS_FAILED` | non-zero exit, no queue result | `exit code <N>` |
| `NO_HEADER` | `detect_header_row` returns `None` | `best candidate row <R> score <S>` if available |
| `TYPE_PARSE_WARNING` | parser couldn't decode a `data_type_raw` | the raw string |

Skipped tabs (`SKIP_TABS` from `config.py` — Instructions, Messages, reference, LOV, Lookups, Revision History, Dependencies) are silently skipped. Expected, not issues.

Every tab that did produce catalog rows still appears in the catalog even if it also generated `TYPE_PARSE_WARNING` entries. Issues are additive visibility, not substitutes.

## Testing

Synthetic `openpyxl.Workbook()` fixtures built inline per test. Same pattern as the existing 54-test suite — no shared fixtures.

**`tests/test_catalog.py`:**
- Rich template happy path (R1=Name / R2=Description / R3=Data Type / R4=Required or Optional / R5=Technical) — assert all fields populate, positions correct.
- Thin template happy path — asterisk-prefixed labels only; assert label normalized, `required` from `*`, technical/type/length blank.
- Rich template missing a role row — `Data Type` row absent; assert type fields blank, others populated.
- Column-A variants — BOM prefix, case differences, leading/trailing whitespace all match.
- End-to-end on two tiny 2-file "releases" (test releases `TESTA` and `TESTB`) under `tests/fixtures/catalog/` — run `generate_catalog`, assert workbook has `TESTA`, `TESTB`, `Issues`, `Drift` tabs; assert `Drift` contains only changed positions with correct `change_type` classification including at least one case each of `ADDED`, `RENAMED`, `TYPE_CHANGED`, `LENGTH_CHANGED`, `REQUIRED_CHANGED`, and `MULTI`; assert `Issues` records an intentionally-broken fixture.
- Subprocess robustness — a worker that raises produces a `SUBPROCESS_FAILED` issue, not a crash.
- Idempotence — running `generate_catalog` twice produces content-identical workbooks (assert by reading both and comparing row data tab-by-tab, not by byte-compare).

**`tests/test_type_parser.py`:**
- One test per observed format (VARCHAR2 variants with/without `CHAR`, with/without case, NUMBER with and without scale, DATE, CLOB, BLOB).
- One test for a garbage string → `parse_warning=True` and parsed fields blank.

**`tests/test_catalog_normalize.py`:**
- `*Source Budget Type` → `Source Budget Type`
- `$Weird, Chars!` → `Weird Chars`
- `COLUMN_NAME` → `COLUMN_NAME`
- `  *Foo  Bar  ` → `Foo Bar`
- Unicode non-ASCII handling (pass through alphanumerics, drop punctuation).

## Validation plan

- After v1 lands, Brad runs the manual VBA process for 26A → 26B and compares its output against the existing `Comparison_Report_26A_26B.xlsx`. This validates the comparison engine (unchanged by this work) against VBA.
- The catalog itself has no VBA equivalent. Its validation is by spot-check: open `FBDI_Master_Catalog.xlsx`, verify representative rich tabs (e.g., `AttachmentsImportTemplate / Attachment Details`) have populated type/length/required; verify representative thin tabs (e.g., `BudgetImportTemplate / XCC_BUDGET_INTERFACE`) have labels only with blanks in other fields; verify the `Drift` tab reproduces the changes seen in `Comparison_Report_26A_26B.xlsx` and additionally surfaces any type/length drift the existing report can't see.

## Open questions

None blocking v1. Flagged for future discussion:

1. **FILE_ERROR stylesheet repair.** 6 templates are unreadable. Could be addressed by unzipping the `.xlsm`, dropping the corrupt `styles.xml`, re-zipping, and re-opening. Separate workstream.
2. **CSV companion exports.** Add once `applaud-mcp`'s preferred ingestion format is known.
3. **Shared `template_reader.py` refactor.** If duplication between `compare.py` and `catalog.py` becomes painful, extract shared reading logic. Not needed today.
