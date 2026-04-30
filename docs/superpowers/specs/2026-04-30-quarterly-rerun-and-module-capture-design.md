# Quarterly Rerun + Module Capture — Design Spec

**Date:** 2026-04-30
**Author:** Brad Hagstrom (with Claude Code)
**Status:** Approved for implementation planning

## Purpose

Two coupled goals:

1. **Validate the FBDI pipeline end-to-end** by wiping the 26A and 26B baselines and rerunning the full `download → clear → compare → catalog → mapping` flow via the `fbdi-compare-release` skill. This exercises the recently-committed `detect_header.py` fix (sparse-data-row filtering via `TIER1_MIN_FILL` plus the `non_empty` Tier-2 tiebreak) against a clean re-download, and it tests the skill itself end-to-end.
2. **Add a Module column to the mapping workflow.** Capture each FBDI file's source Oracle module (Financials, Procurement, Supply Chain & Manufacturing, Project Management) at download time, persist it as a per-release JSON, and surgically populate the existing Module column in `FBDI_to_ApplaudTables_Mapping.xlsx` without disturbing Brad's 639 hand-edited rows.

## Background

`fbdi/build_mapping.py` is a legacy generator wired to stale paths (`25d`, `26a`) and an obsolete output filename (`fbdi_applaud_mapping.xlsx`). The current source-of-truth mapping spreadsheet is `FBDI_to_ApplaudTables_Mapping.xlsx` with sheets `FBDI Mapping` and `Applaud Tables Reference`, columns `FBDI Template, FBDI Tab, Applaud Table, Prefix, Status, Module, In Base System?`. Brad has populated 639 rows manually; 0 of them have Module filled in. We cannot regenerate this file from scratch — we need a surgical updater that touches column F only.

The Selenium downloader (`tools/download_and_clear.py`) already iterates over four module-scoped Oracle docs URLs, so the per-file module is known *at download time*. We just don't currently capture it.

The `fbdi-compare-release` skill ships an 8-stage orchestrator with 6 HITL checkpoints. We need to wedge in a new Stage 6.5 between catalog and summary to run the module-column updater.

## Decisions (closed — do not re-litigate)

| Decision | Choice | Rationale |
|---|---|---|
| Module data source | Capture during download from base URL | Authoritative, automatic, future-proof |
| Wipe scope | Both 26A and 26B | True end-to-end test; exercises both module-capture passes; Brad explicitly asked for full wipe |
| Validation rigor | TDD on new code only; soft macro-signal validation on the rerun | Header-detection fix means byte-regression isn't viable; macro signals are sufficient |
| Mapping update strategy | Surgical updater (column F only) | Brad's 639 hand edits must survive |
| Module taxonomy | 4 modules: `Financials`, `Procurement`, `Supply Chain & Manufacturing`, `Project Management` | Matches `MODULE_URL_TEMPLATES` and existing `KNOWN_MAPPINGS` style (`&` not "and") |
| FSM file (`RapidImplementationForCashManagement.xlsm`) | Hardcoded to `"Financials"` | Banks/branches/accounts is a Financials concept; not auto-downloadable |
| Skill extension | New Stage 6.5 + HITL #7 (backup prompt) | Keeps the existing 8-stage flow intact |

## Architecture

Six independently-testable pieces:

### 1. Module URL classifier — `tools/download_and_clear.py`

Pure function `module_from_base_url(url: str) -> str`. Maps the four `MODULE_URL_TEMPLATES` URL patterns to canonical module names via a `URL_TO_MODULE` dict keyed on URL slug. Raises `ValueError` for unknown URLs. The lookup uses `f"/saas/{slug}/"` to avoid false positives.

### 2. Per-release module capture — `tools/download_and_clear.py`

Inside the existing `for base_url in base_urls:` loop, accumulate `{filename: module}`. After the loop, add the hardcoded entry for `RapidImplementationForCashManagement.xlsm → "Financials"` and write `baselines/<ver>/file_modules.json` (sorted, indented). The JSON is only written on a successful download pass; `--clear-only` and `--skip-clear` do not touch it.

If `module_from_base_url()` raises (Oracle changed a URL pattern), catch the `ValueError`, log it, default the module to `"Unknown"`, and continue. The post-download summary surfaces the count of `"Unknown"` entries.

### 3. Module column updater — `fbdi/populate_module.py` (new)

Function: `populate_module_column(mapping_path, new_modules, old_modules) -> dict[str, int]`. Opens the mapping spreadsheet with `load_workbook()` (full mode, not read_only), iterates rows on the `FBDI Mapping` sheet, looks up each row's FBDI Template (col A) against the merged `{**old_modules, **new_modules}` dict (NEW wins), writes column F. Saves in place. All other cells, formatting, merged cells, formulas, validations, and freeze-panes are preserved.

Lookup tolerates `.xlsm` suffix mismatches: both sides are normalized to the stem before comparison.

CLI wrapper: `python -m fbdi populate-module --new 26B --old 26A [--mapping FBDI_to_ApplaudTables_Mapping.xlsx]`. Default `--mapping` is the working file at repo root.

Returns / prints a summary: `{populated: N, blank: M, overwritten: K}`.

### 4. Skill Stage 6.5 — `.claude/skills/fbdi-compare-release/SKILL.md`

New stage between Stage 6 (catalog) and Stage 7 (summary). Includes HITL #7 (backup prompt). Skips with a notice (not a failure) if the mapping file is absent. Captures the JSON summary for Stage 7.

If a backup file with the same name already exists, append a timestamp: `FBDI_to_ApplaudTables_Mapping.bak.<YYYYMMDD-HHMMSS>.xlsx`.

### 5. Tests (TDD-first)

- `tests/test_module_classifier.py` — 5 cases covering each `URL_TO_MODULE` branch + `ValueError` for unknown URL.
- `tests/test_populate_module.py` — ~7 cases: happy path, OLD-fallback, blank-when-missing, other-columns-preserved, idempotency, file-locked PermissionError, .xlsm suffix normalization.
- `tests/test_verify_rerun.py` — ~5 cases covering each threshold branch in the post-run validator.

All tests build synthetic xlsx fixtures inline (matching existing test style per `CLAUDE.md`). New test count: ~17. Existing 255-test suite must stay green.

### 6. Post-rerun validator — `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py`

Reads the new comparison report, catalog, and mapping file. Computes macro signals and compares to thresholds. Emits structured JSON. Exits 0 on all-green, 1 with warnings if any threshold breached. Thresholds are configurable constants at the top of the script.

| Signal | Threshold | Severity |
|---|---|---|
| Catalog row count delta | ±5% vs baseline | Warning |
| Compare report changes | 706 ±50 | Warning |
| Issues count | ≤9 | Warning |
| NO_HEADER count | == 0 | **Strong** warning (likely regression) |
| Module column populated | ≥95% of rows with a non-blank FBDI Template (col A) | Warning |

The skill surfaces warnings in the Stage 7 summary but does not fail.

## Data flow

```
user invocation → skill triggers
  Stage 1  preflight
  Stage 2  resolve OLD=26A, NEW=26B
  Stage 3  HITL #1 (26A missing) → download both
             ├─ download 26A: scrape, capture module per file, write file_modules.json
             └─ download 26B: same
           HITL #2 if FSM file missing
  Stage 4  smart-clear both
  Stage 5  compare → Comparison_Report_26A_26B.xlsx
  Stage 6  catalog → FBDI_Master_Catalog.xlsx
  Stage 6.5  HITL #7 (backup mapping?) → populate-module → in-place update
  Stage 7  summary (includes module stats)
  Stage 8  verify_run.py + verify_rerun.py → warnings if any
user reviews → push to master
```

### Artifacts created/modified

| Artifact | Status |
|---|---|
| `baselines/26A/originals/*.xlsm` | wiped + re-downloaded |
| `baselines/26A/blanks/*.xlsm` | wiped + re-cleared |
| `baselines/26A/file_modules.json` | **new**, ~212 entries |
| `baselines/26B/originals/*.xlsm` | wiped + re-downloaded |
| `baselines/26B/blanks/*.xlsm` | wiped + re-cleared |
| `baselines/26B/file_modules.json` | **new**, ~213 entries |
| `Comparison_Report_26A_26B.xlsx` | regenerated (header-fix differences expected) |
| `FBDI_Master_Catalog.xlsx` | regenerated |
| `FBDI_to_ApplaudTables_Mapping.xlsx` | Module column populated; all other cells preserved |
| `FBDI_to_ApplaudTables_Mapping.bak.xlsx` | **new** — pre-rerun snapshot |

`file_modules.json` files are gitignored along with the rest of `baselines/` per existing convention.

## Error handling

### Download / scrape (Stage 3)

Existing skill HITLs (#5, #6, #1, #2, #3) cover this. The new module-capture code piggybacks on `verify_download.py`: if a download fails, the missing-files check fires before we ever read `file_modules.json`. Partial JSON cannot reach the populate step.

Edge case: if `module_from_base_url()` raises `ValueError` (Oracle changed a URL pattern), default to `"Unknown"` and surface the count.

### Populate-module (Stage 6.5)

| Failure | Response |
|---|---|
| `file_modules.json` missing for one or both releases | Exit 2, instruct user to re-run download |
| Mapping spreadsheet missing | Skip with notice, continue |
| Mapping spreadsheet has 0 rows | Skip with warning |
| FBDI Template name in mapping doesn't match any JSON entry | Leave Module blank, count in summary |
| openpyxl raises `PermissionError` (file locked by Excel) | Exit 3, instruct user to close Excel |
| Backup destination already exists | Timestamp-suffix the new backup |

### Validation (Stage 8)

`verify_rerun.py` reports regressions but never blocks. Brad decides whether to investigate or accept. The skill does not auto-rollback. The `.bak.xlsx` covers the mapping spreadsheet; the comparison report and catalog are git-recoverable.

## Build sequence

**Phase 1 — code (local, ~2-3 hours, no downloads):**

1. Write all tests (TDD red).
2. Implement `module_from_base_url()` + `URL_TO_MODULE` dict.
3. Implement download-loop capture + JSON writing.
4. Implement `fbdi/populate_module.py` + CLI subcommand.
5. Implement `verify_rerun.py`.
6. Update `SKILL.md` with Stage 6.5 + HITL #7.
7. Full pytest pass → all ~272 tests green.
8. Commit Phase 1.

**Phase 2 — rerun (~2-3 hours wall clock, mostly unattended):**

9. Wipe `baselines/26A/` and `baselines/26B/` (with confirmation).
10. Invoke skill: `compare 26A to 26B`.
11. Walk through 8 stages + Stage 6.5 with HITL prompts.
12. Validate: eyeball Module column, spot-check previously-misdetected tabs in the catalog, review `verify_rerun.py` output.
13. Commit Phase 2.
14. Push to master.

## Out of scope

- **No changes to `compare.py`, `catalog.py`, `diagnose.py`, `clear.py`, `detect_header.py`.** Header-fix is already committed.
- **No reconciliation of `build_mapping.py`'s stale paths/sheet structure.** Separate cleanup, not blocking.
- **No changes to existing HITL #1–#6.** Add HITL #7 only.
- **No `/requesting-code-review` invocation.** Brad picked option (c) in brainstorming, which excludes it.
- **No backup of `Comparison_Report_*.xlsx` or `FBDI_Master_Catalog.xlsx`.** Both are git-tracked.

## Success criteria

- All 17 new tests pass; existing 255 tests still pass.
- Skill completes all 8 stages + Stage 6.5 without unhandled errors.
- `baselines/26A/file_modules.json` and `baselines/26B/file_modules.json` exist with ≥210 entries each.
- `FBDI_to_ApplaudTables_Mapping.xlsx` Module column ≥95% populated (denominator: rows where col A is non-blank — currently 639); column A and other manually-edited columns unchanged from pre-run.
- `verify_rerun.py` macro signals all green: NO_HEADER == 0, Issues ≤ 9, catalog row count within ±5% of the existing (post-fix) catalog. The row-count check doubles as a download-consistency sanity check — Oracle hasn't changed 26A or 26B between yesterday's catalog regeneration and now, so a near-identical row count confirms the fresh download produced consistent data.

## Open questions / risks

- **The `detect_header.py` fix is not re-validated by this rerun.** It was already validated implicitly by the catalog regeneration on 2026-04-29 (the current `FBDI_Master_Catalog.xlsx` at HEAD reflects the post-fix logic). Diffing the rerun's catalog against HEAD's catalog tests download consistency, not fix correctness. If we ever need to re-prove the fix, that requires regenerating a pre-fix catalog from an earlier commit — out of scope here.
- **Wall clock for the rerun is ~2-3 hours.** If something fails partway through Stage 3 downloads, partial baselines remain. The skill is designed to resume from the next stage on re-invocation.
- **Oracle URL stability.** If Oracle restructures `MODULE_URL_TEMPLATES` between now and next quarter, `module_from_base_url()` will surface "Unknown" entries. Acceptable for this rerun; future-proofing is out of scope.
