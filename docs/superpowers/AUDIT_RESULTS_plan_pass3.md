# Audit Results — Pass 3 (Spot-Check of Revised Plan)

**Audience:** Claude Code, before implementing the revised `2026-06-02-applaud-compliance-audit.md`.
**Source:** Third-pass spot-check, scoped per pass-2 §5 to the two blockers + the §3.2 integration fix. Re-verified live against `ORACLE_MASTER/AP0STE.mdb` via `applaud-mcp`.
**Date:** 2026-06-02
**Verdict:** **PASS — cleared to implement.** Both release-blocking defects from pass 2 are correctly fixed in the plan text (not merely asserted in the header), the §3.2 normalization is wired through every matching dimension, and the smaller items are addressed. One residual dependency remains unverifiable from the audit side and must be confirmed by the implementer on first run — called out in §3. No further audit round is required before implementation; do the §3 confirmation as the first implementation step.

---

## 1. Blocker fixes — verified in the task code

### Blocker 1 (Dim 1 sizing source) — FIXED, correctly, on both sides
- **Assembly side (Task 2):** `build_table` now takes `dd_by_ddid` and sources `data_type/size/dec_places` from the DataDictionary slice, keeping only `Row`/`DDID`/`ODBCName` from `DatabaseDetail`. Matches the live reality I re-confirmed: `DatabaseDetail` for `T_BANKS_BRANCHES` returns blank `DataType`/`Size`/`DecPlaces`/`ODBCName` on all columns; the type lives on `DataDictionary` (`T32BANK_NAME` → `X`/100).
- **Extraction side (Task 16 doc):** Step 2d now pulls `SELECT Name,DataType,Size,DecPlaces FROM DataDictionary WHERE Name LIKE 'P%'` with the COUNT guard, and Step 2b explicitly warns not to read type from `DatabaseDetail`. The closing note is corrected to "DataDictionary IS pulled; DatabaseDetail has no type data." The earlier inverted instruction is gone.
- **Regression test (Task 7, `test_actual_shape_reflects_datadictionary_not_blank_databasedetail`):** feeds a blank-`DatabaseDetail` row + populated DD and asserts `actual_shape` is `("char", 100, None)`, not `("", 0, None)`. This is exactly the test whose absence let the bug hide last round. Good.

### Blocker 2 (`@`-audit field exclusion) — FIXED, correctly, everywhere it matters
- `is_audit_field()` added (Task 2) and applied in **both** `build_file_fields` and `build_table`, so `@`-fields never enter the snapshot's matching universe.
- LCP prefix fallback (Task 3) now filters `@`-fields before computing the common prefix (`business = [d for d in column_ddids if not d.lstrip().startswith("@")]`), with a dedicated test. This closes the skew I flagged (a mix of `@T32…`/`T32…` would otherwise collapse the LCP).
- Extraction `LIKE 'P%'` DD pull also naturally excludes `@`-fields — defense in depth.

## 2. §3.2 (Oracle name ↔ Applaud bare name) — promoted to required and wired through ALL matching dims
This was the residual risk from pass 2, and the revision handles it correctly:
- `oracle_match_key(of)` (Task 7) returns `technical.upper()` when present, else `_label_to_technical(label).upper()`.
- It is now used in **every** matching dimension, which I checked individually:
  - Dim 1 (`run_audit` builds `oracle_by_bare` via `oracle_match_key`, line ~1847).
  - Dim 2/3 (`check_file_coverage` uses `oracle_match_key` for both `oracle_order` and the PRESENCE loop, lines ~1092/1101 — replacing the old raw `(technical or label)`).
  - Dim 4 (`check_table_coverage` uses `oracle_match_key(of)`, line ~1250 — same replacement).
- Integration test (Task 15) runs `run_audit` on a thin `Bank Account` tab where every `AlignedField.technical=None` (label-only) and asserts zero spurious `2-IF` PRESENCE findings — i.e. it exercises the exact failure mode (label-only Oracle fields) end-to-end, not just the helper in isolation.

This is the right shape. The fix is no longer a dangling helper; it's the single chokepoint all four dimensions share.

## 3. Residual dependency the implementer MUST confirm on first run (audit cannot verify this)
The entire §3.2 fix rests on one fact I cannot check from the audit side: **`_label_to_technical("Bank Name")` actually returns `"BANK_NAME"`.** That function lives in `fbdi/audit.py`, which is not visible to the MCP audit session, and the FBDI master catalog (`FBDI_Master_Catalog.xlsx`) was not available to me either, so I could not confirm the catalog's `technical`/`label` forms for the `Bank Account` tab. If `_label_to_technical` normalizes differently (e.g. title-case, hyphenation, or dropping words), every normalized match silently fails and the report reverts to ~all-false-positives — the same failure shape as Blocker 1, from the Oracle side.

**Required as the first implementation step (before trusting any output):** run the Task 15 integration test against the *real* catalog + live snapshot for `T_BANKS_BRANCHES` / `Bank Account` and confirm the business fields match. The plan's own self-review (lines ~2081-2084) predicts **~22/23 clean matches plus one genuine divergence**: Oracle "EDI ID Number" (→`EDI_ID_NUMBER`) vs Applaud `EFT_ID_NUMBER`. I confirmed the Applaud side of that claim live — the IF carries `T32EFT_ID_NUMBER` (row 11) and `T32EDI_LOCATION` (row 12), and there is **no** `T32EDI_ID_NUMBER` — so a real `EDI_ID_NUMBER` divergence is structurally consistent with the data. (I could not confirm the Oracle side has exactly "EDI ID Number"; that's part of what the first run validates.)

Acceptance for that first run: **~22 clean matches + that single reviewable HIGH "missing"/INFO "extra" pair, NOT 23 false positives.** If you get 23 PRESENCE findings, `_label_to_technical` is not producing UPPER_SNAKE_CASE and you need a normalization adjustment in `oracle_match_key` (or a small wrapper) before any dimension's output is trustworthy. That single divergence is signal the report *should* surface, not noise to suppress.

## 4. Smaller items (all addressed; no action)
- §3.1 ODBCName-empty: Task 10 docstring/note now states bare-name is the effective Dim 4 key and the ODBCName branch is defensive only. The `test_check_table_coverage_matches_on_odbcname` test still exists (harmless; documents intent) — fine, just don't treat ODBCName as load-bearing.
- §3.3 date-vs-char: `test_check_sizing_date_stored_as_char_is_type_class_finding` added; confirms an Oracle DATE stored as Applaud char yields a real TYPE_CLASS HIGH, not a silent skip.

## 5. Items explicitly preserved from pass-1/2 "keep as-is" (re-confirmed intact)
Per-object COUNT guard (re-verified live: ImportDetail 10,137 / ExportDetail 39,844, both cap ~100 unbounded; per-object pulls complete), app-map DBID bridge, EF-via-`get_application`-steps, confirmed-wins merge, finding model + finding_id, 4-sheet workbook, Dims 2/3/4/5/6b/6c logic, and the `align_tabs`-only-for-6b reasoning. All unchanged.

## 6. Handback
No further audit gate. Proceed to implementation. Make the §3 first-run confirmation (the `T_BANKS_BRANCHES` / `Bank Account` integration run, ~22 matches + 1 divergence) the very first executable check after the code compiles — it is the one assumption the audit could not close. If it fails, fix `oracle_match_key`/`_label_to_technical` normalization before building further on the dimension outputs.
