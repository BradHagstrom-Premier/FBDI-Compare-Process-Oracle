# Audit Results — Pass 2 (Implementation Plan)

**Audience:** Claude Code, before implementing `2026-06-02-applaud-compliance-audit.md`.
**Source:** Second-pass technical audit of the implementation plan, validated live against `ORACLE_MASTER/AP0STE.mdb` via `applaud-mcp` and against `FBDI_to_ApplaudTables_Mapping.xlsx`.
**Date:** 2026-06-02
**Verdict:** **Do not implement as written.** Two release-blocking defects in the data layer; both are confirmed against live data. The architecture, task decomposition, TDD structure, and Dims 2/3/4/5/6b/6c logic are sound and should be kept. Fix the two blockers and the three smaller items below, then this is good to build.

Same precedence rule as before: where this file conflicts with the plan, **this file wins**.

---

## 0. What's good (keep as-is)

- Two-layer split (agent-driven Step A extraction → pure offline Step B engine) is correct and testable.
- Per-object extraction with the `assert_complete()` COUNT guard (Task 2, Task 16) correctly implements Pass-1 finding 1.3. Verified again live: `ImportDetail`=10,137 rows / `ExportDetail`=39,844 rows, both silently capped ~100 on an unbounded select; per-object `WHERE Name=…` pulls return complete (`I_T_BANKS_BRANCHES`=23, EF `T_BANKS_BRANCHES`=23, all returned).
- App-map bridge (Tasks 4–5), EF-via-`get_application`-steps (no `X_` filename assumption), confirmed-wins merge, finding model + stable `finding_id`, 4-sheet workbook, Dims 2/3/4/5/6b/6c logic, and the `align_tabs`-only-for-6b reasoning are all correct.
- The `DatabaseDetail` extraction query names columns that exist (verified: `Name, Row, DataType, Size, DecPlaces, DDID, ODBCName` all present).

---

## 1. BLOCKER 1 — Dim 1 sizing reads the wrong table; `DatabaseDetail` carries NO type data

**This is the headline check and it is built on an empty column.**

The plan (Task 16 reference doc, line ~1879) states: *"DataDictionary is NOT pulled in Phase 1 (sizing comes from DatabaseDetail)."* This is inverted. Verified live on `T_BANKS_BRANCHES`:

- `SELECT Name,Row,DataType,Size,DecPlaces,DDID,ODBCName FROM DatabaseDetail WHERE Name='T_BANKS_BRANCHES'` → **every one of the 49 columns** returns `DataType=""`, `Size=0`, `DecPlaces=0`, `ODBCName=""`. `DatabaseDetail` stores row order and DDID linkage, **not** type/size/scale.
- The actual type data is on **DataDictionary**: `SELECT … FROM DataDictionary WHERE Name='T32BANK_NAME'` → `DataType='X', Size=100, DecPlaces=0`.

**Consequence if shipped as written:** `actual_shape()` (Task 7) receives `data_type="", size=0` for every column. Dim 1 compares Oracle (e.g. `char 100`) against Applaud `("", 0, 0)`, so:
- Every mapped field emits a false **HIGH** finding (type-class mismatch or `char 0 < char N` undersized).
- The "High Priority" worklist — the consultant's primary deliverable — is 100% false positives.
- The bug is invisible to the test suite: Task 7's unit tests hand-construct `DataColumn(..., size=100, ...)` directly, so they never exercise the empty-`DatabaseDetail` reality. Green tests, broken audit.

**Required fix:**
1. **Step A must pull DataDictionary** for every in-scope DDID and source `data_type/size/dec_places` from there, not from `DatabaseDetail`. Two viable shapes:
   - Pull the per-table DD slice by prefix: `SELECT Name,DataType,Size,DecPlaces FROM DataDictionary WHERE Name LIKE 'T32%'` (apply the COUNT guard — DD slices can exceed the ~100 cap for big tables; assert against `SELECT COUNT(*) … WHERE Name LIKE 'T32%'`), **or**
   - Pull DD per DDID as each table's columns are enumerated.
   - `DatabaseDetail` is still pulled — but only for **column presence + Row order + DDID + ODBCName**, not type. Populate `DataColumn.data_type/size/dec_places` by joining each DDID to its DataDictionary entry in the pure-Python assembly step.
2. `build_table()` (Task 2) gains a `dd_by_ddid: dict[str, dict]` argument and fills type/size/scale from it; keep `Row`/`ODBCName` from `DatabaseDetail`.
3. **Add a real-data regression test** that feeds a `DatabaseDetail` row with blank `DataType/Size` plus a DataDictionary entry with the true type, and asserts `actual_shape()` reflects the DD type — this is the test that would have caught the bug.
4. Update the Task 16 reference doc line to the opposite: *"DataDictionary IS pulled in Phase 1; sizing comes from DataDictionary, not DatabaseDetail (which carries no type data)."*

Note: this aligns with existing project knowledge — DataDictionary sizing is keyed `Name LIKE '<prefix>%'`, `DataType` column (not `Type`) holds the code. The plan contradicted that; restore it.

---

## 2. BLOCKER 2 — `@`-prefixed audit/tracking fields are never excluded

The design's §2 explicitly learned: **`@`-prefixed fields are internal Definian/Applaud audit fields and must be excluded from FBDI overlap scoring.** The plan implements no such exclusion anywhere.

Verified live: `T_BANKS_BRANCHES` `DatabaseDetail` rows 24–49 are `@T32DO_NOT_LOAD`, `@T32DO_NOT_LOAD_REASON`, `@T32LEGACY_HEADER1..10`, `@T32LEGACY_FIELD1..10`, `@T32SITE`, `@T32TARGET_TABLE`, `@T32EXPORT_NUMBER` — 26 of the table's 49 columns are `@`-audit fields.

The plan's `_strip_prefix(ddid, "T32")` does **not** strip these, because they start with `@`, not `T32`. Demonstrated: `_strip_prefix("@T32LEGACY_HEADER1","T32")` returns `"@T32LEGACY_HEADER1"` unchanged. So they enter the matching universe with mangled bare names.

**Consequence:**
- Dim 4 (Oracle→table) is *accidentally* safe here only because the table is a superset of Oracle fields, so the junk columns are never the thing that's missing. But they inflate the `present` set and any coverage counts.
- The real exposure is Dim 2/3 (and any future direction): any `@`-field that appears in an IF/EF — which happens on the validation/“FBDI Fields” exports — would be surfaced as an **INFO extra field**, cluttering the report with noise the design said to suppress. It also corrupts ORDER LCS if an `@`-field sits between business fields.
- `prefix_fallback` LCP derivation (Task 3) is also skewed: if a table's columns are a mix of `@T32…` and `T32…`, the longest common prefix is `""` or `"@"`-dependent, silently producing a wrong/empty prefix on the fallback path.

**Required fix:**
1. Add a single exclusion predicate, e.g. `is_audit_field(ddid) -> bool` (`ddid.lstrip().startswith("@")`), applied in `build_file_fields` and `build_table` assembly (drop `@`-fields, or tag them `is_audit=True` and filter in every dimension's matching set).
2. Exclude `@`-fields from the LCP prefix-fallback input in Task 3.
3. Add a test: a table/IF/EF containing an `@`-field must not produce PRESENCE/ORDER/orphan findings for it.

(If there's a reason to retain `@`-fields in the snapshot for Phase-3 write fidelity, tag-and-filter rather than drop — but they must be out of all Dim 1–6 matching.)

---

## 3. SMALLER ITEMS (fix in the plan; not blockers)

### 3.1 — `ODBCName` is empty too; Dim 4's ODBCName match path is dead on this data
Task 10 matches Oracle technical names against `column.bare` **or** `column.odbc_name`. On `T_BANKS_BRANCHES` every `ODBCName=""`. The bare-name path still works, so Dim 4 functions, but the test `test_check_table_coverage_matches_on_odbcname` (line ~1127) asserts behavior that real data won't exercise. Keep the code (defensive), but note in the plan that ODBCName is empty in ORACLE_MASTER so bare-name is the effective match key — don't let a future maintainer think ODBCName is load-bearing.

### 3.2 — The Oracle-technical-name ↔ Applaud-bare-name equality is unverified and is the spine of Dims 1/2/4
I could not validate this half: the FBDI master catalog (`FBDI_Master_Catalog.xlsx`) was not available to me, so I could not confirm that `AlignedField.technical` (what `report.load_catalog_release` produces for the `RapidImplementationForCashManagement` / `Bank Account` tab) equals the Applaud bare names (`COUNTRY`, `BANK_NAME`, `BANK_CODE`, …). Every match in Dims 1/2/4 assumes `technical.upper() == bare.upper()`. If the catalog's `technical` is a display label ("Bank Name") or a differently-cased/punctuated form, **every dimension silently mis-matches** and the report is again mostly false positives — the same failure shape as Blocker 1, just from the Oracle side.

**Required of the implementer (since I can't):** before trusting any audit output, run the audit on the `T_BANKS_BRANCHES` / "Bank Account" pair and confirm the in-scope business fields match (zero spurious PRESENCE findings on the 23 known-good IF fields). Add an integration test that pins the catalog `technical` form against the known T32 bares for this tab. If they don't match, a normalization step (label→technical, case, underscores) is needed in the matching layer — `_label_to_technical` from `audit.py` was cited in the design for exactly this; wire it in.

### 3.3 — Dim 1 type-class guard has a latent hole
`check_sizing` (Task 7) line `if exp_cls != act_cls and exp_cls not in ("", act_cls):` skips the type-class finding when `exp_cls == ""`. After Blocker 1 is fixed, `act_cls` comes from real DD data (`X`→char, `N`→numeric). But `exp_cls` comes from `applaud_type_for()`; if it ever returns `"date"` while Applaud stores the date as `X`, confirm the intended verdict (likely a real TYPE_CLASS finding, not a skip). Add a date-vs-char test once the DD source is wired, so the date class isn't silently swallowed.

---

## 4. ACCEPTANCE CRITERIA FOR THE REVISED PLAN

1. **Dim 1 sources type/size/scale from DataDictionary**, not `DatabaseDetail`; Step A pulls DD (with COUNT guard); `build_table` joins DD type data onto columns; a regression test feeds blank-`DatabaseDetail` + populated-DD and asserts correct `actual_shape()`.
2. **`@`-prefixed fields are excluded** from all Dim 1–6 matching (assembly-level predicate + LCP-fallback exclusion), with a test.
3. Task 16 reference doc corrected to "DataDictionary IS pulled; DatabaseDetail has no type data."
4. Plan notes `ODBCName` is empty in ORACLE_MASTER (bare-name is the effective Dim 4 key).
5. Plan adds an **integration check** on the `T_BANKS_BRANCHES` / "Bank Account" pair confirming the 23 known-good IF fields produce zero spurious PRESENCE findings (validates §3.2 Oracle-name↔bare-name equality end-to-end), with normalization wired in if the forms differ.
6. Everything in §0 is preserved unchanged.

---

## 5. HANDBACK INSTRUCTION

Revise the plan to satisfy §4, then route it back to the Applaud-MCP audit session for a third-pass spot-check (focused only on the two blocker fixes and the §3.2 integration check). Do not begin implementation until Blockers 1 and 2 are resolved in the plan text — they are data-layer facts, not style preferences, and both were confirmed against the live database.
