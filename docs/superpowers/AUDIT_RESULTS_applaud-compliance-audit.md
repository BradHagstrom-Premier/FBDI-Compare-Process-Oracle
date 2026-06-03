# Audit Results — Applaud Compliance Audit Design

**Audience:** Claude Code, before writing the implementation plan for `feat/applaud-compliance-audit`.
**Source:** Technical audit of `2026-06-02-applaud-compliance-audit-design.md`, validated live against `ORACLE_MASTER/AP0STE.mdb` via the `applaud-mcp` MCP server and against `FBDI_to_ApplaudTables_Mapping.xlsx`.
**Date:** 2026-06-02

This file is the authoritative set of corrections and constraints. Where it conflicts with the design doc, **this file wins**. Read it fully, then write the implementation plan. Do not start the plan until the four REQUIRED-CHANGE items below are reflected in your plan's assumptions.

---

## 0. TL;DR for the planner

The architecture (snapshot → offline audit engine → Excel findings; Phase-3-mechanical findings model) is sound and should be preserved. But the design contains one foundational scope error, one silent-data-loss bug in the extraction strategy, one prefix-derivation gap, and one stale-premise factual error. All four were confirmed against the live database. Fix them in the plan's assumptions before designing tasks.

Scope is now settled (confirmed with Brad): **the Applaud side of the audit is the `T_*` target tables only.** `O_*` objects do not map to any Oracle FBDI and are out of scope.

---

## 1. REQUIRED CHANGES (must be reflected in the plan)

### 1.1 — Scope is `T_*` target tables; drop the `O_*`/divergent-prefix premise

**Confirmed:** `FBDI_to_ApplaudTables_Mapping.xlsx` maps every FBDI tab to a `T_*` Applaud table and contains **zero** `O_*` tables (154 MAPPED rows, 485 UNMAPPED, 0 `O_*`). The `O_*` / `O33` family (`O_BANKS`, `I_O_BANKS`, etc.) is a separate staging lineage that does **not** feed FBDIs and is **out of scope** for this audit.

**Consequence for the design's central premise:** The design (§2) frames bare-name matching as solving a cross-prefix puzzle, illustrated as `O33BANK_NAME` (import) ↔ `T31BANK_NAME` (table). That illustration is wrong for the in-scope family. Within the `T_*` family, **the IF, the EF, and the target table all share the same TableId prefix.** Confirmed on the corrected example:

- Table `T_BANKS_BRANCHES` → prefix `T32`
- Import `I_T_BANKS_BRANCHES` fields → `T32COUNTRY`, `T32BANK_NAME`, `T32BANK_CODE`, … (prefix `T32`)
- Export `T_BANKS_BRANCHES` (EF) fields → `T32COUNTRY`, `T32BANK_NAME`, … (prefix `T32`)

So bare-name matching is still needed, but **only on the Oracle↔Applaud boundary** (Oracle technical name `BANK_NAME` ↔ Applaud bare name after stripping `T32`). The **IF↔table and EF↔table** comparisons within Applaud are a *trivial exact DDID match* (same prefix), not a cross-prefix reconciliation. The plan should state this explicitly so the matching logic isn't over-engineered for a problem that doesn't exist in scope, and so Dim 5 (below) is implemented correctly.

**Action:** Remove/rewrite every `O33`/`I_O_BANKS`/divergent-prefix assertion in the design. State that audit scope = `T_*` tables present in the FBDI mapping workbook, resolved to IFs/EFs via the `Application.DBID` bridge.

### 1.2 — Replace the canonical worked example with `T_BANKS_BRANCHES` (validated end-to-end)

`T_BANKS` was a bad example: it maps to **no** Oracle FBDI, so it would never appear in a consultant-facing audit. Use `T_BANKS_BRANCHES` throughout. Full chain, validated live:

| Link | Value | Source (verified) |
|---|---|---|
| FBDI template . tab | `RapidImplementationForCashManagement` . `Bank Account` | mapping workbook |
| Applaud target table | `T_BANKS_BRANCHES` (prefix `T32`) | mapping workbook + `get_table_definition` |
| Bridge | `Application.DBID = 'T_BANKS_BRANCHES'` → rows `CQ_T_BANKS_BRANCHES`, `I_T_BANKS_BRANCHES`, `X_T_BANKS_BRANCHES` | `Application` query |
| Import file (IF) | `I_T_BANKS_BRANCHES` → `get_application` step `I_T_BANKS_BRANCHES (IF)` | `get_application` |
| Export app | `X_T_BANKS_BRANCHES` → steps `T_BANKS_BRANCHES (EF)`, `X_T_BANKS_BRANCHES_VAL (EF)` | `get_application` |

Note the EF naming asymmetry the design should encode: the *export application* is `X_T_BANKS_BRANCHES`, but its first *export-file step* is named `T_BANKS_BRANCHES` (no `X_` prefix), and there's a second `_VAL` validation EF. The design's §4 prose ("the `X_T_*` exports") is only half right — resolve EFs by reading `get_application` steps, **not** by assuming an `X_` filename.

### 1.3 — `execute_query` SILENTLY TRUNCATES at ~100 rows; the "one-shot bulk pull" snapshot strategy is unsafe

This is the highest-risk technical defect. The design (§4) specifies each snapshot collection as "a one-shot bulk `execute_query` pull (no per-object round trips)." Confirmed against live data:

| Table | `SELECT COUNT(*)` | Rows returned by unbounded `SELECT … FROM <table>` |
|---|---|---|
| `ImportDetail` | **10,137** | ~100 (silent cap, no error) |
| `ExportDetail` | **39,844** | ~100 (silent cap, no error) |

An unbounded select returns ~100 rows with **no error and no truncation signal**. For an audit whose entire value is completeness ("so silence is never mistaken for a pass," §7), a snapshot that silently drops >99% of detail rows produces a clean-looking, confidently-wrong report. This is a release-blocking bug if shipped as designed.

**Prescribed pattern (validated):** Pull detail tables **per resolved object**, driven off the confirmed app-map, not in one bulk select. Per-object pulls are complete and well under the cap:

- `SELECT … FROM ImportDetail WHERE Name = 'I_T_BANKS_BRANCHES' ORDER BY Row` → all rows returned.
- `SELECT … FROM ExportDetail WHERE Name = 'T_BANKS_BRANCHES' ORDER BY Row` → all 23 rows returned (matches `COUNT(*)`).

**Plan must include:**
1. Per-object detail extraction (loop over the IFs/EFs/tables the confirmed app-map names), **not** a bulk pull. This also naturally scopes the snapshot to what's actually audited.
2. A **row-count assertion** after every pull: compare returned row count to `SELECT COUNT(*) FROM <table> WHERE Name = '<obj>'`; if they differ, fail loud (do not silently proceed). Generalize this guard for any pull that could exceed the cap (e.g., `DataDictionary`, `DatabaseDetail` pulled per-table).
3. Re-confirm the cap value at implementation time and treat the exact number as unknown/environment-dependent — assert against `COUNT(*)`, don't hardcode 100.

### 1.4 — Prefix derivation from the description parenthetical is not reliable; add a fallback

The design (§2) states prefixes are "read from the table description `"T_BANKS (T31)"` … never guessed." This fails for real in-scope-adjacent objects:

- `T_BANKS_BRANCHES` description = `"T_BANKS_BRANCHES (T32)"` → parenthetical present ✅
- `O_BANKS` description = `"O_BANKS"` → **no parenthetical**, yet prefix is `O33` (its key field is `O33BANK_NAME`)

While `O_*` is out of scope, the loader must not crash or mis-derive on objects lacking the parenthetical. **Plan must specify a documented prefix-derivation fallback:** parse the parenthetical when present; otherwise derive the prefix from the table's own column/key DDIDs (`get_table_definition` key sequence, or first `DatabaseDetail.DDID`). Note this *is* the "guessing" the design warned against, so make the fallback explicit and logged, not silent.

The mapping workbook also carries an authoritative `Prefix` column (fully populated, 0 blanks across 639 rows). For the **Oracle/mapping side**, use that column. The parenthetical/fallback issue applies only to **Applaud-side table-definition** prefix reads.

---

## 2. FACTUAL CORRECTION (stale premise — relaxes a prerequisite)

The design (§2, §11) asserts the MCP has **no named systems** (`list_systems` returns none) and a stale default path, making the §11 `/update-config` step a Step-A prerequisite. **This is now false.** `list_systems` returns two configured aliases:

```text
AWC_MASTER     — C:/Users/10193/Definian/MDB_for_ApplaudMCP/AWC_MASTER/AP0STE.mdb
ORACLE_MASTER  — C:/Users/10193/Definian/MDB_for_ApplaudMCP/ORACLE_MASTER/AP0STE.mdb
```

**Consequences for the plan:**
- The `--system` flag can delegate name→path resolution to the MCP server today (pass `system: 'ORACLE_MASTER'` directly). The design's "preferable" option in §11 is already the live state.
- `config.py` does **not** need to duplicate a name→path map. Prefer passing `system` through to MCP calls.
- The §11 `/update-config` task is **no longer a blocker** for Step A. Downgrade it from prerequisite to optional cleanup (or drop it). Keep a defensive note that a bare default-path call still errors, so every call must pass `system` (or `file_path`).

---

## 3. CONFIRMED-CORRECT (preserve these; no change needed)

- **`Application.DBID` bridge works.** `DBID = '<target_table>'` returns the `I_*` (import app), `X_*` (export app), and `CQ_*` (CTQ) applications for that table. Verified for `T_BANKS_BRANCHES`.
- **`get_application` cleanly labels steps `IF` / `EF` / `CS`.** Verified: `X_T_BANKS_BRANCHES` → two `EF` steps; `I_T_BANKS_BRANCHES` → one `IF` step. The design's step-resolution approach is correct.
- **Per-object detail pulls are complete and ordered** via `ORDER BY Row` (see 1.3).
- **`ImportDetail` schema** carries `Name, Row, DDID, Pic, InputType` (design §4 correct). **`ExportDetail`** carries `Name, Row, DDID, ColumnHeader` — but see note in §4 below.
- **The phased model and `Finding` record shape** (addressable `(object_type, object_name, field, attribute, current→expected)` tuple; idempotent by `finding_id`; snapshot-as-pre-image) are good and keep Phase 3 mechanical. Preserve as-is.
- **Mapping workbook structure:** sheet `FBDI Mapping` with columns `FBDI Template | FBDI Tab | Applaud Table | Prefix | Status | Module | In Base System?`; sheet `Applaud Tables Reference`. 639 mapping rows (154 MAPPED / 485 UNMAPPED). Some rows carry a multi-FBDI caveat in the "In Base System?" column (see §4).

---

## 4. SECONDARY NOTES (address in the plan, not blockers)

- **Dim 5 ("orphaned data element") needs re-grounding.** With scope = `T_*` only and IF/EF/table sharing one prefix, an in-scope IF/EF field should normally match a table column on exact DDID. Dim 5 should fire only on genuine intra-Applaud orphans (IF/EF field with no table column), not on the cross-prefix mismatches the old design implied. Re-validate Dim 5's trigger against the corrected single-prefix reality so it doesn't either over-fire or become a no-op.
- **Dim 3 `ColumnHeader` is empty on real EFs.** On `T_BANKS_BRANCHES` (EF), every row's `ColumnHeader = ""`. Dim 3 must derive the Oracle-comparison name from the **bare DDID**, not `ColumnHeader`. Don't assume `ColumnHeader` is populated.
- **One FBDI tab → multiple Applaud tables is real.** The "BANKS" search returned `T_BANKS`, `T_BANKS_BRANCHES`, `T_RA_CUSTOMER_BANKS_INT_ALL`, `T_IBY_TEMP_EXT_BANK_ACCT`, plus out-of-scope `O_BANKS`/`HD_T_BANKS`/`^CLOUD_BANKS`. The mapping workbook also flags rows where "Multiple mapping is possible … Would need unique IFs and EFs for each Oracle FBDI." Confirm the app-map workbook schema (`target_table | import_files | export_files | …`) can represent many-tables-per-FBDI-tab and many-IFs/EFs-per-table. The design's per-table-row model is probably fine, but verify the FBDI→table join isn't assumed 1:1.
- **EF resolution must not assume an `X_` filename** (see 1.2): read `get_application` steps.

---

## 5. ACCEPTANCE CRITERIA FOR THE PLAN YOU ARE ABOUT TO WRITE

The plan is acceptable when it:

1. States scope as `T_*` target tables from the FBDI mapping workbook; contains no `O_*`/divergent-prefix matching logic.
2. Uses `T_BANKS_BRANCHES` as the canonical worked example, with EF resolution via `get_application` steps (not `X_`-filename assumption).
3. Specifies **per-object** detail extraction with a **post-pull `COUNT(*)` row-count assertion that fails loud** on mismatch — explicitly replacing the "one-shot bulk pull." Asserts against `COUNT(*)`, does not hardcode the cap.
4. Specifies the prefix-derivation fallback (parenthetical → else DDID-derived, logged) for Applaud-side reads, and uses the mapping workbook's `Prefix` column for the Oracle side.
5. Treats the §11 MCP-config task as optional cleanup, not a Step-A blocker; passes `system: 'ORACLE_MASTER'` to MCP calls rather than duplicating a path map in `config.py`.
6. Re-grounds Dim 5 and Dim 3 per §4 above.
7. Keeps the `Finding` model, phased plan, finding-id reconciliation, and Coverage sheet (silence-is-not-a-pass) intact.
8. Keeps the testing approach (offline synthetic snapshots; MCP mocked at the data boundary) and adds a test for the row-count-assertion guard and for the prefix-fallback path.

---

## 6. HANDBACK INSTRUCTION

Write the implementation plan incorporating the above. When the plan is drafted, **hand it back to the auditing session (the Claude project with `applaud-mcp` access)** for a second-pass technical audit against the live database before implementation begins. Do not begin implementation until the plan has passed that audit.
