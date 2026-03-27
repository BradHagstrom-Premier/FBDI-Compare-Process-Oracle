# FBDI–Applaud Mapping Completion

## Context

Prior work (already done): 87 FBDI tabs have been mapped to Applaud tables with `in_scope = YES`. The remaining 512 FBDI tabs are marked `TBD`.

**This task:** Complete the mapping — assign an Applaud table to every FBDI tab where a match exists. Many FBDI tabs will have no Applaud match (that is expected and correct). Assign as many Applaud tables as you confidently can to one FBDI tab if a match exists. It shouldn't happen, but flag if an Applaud tables matches >1 FBDI tab.

---

## Input Files

All files are in the project directory:

| File | Description |
|------|-------------|
| `fbdi_applaud_mapping.xlsx` — sheet: `FBDI Mapping` | Master mapping workbook. 610 rows. Columns: `fbdi_file`, `fbdi_tab`, `applaud_table`, `prefix`, `in_scope`, `module`, `notes` |
| `applaud_table_coverage.csv` | 187 Applaud tables. Columns: `applaud_table`, `status` (`MAPPED`/`UNMAPPED`), `prefix`, `fbdi_mappings` |
| `AP5TBFLD.CSV` | Applaud field-level reference. Columns: `DB_Name` (table name), `DD_Name`, `DD_Type`, `DD_Size`, `DD_Decimals` |

**Current state:**
- 87 rows: `in_scope = YES` (already mapped — do not touch)
- 512 rows: `in_scope = TBD` (target of this task)
- 6 rows: `in_scope = FILE_TOO_LARGE` — skip
- 5 rows: `in_scope = FILE_ERROR` — skip
- 64 Applaud tables already `MAPPED`, 120 `UNMAPPED`

---

## Scope

**Files modified:**
- `fbdi_applaud_mapping.xlsx` — add `applaud_table` and `prefix` values to TBD rows where a match is found; leave `applaud_table` blank where no match exists; update `in_scope` on matched rows from `TBD` to `YES`; add `match_type` and `confidence` columns (see below); save in place
- `applaud_table_coverage.csv` — update `status` and `fbdi_mappings` columns in place. Change to xlsx file for simplicity whenever you deem appropriate.

**Files not modified:**
- `AP5TBFLD.CSV` — read-only reference, it can be removed from repo if not needed.
- Any row with `in_scope = YES`, `FILE_TOO_LARGE`, or `FILE_ERROR`

---

## Step-by-Step Instructions

### Step 1 — Invoke the `brainstorming` skill (superpowers)

Before writing any code, use the `brainstorming` skill to reason through the matching strategy. Consider:
- The naming pattern already established: Applaud tables almost always equal `T_` + FBDI tab name (e.g., FBDI tab `EGO_CHANGES_INT` → Applaud table `T_EGO_CHANGES_INT`)
- Some Applaud tables use slightly different suffixes or truncations
- Some FBDI tabs use human-readable names (e.g., `"Statement Headers"`, `"Project Billing Events"`) that require semantic inference
- Some Applaud tables have no FBDI counterpart (staging-only, reference, or utility tables)

### Step 2 — Invoke the `writing-plans` skill (superpowers)

Before implementing, write a short execution plan. Confirm:
- How you'll handle exact name matches
- How you'll handle fuzzy/semantic matches
- How you'll surface confidence levels
- How you'll handle one-to-many (same Applaud table → multiple FBDI tabs — this is valid and already present in the existing YES rows)

### Step 3 — Add columns to the working dataframe

Add two new columns to the `FBDI Mapping` sheet **before** writing any match logic:

| Column | Values |
|--------|--------|
| `match_type` | `EXACT` — Applaud table name = `T_` + fbdi_tab (case-insensitive) |
| | `INFERRED` — match derived from semantic similarity, prefix alignment, or field-level cross-reference |
| | `NO_MATCH` — no Applaud table found for this FBDI tab |
| `confidence` | `HIGH` — structural name match or near-certain semantic match |
| | `MEDIUM` — plausible semantic match, one degree of inference |
| | `LOW` — speculative; multiple candidates, weak signal |
| | _(blank)_ — for `NO_MATCH` rows |

### Step 4 — Run exact matching pass

For every TBD row:
1. Construct candidate: `"T_" + fbdi_tab.upper()`
2. Check if that string exists in `applaud_table_coverage.csv` → `applaud_table` column (case-insensitive)
3. If yes: set `applaud_table`, `prefix` (from coverage file), `match_type = EXACT`, `confidence = HIGH`, `in_scope = YES`

Known pattern exceptions already present in the YES rows (use these as examples, do not re-map them):
- `FA_RETIREMENTS_T` → `T_FA_RETIREMENTS_T` (suffix moved)
- `HZ_IMP_ACCOUNTRELS` → `T_HZ_IMP_ACCOUNTRELS_T` (suffix added)

Account for these suffix variants in exact matching:
- Try `T_` + fbdi_tab
- Try `T_` + fbdi_tab + `_T`
- Try stripping trailing `_INT`, `_ALL`, `_V` from fbdi_tab before prefixing

### Step 5 — Run semantic/inferred matching pass

For TBD rows that did not get an exact match, apply semantic inference. Use FBDI tab name, FBDI file name, and the field list from `AP5TBFLD.CSV` to reason about which Applaud table is the correct staging counterpart.

Key patterns to apply:

| Signal | Inference |
|--------|-----------|
| Human-readable FBDI tab names (e.g., `"Award Keywords"`) | Match to Applaud table whose name maps semantically (e.g., `T_AWARD_KEYWORDS`) |
| FBDI file = `ImportAwards` | Map Award-prefixed tabs to corresponding `T_AWARD_*` Applaud tables |
| FBDI file = `WorkOrderTemplate` or `ProcessWorkOrderTemplate` | Map to `T_WO_*` tables by operation type |
| FBDI file = `MaintenanceWorkDefinitionTemplate` or `WorkDefinitionTemplate` | Map to `T_WIS_WD_*` or `T_WORK_DEFINITION_*` tables |
| FBDI file = `ProjectImportTemplate` | Map to `T_PROJECTS`, `T_PROJECT_CLASSIFICATIONS`, `T_PROJECT_TEAM_MEMBERS` |
| FBDI tab = `DOO_ORDER_*` | Map to corresponding `T_DOO_ORDER_*` Applaud table |
| FBDI tab = `FA_*` | Map to corresponding `T_FA_*` Applaud table |
| FBDI tab = `GL_*` | Map to corresponding `T_GL_*` Applaud table |
| FBDI file = `SupplierAddressImportTemplate` | Map to `T_POZ_SUP_ADDRESSES_INT` |
| FBDI file = `SupplierContactImportTemplate` | Map to `T_POZ_SUP_CONTACTS_INT`, `T_POZ_SUP_CONTACT_ADDRESS_INT` |
| FBDI file contains `Worker` / `Person` / `Assignment` | Map to `T_WORKER`, `T_PERSONNAME`, `T_PERSONADDRESS`, etc. |
| FBDI tab = `RA_INTERFACE_DISTRIBUTIONS_ALL` | Map to `T_RA_INTERFACE_DISTRIBUTIONS` |
| FBDI file = `CycleCountImportTemplate` | Map to `T_INV_INVENTORY_*` or `T_INV_ITEM_*` tables |
| FBDI file = `ScpSourcingImportTemplate` | Map to `T_MSC_ST_SOURCING_RULES` or `T_MSC_ST_ASSIGNMENT_SETS` |
| FBDI file = `ScpBookingHistoryImportTemplate` | Map to `T_MSC_ST_MEASURE_DATA_BOOKINGS` |
| FBDI file = `ScpSalesOrderImportTemplate` | Map to `T_SCP_SALESORDER` |
| FBDI file = `ScpExternalForecastImportTemplate` | Map to `T_SCP_EXTERNALFORECAST` |
| FBDI file = `ScpUOMImportTemplate` | Map to `T_SCP_UOM_CONVERSION` |
| FBDI file = `ScpSafetyStockLevelImportTemplate` | Map to `T_SAFETYSTOCKLEVEL` |
| FBDI tab = `PJT_PRJ_ENT_RES_INTERFACE` | Map to `T_PROJ_ENT_RES_INTERFACE` |
| FBDI file = `ProjectResourceRequestImportTemplate` | Map to `T_PROJ_RES_REQ_INTERFACE` |
| FBDI file = `CreateBillingEventsTemplate` | Map to `T_PJB_BILLING_EVENTS_INT` |
| FBDI file = `StandardCostImportTemplate`, tab = `CST_INTERFACE_STD_COST_DETAILS` | Map to `T_CST_INTR_STD_COST_DETAIL` |
| FBDI file = `StandardCostImportTemplate`, tab = `CST_INTERFACE_STD_COST_HEADERS` | Map to `T_CST_INTR_STD_COST_HEADERS` |
| FBDI file contains `BankStatement` | No Applaud table likely exists — mark `NO_MATCH` |
| FBDI tab is a human-readable label with no Oracle table name (e.g., `"Genealogy"`, `"Batches"`, `"Contacts"`) | Use FBDI file context + field inspection to infer; mark `INFERRED / MEDIUM` or `LOW` |

Set `match_type = INFERRED` and assign `confidence = HIGH`, `MEDIUM`, or `LOW` based on your certainty. Populate `notes` with a brief rationale for every `INFERRED` row (e.g., `"Semantic match: tab name maps directly to Applaud table name pattern"`).

### Step 6 — Mark remaining rows NO_MATCH

Any TBD row that did not receive a match after Steps 4 and 5:
- Set `match_type = NO_MATCH`
- Leave `applaud_table` and `prefix` blank
- Leave `in_scope = TBD` (do not change — Brad will review these)
- Do not add notes unless there's something Brad needs to know

### Step 7 — Update `applaud_table_coverage.csv`

For every Applaud table that received at least one new mapping in Steps 4–5:
- Set `status = MAPPED`
- Set `fbdi_mappings` = count of FBDI tabs mapped to this table (across all rows, including pre-existing YES rows)

### Design Step - Do not skip
Utilize your 'frontend-design' skill to polish and improve the human readability of these files you are modifying and creating. Brad should open these files and immediately understand along with being impressed. He may even present the end results to his manager.

### Step 8 — Use the `verification-before-completion` skill (superpowers)

Before writing final output, verify:
- [ ] No YES row was modified
- [ ] Every EXACT match row has `confidence = HIGH`
- [ ] Every INFERRED row has a non-blank `notes` value
- [ ] `prefix` on all new matches aligns with the prefix in `applaud_table_coverage.csv`
- [ ] No Applaud table was assigned to a FBDI tab from a mismatched domain (e.g., a Financials Applaud table assigned to a Supply Chain FBDI tab)
- [ ] `fbdi_mappings` counts in coverage file are consistent with the mapping sheet

### Step 9 — Write output files

Save `fbdi_applaud_mapping.xlsx` in place using `openpyxl`:
- Preserve all existing formatting and column structure
- Add `match_type` and `confidence` columns immediately after `notes`
- Apply conditional formatting:
  - `match_type = EXACT` → light green fill (`#C6EFCE`)
  - `match_type = INFERRED` → light yellow fill (`#FFEB9C`)
  - `match_type = NO_MATCH` → no fill (white)
  - `confidence = LOW` → bold font

Save `applaud_table_coverage.csv` in place using pandas.

Run `scripts/recalc.py` on the xlsx and fix any formula errors before proceeding.

### Step 10 — Use the `commit-commands` plugin

Commit all changes with a descriptive message after output files are verified.

---

## Verification Criteria

Brad will review the completed file. His criteria:

1. **EXACT matches are obvious at a glance** — green fill, `HIGH` confidence, no notes needed
2. **INFERRED matches are clearly explained** — yellow fill, notes column tells him why
3. **LOW confidence rows stand out** — bold font signals "please review this one"
4. **NO_MATCH rows are not noise** — white/blank is the expected majority; Brad expects many FBDI tabs with no Applaud counterpart
5. **Every Applaud table that should be mapped is mapped** — especially the 123 currently `UNMAPPED` tables; if after this task an Applaud table is still `UNMAPPED`, that is a deliberate finding, not an oversight
6. **Coverage file is consistent** — `fbdi_mappings` counts match what's in the mapping sheet

---

## Key Domain Knowledge

- Applaud table names almost always follow the pattern `T_` + Oracle interface table name
- The `prefix` column in the coverage file is Applaud's internal staging prefix — it should be carried into the mapping sheet for every matched row
- One Applaud table can map to multiple FBDI tabs (already true for `T_PJC_TXN_XFACE_STAGE_ALL`, `T_INV_SERIAL_NUMBERS_INTERFACE`, etc.)
- Human-readable FBDI tab names (no underscores, title case) are Oracle UI labels — they do not have a direct Oracle table name and require inference from FBDI file context
- Tables prefixed `TABLE_AF`, `TABLE_B4`, `T_WORDS_FREQ`, `TEST_CONV_EXPORT` are likely utility/reference tables with no FBDI counterpart — confirm and mark `NO_MATCH`
- `T_POZ_SUP_THIRDPARTY_INT` has prefix `TE1` which conflicts with `T_AWARDS` (also `TE1`) — flag this in notes on both rows

---

## Migration Notes

No database migrations. Output is Excel + CSV only.

```bash
# Install dependencies if needed
pip install openpyxl pandas --break-system-packages

# Run formula recalculation after saving xlsx
python scripts/recalc.py fbdi_applaud_mapping.xlsx
```
