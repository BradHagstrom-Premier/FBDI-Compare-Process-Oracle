# Handoff: FBDI–Applaud Mapping Enrichment

## Context

`fbdi_applaud_mapping.xlsx` is a lookup table Brad maintains that maps each FBDI
template tab to its corresponding Applaud database table (`applaud_table`), the
Applaud field prefix for that table (`prefix`), and a scoping flag (`in_scope`).
Most rows are currently `TBD` — no Applaud table or prefix is filled in.

`AP5TBFLD.CSV` is Applaud's field catalog: 11,944 rows across 187 tables, each
row being one Applaud field. Table names in the CSV follow the convention
`T_<ORACLE_INTERFACE_TABLE_NAME>` (e.g., `T_RA_INTERFACE_LINES_ALL`). Field
names follow the convention `<PREFIX><ORACLE_FIELD_NAME_TRUNCATED_AT_30_CHARS>`
(e.g., `TA4CONS_BILLING_SET_ID`).

The join key between the two files is:
`CSV.DB_Name == 'T_' + mapping.fbdi_tab`

This script uses that key to enrich the mapping file in place — no new file, no
new columns. It only updates rows currently marked `TBD`.

---

## Decisions (Closed)

| Decision | Choice |
|---|---|
| Match strategy | `T_` + `fbdi_tab` exact match; fuzzy matching allowed when confidence is >=95% |
| in_scope for confirmed matches | Set to `YES` |
| Inference pass | allowed when confidence is >=95% |
| Output | Update `fbdi_applaud_mapping.xlsx` in place |
| Rows already filled (YES / FILE_TOO_LARGE / FILE_ERROR) | Leave untouched |

---

## Input Files

Both files are at fixed paths. Do not hardcode any other paths.

```
AP5TBFLD.CSV              → Applaud field catalog
fbdi_applaud_mapping.xlsx → Mapping file to enrich (update in place)
```

Locate both via `Path(__file__).resolve().parent` — assume the script lives in
the repo root alongside these files, which is where Brad will place them.

---

## Reconnaissance Findings (Do Not Re-derive)

These facts were established in Claude Chat before writing this spec. The
implementation must match them exactly.

**CSV structure:**
- Columns: `DB_Name`, `DD_Name`, `DD_Type`, `DD_Size`, `DD_Decimals`
- Valid `DD_Type` values: `X` (char), `N` (numeric), `D` (date), `U` (unknown)
- 27 rows have malformed `DD_Type` values (CSV parse artifacts — embedded commas
  in description columns). Filter to valid types only before processing.
- Filter: `df[df['DD_Type'].isin(['X', 'N', 'D', 'U'])]` → 11,917 clean rows

**Mapping file structure:**
- Single sheet: `FBDI Mapping`
- Columns (row 1 header): `fbdi_file`, `fbdi_tab`, `applaud_table`, `prefix`,
  `in_scope`, `module`, `notes`
- 610 data rows (row 2 onward)
- Existing `in_scope` values: `YES`, `TBD`, `FILE_TOO_LARGE`, `FILE_ERROR`
- Only rows where `in_scope == 'TBD'` are eligible for enrichment

**Auto-match results (78 rows confirmed):**
- 77 of 78 auto-match candidates have a single, consistent prefix across all
  fields in the table
- 1 edge case: `T_GL_INTERFACE` (tab: `GL_INTERFACE`, file:
  `JournalImportTemplate`) has two prefixes — `T01` (139 fields, majority) and
  `T02` (10 attribute fields, minority). Resolution: use `T01` as the prefix;
  write `"Secondary prefix T02 (10 fields) also present"` into the `notes`
  column for that row (do not overwrite any existing notes — append).

**Prefix extraction algorithm:**
- Pattern: `^([A-Z][A-Z0-9]{1,2})([A-Z_].*)` applied to `DD_Name`
- The prefix is the first capture group (2–3 chars)
- Extract prefix from the majority of fields in the table (mode of extracted
  prefixes). If only one prefix exists, use it directly.
- If no prefix can be extracted from any field, leave `prefix` blank and write
  `"No prefix found in CSV fields"` to `notes`.

**164 TBD rows with Oracle-style tab names that do NOT match any CSV table:**
These have no `T_` + tab entry in the CSV. Leave them `TBD` — do not touch.

**~200 TBD rows with non-Oracle-style tab names** (e.g., `"Work Definition
Headers"`, `"Transfer Order Lines"`): these cannot be joined mechanically.
Leave them `TBD` — do not touch.

---

## Implementation

### Step 0 — Verify inputs before writing any code

Activate the superpowers brainstorming skill. Activate the superpowers executing a plan skill.

Read both files and confirm:
1. `AP5TBFLD.CSV` loads and has the 5 expected columns
2. After filtering to valid `DD_Type`, row count is ~11,917
3. `fbdi_applaud_mapping.xlsx` opens, has sheet `FBDI Mapping`, columns match
   the 7-column spec above
4. Count TBD rows — should be ~599 (610 total minus 11 non-TBD rows)
5. Spot-check: `T_RA_INTERFACE_LINES_ALL` is present in the CSV with prefix `TA4`

Do not proceed to implementation until these checks pass.

### Step 1 — Build the enrichment script (`enrich_mapping.py`)

Use `openpyxl` for the Excel file (already a project dependency). Use `pandas`
for CSV loading and prefix extraction. Do not use `xlwings` or `xlrd`.

```
Script: enrich_mapping.py (repo root)
Inputs: AP5TBFLD.CSV, fbdi_applaud_mapping.xlsx (same directory as script)
Output: fbdi_applaud_mapping.xlsx (updated in place)
```

**Algorithm:**

```
1. Load CSV, filter to valid DD_Type rows, build lookup:
      csv_lookup: { DB_Name -> list[DD_Name] }
      csv_tables: set of all DB_Name values

2. Load mapping xlsx, read all rows into memory as list of dicts keyed by
   column name. Preserve original row order.

3. For each mapping row where in_scope == 'TBD':
      tab = row['fbdi_tab']
      if tab is None or not re.match(r'^[A-Z][A-Z0-9_]+$', tab.strip()):
          skip (non-Oracle-style tab name)
      candidate = 'T_' + tab.strip()
      if candidate not in csv_tables:
          skip
      # Confirmed match — extract prefix
      fields = csv_lookup[candidate]
      prefix = extract_majority_prefix(fields)
      # Update row
      row['applaud_table'] = candidate
      row['prefix'] = prefix  (or '' if none found)
      row['in_scope'] = 'YES'
      # Notes: append if T_GL_INTERFACE edge case, else leave notes unchanged

4. Write all rows back to fbdi_applaud_mapping.xlsx using openpyxl,
   preserving existing formatting (bold header row, column widths if set).
   Do not reformat or strip styles from rows that were not changed.
```

**`extract_majority_prefix(fields)` function:**

```python
import re
from collections import Counter

PREFIX_PATTERN = re.compile(r'^([A-Z][A-Z0-9]{1,2})([A-Z_].*)')

def extract_majority_prefix(fields: list[str]) -> str:
    prefixes = []
    for f in fields:
        m = PREFIX_PATTERN.match(f)
        if m:
            prefixes.append(m.group(1))
    if not prefixes:
        return ''
    return Counter(prefixes).most_common(1)[0][0]
```

**openpyxl write-back pattern:**

Load with `data_only=True` (not `read_only=True` — we need to write back).
Iterate `ws.iter_rows(min_row=2)` to find rows to update. Match by row index
(safe here since we're reading the same file we just loaded — no insertions or
deletions). Update only the cells for columns `applaud_table` (col 3),
`prefix` (col 4), `in_scope` (col 5), and `notes` (col 7) on matched rows.
Do not touch other cells. Save with `wb.save(path)`.

Column index reference (1-based):
- col 1: `fbdi_file`
- col 2: `fbdi_tab`
- col 3: `applaud_table`
- col 4: `prefix`
- col 5: `in_scope`
- col 6: `module`
- col 7: `notes`

### Step 2 — Console output

Print a summary when the script finishes:

```
Enrichment complete.
  Rows updated:   78
  Rows skipped (no match): 521
  Rows skipped (already filled): 11

Updated rows:
  JournalImportTemplate | GL_INTERFACE -> T_GL_INTERFACE (prefix=T01)
  ... (all 78, sorted by fbdi_file then fbdi_tab)
```

### Step 3 — Verification

After running the script, verify:
1. Open the updated xlsx and confirm row count is unchanged (610 data rows)
2. Spot-check 5 specific rows against the known-correct mapping:
   - `JournalImportTemplate | GL_INTERFACE` → `T_GL_INTERFACE`, prefix `T01`,
     in_scope `YES`, notes contains "Secondary prefix T02"
   - `POPurchaseOrderImportTemplate | PO_HEADERS_INTERFACE` → `T_PO_HEADERS_INTERFACE`,
     prefix `T77`, in_scope `YES`
   - `CustomerImportTemplate | HZ_IMP_PARTIES_T` → `T_HZ_IMP_PARTIES_T`,
     prefix `T50`, in_scope `YES`
   - `SourceSalesOrderImportTemplate | DOO_ORDER_CHARGES_INT` → `T_DOO_ORDER_CHARGES_INT`,
     prefix `TC2`, in_scope `YES`
   - `InventoryTransactionImportTemplate | INV_TRANSACTIONS_INTERFACE` → `T_INV_TRANSACTIONS_INTERFACE`,
     prefix `T47`, in_scope `YES`
3. Confirm a known non-matching row is untouched:
   - `AutoInvoiceImportTemplate | RA_INTERFACE_DISTRIBUTIONS_ALL` → still `TBD`
4. Confirm a non-Oracle-style tab is untouched:
   - `WorkOrderMaterialTransactionTemplate | Interface Batch` → still `TBD`

Use `pyright-lsp` to type-check after implementation.
Use `commit-commands` plugin to commit when verification passes.

---

## Known Hazards

- **`openpyxl` + `.xlsx` write-back strips some formatting** if you load with
  `read_only=True` then try to save. Load with neither `read_only` nor
  `write_only` (the default) to allow both read and write.
- **`data_only=True` is required** to read cell values (not formulas) from
  cells that may have been formula-populated in a prior version of the file.
- **The T02 edge case must not corrupt the notes column** for rows that already
  have notes content. Use append logic: `existing + "; " + new_note` if
  `existing` is non-empty, else just `new_note`.
- **Tab names in the mapping file have occasional leading spaces** — a few rows
  like `' DOO_ORDER_LOT_SERIALS_INT'` have a leading space. `.strip()` before
  matching.
- **CSV has trailing spaces in some DB_Name values** — build the `csv_tables`
  set using `.strip()` on `DB_Name`.

---

## Plugins

- `pyright-lsp` — after writing `enrich_mapping.py`, before first run
- `commit-commands` — after verification passes
- supeerpowers