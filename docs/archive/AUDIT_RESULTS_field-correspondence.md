# AUDIT RESULTS — Field-Correspondence Layer Design (Second-Pass, Live-Verified)

**Source:** Technical audit of `2026-06-10-applaud-field-correspondence-design.md`, validated live
against `ORACLE_MASTER` via the `applaud-mcp` MCP server. All ten pilot tables were pulled in full
(per-object `WHERE Name = '<table>'` pulls, each preceded by a `COUNT(*)` assertion; all pulls
complete).
**Date:** 2026-06-11
**Pilot scope (confirmed with Brad):** T_AP_INVOICE_INT (TA1, 155 cols), T_AP_INVOICE_LINES (T99, 189),
T_BANKS_BRANCHES (T32, 49), T_BPA_PO_LINES_INTERFACE (T64, 132), T_EGP_COMPONENTS_INTERFACE (T91, 126),
T_EGP_ITEM_CATEGORIES_INT (T87, 40), T_EGO_ITEM_INTF_EFF_B (T86, 156), T_MSC_ST_ASSIGNMENT_SETS (T04, 53),
T_POZ_SUPPLIERS_INT (T07, 184), T_POZ_SUPPLIER_SITES_INT (T09, 226).

This file is the authoritative set of corrections and constraints for the field-correspondence
implementation plan. Where it conflicts with the design doc, **this file wins** (per the standing
handback convention). Read it fully, then write the plan. Do not start the plan until the four
REQUIRED-CHANGE items are reflected in the plan's assumptions.

---

## 0. TL;DR for the planner

The architecture is sound and should be preserved: derive → confirm → audit as three re-runnable
commands, committed `FBDI_to_Applaud_FieldMap.xlsx` cloning the app-map derived→confirmed pattern,
aliasing applied only inside `run_audit` with the four check functions untouched, DDID left alone
for Dim 5, `confirmed` as the default acceptance gate, no label fuzzing.

But four of the spec's §8 assumptions are wrong or incomplete against live data, and one new
structural defect was found that affects the existing audit engine too:

1. **Truncation is not always right-truncation** — digit-preserving mid-name truncation exists
   (§1.1). Pure prefix matching misses it.
2. **DataType codes include `U`**, not just X/N/D — including on the *first business column of a
   pilot table* (§1.2). The type-class veto will misfire without a `U→char` mapping.
3. **The @-field exclusion must be restated for the new module** — `correspondence.py` is a new
   consumer of column lists and the spec never binds it to the filtered set (§1.3).
4. **Non-prefixed working columns exist inside T_* tables** (`X_PHANTOM` on
   T_EGP_COMPONENTS_INTERFACE) — bare-name derivation by prefix-strip mangles them, and they must
   be excluded from correspondence candidates (§1.4). This likely already pollutes the pilot
   findings as a phantom "extra column."

The §8.1 truncation window resolves to **27 = 30 − len(prefix)**; derive it, don't hardcode it
(§2.1). Suffix stripping must be truncation-aware (§2.2). The abbreviation table now has a
data-grounded seed (§3). Two open design questions were resolved per delegated judgment (§4).

---

## 1. REQUIRED CHANGES (must be reflected in the plan)

### 1.1 — Truncation rule must preserve trailing digit runs (spec §8.2 assumption violated)

**Confirmed live, in two pilot tables:** `T09GLOBAL_ATTRIBUTE_TIMESTAM10` (T_POZ_SUPPLIER_SITES_INT
row 184) and `T07GLOBAL_ATTRIBUTE_TIMESTAM10` (T_POZ_SUPPLIERS_INT row 139). The Oracle name is
`GLOBAL_ATTRIBUTE_TIMESTAMP10`; the `P` was dropped from the **middle** of the name to keep the
trailing ordinal `10` within the 30-char cap. The truncated bare is **not a prefix** of the Oracle
name, so prefix-match-within-window misses it — while `TIMESTAMP1`–`TIMESTAMP9` in the same block
match exactly. The result with the spec as written: exactly one false "missing field" per long
numbered series, repeated across the 147-table set — a pattern any consultant will spot, defeating
the project's trust goal.

Note the contrast that proves the mechanism: `T09ATTRIBUTE_TIMESTAMP10` (no `GLOBAL_` prefix, fits
at 30) is stored **untruncated** in DataDictionary. Truncation kicks in only when needed, and when
the name ends in digits the digits survive.

**Required rule:** when both the Oracle key and the Applaud bare end in a digit run, strip the
digit runs, require them to be **equal**, and apply the truncation-aware stem match to the
remainders. Add `GLOBAL_ATTRIBUTE_TIMESTAM10 ↔ GLOBAL_ATTRIBUTE_TIMESTAMP10` as a named test case
alongside the spec's `PROCUREMENT_BU` regression test (build-sequence step 6).

### 1.2 — DataType code `U` must map to character class (spec §8.7 assumption violated)

**Confirmed live:** 1,219 DataDictionary rows in ORACLE_MASTER carry DataType `U` (Unicode text);
297 are on T-prefixed names. This is not adjacent to the pilot — it is **inside it**:
`T07VENDOR_NAME` is `U(100)`, i.e. the first business column of T_POZ_SUPPLIERS_INT. The T05/T06/T08
families are also heavily `U`.

If `actual_shape` (audit_applaud.py:134-142) treats `U` as unknown and the type-class veto fires on
unknowns, every `U` column loses its correct name-match candidate. **Required:** extend the shape
mapping with `U → character class` (same bucket as `X`).

**Companion constraint, verified:** Applaud stores Oracle TIMESTAMP columns as `X(150)` and DATE
columns as `D(8)` (confirmed on the T09 attribute blocks). The veto must therefore remain strictly
**char-vs-numeric** as the spec says — never date-vs-char — or every timestamp column gets vetoed.
State this explicitly in the plan so it survives implementation.

### 1.3 — Bind the correspondence derivation to the @-excluded column set

The prior plan audit's second release blocker (exclude `@`-prefixed internal audit fields) was
fixed in the audit engine, but `correspondence.py` is a **new consumer** of Applaud column lists
and the spec never states its input is the filtered set. The pilot tables carry 26–49 @-fields each
(e.g., T09 rows 200–226, including extras beyond the standard block: `@T07PARTY_ID`,
`@T07EXPORT_BUCKET`, `@T91HEADER_NUM`). If derivation runs over raw snapshot columns, fields like
`@T07PARTY_ID` enter the residual and can fuzzy-hit real Oracle keys (Oracle suppliers FBDI does
carry party identifiers).

**Required:** one explicit sentence in the plan — derivation input is the same @-excluded column
set the four audit checks see, and a test asserting no `@`-origin bare ever appears in a derived
candidate.

### 1.4 — Handle non-prefixed columns: `X_PHANTOM` (new finding; affects the existing audit too)

**Confirmed live:** `T_EGP_COMPONENTS_INTERFACE` row 126 is `X_PHANTOM` — a working-variable-style
DDID registered as a physical table column with **no TableId prefix**. The snapshot's bare-name
derivation (strip the table prefix) will mangle it: stripping 3 chars yields the garbage bare
`HANTOM`, which then circulates through coverage checks and the correspondence residual.

**Required:**
- The snapshot/bare derivation must detect columns whose name does not start with the table's
  prefix and either pass them through unstripped or (preferred) tag them.
- Correspondence derivation must **exclude** non-prefixed columns from the candidate pool, same as
  @-fields.
- Recommended: scan the existing pilot findings — if `HANTOM` (or `X_PHANTOM` mis-stripped) appears
  as an extra-column finding in the PR #3 workbook, this defect already exists in the audit engine
  and the fix belongs there, not only in `correspondence.py`.

---

## 2. HIGH-PRIORITY corrections to the derivation ladder (spec §5)

### 2.1 — Truncation window: derive as `30 − len(prefix)`, don't hardcode

Confirmed: Applaud's name cap is 30 characters at the application level (the physical
`DataDictionary.Name` column is `TEXT(60)` — do not infer the cap from schema). The longest
observed bares are exactly 27 with 3-char prefixes (`ALLOW_SUBSTITUTERECEIPTSFLA`,
`PROCUREMENT_BUSINESSUNITNAM`, `UNIQUE_REMITTANCEIDENTIFIER`, `CONSUMPTION_ADVICELINENUMBE`,
`PARENT_SOURCESYSTEMREFERENC`, `LINE_ATTRIBUTECATEGORYLINES`, `REMIT_ADVICEDELIVERY_METHOD`,
`ORGANIZATION_TYPELOOKUPCODE`, `FINAL_DISCHARGELOCATIONCODE`, `DEF_ACCRUALCODECONCATENATED`).
Compute `TRUNCATION_WINDOW = 30 − len(prefix)` per table.

### 2.2 — Suffix stripping must be truncation-aware and run after underscore-collapse

The `_FLAG/_FLG/_F` list misses what's live:

- `ALLOW_SUBSTITUTERECEIPTSFLA` — the `FLAG` suffix itself truncated to `FLA`, with no underscore
  in front of it after collapse.
- `PROCUREMENT_BUSINESSUNITNAM` — a truncated `NAME` suffix, not in the list at all.
- `CONSUMPTION_ADVICELINENUMBE` — truncated `NUMBER`.
- `PARENT_SOURCESYSTEMREFERENC` — truncated `REFERENCE`.

**Recommended formulation that subsumes spec steps 2+3:** after normalization (underscore squash +
bidirectional abbreviation expansion + digit-run handling per §1.1), accept *"one normalized name
is a prefix of the other"* with a bounded length delta justified by the 27-char window plus known
suffix lengths (`FLAG`, `NAME`, `NUMBER`, `CODE`). This single rule catches clean right-truncations,
truncated suffixes, and added-then-truncated suffixes (Oracle `PROCUREMENT_BU` → expand →
`PROCUREMENT_BUSINESSUNIT` → Applaud appended `NAME` → truncated to `...UNITNAM`) without
enumerating fragments.

### 2.3 — Underscore collapse is selective and inconsistent across tables: squash both sides fully

Live names keep the first underscore and collapse later ones (`BUYER_MANAGEDTRANSPORTFLAG`,
`EXCLUDE_FREIGHTFROMDISCOUNT`, `HOLD_UNMATCHEDINVOICESFLAG`, `INVOICE_INCLUDESPREPAYFLAG`), collapse
even when the result lands under 27 (several are 26), and render the **same logical field two
different ways in two pilot tables**: `T07REMIT_ADVICEDELIVERY_METHOD` vs
`T09REMIT_ADVICEDELIVERYMETHOD`. Conclusion: collapse position carries no information. Full squash
on both sides is the only safe normalized comparison form; the spec's squash step is right —
the plan should make the squashed form the primary name-equality path, with the window applied to
squashed forms.

### 2.4 — Position signal: Row order reflects addition history, not Oracle layout

`T_EGO_ITEM_INTF_EFF_B` shows the column blocks appended in waves: `ATTRIBUTE_CHAR1–20`,
`NUMBER1–10`, `DATE1–5`, **then** `CHAR21–40`, `NUMBER11–20`, `DATE6–10`, `TIMESTAMP1–10`, then the
`_UOM_NAME` and `_UE` blocks. Applaud `Row` order diverges substantially from Oracle `position`
order whenever a table has absorbed release additions. The spec's choice to make position a 0.15
tiebreak (never sufficient alone) is correct — keep it that way, and the plan should state *why*
so nobody later "improves" the weight upward.

---

## 3. Abbreviation table — data-grounded seed (spec §8.3)

Built from divergences observed in the ten pilot tables. Two cautions first:

- **Abbreviation is a naming choice, not a length-fitting mechanism.** `ALWAYS_TAKE_DISCOUNT_FLAG`
  is only 25 chars and would have fit, yet the builder chose `ALWAYS_TAKE_DISC_FLAG`. Do not
  down-weight abbreviation candidates on short names.
- **Don't bloat the table with Oracle's own abbreviations.** Names like `DEF_ACCTG_START_DATE`,
  `TRX_BUSINESS_CATEGORY`, `VAT_REGISTRATION_NUM`, `AWT_GROUP_NAME` are Oracle FBDI's own spellings
  and will sit in the **exact pre-pass**. Seed the table only from post-exact residual divergences,
  which is exactly where the derive command operates.

Seed entries with live evidence:

| Abbrev | Expansion | Evidence |
|---|---|---|
| `BU` | `BUSINESSUNIT` / `BUSINESS_UNIT` | `O33PROCUREMENT_BU_NAME` ↔ `T09/T10PROCUREMENT_BUSINESSUNITNAM`, `T90PROCUREMENT_BUSINESS_UNIT` |
| `BUS` | `BUSINESS` | `TE1PROCUREMENT_BUS_UNIT_NAME`, `T07BUS_CLASS_NOT_APPLICABLE` |
| `DISC` | `DISCOUNT` | `T09ALWAYS_TAKE_DISC_FLAG` |
| `NUM` | `NUMBER` | `T07CUSTOMER_NUM`, `T91NEW_FROM_END_ITEM_UNIT_NUM`, `T99PO_SHIPMENT_NUM` |
| `DESCR` | `DESCRIPTION` | `T64ALLOW_DESCR_UPDATE_FLAG` |
| `DESC` | `DESCRIPTION` | `T91COMP_SOURCESYSTEMREFERDESC` |
| `AMT` | `AMOUNT` | `TA1AMT_APPL_TO_DISCOUNT`, `TA1ADD_TAX_TO_INV_AMT_FLAG` |
| `INV` | `INVOICE` | `T99PRICE_CORRECT_INV_NUM`, `T09GAPLESS_INV_NUM_FLAG` |
| `COMP` | `COMPONENT` | `T91COMP_SOURCESYSTEMREFERENCE` (table context: components interface) |
| `REFER` | `REFERENCE` | `T91COMP_SOURCESYSTEMREFERDESC` (compound: REFER+DESC) |

The specialist review (Brad) should extend this from the derive command's first residual output —
the review workbook's WEAK tier is effectively a worklist of missing abbreviation entries.

**One known case the table cannot resolve, for the review workbook:** T_BPA_PO_LINES_INTERFACE
names rows 24–38 `LINE_ATTRIBUTE1–15` but rows 39–43 `ATTRIBUTE16–20` (a dropped token, not a
truncation — `LINE_ATTRIBUTE16` at 16 chars would have fit). Whether Oracle's BPA FBDI names these
`LINE_ATTRIBUTE16–20` or `ATTRIBUTE16–20` determines whether these are exact matches or need map
rows. Cannot be resolved from the Applaud side; the reviewer settles it in the workbook. This is a
good canary that the HITL gate is earning its keep.

---

## 4. Resolved design questions (delegated to auditor judgment)

### 4.1 — `Corrected Bare` MUST be validated at confirm time; fail loud

`correspondence-confirm` must check a reviewer-entered `Corrected Bare` against the table's actual
(@-excluded, prefix-stripped) bare set and **reject the merge with a named error** if absent. A
typo'd bare otherwise becomes a permanent committed alias that maps to nothing: `build_alias`
emits an entry no column carries, the Oracle-side finding persists, and the reviewer believes it
was resolved — silent divergence between human intent and audit behavior, in a git-committed file.
Fail-loud at plan/merge time matches the repo's existing write contract.

### 4.2 — `rejected` rows annotate the finding's provenance; severity unchanged

When `run_audit` produces a missing-field finding whose `(table, oracle_key)` has a `rejected` map
row, append a provenance note to the finding (e.g., *"Reviewed — confirmed no Applaud
counterpart"*). Do **not** suppress or downgrade the finding: the gap is real; what changes is that
the consultant can distinguish "engine found nothing" from "a human verified nothing exists." This
is cheap (one map lookup in a loop that already holds the map) and directly serves the project's
stated purpose — a workbook safe to hand to a consultant. It also gives `rejected` rows a visible
payoff, which encourages reviewers to actually mark `N` instead of leaving rows undecided.

---

## 5. Assumptions verified and CONFIRMED (no spec change needed)

- **§8.4** — `DatabaseDetail.ODBCName`, `DataType`, and `Size` are empty in ORACLE_MASTER
  (re-confirmed on T_POZ_SUPPLIER_SITES_INT rows). `bare` from DataDictionary remains the right
  resolution key; no ODBCName-preference branch needed for this system.
- **§8.5 (one-to-one per table)** — no counterexample in ten tables. Lookalike clusters
  (`VENDOR_SITE_CODE` / `_NEW` / `_ALT`; `VENDOR_NAME` / `_NEW` / `_ALT`) are distinct Oracle keys
  that resolve in the exact pre-pass. Keep greedy assignment with conflicts recorded in `notes`.
- **§8.6** — map-key invariant is a code property; cover with the roundtrip tests as planned.
- **Dim-1 side benefit is real and already has a concrete instance:**
  `T09PROCUREMENT_BUSINESSUNITNAM` is `X(25)` while `T10PROCUREMENT_BUSINESSUNITNAM` is `X(40)` —
  the same logical Oracle field with divergent sizing that today's exact-match audit silently
  skips. Aliasing surfaces it. (This also intersects the known `VENDOR_SITE_CODE` resize-to-240
  issue.) The plan's load-bearing regression test (step 6) should assert the sizing check fires
  post-alias.

---

## 6. Additional implementation notes for the planner

- **Test fixture realism:** synthetic test names should include a digit-suffix truncation case, a
  truncated-suffix case (`...FLA`), a full-squash divergence pair (`REMIT_ADVICEDELIVERY_METHOD` vs
  `REMIT_ADVICEDELIVERYMETHOD`), a `U`-typed column, an `X_`-prefixed phantom column, and an
  `@`-field — one fixture table can carry all six.
- **Review workbook ordering:** sorting table → tier → score is right; additionally surface the
  per-table exact-match count in a header row so the reviewer sees denominator context ("212 of
  226 matched exactly; you are deciding the residual 14").
- **`_label_to_technical()` residual carry-over:** the prior pass's one open dependency still
  stands — verify its behavior on the first live catalog run; the correspondence layer reuses it
  (spec §4) and inherits the risk.
- **Re-derive idempotence across releases:** the merge semantics (confirmed/rejected win; derive
  fills only undecided pairs) are correct — add the 26B→next-release re-derive scenario to the
  merge tests so release turnover never clobbers decisions.

---

## 7. Instruction to the planner

Write the implementation plan against the spec **as amended by this file**. The four REQUIRED
items (§1.1–§1.4) must appear in the plan's assumptions and test list. Sections 2–4 are
design-level corrections the plan should incorporate directly. When the plan is written, hand it
back for the next audit pass, which will verify the plan's handling of each item above against the
live database again.
