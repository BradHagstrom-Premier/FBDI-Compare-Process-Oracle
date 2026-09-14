# Applaud Audit — Field-Correspondence Layer — Design

**Date:** 2026-06-10 (amended 2026-06-11 per second-pass audit)
**Status:** Audited — **second-pass audit complete and folded in; ready for the implementation plan**
**Predecessor:** `docs/superpowers/specs/2026-06-09-applaud-audit-first-run-design.md` (pilot run, merged via PR #3)
**Roadmap source:** `docs/superpowers/applaud-audit-first-run-notes.md` §Stage 4–5 (names this the #1 audit-quality blocker)
**Authoritative audit:** `docs/superpowers/AUDIT_RESULTS_field-correspondence.md` — live-verified against `ORACLE_MASTER` over all ten pilot tables. **Where this spec conflicts with the audit, the audit wins.** The audit's corrections are folded into the sections below; §8 records each §8 assumption's verified outcome.

---

## 1. Purpose

The first end-to-end pilot (PR #3) proved the audit engine is mechanically sound but its findings
are **not trustworthy**: 1,567 findings (957 HIGH) across 10 tables, overwhelmingly *false*
"missing field" positives.

**Root cause.** The audit matches Oracle FBDI fields to Applaud fields by **exact bare-name
equality** (`oracle_match_key` ∩ Applaud `bare`). Oracle and Applaud name the same logical field
differently, so every diverging field reads as missing:

| Oracle | Applaud | Divergence kind |
|---|---|---|
| `PROCUREMENT_BU` | `PROCUREMENT_BUSINESSUNITNAM` | abbreviation + right-truncation |
| `ALWAYS_TAKE_DISCOUNT` | `ALWAYS_TAKE_DISC_FLAG` | abbreviation + `_FLAG` suffix |
| `ALLOW_SUBSTITUTE_RECEIPTS` | `ALLOW_SUBSTITUTERECEIPTSFLA` | underscore-collapse + truncated `_FLAG` (→`FLA`) + truncation |
| `GLOBAL_ATTRIBUTE_TIMESTAMP10` | `GLOBAL_ATTRIBUTE_TIMESTAM10` | **digit-preserving mid-name truncation** — `P` dropped from the middle to keep the trailing ordinal `10`; the bare is **not a prefix** of the Oracle name (audit §1.1) |

This project builds a **field-correspondence layer** so confirmed Oracle↔Applaud field pairs count
as matches, and only *genuinely* unmatched fields become findings — making the workbook safe to
hand to a consultant.

**Hard constraint that shapes the whole design.** The Applaud snapshot carries **no
human-readable labels/descriptions** — Applaud's `DataDictionary` exposes only
Name/DataType/Size/DecPlaces (`tools/extract_applaud_snapshot.mjs`). So correspondence **cannot**
fuzzy-match on label text; it must work from bare technical names plus structural signals
(position/order within a table or IF/EF, type class, size). This rules out label-fuzzing.

**Chosen approach: Hybrid derive → confirm.** The engine *proposes* candidate correspondences
with confidence tiers; a human/specialist *confirms or overrides* via a review workbook; confirmed
pairs persist in a committed map. This is the **field-grain analogue of the existing app-map
`derived → confirmed` pattern** (`fbdi/applaud_appmap.py`) — clone that shape, don't invent a new
one. (Alternatives considered and rejected: *curated-only* — fully manual across 147 tables, too
much effort; *algorithmic-only* — no human gate, produces silent false matches and misses.)

---

## 2. Scope

| Dimension | Decision |
|---|---|
| Target system | **ORACLE_MASTER** (same reference master as the pilot) |
| Table scope | The **same 10 pilot tables** first, then generalizes to the 147-table set |
| What it matches | **Oracle field ↔ Applaud field** correspondence only (the audit boundary). Intra-Applaud matching stays exact-DDID (unchanged). |
| Write-back | **None.** The map is curated review state, not a write to the MDB. |
| Findings model | Unchanged — same 6 dimensions, same severities; only the *match set* each dimension sees changes. |

**Non-goals:** no new audit dimensions; no orchestrator skill (still Candidate C, separate); no
label-based fuzzing (Applaud has no labels); no auto-confirmation (confirmed-only is the default
gate).

---

## 3. The three-phase command model

Derive / confirm / audit stay separate, re-runnable commands — the same split already used for
snapshot / appmap / audit:

1. `python -m fbdi correspondence-derive --release 26B --system ORACLE_MASTER [--tables ...]`
   → emits a **review workbook** `Applaud_FieldMap_Review_<release>_<system>.xlsx` (gitignored,
   disposable). Loads the existing committed map first so already-decided pairs are not re-proposed.
2. `python -m fbdi correspondence-confirm --review <review.xlsx> [--map FBDI_to_Applaud_FieldMap.xlsx]`
   → merges confirmed / corrected / rejected rows into the committed map.
3. `python -m fbdi audit-applaud … --fieldmap FBDI_to_Applaud_FieldMap.xlsx [--accept-confidence confirmed]`
   → audit consumes the map automatically (loaded if present, mirroring `--appmap`).

---

## 4. New module — `fbdi/correspondence.py`

Pure-Python, no MCP/live I/O — same discipline as `applaud_appmap.py`. Exports:

- `@dataclass FieldCorrespondence(applaud_table, oracle_key, applaud_bare, applaud_ddid,
  confidence, origin="derived", score=0.0, signals="", notes="")` — `origin ∈ derived|confirmed|rejected`.
- **Normalization:** `normalize_name`, `name_tokens`, `ABBREVIATIONS` (data-grounded constant),
  bidirectional `expand_abbreviations`.
- **Derivation:** `score_candidate`, `derive_table_correspondences`, `derive_correspondences`.
- **Workbook I/O** (clone `applaud_appmap` signatures): `write_review_workbook`,
  `load_review_workbook`, `write_fieldmap_workbook`, `load_fieldmap_workbook`, `merge_fieldmap`.
- **Resolver:** `build_alias(fieldmap_for_table, accept_confidence) -> {applaud_bare_upper: oracle_key_upper}`.

Reuses, unchanged: `AlignedField` (`align.py:19`), `DataColumn`/`FileField`
(`applaud_snapshot.py:26-44`), `_label_to_technical` (`audit.py:39`), `expected_shape`/
`actual_shape` (`audit_applaud.py:127/134`), `align._lcs_match` (`align.py:61`). To avoid an import
cycle, `audit_applaud` does **not** import `correspondence` at module top — the fieldmap is loaded
in `cli.py` and passed into `run_audit`.

**Derivation input = the filtered column set (audit §1.3, §1.4).** `correspondence.py` is a *new*
consumer of Applaud column lists and must see exactly the set the four audit checks already see —
**not** the raw snapshot. Concretely, before derivation runs:

- **Exclude `@`-prefixed internal audit fields** (the prior plan's release blocker, fixed in the
  audit engine). The pilot tables carry 26–49 `@`-fields each, some beyond the standard block
  (`@T07PARTY_ID`, `@T07EXPORT_BUCKET`, `@T91HEADER_NUM`); if they reach the residual they can
  fuzzy-hit real Oracle keys (Oracle suppliers FBDI does carry party identifiers).
- **Exclude non-prefixed working columns** — e.g. `X_PHANTOM` on `T_EGP_COMPONENTS_INTERFACE`
  (row 126), a DDID registered as a physical column with **no TableId prefix**. Bare-name
  derivation (strip the 3-char prefix) mangles it into the garbage bare `HANTOM`. The
  snapshot/bare-derivation step must detect columns whose name does not start with the table's
  prefix and pass them through unstripped (preferred: tag them) rather than blind-stripping; the
  correspondence candidate pool then excludes them, same as `@`-fields.
- This X_PHANTOM mis-strip is an **existing defect in the audit engine**, not new to this layer:
  scan the PR #3 pilot workbook for `HANTOM` (or mis-stripped `X_PHANTOM`) appearing as an
  extra-column finding; if present, the prefix-strip fix belongs in the snapshot/audit path and
  `correspondence.py` merely inherits it.

---

## 5. Derivation algorithm (per Applaud table)

1. **Exact pre-pass.** `exact = oracle_keys ∩ applaud_bares` — these need **no** map entry (the
   audit already matches them). An empty map ⇒ today's pure-exact behavior. Derive only over the
   residual `oracle_unmatched × applaud_unmatched`. (The abbreviation table in step 4 is seeded
   from *post-exact residual* divergences only — see §8.3.)
2. **Name ladder.** Upper / strip `*`. **Full underscore squash on both sides** is the primary
   name-equality form (audit §2.3): collapse position carries no information — live names keep the
   first underscore and drop later ones inconsistently, and render the *same* logical field two
   ways across tables (`T07REMIT_ADVICEDELIVERY_METHOD` vs `T09REMIT_ADVICEDELIVERYMETHOD`). Squash
   both, then tokenize. Boolean-suffix stripping (`_FLAG`/`_FLG`/`_F`) must run **after** squash and
   be **truncation-aware** — the live data carries truncated suffixes the literal list misses
   (`...FLA` for `FLAG`, `...NAM` for `NAME`, `...NUMBE` for `NUMBER`, `...REFERENC` for
   `REFERENCE`). See step 3 for the rule that subsumes suffix enumeration.
3. **Truncation-aware match (audit §1.1, §2.1, §2.2).**
   - **Window is derived, not hardcoded:** `TRUNCATION_WINDOW = 30 − len(prefix)`. Applaud's cap is
     30 chars at the *application* level (do not infer it from `DataDictionary.Name`'s `TEXT(60)`
     schema). The longest observed bares are exactly 27 with 3-char prefixes.
   - **Core rule:** after normalization (squash + bidirectional abbreviation expansion + digit-run
     handling below), accept *"one normalized name is a prefix of the other"* within a length delta
     bounded by the window plus known suffix lengths (`FLAG`/`NAME`/`NUMBER`/`CODE`). This single
     rule catches clean right-truncation, truncated suffixes, and appended-then-truncated suffixes
     (`PROCUREMENT_BU` → expand → `PROCUREMENT_BUSINESSUNIT` → Applaud appends `NAME` → truncates to
     `...UNITNAM`) without enumerating fragments. A length-class gate still guards against a
     genuinely short coincidental prefix being read as a truncation hit.
   - **Digit-run preservation (REQUIRED — audit §1.1):** truncation is *not* always right-truncation.
     When the trailing token is an ordinal, Oracle drops a letter from the **middle** to keep the
     digits within the cap (`GLOBAL_ATTRIBUTE_TIMESTAMP10` → `GLOBAL_ATTRIBUTE_TIMESTAM10`), so the
     bare is not a prefix. Rule: when **both** the Oracle key and the Applaud bare end in a digit
     run, strip the digit runs, **require them to be equal**, and apply the truncation-aware stem
     match to the remainders. Without this, every long numbered series throws exactly one false
     "missing field" — a pattern any consultant spots, defeating the trust goal.
4. **Abbreviation table (bidirectional).** A committed constant expanded on both sides before token
   equality — **data-grounded seed in §8.3** (`BU→BUSINESSUNIT`, `DISC→DISCOUNT`, `NUM→NUMBER`,
   `DESC(R)→DESCRIPTION`, `AMT→AMOUNT`, `INV→INVOICE`, `COMP→COMPONENT`, `REFER→REFERENCE`,
   `BUS→BUSINESS`, …). **Highest-risk correctness input; specialist-reviewed and extended from the
   derive command's first residual output (§8.3).** Two guardrails: abbreviation is a *naming
   choice, not a length-fitting mechanism* (so do **not** down-weight abbreviation candidates on
   short names — `ALWAYS_TAKE_DISC_FLAG` is only 25 chars and would have fit), and the table is
   seeded only from post-exact residual divergences, **not** Oracle's own spellings (`DEF_ACCTG_…`,
   `VAT_REGISTRATION_NUM`) which already match in the exact pre-pass.
5. **Type/size as tiebreak, never sole evidence.** Type-class agreement raises confidence; a
   char-vs-numeric clash **vetoes** a name-only candidate. The veto is **strictly char-vs-numeric,
   never date-vs-char** (audit §1.2): Applaud stores Oracle TIMESTAMP columns as `X(150)` and DATE
   columns as `D(8)`, so a date-vs-char veto would kill every timestamp column. The shape mapping
   **must include `U → character class`** (Unicode text, same bucket as `X`) — `U` is live *inside*
   the pilot (`T07VENDOR_NAME` is `U(100)`, the first business column of T_POZ_SUPPLIERS_INT; 1,219
   `U` rows in ORACLE_MASTER, 297 on T-prefixed names). Without `U→char`, the veto fires on every
   `U` column and kills its correct name-match candidate.
6. **Position/order alignment.** `align._lcs_match` over the residual lists (Oracle `position`
   order vs Applaud `row` order) — a tiebreak bonus only (0.15 weight), **never sufficient alone,
   and the weight must not be raised** (audit §2.4). Applaud `Row` order reflects *addition history*,
   not Oracle layout: tables that absorbed release additions show appended waves
   (`ATTRIBUTE_CHAR1–20`, then later `CHAR21–40`, `TIMESTAMP1–10`, …) that diverge substantially
   from Oracle `position` order. Keep position weak so nobody "improves" it upward.
7. **Confidence tiers** from a single weighted score (name 0.6 / type 0.25 / position 0.15),
   bucketed into `EXACT` (not persisted), `HIGH`, `PROBABLE`, `WEAK`. Bands are tunable in one place.
8. **One-to-one bijection per table.** Greedy assignment: sort candidate pairs by (tier, score)
   desc; accept a pair only if both its Oracle key and Applaud bare are still free; record losers
   in `notes`/conflicts so the reviewer sees why a plausible link was dropped. Leftover Oracle keys
   remain genuine findings; leftover Applaud columns remain genuine extras.

---

## 6. Persistence & review workflow

### Committed map
`FBDI_to_Applaud_FieldMap.xlsx` at repo root, sheet `"Field Map"`, columns:
`Applaud Table | Oracle Key | Applaud Bare | Applaud DDID | Confidence | Origin | Notes`.
- **A workbook, not JSON** — because the HITL reviewer edits it in Excel (same rationale the
  app-map is a workbook while the machine-extracted snapshot is JSON).
- **Committed** (mirrors the committed `FBDI_to_Applaud_AppMap.xlsx`): human-curated decisions must
  be git-auditable. Only divergent fields get rows, so it stays small. Add a `.gitignore`
  allow-comment for the map and an ignore-glob for the disposable `Applaud_FieldMap_Review_*.xlsx`.
- **Invariant:** `Oracle Key` is always exactly `oracle_match_key(of)` so the aliased Applaud bare
  set-intersects with the Oracle side.

### Review workbook (the HITL gate)
`Applaud_FieldMap_Review_<release>_<system>.xlsx` (gitignored), one row per derived candidate,
sorted table → tier → score. Columns include `Oracle Type`, `Candidate Applaud Bare`,
`Applaud Type`, `Confidence`, `Score`, `Signals` (human-readable breakdown),
`Conflicts/Alternatives` (runner-up bares for override), and the reviewer inputs
`Confirm?` (`Y`/`N`) + `Corrected Bare`.

`correspondence-confirm` merges decisions back: `Y` → `Origin=confirmed`; `Corrected Bare` →
confirmed with the substitute; `N` → `Origin=rejected` (suppresses re-proposal). Merge semantics
clone `merge_appmap` (`applaud_appmap.py:165`): **confirmed/rejected rows win**; a fresh derive
fills only undecided `(table, oracle_key)` pairs, so derivation is re-runnable across releases
without clobbering human decisions.

**`Corrected Bare` is validated at confirm time — fail loud (audit §4.1).** `correspondence-confirm`
must check any reviewer-entered `Corrected Bare` against the table's actual (@-excluded,
prefix-stripped) bare set and **reject the merge with a named error** if absent. A typo'd bare
otherwise becomes a permanent committed alias mapping to nothing: `build_alias` emits an entry no
column carries, the Oracle-side finding silently persists, and the reviewer believes it resolved.
Fail-loud at merge time matches the repo's existing write contract.

**`rejected` rows annotate finding provenance — severity unchanged (audit §4.2).** When `run_audit`
produces a missing-field finding whose `(table, oracle_key)` carries a `rejected` map row, append a
provenance note (e.g. *"Reviewed — confirmed no Applaud counterpart"*). Do **not** suppress or
downgrade it — the gap is real; what changes is that the consultant can distinguish "engine found
nothing" from "a human verified nothing exists." One map lookup in a loop that already holds the
map; it gives `rejected` rows a visible payoff so reviewers mark `N` instead of leaving rows
undecided.

---

## 7. Audit integration (minimal — one function changes, the four checks untouched)

The single-point match-key design lets us alias the **Applaud side** before set-intersection. In
`run_audit` (`audit_applaud.py:513-579`), inside the per-table loop right after
`table = snapshot.tables.get(table_name)` (`:534`):

1. `alias = build_alias(fieldmap[table_name], accept_confidence)` → `{applaud_bare_upper: oracle_key_upper}`.
2. Build **aliased copies** of `table.columns` and the IF/EF `FileField`s via
   `dataclasses.replace(c, bare=alias.get(c.bare.upper(), c.bare))`; pass those copies into the
   existing `check_sizing` / `check_table_coverage` / `check_file_coverage` / `check_release_delta`
   calls (`:545-568`). The four check functions keep their signatures — all new logic lives in
   `run_audit`.
3. **Leave DDID untouched** — Dim 5 orphans match on DDID, which is correct; only `bare` is aliased.
4. Default gate `--accept-confidence confirmed`; allow `HIGH` for a pre-review noise-reduction pass.

**Two engine fixes this layer depends on (audit §1.2, §1.4) live in the audit/snapshot path, not in
`correspondence.py`:**
- `actual_shape` (`audit_applaud.py:134-142`) must map `U → character class` (same bucket as `X`),
  or the type-class veto misfires on the 297 T-prefixed `U` columns. Keep the veto char-vs-numeric
  only (never date-vs-char) — Applaud stores TIMESTAMP as `X(150)`, DATE as `D(8)`.
- The snapshot/bare-derivation prefix-strip must not blind-strip non-prefixed columns (`X_PHANTOM`
  → `HANTOM`). Fix it where the strip happens; check whether the PR #3 workbook already shows
  `HANTOM` as an extra-column finding (it likely does — pre-existing defect).

**Side benefit (intended), now with a concrete instance (audit §5):** Dim 1 `check_sizing` matches
on `bare`, so aliasing also **enables sizing checks on renamed fields** that were previously
silently skipped. Live example: `T09PROCUREMENT_BUSINESSUNITNAM` is `X(25)` while
`T10PROCUREMENT_BUSINESSUNITNAM` is `X(40)` — the same logical Oracle field with divergent sizing
that today's exact-match audit skips; aliasing surfaces it (this intersects the known
`VENDOR_SITE_CODE` resize-to-240 issue). The load-bearing regression test (build step 6) should
assert the sizing check **fires post-alias**.

Only `run_audit` gains a `fieldmap` param; `cli.py` `_run_audit_applaud` (`:453-517`) loads the map
(if present, like `--appmap`) and passes it through.

---

## 8. Second-pass audit outcomes (live-verified against ORACLE_MASTER)

The original §8 listed seven assumptions to verify. The audit
(`AUDIT_RESULTS_field-correspondence.md`) verified all ten pilot tables live. Outcomes:

1. **Truncation width — CORRECTED to a derived value.** Cap is **30 chars at the application
   level** (the `DataDictionary.Name` schema is `TEXT(60)` and must *not* be used to infer it);
   longest observed bares are exactly 27 with 3-char prefixes. Use `TRUNCATION_WINDOW = 30 −
   len(prefix)` per table — folded into §5 step 3.
2. **Truncation is NOT always right-truncation — CORRECTED (REQUIRED).** Digit-preserving mid-name
   truncation exists (`GLOBAL_ATTRIBUTE_TIMESTAMP10` → `…TIMESTAM10`); pure prefix matching misses
   it. Digit-run-equality rule folded into §5 step 3 (audit §1.1).
3. **`ABBREVIATIONS` — data-grounded seed provided below (§8.3).** Highest-risk input; seeded from
   post-exact residual divergences in the ten pilot tables, extended by the reviewer from the
   derive command's first residual (the WEAK tier is effectively the missing-abbreviation worklist).
4. **`bare` is the right key — CONFIRMED.** `DatabaseDetail.ODBCName`, `DataType`, and `Size` are
   empty in ORACLE_MASTER. No ODBCName-preference branch needed for this system.
5. **One-to-one per table — CONFIRMED.** No counterexample in ten tables; lookalike clusters
   (`VENDOR_SITE_CODE`/`_NEW`/`_ALT`) are distinct Oracle keys that resolve in the exact pre-pass.
   Keep greedy assignment with conflicts recorded in `notes`.
6. **Map `Oracle Key == oracle_match_key(of)` — CONFIRMED** as a code property; covered by the
   roundtrip tests.
7. **`DataType` codes are only X/N/D — CORRECTED (REQUIRED).** Code **`U`** (Unicode text) is live
   *inside* the pilot (1,219 rows in ORACLE_MASTER, 297 T-prefixed; `T07VENDOR_NAME` is `U(100)`).
   Add `U → character class` and keep the veto char-vs-numeric only — folded into §5 step 5 and §7
   (audit §1.2).

**New structural finding not in the original §8 — non-prefixed working columns (REQUIRED, audit
§1.4):** `X_PHANTOM` (no TableId prefix) is mis-stripped to `HANTOM`. Excluded from the candidate
pool and fixed at the snapshot/audit strip — see §4 and §7.

### 8.3 — Abbreviation table seed (data-grounded, audit §3)

Seeded from post-exact residual divergences in the ten pilot tables. Bidirectional; expanded on
both sides before token equality. **Do not** add Oracle's own spellings (those match in the exact
pre-pass), and **do not** down-weight abbreviation candidates on short names (abbreviation is a
naming choice, not length-fitting — `ALWAYS_TAKE_DISC_FLAG` is 25 chars and would have fit).

| Abbrev | Expansion | Evidence |
|---|---|---|
| `BU` | `BUSINESSUNIT` / `BUSINESS_UNIT` | `O33PROCUREMENT_BU_NAME` ↔ `…PROCUREMENT_BUSINESSUNITNAM` |
| `BUS` | `BUSINESS` | `TE1PROCUREMENT_BUS_UNIT_NAME`, `T07BUS_CLASS_NOT_APPLICABLE` |
| `DISC` | `DISCOUNT` | `T09ALWAYS_TAKE_DISC_FLAG` |
| `NUM` | `NUMBER` | `T07CUSTOMER_NUM`, `T99PO_SHIPMENT_NUM` |
| `DESCR` | `DESCRIPTION` | `T64ALLOW_DESCR_UPDATE_FLAG` |
| `DESC` | `DESCRIPTION` | `T91COMP_SOURCESYSTEMREFERDESC` |
| `AMT` | `AMOUNT` | `TA1AMT_APPL_TO_DISCOUNT`, `TA1ADD_TAX_TO_INV_AMT_FLAG` |
| `INV` | `INVOICE` | `T99PRICE_CORRECT_INV_NUM`, `T09GAPLESS_INV_NUM_FLAG` |
| `COMP` | `COMPONENT` | `T91COMP_SOURCESYSTEMREFERENCE` (components-interface context) |
| `REFER` | `REFERENCE` | `T91COMP_SOURCESYSTEMREFERDESC` (compound: REFER+DESC) |

**One case the table cannot resolve, for the review workbook:** `T_BPA_PO_LINES_INTERFACE` names
rows 24–38 `LINE_ATTRIBUTE1–15` but rows 39–43 `ATTRIBUTE16–20` (a *dropped token*, not a
truncation — `LINE_ATTRIBUTE16` at 16 chars would have fit). Whether Oracle's BPA FBDI names these
`LINE_ATTRIBUTE16–20` or `ATTRIBUTE16–20` decides exact-match vs map-row, and cannot be resolved
from the Applaud side — the reviewer settles it in the workbook. A good canary that the HITL gate
earns its keep.

Per the handback convention: where this design conflicts with `AUDIT_RESULTS_field-correspondence.md`,
the audit-results file wins.

---

## 9. Build sequence (each step independently testable)

0. **Spec written and second-pass-audited (this doc). GATE cleared — ready for the plan.**
1. `correspondence.py` normalization + abbreviation primitives. (tests: full-squash equality,
   truncation-aware suffix strip incl. `…FLA`/`…NAM`, normalize-equal, `@`-origin bare never enters
   a candidate, non-prefixed `X_PHANTOM` excluded/not mis-stripped)
2. Scoring + tiers + `derive_table_correspondences`. (tests: right-truncation,
   **digit-run case `GLOBAL_ATTRIBUTE_TIMESTAM10 ↔ GLOBAL_ATTRIBUTE_TIMESTAMP10`**,
   `PROCUREMENT_BU`↔`…BUSINESSUNITNAM`, `U`-column survives the veto, char-vs-numeric veto fires,
   date-vs-char does **not** veto, coincidental short prefix does **not** match, bijection, tier bands)
3. Fieldmap workbook I/O + `merge_fieldmap`. (tests: roundtrip; confirmed/rejected win over
   re-derive; **26B→next-release re-derive idempotence** — release turnover never clobbers decisions)
4. Review workbook emit + `load_review_workbook` + confirm-merge. (tests: `Y` / `Corrected Bare` /
   `N`; **`Corrected Bare` absent from the table's bare set → named error, merge rejected**;
   `rejected` row → finding gains a provenance note, severity unchanged)
5. `build_alias` resolver + confidence gate.
6. Wire `run_audit` aliasing. **Load-bearing regression test:** `PROCUREMENT_BU` Oracle field vs
   `PROCUREMENT_BUSINESSUNITNAM` Applaud column → HIGH with an empty map; a confirmed alias makes
   the finding vanish **and** fires the Dim-1 sizing check post-alias (the `X(25)` vs `X(40)`
   divergence from §8 surfaces).
7. CLI: `correspondence-derive`, `correspondence-confirm`, `audit-applaud --fieldmap/--accept-confidence`.
8. Operational (post-gate): derive over the 10 pilot tables, confirm via the review workbook, commit
   `FBDI_to_Applaud_FieldMap.xlsx`; add the `.gitignore` entries.

---

## 10. Testing approach

Follows the repo convention (synthetic objects inline per test, lowercase free-text data per the
CLAUDE.md header-detection gotcha; identifiers stay UPPER as they are real). New
`tests/test_correspondence.py` covers normalization, abbreviation, truncation (incl. a coincidental
short prefix that must **not** match), type-veto, bijection, tier bucketing, workbook roundtrip,
merge precedence, and the review/confirm flow. The audit-integration regression lives in
`tests/test_audit_applaud.py` (step 6). Full suite (currently 389) stays green.

**Fixture realism (audit §6).** One synthetic fixture table can carry all six edge cases at once —
include them so the ladder is exercised end-to-end: a digit-suffix truncation case
(`…TIMESTAM10`), a truncated-suffix case (`…FLA`), a full-squash divergence pair
(`REMIT_ADVICEDELIVERY_METHOD` vs `REMIT_ADVICEDELIVERYMETHOD`), a `U`-typed column, an
`X_`-prefixed non-prefixed phantom column, and an `@`-field. Surface the per-table exact-match
count in the review workbook header so the reviewer sees denominator context ("212 of 226 matched
exactly; you are deciding the residual 14"). Carry-over risk to watch: `_label_to_technical()`'s
residual behavior on the first live catalog run — the correspondence layer reuses it and inherits
the risk.

---

## 11. Verification (end-to-end)

After the gate clears and implementation lands: `correspondence-derive` → confirm → re-run
`audit-applaud` on the 10 pilot tables, and confirm the HIGH "missing field" count collapses from
957 toward the genuine residual — `T_BANKS_BRANCHES` keeps only its real EDI/EFT_ID divergence;
POZ/BPA tables drop their false missings. Compare HIGH counts before/after against the first-run
notes distribution table.
