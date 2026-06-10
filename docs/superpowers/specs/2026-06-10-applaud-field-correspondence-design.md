# Applaud Audit — Field-Correspondence Layer — Design

**Date:** 2026-06-10
**Status:** Draft (brainstorming) — **awaiting second-pass audit by the Applaud-specialist project before implementation**
**Predecessor:** `docs/superpowers/specs/2026-06-09-applaud-audit-first-run-design.md` (pilot run, merged via PR #3)
**Roadmap source:** `docs/superpowers/applaud-audit-first-run-notes.md` §Stage 4–5 (names this the #1 audit-quality blocker)

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
| `ALLOW_SUBSTITUTE_RECEIPTS` | `ALLOW_SUBSTITUTERECEIPTSFLA` | underscore-collapse + `_FLAG` + truncation |

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

---

## 5. Derivation algorithm (per Applaud table)

1. **Exact pre-pass.** `exact = oracle_keys ∩ applaud_bares` — these need **no** map entry (the
   audit already matches them). An empty map ⇒ today's pure-exact behavior. Derive only over the
   residual `oracle_unmatched × applaud_unmatched`.
2. **Name ladder.** Upper / strip `*`; strip boolean suffixes `_FLAG`/`_FLG`/`_F`; collapse
   underscores to a squashed form; tokenize on `_`.
3. **Truncation-aware match.** Prefix match within a `TRUNCATION_WINDOW` (≈25–27 chars — **verify
   live, §8.1**) so Applaud's right-truncated names match their full Oracle counterparts. A
   length-class gate avoids treating a genuinely short coincidental prefix as a truncation hit.
4. **Abbreviation table (bidirectional).** A committed constant — `BU→BUSINESSUNIT`,
   `DISC→DISCOUNT`, `ORG→ORGANIZATION`, `NUM→NUMBER`, `DESC→DESCRIPTION`, `AMT→AMOUNT`,
   `QTY→QUANTITY`, … — expanded on both sides before token equality. **Highest-risk correctness
   input; must be specialist-reviewed (§8.3).**
5. **Type/size as tiebreak, never sole evidence.** Type-class agreement raises confidence; a
   char-vs-numeric clash **vetoes** a name-only candidate.
6. **Position/order alignment.** `align._lcs_match` over the residual lists (Oracle `position`
   order vs Applaud `row` order) — a tiebreak bonus only, never sufficient alone.
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

**Side benefit (intended):** Dim 1 `check_sizing` matches on `bare`, so aliasing also **enables
sizing checks on renamed fields** that were previously silently skipped — a latent gap this fixes.

Only `run_audit` gains a `fieldmap` param; `cli.py` `_run_audit_applaud` (`:453-517`) loads the map
(if present, like `--appmap`) and passes it through.

---

## 8. Assumptions to verify in the second-pass audit (live Applaud data)

These drive correctness and **must be checked against the real MDB before implementation**:

1. **Truncation width** is ~27 chars and consistent (drives `TRUNCATION_WINDOW`). Verify the actual
   max bare-name length in `DataDictionary` for ORACLE_MASTER.
2. **Truncation is right-truncation** of the full logical name, not a different abbreviation scheme.
   (`PROCUREMENT_BUSINESSUNITNAM` dropping trailing `E` supports this — confirm across more samples.)
3. **The `ABBREVIATIONS` table is correct/complete** for the 10 pilot tables — the single
   highest-risk input. The specialist should review and extend it against real column names.
4. **`bare` is the right resolution key** and `odbc_name` is empty for ORACLE_MASTER. If some tables
   *do* populate ODBCName, the resolver should prefer it (Dim 4 already matches on it,
   `audit_applaud.py:298`).
5. **One-to-one per table holds** — no Oracle field legitimately maps to two Applaud columns (or
   vice versa) within a table. The greedy bijection assumes this.
6. **Map `Oracle Key` == `oracle_match_key(of)`** exactly, so aliased bares set-intersect.
7. **Applaud `DataType` codes are only X/N/D** as `actual_shape` (`audit_applaud.py:134-142`)
   assumes — else the type-class veto misfires.

Per the handback convention: where this design conflicts with the specialist's
`AUDIT_RESULTS_*.md`, the audit-results file wins.

---

## 9. Build sequence (each step independently testable)

0. **Write this spec → route for second-pass audit. GATE — implementation waits on it.**
1. `correspondence.py` normalization + abbreviation primitives. (tests: suffix strip, normalize-equal)
2. Scoring + tiers + `derive_table_correspondences`. (tests: truncation,
   `PROCUREMENT_BU`↔`…BUSINESSUNITNAM`, type-veto, bijection, tier bands)
3. Fieldmap workbook I/O + `merge_fieldmap`. (tests: roundtrip; confirmed/rejected win over re-derive)
4. Review workbook emit + `load_review_workbook` + confirm-merge. (tests: `Y` / `Corrected Bare` / `N`)
5. `build_alias` resolver + confidence gate.
6. Wire `run_audit` aliasing. **Load-bearing regression test:** `PROCUREMENT_BU` Oracle field vs
   `PROCUREMENT_BUSINESSUNITNAM` Applaud column → HIGH with an empty map; a confirmed alias makes
   the finding vanish **and** enables the Dim-1 sizing check.
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

---

## 11. Verification (end-to-end)

After the gate clears and implementation lands: `correspondence-derive` → confirm → re-run
`audit-applaud` on the 10 pilot tables, and confirm the HIGH "missing field" count collapses from
957 toward the genuine residual — `T_BANKS_BRANCHES` keeps only its real EDI/EFT_ID divergence;
POZ/BPA tables drop their false missings. Compare HIGH counts before/after against the first-run
notes distribution table.
