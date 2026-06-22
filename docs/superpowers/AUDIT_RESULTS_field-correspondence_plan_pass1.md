# AUDIT RESULTS — Field-Correspondence Implementation Plan (Pass 1)

**Source:** Technical audit of `2026-06-11-applaud-field-correspondence.md` against the spec, the
prior audit handback (`AUDIT_RESULTS_field-correspondence.md`), and the live ORACLE_MASTER data
behind it. The plan's matching algorithm was traced line-by-line and executed against its own test
cases plus the audit's live counterexamples.
**Date:** 2026-06-11

This file is the authoritative correction set for the plan. Where it conflicts with the plan,
this file wins. Fix the two BLOCKERS and the MEDIUMs below, then hand the revised plan back for
the clearance pass.

---

## 0. TL;DR

The plan is structurally strong: TDD throughout, names locked across tasks, no placeholders, no
import cycle, and all four REQUIRED audit items have dedicated tasks with tests. The matching core
is **correct** — I executed `normalize_name`/`names_correspond` as written against the live
counterexamples and all pass: digit-run (`GLOBAL_ATTRIBUTE_TIMESTAM10`), truncated `FLA` suffix,
`DISC`+`_FLAG` combo, appended-then-truncated `NAM`. The X_PHANTOM deviation (audit §1.4 corrected
to a lock-in test after live verification of `build_table`/`_strip_prefix` and a zero-`HANTOM`
workbook scan) is a sound, well-evidenced resolution — accepted.

Two defects block release, both in the workflow wiring rather than the algorithm:

1. **Reviewer decisions are immutable after first commit** — `correspondence-confirm` merges with
   committed-wins precedence, so a wrong confirmation can never be corrected through the tooling
   (§1.1).
2. **Rejected-provenance may never reach the consultant** — the plan asserts `f.notes` on Finding
   and never touches the workbook writer; if Finding lacks a notes field or the writer doesn't
   emit it, audit §4.2 silently becomes a no-op (§1.2).

Four MEDIUMs (§2) and a set of LOW/cleanup items (§3) follow.

---

## 1. BLOCKERS

### 1.1 — `correspondence-confirm` cannot change a prior decision (merge precedence inverted at confirm time)

`_run_correspondence_confirm` calls `merge_fieldmap(decisions, committed)`. `merge_fieldmap` gives
the **committed** side absolute precedence — incoming rows fill only keys absent from the
committed map. That is the correct semantic at **derive** time ("a fresh derive fills only
undecided pairs"). At **confirm** time it is backwards: once `(table, oracle_key)` has any row in
the committed map, every future reviewer decision for that key is **silently discarded** — the
command even prints "Merged N decision(s)" while dropping them.

Compounding it: `correspondence-derive` skips `decided` pairs, so the pair never reappears in a
review workbook either. Net effect: the only way to fix a wrong confirmation (a realistic event —
the `LINE_ATTRIBUTE16–20` canary is exactly the kind of call a reviewer may revise) is hand-editing
the committed xlsx, defeating the tooling and the git audit trail.

**Required fix:** at confirm time, **new human decisions win**. Either add a
`merge_decisions(decisions, committed)` (decisions override; untouched committed rows carry
forward), or give `merge_fieldmap` an explicit precedence parameter — but do not reuse
derive-precedence for confirm. Add the missing test:

```python
def test_confirm_overrides_prior_decision():
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "OLD_BARE", origin="confirmed")]}
    new = [_fc("T_POZ", "PROCUREMENT_BU", "NEW_BARE", origin="confirmed",)]
    merged = merge_decisions(new, committed)
    assert merged["T_POZ"][0].applaud_bare == "NEW_BARE"
```

Also make the precedence invariant explicit: the committed map should only ever contain
`confirmed`/`rejected` rows (the confirm flow guarantees this today, but nothing enforces it).
Either assert/strip `origin="derived"` rows in `load_fieldmap_workbook`, or make the merge
origin-aware so an incoming confirmed/rejected always replaces a committed derived. Otherwise a
hand-edited or future-flow derived row in the map becomes another silent decision-blocker.

### 1.2 — Rejected-provenance (audit §4.2) must demonstrably reach the Excel workbook

Task 7 annotates findings via `f.notes = "Reviewed — confirmed no Applaud counterpart"` and the
test asserts `miss[0].notes`. Two unverified dependencies, neither addressed by any task:

1. **Does `Finding` have a mutable `notes` field?** No task modifies the findings model. If the
   dataclass is frozen or lacks the field, Task 7 Step 3 doesn't compile against reality.
2. **Does the workbook writer emit it?** The entire point of §4.2 is *consultant-visible*
   provenance. If the Excel writer has no Notes column, the annotation exists only in memory and
   the feature is a silent no-op — the in-memory test would pass while the deliverable stays
   unchanged.

**Required fix:** Task 7 must (a) verify or add `notes` on `Finding` (with the writer column), and
(b) extend the regression test to assert the provenance string appears **in the written workbook**
(re-open `out_path` with openpyxl and find the cell), not only on the in-memory finding. If the
findings model already has notes end-to-end, this collapses to a one-line confirmation in the
plan's assumptions — but it must be stated, because right now the plan asserts an attribute it
never establishes.

---

## 2. MEDIUM — fix before implementation

### 2.1 — Gate and score use different Oracle strings (`okey` vs `of.technical`)

`derive_table_correspondences` gates candidates with `_name_score(okey, ...)` but
`score_candidate` recomputes the name score from `of.technical or ""`. When
`oracle_match_key(of) != of.technical` — exactly the `_label_to_technical()` path flagged as the
standing residual, where `technical` is empty and the key is derived from the label — the
candidate **passes the gate and then scores 0.0 on name** (verified by execution: an empty
oracle string matches nothing). Result: a correct correspondence is mis-tiered into WEAK with
`name=0.00` signals, confusing the reviewer. It also computes the name score twice per pair.

**Fix:** compute the name score once in the loop, pass it (or `okey`) into `score_candidate` —
never re-derive from `of.technical`. Add a test where `technical=None` and the key comes from the
label, asserting the tier is computed from the key.

### 2.2 — `assemble_derivation_inputs` silently drops tabs on many-tabs→one-table mappings

`out[table_name] = (prefix, oracle_by_key, cols)` overwrites: if two `(template, tab)` mapping
rows resolve to the same Applaud table, only the last tab's Oracle keys survive, and the dropped
tab's divergent fields silently never get correspondence candidates. The mapping workbook is known
to contain multi-mapping rows (the "unique IFs and EFs for each Oracle FBDI" flags from the
first-pass audit). `run_audit` handles this naturally because it loops per `(template, tab)`;
the derivation assembly must not be lossier than the audit it serves.

**Fix:** merge instead of overwrite —
`out[table_name][1].update(oracle_by_key)` (keep first prefix/cols; assert prefix agreement).
Add a test with two tabs mapping to one table, asserting keys from **both** appear.

### 2.3 — Rejected-annotation filter misses Dim 2-IF / 3-EF findings, and never appends

The annotation loop requires `f.applaud_object_name == table_name`, but the same rejected Oracle
key also produces missing-field findings in `check_file_coverage` (Dim 2-IF / 3-EF), which carry
the **IF/EF object name** — those findings stay un-annotated, so the workbook shows the same key
as both "reviewed" and "unreviewed" across dimensions. The `findings[n_before:]` slice already
scopes to this table's iteration; drop the object-name condition and match on
`current_value == "absent" and oracle_field in rejected_keys`. Secondly,
`f.notes = (... if not f.notes else f.notes)` silently discards the provenance whenever a note
already exists — append (`"; ".join` style) instead of keep-old.

### 2.4 — Position signal: dead code and an order inconsistency

Task 3 imports `_lcs_match` and builds `o_order`/`a_order` (row-sorted) — then uses none of them:
`_position_score` is an index-distance gradient over `enumerate(residual_cols)` in **original list
order**, not the row-sorted order it just computed. The weak-tiebreak design (audit §2.4) is fine
and the simple gradient is acceptable — but pick one implementation: either (a) delete the
`_lcs_match` import and `o_order`/`a_order`, and feed `a_idx` from the row-sorted list so the
gradient means what it claims, or (b) actually use `_lcs_match` as the spec stated. As written,
the unused import is a lint failure and the unsorted `a_idx` quietly changes scores if snapshot
columns ever arrive out of row order.

---

## 3. LOW / cleanup (fix opportunistically)

1. **The bijection test doesn't test bijection.** Verified by execution: `PROCUREMENT_BUS` →
   `PROCUREMENTBUSINESS` does not match `PROCUREMENTBUSINESSUNITNAM` (append-delta 7 >
   `MAX_SUFFIX_SLACK` 6), so `test_bijection_one_to_one_per_table` has exactly one viable
   candidate and passes even with the greedy assignment deleted. Use two Oracle keys that both
   genuinely match one column (e.g., `PROCUREMENT_BU` and `PROCUREMENT_BU_NAME`) and assert the
   higher-scoring one wins and the loser is absent.
2. **Missing named test for the truncated-`FLA` suffix.** The logic handles it (verified:
   `ALLOW_SUBSTITUTE_RECEIPTS` ↔ `ALLOW_SUBSTITUTERECEIPTSFLA` matches via the append path), but
   the audit named it and no test pins it. One-liner next to `test_digit_run_truncation`.
3. **Exact pre-pass key-case nit:** `residual_cols` tests `c.bare.upper() not in oracle_by_key` —
   correct only if `oracle_match_key` always returns uppercase. Either normalize the key set once
   (`{k.upper() for k in oracle_by_key}`) or state the invariant.
4. **Unused imports:** `normalize_name` in `_run_correspondence_derive`; `build_file_fields` in the
   Task 8 test snippet. Trim, or lint will.
5. **Alias-collision note:** if a confirmed row's `oracle_key` also exists verbatim as another
   column's bare in the same table (possible only via a stale map after a rename), aliasing
   produces two columns with the same bare. Cheap guard: `build_alias` warns when an alias target
   collides with an existing exact bare for the table. Not required for the pilot.
6. **`assemble_derivation_inputs` silently skips** tables with no catalog fields or no snapshot
   entry — log at INFO so an empty review workbook is explainable (fail-loud convention, soft
   form).
7. **Task 8 `build_table(...)` call signature is assumed** — the step note "match the existing
   import style" should extend to the constructor signature; the implementing agent must read the
   real signature first rather than trusting the snippet.

---

## 4. VERIFIED CORRECT (no action — keep as written)

- **Task 1 `U → char`** — patch and tests exactly implement audit §1.2; veto remains strictly
  char-vs-numeric via set equality (date-vs-char cannot fire it).
- **Digit-run rule** — implementation traced and executed against the live counterexample;
  equal-digits requirement correctly rejects `TIMESTAMP10` vs `TIMESTAMP1`.
- **Normalization ladder** — full squash both sides, `*`-strip, `_FLAG/_FLG/_F` strip,
  token-wise idempotent abbreviation expansion; all audit §2.2/§2.3 cases pass by execution,
  including the `DISC`+`_FLAG` combination and the coincidental-short-prefix rejection.
- **`MAX_SUFFIX_SLACK=6` + cap-escape (`applaud_bare_len >= window-1`)** — sound formulation of
  the bounded-delta prefix rule; window derived as `30 − len(prefix)` per audit §2.1.
- **@/non-prefix exclusion** — defensive `_candidate_excluded` plus Task 8 lock-in is the right
  belt-and-suspenders; the X_PHANTOM deviation from audit §1.4 is accepted as live-verified.
- **Fail-loud `Corrected Bare`** (audit §4.1) — validated against the table's bare set with a
  named exception and a clear message; CLI exits non-zero.
- **`build_alias` gate** — confirmed-always + tier-admits-derived semantics match the spec;
  rejected never aliased; empty-bare guarded.
- **Aliasing wiring** — bare-only replacement via `dataclasses.replace` preserves DDID (Dim 5
  untouched); `fieldmap=None` is a strict no-op so the existing suite stays green; no top-level
  import cycle.
- **Load-bearing regression** (Task 7 Step 1) — correctly asserts both halves: the missing-field
  finding vanishes AND Dim-1 sizing fires on the live `X(25)` vs Oracle 40 divergence.
- **Re-derive idempotence, roundtrips, .gitignore split** — all match the audit's §6 items.

---

## 5. Instruction to the planner

Revise the plan to resolve §1.1 and §1.2 (blockers) and §2.1–§2.4, incorporating the named tests
above. §3 items may be folded into the relevant tasks without new task numbers. Do not change the
matching algorithm, the normalization ladder, or the Task 1 engine fix — they are verified
correct. Hand the revised plan back for the clearance pass, which will re-trace the merge
precedence, the provenance path to the written workbook, and the scoring consistency fix.
