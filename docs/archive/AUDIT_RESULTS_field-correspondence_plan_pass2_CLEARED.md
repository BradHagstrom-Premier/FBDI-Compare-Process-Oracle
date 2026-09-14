# AUDIT RESULTS — Field-Correspondence Implementation Plan (Pass 2 — CLEARANCE)

**Source:** Technical audit of the revised `2026-06-11-applaud-field-correspondence.md` against
`AUDIT_RESULTS_field-correspondence_plan_pass1.md` (the two blockers and four mediums), the design
audit, and the live ORACLE_MASTER data. The revised merge precedence, scoring path, annotation
scope, and bijection test were re-traced; the new bijection contest and the match-key scoring fix
were executed to confirm deterministic behavior.
**Date:** 2026-06-12

---

## 0. Verdict

**The plan is CLEARED FOR IMPLEMENTATION.** Both Pass-1 blockers and all four mediums are
correctly resolved, every LOW item is addressed, and no new defects were introduced by the
revisions. Three non-blocking residuals are listed in §3 for the implementer/operator to carry —
none requires another plan revision.

---

## 1. Blocker resolutions — VERIFIED

### 1.1 — Confirm-time merge precedence (Pass-1 §1.1): RESOLVED

The plan now defines **two merge functions with opposite, named precedences** and wires each into
the correct command:

- `merge_fieldmap(derived, committed)` — derive-time, committed wins (idempotence across
  re-derives preserved; tests unchanged and still correct).
- `merge_decisions(decisions, committed)` — confirm-time, **incoming human decisions win**;
  untouched committed rows carry forward. `_run_correspondence_confirm` calls `merge_decisions`
  (verified at the call site, with the audit reference in an inline comment).

The prescribed test `test_confirm_overrides_prior_decision` is present, plus
`test_merge_decisions_carries_forward_untouched_committed` covering the non-interference side.
The precedence invariant is **enforced, not just documented**: `load_fieldmap_workbook` drops any
stray `origin=derived` row with a WARNING, with its own test
(`test_load_fieldmap_drops_stray_derived_rows`). A reviewer can now revise any prior decision
through the tooling — the failure mode (silently discarded decisions behind a "Merged N
decision(s)" success message) is closed.

### 1.2 — Rejected-provenance to the written workbook (Pass-1 §1.2): RESOLVED

Resolved the right way — by **verifying the codebase facts and asserting end-to-end**:

- The new "Verified codebase facts" assumptions block establishes (with file/line citations) that
  `Finding` is a mutable dataclass with `notes: str = ""` and that the findings writer already
  emits a "Notes" column. The plan no longer asserts an unestablished attribute.
- `test_rejected_key_annotates_finding_without_changing_severity` now **re-opens the written
  `out.xlsx`**, locates the Notes column by header, and asserts the provenance string in the
  written cell — the feature can no longer pass as an in-memory no-op.

---

## 2. Medium resolutions — VERIFIED

- **§2.1 (score from the match key):** `score_candidate(oracle_key, of, col, window,
  position_score)` — name score computed from the key, never re-derived from `of.technical`. The
  derive loop passes `okey`; `test_score_uses_match_key_not_technical` exercises the
  `technical=None` label-derived-key path and asserts a non-zero score and non-`name=0.00`
  signals. Executed the arithmetic: the label-key candidate scores 0.88 → HIGH, exactly the
  mis-tier the fix prevents. The verified fact that `oracle_match_key` always returns uppercase
  also discharges the Pass-1 LOW #3 invariant.
- **§2.2 (multi-tab merge):** `assemble_derivation_inputs` merges `oracle_by_key` across tabs
  mapping to one table (`prev_keys.update`), warns on prefix disagreement, and INFO-logs skipped
  tables (LOW #6). `test_assemble_merges_multiple_tabs_to_one_table` asserts keys from **both**
  tabs survive.
- **§2.3 (annotation scope + append):** the rejected-annotation block drops the
  `applaud_object_name == table_name` condition — the `findings[n_before:]` slice now annotates
  Dim 4-TABLE **and** Dim 2-IF/3-EF missing-field findings for the same key, so a key is never
  shown as both reviewed and unreviewed across dimensions. Notes are appended
  (`"{existing}; {note}"`), never discarded.
- **§2.4 (position signal):** the unused `_lcs_match` import and dead `o_order`/`a_order` are
  gone; the gradient scores over `residual_cols_by_row` (row-sorted once), so `a_idx` means what
  it claims. The deliberate deviation from the spec's `_lcs_match` is recorded in the
  self-review with the weak-tiebreak rationale — acceptable and now internally consistent.

**LOW items:** all seven addressed. Notably the bijection test now stages a genuine two-way
contest — executed: `PROCUREMENT_BUSINESSUNIT_NAM` scores 1.0 vs `PROCUREMENT_BU` at 0.805, a
deterministic winner with both candidates viable, so deleting the greedy assignment would now
fail the test (the Pass-1 gap). The truncated-`FLA` named test is added; unused imports trimmed;
the alias-collision guard is recorded as deferred with rationale; `build_table`'s 7-arg signature
is cited as verified.

---

## 3. Non-blocking residuals (carry forward; no plan change required)

1. **`--accept-confidence` is inert through the CLI as wired.** `build_alias` supports admitting
   `derived` rows at a tier, but the audit CLI only ever loads the **committed** map — which the
   confirm flow populates exclusively with confirmed/rejected rows, and which
   `load_fieldmap_workbook` now actively strips of `derived` rows. So no CLI invocation can ever
   exercise the HIGH/PROBABLE/WEAK gates; the "pre-review noise-reduction pass" exists at the
   library level only. This does not affect the pilot (the real workflow is the `confirmed`
   default), but either document the flag as not-yet-operational or, later, give `audit-applaud`
   an optional in-memory derive (or review-workbook input) when the gate is a tier. Decide
   post-pilot; do not block on it.
2. **`_label_to_technical()` first-live-run verification still stands** (carried from the design
   audit). The §2.1 fix means label-derived keys are now *scored* correctly, but the keys
   themselves still depend on that helper's behavior against the real catalog. The first
   operational derive (follow-up step 2) is the verification point — eyeball the review workbook
   for label-shaped Oracle keys.
3. **Workbook header literals in the Task 7 test** ("Notes", "Oracle Field", sheet name
   "Findings") are taken from the verified-facts reading; if any literal differs in the real
   writer the test fails loudly at implementation time and the implementer should fix the test's
   lookup, not the writer. Self-correcting; noted so the failure isn't misread.
4. Trivial: `test_truncated_bool_suffix_fla` passes `applaud_bare_len=26` for a 27-char bare —
   harmless (the append path never reads that parameter), fix opportunistically or leave.

---

## 4. Clearance statement

All Pass-1 REQUIRED items are correctly reflected in tasks, code, and tests; the matching
algorithm remains the verified-correct Pass-1 core, untouched except for the scoring-signature
fix; the workflow plumbing defects are closed with enforcement and end-to-end assertions rather
than documentation alone. **Proceed with implementation in task order.** The operational
follow-up's step 6 remains the project's success measure: HIGH "missing field" findings should
collapse from 957 toward the genuine residual on the ten pilot tables, with
`T_BANKS_BRANCHES` retaining only its real EDI/EFT divergence — and the new Dim-1 sizing findings
(e.g., `T09PROCUREMENT_BUSINESSUNITNAM` X(25) vs Oracle 40) appearing as the intended side
benefit, not as regressions.
