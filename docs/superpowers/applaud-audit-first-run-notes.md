# Applaud Audit — First Run Notes (Pilot)

**Date:** 2026-06-09
**Run:** ORACLE_MASTER, release 26B (26A→26B for Dim 6b), 10-table pilot.
**Spec/plan:** `docs/superpowers/specs/2026-06-09-applaud-audit-first-run-design.md`,
`docs/superpowers/plans/2026-06-09-applaud-audit-first-run.md`.

This is the Stage 4 (findings inspection) + Stage 5 (orchestrator requirements) output —
the primary deliverable of the first run. **The pilot succeeded at its real purpose:** prove
the engine end-to-end at scale and surface what the orchestrator/engine actually need. The
findings themselves are mostly *not* yet actionable (see §Stage 4) — and discovering precisely
why is the win.

---

## Stage 4 — Findings inspection

**Headline:** 1,567 findings (957 HIGH) across 10 tables — implausible as 957 real defects, and
indeed dominated by false positives. Distribution by dimension (HIGH / MED / INFO):

| Dim | HIGH | MED | INFO |
|---|---|---|---|
| 1-SIZING | 205 | 0 | 0 |
| 2-IF (coverage/order) | 333 | 1 | 259 |
| 3-EF (coverage/order) | 219 | 1 | 187 |
| 4-TABLE (coverage) | 199 | 0 | 0 |
| 5-ORPHAN | 0 | 162 | 0 |
| 6b-RELEASE | 1 | 0 | 0 |

**The mechanics work.** Extraction is byte-exact (every object matched MCP `COUNT(*)`), the
`--tables` scope kept the Coverage sheet to exactly the 10 pilot tables (zero out-of-scope
noise), all six dimensions computed, the workbook rendered, and the known-good
`T_BANKS_BRANCHES` reproduced its prior acceptance result (2 HIGH — the genuine
EDI/EFT_ID divergence). So the engine is sound; the **matching model** is the problem.

### Root cause of the noise (in priority order)

1. **Genuine Oracle↔Applaud name divergence — dominant, ~750+ of the HIGH.** The audit matches
   Oracle fields to Applaud fields by *exact bare-name equality*. That only works where the two
   systems happen to name a field identically. Measured per-table overlap (Oracle keys ∩ Applaud
   bares):

   | Table | Oracle | Applaud | Overlap |
   |---|---|---|---|
   | T_EGO_ITEM_INTF_EFF_B | 130 | 130 | 130 (clean) |
   | T_AP_INVOICE_LINES | 165 | 162 | 159 |
   | T_BANKS_BRANCHES | 23 | 23 | 22 (clean) |
   | T_AP_INVOICE_INT | 136 | 129 | 124 |
   | T_BPA_PO_LINES_INTERFACE | 107 | 106 | 73 |
   | T_EGP_COMPONENTS_INTERFACE | 103 | 98 | 79 |
   | T_POZ_SUPPLIERS_INT | 155 | 155 | 113 |
   | T_POZ_SUPPLIER_SITES_INT | 199 | 199 | 122 |

   The gaps are real naming differences, not missing fields: abbreviations and suffixes
   (`PROCUREMENT_BU` vs `PROCUREMENT_BUSINESSUNITNAM`, `ALLOW_SUBSTITUTE_RECEIPTS` vs
   `ALLOW_SUBSTITUTERECEIPTSFLA`, `ALWAYS_TAKE_DISCOUNT` vs `ALWAYS_TAKE_DISC_FLAG`), Applaud
   `_FLAG`/length-truncated identifiers, etc. Every non-overlapping Oracle field becomes a
   HIGH "missing field/column" in Dims 2/3/4 — almost all false. **Exact bare-name matching is
   insufficient; a field-correspondence layer is required.** This is the headline finding.

2. **Catalog stores display headers in `technical` for some templates — partially fixed.** The
   `SupplierSiteImportTemplate` tab (and similar) have `label=None` and `technical='Supplier
   Name*'`, `'Import Action *'` — the human display header (spaces, trailing `*`), not a real
   technical name. `oracle_match_key` previously returned `technical` raw, so these never matched.
   **Fixed this run:** `oracle_match_key` now normalizes `technical` through `_label_to_technical`
   the same way it does `label` (`'Supplier Name*'` → `SUPPLIER_NAME`); clean technicals pass
   through unchanged. Recovered ~30 matches, HIGH 1028→957. The deeper inconsistency (why the
   catalog captured headers vs technicals for different templates) is an upstream
   catalog/header-detection issue worth auditing.

3. **Dim 1 sizing mixes real undersizing with type-representation artifacts.** Examples:
   - Real-ish: `OPERATING_UNIT char 60 → char 240`, `GROUP_ID char 40 → char 80` (Applaud
     narrower than Oracle — though Oracle's 240 is often a generous max).
   - Artifact: `INVOICE_DATE d 8 → date` (Applaud date stored as type `d`/size 8 vs Oracle
     `date` — equivalent, false positive); `INVOICE_ID char 13 → numeric 15` (char-vs-numeric
     for an ID Applaud deliberately stores as char). Dim 1 needs type-equivalence rules
     (date forms, char-keyed numerics) and possibly tolerance on generous Oracle maxima.

### Verdict
Findings are **not yet trustworthy/actionable at scale.** Do not hand this workbook to a
consultant as-is. The engine is mechanically correct; it needs (a) a field-correspondence layer
and (b) Dim-1 type-equivalence rules before its output is meaningful beyond cleanly-named tables.

---

## Stage 5 — Orchestrator / engine requirements (project B and the audit roadmap)

Ordered by impact:

1. **Extraction must be programmatic — agent-driven Step A does not scale.** The mandated
   "agent calls MCP, agent writes results" model is bottlenecked on the agent re-typing every
   row (the only MCP-result→disk path). ~4,500 rows blew the context budget and couldn't be
   fidelity-checked against the source. **Resolved this run** with `tools/extract_applaud_snapshot.mjs`
   (reads the `.mdb` directly via `mdb-reader` — the same pure-JS lib applaud-mcp uses, no driver
   install — and cross-checks every object's count against MCP `COUNT(*)`). The orchestrator should
   adopt/generalize this: drive the table/IF/EF list from the confirmed app-map, parameterize the
   system, drop the obsolete pagination/`TOP 999` concern (direct read returns everything).

2. **Build a field-correspondence layer (the #1 audit-quality blocker).** Exact bare-name
   matching produces mostly-false "missing" findings. Needs its own brainstorm→spec. Options to
   weigh: a curated Oracle-technical ↔ Applaud-DDID alias map (authoritative, high effort);
   normalized/fuzzy matching (cheap, imperfect); or exploiting any field-grain data in the FBDI
   mapping workbook. Until this exists, audit only tables with known-clean naming, or present
   findings with an explicit "unmatched — may be a naming difference, not a gap" caveat.

3. **Dim 1 type-equivalence + severity tiering.** Treat Applaud `d`(8)/`date` and Oracle `DATE`
   as equal; handle deliberately char-keyed numerics; consider INFO/tolerance for Applaud sizes
   that are narrower than Oracle's generous maxima (240) but adequate in practice.

4. **Audit catalog header detection for `*ImportTemplate` tabs.** Some land the display header in
   `technical` with `label=None`. The `oracle_match_key` normalization is a band-aid; the catalog
   itself should capture a consistent technical identity.

5. **App-map confirmation is cheap and reliable.** All 10 derived rows were correct against the
   live `Application` bridge (single IF each; EF step names matched incl. the `T_BANKS_BRANCHES`
   no-`X_` asymmetry and the `T_AP_INVOICE_INT` dual-EF). The HITL gate took minutes. Note for the
   orchestrator: tables also carry non-`I_`/`X_` apps (`UPD_`, `CMP_…-B4`, `CQ_…_STRUCT`) that the
   classifier ignores — fine for now, but the orchestrator should surface them so a consultant can
   opt one in.

6. **Phantom/system columns recur** (`X_PHANTOM`, "Phantom Run?"). **Resolved this run:**
   `build_table` now excludes non-prefix system columns with a log while preserving the
   truncation fail-loud for prefix-matching DDIDs. The orchestrator/full run will hit more of
   these across 147 tables — the rule (real data elements share the table prefix) generalizes.

7. **`--tables` scoping + Coverage sheet worked exactly as designed** — the first run produced
   zero out-of-scope coverage noise. No change needed.

### Engine changes already landed on this branch (feat/applaud-audit-first-run)
- `--tables` scope filter (`filter_mapping_to_tables` + CLI).
- Programmatic extractor (`tools/extract_applaud_snapshot.mjs` + `tools/assemble_applaud_snapshot.py`).
- `build_table` excludes non-prefix phantom/system columns (truncation guard preserved).
- `oracle_match_key` normalizes `technical` (recovers display-header tabs).
