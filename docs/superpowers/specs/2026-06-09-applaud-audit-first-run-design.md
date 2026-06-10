# Applaud Audit — First End-to-End Run (Pilot) — Design

**Date:** 2026-06-09
**Status:** Approved (brainstorming) — ready for implementation planning
**Predecessor:** `docs/superpowers/specs/2026-06-02-applaud-compliance-audit-design.md` (engine design, now built + merged via PR #2)

---

## 1. Purpose

The Applaud compliance audit **engine** (`audit-applaud`, dimensions 1–5 + 6b + 6c) is built and
merged. It has only ever been exercised on a **single table** (`T_BANKS_BRANCHES`) as an
acceptance check. It has never been run end-to-end across a real system at scale.

This project is the **first real end-to-end run** of that engine against a live Applaud system.
Its goals, in order:

1. **Prove the engine at scale** — drive a real, multi-table audit from zero (live extraction) to
   a findings workbook, and judge whether the findings are trustworthy.
2. **Discover what the orchestrator needs** — every friction point in doing this by hand becomes a
   requirement for the future orchestrator skill (the "Candidate C" named in the engine spec).

This is **approach A** in a deliberate **A-then-B** sequence: do one real run manually/agent-driven
first (so the friction is visible), *then* build the orchestrator skill (B) informed by it.
Building the orchestrator is explicitly **out of scope** for this project.

---

## 2. Scope

| Dimension | Decision |
|---|---|
| Target system | **ORACLE_MASTER** (the reference master; cleanest data, lowest risk for a first run) |
| Oracle release | **26B** (latest catalogued); **26A→26B** delta drives dimension 6b |
| Table scope | A **10-table pilot** (§2.1), not the full 147-table mapped set |
| Automation | **None.** Manual / agent-driven throughout — deliberately, so friction is observable |
| Write-back | **None.** Report-only (Phase 1 of the engine spec) |

### 2.1 The pilot set (10 tables)

Chosen by Brad for module coverage; they also happen to span the engine's tricky shapes.

| # | Table | Notable shape (from derived app-map) |
|---|---|---|
| 1 | `T_AP_INVOICE_INT` | **Multi-EF**: `X_T_AP_INVOICE_INT` + `X_T_AP_INVOICE_INT_TXT` |
| 2 | `T_AP_INVOICE_LINES` | Standard IF + EF |
| 3 | `T_BANKS_BRANCHES` | **EF naming asymmetry**: export file is `T_BANKS_BRANCHES` (no `X_`). Known-good anchor. |
| 4 | `T_BPA_PO_LINES_INTERFACE` | Standard IF + EF |
| 5 | `T_EGP_COMPONENTS_INTERFACE` | Standard IF + EF |
| 6 | `T_EGP_ITEM_CATEGORIES_INT` | Standard IF + EF |
| 7 | `T_EGO_ITEM_INTF_EFF_B` | Standard IF + EF |
| 8 | `T_MSC_ST_ASSIGNMENT_SETS` | Standard IF + EF |
| 9 | `T_POZ_SUPPLIERS_INT` | **IF/EF name divergence**: export `X_T_POZ_SUPPLIERS` (`_INT` dropped) |
| 10 | `T_POZ_SUPPLIER_SITES_INT` | **IF/EF name divergence**: export `X_T_POZ_SUPPLIER_SITES` |

All 10 are present in `FBDI_to_Applaud_AppMap.xlsx` with `origin=derived` (machine-guessed,
**not yet human-confirmed**). Confirming these 10 rows is Stage 1 of the run (§4).

---

## 3. The one engine change — a `--tables` scope filter

### Why it's needed

`run_audit` (`fbdi/audit_applaud.py`) loops over the **full** FBDI→table mapping. Per-table checks
are correctly gated on snapshot presence, so a scoped (10-table) snapshot + a scoped (10-row)
app-map already yields **zero spurious findings** for the other ~137 tables. **However**,
`mapped_tables` is populated unconditionally for every mapped table (before the snapshot-presence
guard), so `coverage_gaps` would report **all ~137 non-extracted tables as "no IF/EF resolved."**
For a 10-table pilot, that floods the **Coverage** sheet with ~137 noise rows and destroys the very
distinction the Coverage sheet exists to make: *deliberately out of scope* vs. *genuinely could not
be checked*.

### The change

Add a `--tables` option to the `audit-applaud` CLI:

- **Value:** comma-separated list of target-table names (e.g.
  `--tables T_AP_INVOICE_INT,T_BANKS_BRANCHES,...`).
- **Semantics:** when present, filter the loaded `mapping` to rows whose Applaud target table is in
  the list, *before* calling `run_audit`. Everything downstream — findings, `mapped_tables`,
  `coverage_gaps`, the Summary counts — then reflects exactly the pilot set.
- **Default (omitted):** unchanged full-mapping behavior. Existing full-run callers are unaffected.
- **Validation:** if a named table is not found in the mapping, **fail loud** (print the unknown
  names and exit non-zero) rather than silently auditing fewer tables.

This is a small, reusable knob — the orchestrator (B) and any future subset audit will want the
same filter. It is preferred over hand-crafting throwaway scoped copies of the mapping/app-map
workbooks, which is error-prone and reusable for nothing.

**Filtering lives at the CLI/mapping-load boundary, not inside `run_audit`.** `run_audit`'s
signature and behavior are unchanged; it simply receives a smaller `mapping` dict. This keeps the
engine's core untouched and the change unit-testable in isolation.

---

## 4. The run — five stages

**Stage 1 — Pick & confirm the app-map (Brad + agent).**
For each of the 10 tables, review its derived `FBDI_to_Applaud_AppMap.xlsx` row (Import Files /
Export Files / Source Applications). Brad confirms or corrects the IF/EF resolution; flip
`Origin` → `confirmed`. Pay specific attention to the three tricky shapes flagged in §2.1
(multi-EF, EF naming asymmetry, IF/EF name divergence) — these are exactly where the derived guess
is most likely wrong. Confirmed rows win on any future re-derivation (existing app-map merge rule).

**Stage 2 — Extract (agent-driven Step A).**
Drive `applaud-mcp` per-object pulls for the 10 tables and their confirmed IFs/EFs, feeding raw
results to the `applaud_snapshot.py` assembly helpers. Each pull carries the `COUNT(*)`
completeness assertion (**fail loud** on truncation — the ~100-row silent-truncation hazard).
Output: a scoped `applaud_snapshot.json` (ORACLE_MASTER) containing exactly these 10 tables and
their files.

**Stage 3 — Audit (Step B).**
Run the CLI scoped to the pilot:

```bash
py -m fbdi audit-applaud --release 26B --old-release 26A --system ORACLE_MASTER \
   --tables T_AP_INVOICE_INT,T_AP_INVOICE_LINES,T_BANKS_BRANCHES,T_BPA_PO_LINES_INTERFACE,T_EGP_COMPONENTS_INTERFACE,T_EGP_ITEM_CATEGORIES_INT,T_EGO_ITEM_INTF_EFF_B,T_MSC_ST_ASSIGNMENT_SETS,T_POZ_SUPPLIERS_INT,T_POZ_SUPPLIER_SITES_INT
```

Output: `Applaud_Compliance_Report_26B_ORACLE_MASTER.xlsx` (Summary / Findings / High Priority /
Coverage).

**Stage 4 — Inspect together (Brad + agent).**
Read the findings as the consultant would. Judge:
- Are the HIGH findings *real* misalignments, or engine artifacts?
- Is the Coverage sheet honest — exactly 10 tables, no noise, gaps explained?
- Is 6b (release-delta) firing correctly against the 26A→26B changes?
- Do the three tricky-shape tables produce sane IF/EF coverage, or expose a bridge/matching bug?

This stage is the real payoff — it's where we decide whether the engine is trustworthy at scale.

**Stage 5 — Capture orchestrator requirements (agent).**
Maintain a running notes file (`docs/superpowers/applaud-audit-first-run-notes.md`) recording every
friction point: app-map confirmation effort, extraction volume/time, noisy or low-value dimensions,
matching edge cases, anything manual that the orchestrator should automate or checkpoint. This note
is the **primary deliverable feeding project B**.

---

## 5. Deliverables

1. The `--tables` filter on `audit-applaud` (+ tests).
2. The 10 confirmed app-map rows (`origin=confirmed`) committed to `FBDI_to_Applaud_AppMap.xlsx`.
3. The scoped `applaud_snapshot.json` (gitignored; an artifact, not committed).
4. `Applaud_Compliance_Report_26B_ORACLE_MASTER.xlsx` — the pilot findings workbook.
5. `docs/superpowers/applaud-audit-first-run-notes.md` — observed findings + orchestrator
   requirements for project B.

---

## 6. Testing approach

Follows the repo convention (synthetic fixtures inline per test, `py -m pytest tests/`):

- **`--tables` filter** — given a mapping with N tables and `--tables` naming a subset, the audit
  processes exactly that subset; `mapped_tables` / coverage reflect only the subset.
- **Unknown table name** — a `--tables` entry absent from the mapping **fails loud** (non-zero
  exit, names the offender), never silently narrows scope.
- **Omitted `--tables`** — full-mapping behavior is byte-for-byte unchanged (regression guard for
  existing full-run callers).

Stages 1–5 of the *run* are operational, not unit-tested; the engine paths they exercise are
already covered by the existing 37 audit tests. Only the new `--tables` knob needs new tests.

---

## 7. Non-goals (explicit)

- **No orchestrator skill / automation.** That is project B, informed by this run's Stage 5 notes.
- **No write-back to the MDB.** Report-only.
- **No full 147-table run.** Pilot only; expansion comes after the engine is judged trustworthy.
- **No new dimensions** (6a CTQ required-field checks remain deferred per the engine spec).
- **No HTML/PDF.** Excel-first, per the engine spec's output decision.

---

## 8. Open questions

1. **App-map corrections at scale.** If Stage 1 reveals the *derived* IF/EF guesses are frequently
   wrong (especially on the divergence/asymmetry tables), that itself is a major orchestrator
   requirement — capture it in Stage 5 rather than treating it as a one-off fix. *(Resolved:
   handled by Stage 5.)*
2. **Snapshot reuse vs. full extraction.** This run extracts a fresh 10-table snapshot. Whether a
   later full run reuses/merges per-table snapshots is a project-B concern, out of scope here.
   *(Resolved: out of scope.)*
