# Applaud Audit First Run (Pilot) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a reusable `--tables` scope filter to `audit-applaud`, then run the first real end-to-end Applaud audit on a 10-table pilot (ORACLE_MASTER / 26B) and capture orchestrator requirements.

**Architecture:** One small, isolated code change — a pure `filter_mapping_to_tables()` helper in `fbdi/audit_applaud.py`, wired into the `audit-applaud` CLI at the mapping-load boundary. `run_audit` is untouched; it simply receives a smaller mapping dict. The rest of the project is an operational run (confirm app-map → agent-driven MCP extraction → audit → inspect → capture notes) — not code, executed with human-in-the-loop checkpoints.

**Tech Stack:** Python 3.14+, `openpyxl`, `argparse`, `pytest`; `applaud-mcp` MCP server for live extraction (agent-driven). Use the `py` launcher on Brad's Windows setup.

**Spec:** `docs/superpowers/specs/2026-06-09-applaud-audit-first-run-design.md`

---

## File Structure

| File | Responsibility | Change |
|---|---|---|
| `fbdi/audit_applaud.py` | Add `UnknownTableError` + `filter_mapping_to_tables()` (pure mapping filter, fail-loud) | Modify |
| `fbdi/cli.py` | Add `--tables` arg to the `audit-applaud` subparser; apply the filter in `_run_audit_applaud` | Modify |
| `tests/test_audit_applaud.py` | Unit tests for the filter (subset, unknown-name fail-loud, omitted = no-op) | Modify |
| `FBDI_to_Applaud_AppMap.xlsx` | Flip 10 pilot rows `origin` → `confirmed` after Stage 1 review | Modify (data) |
| `baselines/applaud/applaud_snapshot.json` | Scoped 10-table snapshot from Stage 2 | Create (gitignored artifact) |
| `Applaud_Compliance_Report_26B_ORACLE_MASTER.xlsx` | Pilot findings workbook from Stage 3 | Create (artifact) |
| `docs/superpowers/applaud-audit-first-run-notes.md` | Stage 4 findings + Stage 5 orchestrator requirements | Create |

---

## PART 1 — Code: the `--tables` scope filter

### Task 1: Pure mapping filter with fail-loud on unknown names

**Files:**
- Modify: `fbdi/audit_applaud.py` (add near the top, after the imports / before the `Finding` dataclass)
- Test: `tests/test_audit_applaud.py`

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_audit_applaud.py`:

```python
# --- Task (first-run): --tables mapping filter --------------------------------

import pytest
from fbdi.audit_applaud import filter_mapping_to_tables, UnknownTableError


def _sample_mapping():
    # (template, tab) -> info dict, mirroring report.load_mapping's shape
    return {
        ("RapidImplementationForCashManagement", "Bank Account"):
            {"applaud_table": "T_BANKS_BRANCHES"},
        ("PayablesStandardInvoiceImportTemplate", "Invoice Header"):
            {"applaud_table": "T_AP_INVOICE_INT"},
        ("PayablesStandardInvoiceImportTemplate", "Invoice Lines"):
            {"applaud_table": "T_AP_INVOICE_LINES"},
        ("SomeOtherTemplate", "Some Tab"):
            {"applaud_table": "T_OUT_OF_SCOPE"},
    }


def test_filter_mapping_keeps_only_named_tables():
    mapping = _sample_mapping()
    out = filter_mapping_to_tables(mapping, ["T_BANKS_BRANCHES", "T_AP_INVOICE_INT"])
    kept_tables = {info["applaud_table"] for info in out.values()}
    assert kept_tables == {"T_BANKS_BRANCHES", "T_AP_INVOICE_INT"}
    assert ("SomeOtherTemplate", "Some Tab") not in out


def test_filter_mapping_keeps_all_rows_for_a_multi_tab_table():
    mapping = _sample_mapping()
    mapping[("ExtraTemplate", "Extra Tab")] = {"applaud_table": "T_AP_INVOICE_INT"}
    out = filter_mapping_to_tables(mapping, ["T_AP_INVOICE_INT"])
    assert len(out) == 2  # both tabs that map to T_AP_INVOICE_INT survive


def test_filter_mapping_is_case_insensitive_on_table_names():
    mapping = _sample_mapping()
    out = filter_mapping_to_tables(mapping, ["t_banks_branches"])
    assert {info["applaud_table"] for info in out.values()} == {"T_BANKS_BRANCHES"}


def test_filter_mapping_fails_loud_on_unknown_table():
    mapping = _sample_mapping()
    with pytest.raises(UnknownTableError) as exc:
        filter_mapping_to_tables(mapping, ["T_BANKS_BRANCHES", "T_TYPO_NOPE"])
    assert "T_TYPO_NOPE" in str(exc.value)
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `py -m pytest tests/test_audit_applaud.py -k filter_mapping -v`
Expected: FAIL — `ImportError: cannot import name 'filter_mapping_to_tables'`

- [ ] **Step 3: Write the minimal implementation**

In `fbdi/audit_applaud.py`, after the module imports and before the `Finding` dataclass (around line 31), add:

```python
class UnknownTableError(ValueError):
    """Raised when --tables names a target table absent from the FBDI mapping —
    so a typo fails loud instead of silently narrowing audit scope."""


def filter_mapping_to_tables(
    mapping: dict[tuple[str, str], dict],
    table_names: list[str],
) -> dict[tuple[str, str], dict]:
    """Restrict the FBDI->table mapping to rows whose Applaud target table is in
    `table_names` (case-insensitive). Fail loud via UnknownTableError if any
    requested name is absent from the mapping. Returns a new dict; input untouched."""
    requested = {t.strip().upper() for t in table_names if t and t.strip()}
    present = {(info.get("applaud_table") or "").upper() for info in mapping.values()}
    unknown = sorted(requested - present)
    if unknown:
        raise UnknownTableError(
            "Unknown table(s) not in mapping: " + ", ".join(unknown))
    return {key: info for key, info in mapping.items()
            if (info.get("applaud_table") or "").upper() in requested}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `py -m pytest tests/test_audit_applaud.py -k filter_mapping -v`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): add filter_mapping_to_tables for subset scoping

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 2: Wire `--tables` into the audit-applaud CLI

**Files:**
- Modify: `fbdi/cli.py:171` (add the arg after `--output`)
- Modify: `fbdi/cli.py:481-482` (apply the filter after `mapping = load_mapping(...)`)

- [ ] **Step 1: Add the `--tables` argument to the subparser**

In `fbdi/cli.py`, immediately after the `--output` line (currently line 171), add:

```python
    audit_applaud_parser.add_argument(
        "--tables", default=None,
        help="Comma-separated Applaud target tables to scope the audit to "
             "(e.g. T_BANKS_BRANCHES,T_AP_INVOICE_INT). Omit to audit the full mapping. "
             "An unknown table name fails loud.")
```

- [ ] **Step 2: Apply the filter in `_run_audit_applaud`**

In `fbdi/cli.py`, the current line `mapping = load_mapping(args.mapping)` (line 481) is followed by the appmap load. Insert the filter between them:

```python
    mapping = load_mapping(args.mapping)
    if args.tables:
        from fbdi.audit_applaud import filter_mapping_to_tables, UnknownTableError
        names = [t for t in args.tables.split(",") if t.strip()]
        try:
            mapping = filter_mapping_to_tables(mapping, names)
        except UnknownTableError as exc:
            print(f"Error: {exc}")
            sys.exit(1)
        print(f"Scoped audit to {len({i['applaud_table'] for i in mapping.values()})} "
              f"table(s) via --tables.")
    appmap = load_appmap_workbook(args.appmap) if args.appmap.exists() else {}
```

- [ ] **Step 3: Verify the full test suite still passes**

Run: `py -m pytest tests/`
Expected: PASS — all prior tests plus the new filter tests (398 passed, 2 skipped).

- [ ] **Step 4: Smoke-test the CLI help shows the new flag**

Run: `py -m fbdi audit-applaud --help`
Expected: output includes the `--tables` option and its help text.

- [ ] **Step 5: Commit**

```bash
git add fbdi/cli.py
git commit -m "feat(applaud-audit): add --tables scope flag to audit-applaud CLI

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## PART 2 — The operational run (Stages 1–5)

> These tasks are an **operational run**, not TDD. They drive the live `applaud-mcp` server and require human-in-the-loop confirmation. Each has explicit commands and checkpoints. Do not skip the HITL gates in Tasks 3 and 6.

### Task 3: Stage 1 — Confirm the 10 pilot app-map rows (HITL gate)

**Files:**
- Modify: `FBDI_to_Applaud_AppMap.xlsx` (10 rows)

The 10 pilot tables and their **derived** app-map rows (captured 2026-06-09):

| Table | Import Files | Export Files |
|---|---|---|
| `T_AP_INVOICE_INT` | `I_T_AP_INVOICE_INT` | `X_T_AP_INVOICE_INT; X_T_AP_INVOICE_INT_TXT` |
| `T_AP_INVOICE_LINES` | `I_T_AP_INVOICE_LINES` | `X_T_AP_INVOICE_LINES` |
| `T_BANKS_BRANCHES` | `I_T_BANKS_BRANCHES` | `T_BANKS_BRANCHES` |
| `T_BPA_PO_LINES_INTERFACE` | `I_T_BPA_PO_LINES_INTERFACE` | `X_T_BPA_PO_LINES_INTERFACE` |
| `T_EGP_COMPONENTS_INTERFACE` | `I_T_EGP_COMPONENTS_INTERFACE` | `X_T_EGP_COMPONENTS_INTERFACE` |
| `T_EGP_ITEM_CATEGORIES_INT` | `I_T_EGP_ITEM_CATEGORIES_INT` | `X_T_EGP_ITEM_CATEGORIES_INT` |
| `T_EGO_ITEM_INTF_EFF_B` | `I_T_EGO_ITEM_INTF_EFF_B` | `X_T_EGO_ITEM_INTF_EFF_B` |
| `T_MSC_ST_ASSIGNMENT_SETS` | `I_T_MSC_ST_ASSIGNMENT_SETS` | `X_T_MSC_ST_ASSIGNMENT_SETS` |
| `T_POZ_SUPPLIERS_INT` | `I_T_POZ_SUPPLIERS_INT` | `X_T_POZ_SUPPLIERS` |
| `T_POZ_SUPPLIER_SITES_INT` | `I_T_POZ_SUPPLIER_SITES_INT` | `X_T_POZ_SUPPLIER_SITES` |

- [ ] **Step 1: Verify the derived IF/EF against the live `Application` bridge**

For each table, confirm the IF/EF resolution against `applaud-mcp` (ORACLE_MASTER), per the engine spec §4 bridge: `Application.DBID = '<table>'` → classify `I_*`/`X_*` apps → `get_application` steps give the IF/EF step names in order. Pay particular attention to the three tricky shapes:
- **`T_AP_INVOICE_INT`** — confirm it genuinely has **two** EFs (not one + noise).
- **`T_BANKS_BRANCHES`** — confirm the EF step is `T_BANKS_BRANCHES` (no `X_`), the known asymmetry.
- **`T_POZ_SUPPLIERS_INT` / `T_POZ_SUPPLIER_SITES_INT`** — confirm the export resolves to `X_T_POZ_SUPPLIERS` / `X_T_POZ_SUPPLIER_SITES` (the `_INT`-dropped name), and decide whether the listed name is the export *application* or the resolved EF *step*. Correct the row to the EF step name if they differ.

Use per-object pulls with the `COUNT(*)` assertion (truncation hazard).

- [ ] **Step 2: HITL — present the 10 rows to Brad for confirmation**

Show Brad each derived row (corrected where Step 1 found a divergence). Brad confirms or edits. **Do not proceed past this gate without explicit confirmation.**

- [ ] **Step 3: Flip `Origin` → `confirmed` for the 10 rows**

Edit `FBDI_to_Applaud_AppMap.xlsx` in place (openpyxl full mode, preserve formatting) setting `Origin = confirmed` on the 10 pilot rows, applying any corrections from Steps 1–2.

- [ ] **Step 4: Commit the confirmed app-map**

```bash
git add FBDI_to_Applaud_AppMap.xlsx
git commit -m "data(applaud-audit): confirm 10 pilot app-map rows for first run

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 4: Stage 2 — Extract the scoped 10-table snapshot (agent-driven Step A)

**Files:**
- Create: `baselines/applaud/applaud_snapshot.json` (gitignored)

- [ ] **Step 1: Pull per-object for the 10 tables and their confirmed IFs/EFs**

Drive `applaud-mcp` (ORACLE_MASTER) per the engine spec §4 extraction strategy. For each pilot table: the table's `DatabaseDetail` columns + per-table `DataDictionary` slice (`WHERE Name LIKE '<prefix>%'`); for each confirmed IF: `ImportDetail WHERE Name='<if>' ORDER BY Row`; for each confirmed EF: `ExportDetail WHERE Name='<ef>' ORDER BY Row`. Drop `@`-prefixed audit columns at assembly.

- [ ] **Step 2: Assert completeness on every pull**

For each per-object pull, compare the returned row count to `SELECT COUNT(*) FROM <table> WHERE Name='<obj>'`. **Fail loud** on any mismatch (the ~100-row silent-truncation hazard). Do not hardcode the cap.

- [ ] **Step 3: Assemble the snapshot via the pure-Python helpers**

Feed the raw results to `fbdi/applaud_snapshot.py` assembly helpers to write `baselines/applaud/applaud_snapshot.json` with the five indexed collections + extraction metadata (`system=ORACLE_MASTER`, `extracted_at`, `extractor_version`).

- [ ] **Step 4: Verify the snapshot contains exactly the 10 tables**

Run:

```bash
py -c "from fbdi.applaud_snapshot import ApplaudSnapshot; from fbdi.config import applaud_snapshot_path; s=ApplaudSnapshot.load(applaud_snapshot_path('ORACLE_MASTER')); print(sorted(s.tables)); print('tables=',len(s.tables),'imports=',len(s.imports),'exports=',len(s.exports))"
```

Expected: the 10 pilot table names, `tables=10`, imports ≥ 10, exports ≥ 10 (T_AP_INVOICE_INT contributes 2 EFs).

---

### Task 5: Stage 3 — Run the audit scoped to the pilot

**Files:**
- Create: `Applaud_Compliance_Report_26B_ORACLE_MASTER.xlsx`

- [ ] **Step 1: Run `audit-applaud` with `--tables`**

```bash
py -m fbdi audit-applaud --release 26B --old-release 26A --system ORACLE_MASTER \
   --tables T_AP_INVOICE_INT,T_AP_INVOICE_LINES,T_BANKS_BRANCHES,T_BPA_PO_LINES_INTERFACE,T_EGP_COMPONENTS_INTERFACE,T_EGP_ITEM_CATEGORIES_INT,T_EGO_ITEM_INTF_EFF_B,T_MSC_ST_ASSIGNMENT_SETS,T_POZ_SUPPLIERS_INT,T_POZ_SUPPLIER_SITES_INT
```

Expected: prints `Scoped audit to 10 table(s) via --tables.`, a `Findings: N (HIGH=M)` line, and `Output written to: Applaud_Compliance_Report_26B_ORACLE_MASTER.xlsx`.

- [ ] **Step 2: Verify the Coverage sheet is clean (the whole point of `--tables`)**

Run:

```bash
py -c "from openpyxl import load_workbook; wb=load_workbook('Applaud_Compliance_Report_26B_ORACLE_MASTER.xlsx'); ws=wb['Coverage']; rows=[r[0].value for r in ws.iter_rows(min_row=2)]; print('coverage rows:',len(rows)); print(rows)"
```

Expected: only pilot-scoped tables appear — **no** flood of ~137 out-of-scope tables. (Ideally 0 gap rows if every pilot table resolved an IF/EF.)

---

### Task 6: Stage 4 — Inspect the findings together (HITL gate)

**Files:**
- Create/append: `docs/superpowers/applaud-audit-first-run-notes.md`

- [ ] **Step 1: Summarize the findings workbook for review**

Read the Summary, High Priority, and Coverage sheets. Produce a concise digest: counts by dimension × severity, the HIGH findings list, and the three tricky-shape tables' IF/EF coverage outcomes.

- [ ] **Step 2: HITL — judge trustworthiness with Brad**

Walk Brad through: are the HIGH findings real misalignments or engine artifacts? Is 6b firing correctly against the 26A→26B delta? Did the multi-EF / asymmetry / divergence tables produce sane coverage, or expose a bridge/matching bug? Record Brad's verdict per finding/dimension.

- [ ] **Step 3: Record the inspection results**

Write the findings digest + Brad's trustworthiness verdict into `docs/superpowers/applaud-audit-first-run-notes.md` under a "Stage 4 — Findings inspection" heading.

---

### Task 7: Stage 5 — Capture orchestrator requirements for project B

**Files:**
- Append: `docs/superpowers/applaud-audit-first-run-notes.md`

- [ ] **Step 1: Write the orchestrator-requirements section**

Append a "Stage 5 — Orchestrator requirements (project B)" heading capturing every friction point observed across Stages 1–4: app-map confirmation effort and how often derived guesses were wrong; extraction volume/time and any truncation near-misses; noisy or low-value dimensions; matching edge cases; and any manual step that should become an automated step or an HITL checkpoint in the orchestrator skill.

- [ ] **Step 2: Commit the notes**

```bash
git add docs/superpowers/applaud-audit-first-run-notes.md
git commit -m "docs(applaud-audit): first-run findings + orchestrator requirements

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

- [ ] **Step 3: Final full-suite regression check**

Run: `py -m pytest tests/`
Expected: PASS (398 passed, 2 skipped). Confirms the code change is green after the operational run.

---

## Self-Review

**Spec coverage:**
- §2 scope (ORACLE_MASTER, 26B, 26A→26B, 10-table pilot, no automation, no write-back) → Tasks 3–7 commands encode all of it. ✓
- §3 the `--tables` change (CLI boundary, fail-loud, default unchanged, `run_audit` untouched) → Tasks 1–2. ✓
- §4 five run stages → Tasks 3 (S1), 4 (S2), 5 (S3), 6 (S4), 7 (S5). ✓
- §5 deliverables (filter+tests, confirmed app-map, snapshot, findings workbook, notes file) → Tasks 1–2, 3, 4, 5, 6–7. ✓
- §6 testing (subset, unknown-name fail-loud, omitted = no-op) → Task 1 Steps 1–4 + Task 2 Step 3. ✓
- §7 non-goals → respected (no orchestrator, no write-back, no full run, no new dimensions, no HTML/PDF). ✓

**Placeholder scan:** No TBD/TODO; every code step shows complete code; every command shows expected output. ✓

**Type/name consistency:** `filter_mapping_to_tables(mapping, table_names)` and `UnknownTableError` are defined in Task 1 and imported with the same names in Task 2's CLI wiring and the tests. Mapping shape (`(template, tab) -> {"applaud_table": ...}`) matches `report.load_mapping` as consumed by `run_audit` (`info.get("applaud_table")`). ✓
