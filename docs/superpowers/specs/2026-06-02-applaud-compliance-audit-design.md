# Applaud Compliance Audit — Design

**Date:** 2026-06-02
**Status:** Approved (brainstorming), revised per external technical audit — ready for implementation planning
**Branch:** `feat/applaud-compliance-audit`

> **Revision note (2026-06-02):** This spec was audited live against `ORACLE_MASTER/AP0STE.mdb`
> by a separate Applaud-specialist session (`docs/superpowers/AUDIT_RESULTS_applaud-compliance-audit.md`).
> Four corrections were applied and re-verified in this session: (1) scope is `T_*` target
> tables only — the `O_*`/divergent-prefix premise was wrong; within the `T_*` family the IF,
> EF, and table **share one TableId prefix**; (2) the canonical example is now
> `T_BANKS_BRANCHES`; (3) `execute_query` **silently truncates at ~100 rows** — extraction
> is per-object with a `COUNT(*)` assertion, not a bulk pull; (4) prefix derivation needs a
> logged fallback. One audit claim (§11, that the MCP has named systems configured) did **not**
> reproduce in the Claude Code environment — see §11.

---

## 1. Purpose

The existing compliance report (`python -m fbdi report`) answers *"What changed on
Oracle's end?"* by diffing two FBDI releases. This design adds the next layer: an
**Applaud compliance audit** that answers *"Is our Applaud implementation aligned with
what Oracle's FBDIs expect?"*

The deliverable is a consultant-readable Excel workbook listing every field-level
misalignment between an Applaud system (an `.mdb`, reached **only** through the
`applaud-mcp` MCP server) and the Oracle FBDI release it targets.

Write-to-MDB is explicitly **out of scope** for this phase. The design records the
guardrails that keep a future write phase (Phase 3) mechanical.

---

## 2. Scope confirmation (feasibility, validated live against `ORACLE_MASTER/AP0STE.mdb`)

All six requested dimensions are feasible. Findings from live MCP probing:

| Dim | Check | MDB source | Status |
|---|---|---|---|
| 1 | Data element sizing | `DataDictionary` (`DataType`, `Size`, `DecPlaces`) + `DatabaseDetail` | ✅ Direct; reuse `applaud_type.applaud_type_for` |
| 2 | Import File (IF) coverage & ordering | `Import` + `ImportDetail` (`Name`, `Row`, `DDID`, `Pic`) | ✅ Data present; needs the table↔IF bridge (§4) |
| 3 | Export File (EF) coverage & ordering | `Export` + `ExportDetail` (`Name`, `Row`, `DDID`, `ColumnHeader`) | ✅ Same bridge as dim 2 |
| 4 | Target-table field coverage | `DatabaseTable` + `DatabaseDetail` (`DDID`, `Size`, `ODBCName`) | ✅ Direct; FBDI→table mapping already exists |
| 5 | Data element ↔ target-table consistency | `DDID` cross-ref across `ImportDetail`/`ExportDetail`/`DatabaseDetail` | ✅ Data present |
| 6 | Additional dimensions | see §6 | ✅ 6b + 6c in scope; 6a deferred |

### Key MDB facts learned (load-bearing for the design)

- **The configured MCP default path is stale.** It points to
  `C:/Users/10193/Definian/MDB_for_ApplaudMCP/AP0STE.mdb`, which no longer exists. The
  real databases moved into subdirectories: `…/MDB_for_ApplaudMCP/ORACLE_MASTER/AP0STE.mdb`
  and `…/AWC_MASTER/AP0STE.mdb`. Every tool call must pass an explicit `file_path` (or a
  resolved `--system` path); the bare default errors out.
- **Scope is `T_*` target tables only.** `FBDI_to_ApplaudTables_Mapping.xlsx` maps every FBDI
  tab to a `T_*` Applaud table (154 MAPPED, 485 UNMAPPED, **zero** `O_*` rows). The `O_*`/`O33`
  staging lineage (`O_BANKS`, `I_O_BANKS`, …) does **not** feed FBDIs and is **out of scope**.
- **Within the `T_*` family, the IF, EF, and target table all share one TableId prefix.**
  Validated live: table `T_BANKS_BRANCHES` is prefix `T32`; import `I_T_BANKS_BRANCHES` fields
  are `T32COUNTRY`, `T32BANK_NAME`, …; the EF `T_BANKS_BRANCHES` fields are likewise `T32*`.
  Consequences for matching:
  - **Intra-Applaud comparisons (IF↔table, EF↔table) are exact `DDID` matches** — same prefix,
    no reconciliation needed. Do **not** over-engineer cross-prefix matching; it doesn't exist
    in scope. (This is why Dim 5 is re-grounded in §5.)
  - **Bare-name matching is needed only on the Oracle↔Applaud boundary** — Oracle technical
    name `BANK_NAME` ↔ Applaud bare name after stripping `T32`. Reuse `audit.py`'s
    `derive_bare_name` / `extract_prefix` / `_label_to_technical` machinery here.
- **Prefix derivation needs a documented, logged fallback.** The Applaud-side prefix is read
  from the table description parenthetical (`"T_BANKS_BRANCHES (T32)"`) when present; some
  objects lack it (`O_BANKS` description is just `"O_BANKS"` yet its prefix is `O33`). When the
  parenthetical is absent, derive the prefix from the table's own column/key DDIDs
  (`get_table_definition` key sequence, or first `DatabaseDetail.DDID`) and **log** that a
  fallback was used (this is the "guessing" — make it explicit, never silent). For the
  **Oracle/mapping side**, use the mapping workbook's authoritative `Prefix` column
  (fully populated, 0 blanks across 639 rows) — not the parenthetical.
- **`execute_query` has no JOINs and no aggregates beyond `COUNT(*)`.** All joins happen
  client-side in Python after the per-object pulls (§4). It also **silently truncates at
  ~100 rows** — see §4's per-object + `COUNT(*)`-assertion strategy.
- **The Application table is the table↔IF/EF bridge** (see §4).

### The one structural gap

The mapping workbook (`FBDI_to_ApplaudTables_Mapping.xlsx`) maps FBDI tab → Applaud
**target table** only. It says nothing about which Import/Export *File* serves a table. That
bridge is solved in §4 via the `Application` metadata. Note the join is **not 1:1**: one FBDI
tab can map to several `T_*` tables, and one table can have several IFs/EFs — the app-map
schema (§4) must represent both fan-outs.

---

## 3. Architecture (Candidate B — approved)

A new, self-contained audit engine that mirrors the proven shape of the existing
`fbdi/audit.py` (dataclasses → loaders → signal computation → writers). It **extends** the
`fbdi/` package; nothing existing is replaced.

A new orchestrator skill (chaining the steps with human-in-the-loop checkpoints, built via
`/skill-creator`, like `fbdi-compare-release`) is the explicit **next** project after this
engine ships. It is out of scope for this spec.

### New files

| File | Role |
|---|---|
| `fbdi/applaud_snapshot.py` | **Step A.** Extracts the MDB via `applaud-mcp` (per-object pulls with `COUNT(*)` assertions, §4) into `baselines/applaud/applaud_snapshot.json` (gitignored). Also derives the candidate app-map. |
| `fbdi/applaud_appmap.py` | Load / derive / merge the table↔IF/EF application map. The bridge lives here. |
| `fbdi/audit_applaud.py` | **Step B.** Offline audit engine + Excel writer. |
| `fbdi/cli.py` (edit) | Two new subcommands: `snapshot-applaud` and `audit-applaud`. |

### Reused as-is

- `applaud_type.applaud_type_for(ParsedType) -> str` → `"char 50"` / `"numeric 18,4"` / `"date"`
- `type_parser.parse_data_type` → `ParsedType(data_type, length, scale, parse_warning)`
- `audit.py` styling helpers (`_HEADER_FILL`, `_VERDICT_FILLS`, `_style_header_row`) and
  bare-name machinery (`derive_bare_name`, `extract_prefix`, `_label_to_technical`)
- `report.load_catalog_release` / `report.load_mapping` (or equivalents) for the Oracle side
- `align.align_tabs` for ordering-violation analysis (dims 2/3) — gives *what* moved, not
  just *that* something moved

### CLI

```bash
# Step A — extract MDB snapshot + derive candidate app-map
python -m fbdi snapshot-applaud --system ORACLE_MASTER

# Step B — run the audit against the snapshot
python -m fbdi audit-applaud --release 26B --system ORACLE_MASTER
```

`--system` resolves an alias to its `.mdb` path (default `ORACLE_MASTER`). Because the MCP
server has **no named systems configured** (`list_systems` returns none), this name→path map
lives in *our* config (e.g. `config.py`: `ORACLE_MASTER` →
`…/MDB_for_ApplaudMCP/ORACLE_MASTER/AP0STE.mdb`, `AWC_MASTER` → its path); the resolved path
is then passed as `file_path` to every MCP call. The audit targets **one** MDB per run
(single-master, selectable model). AWC_MASTER and future client DBs are just other selectable
targets — no reference-vs-client diff in this phase.

### Data flow

```
                    ┌─ Oracle FBDI catalog (FBDI_Master_Catalog.xlsx, <release> tab)
                    ├─ FBDI→table mapping (FBDI_to_ApplaudTables_Mapping.xlsx)
 applaud-mcp ──A──► applaud_snapshot.json ─┐
       │                                   ├──B──► audit_applaud ──► Applaud_Compliance_Report_<rel>_<sys>.xlsx
       └──A──► FBDI_to_Applaud_AppMap.xlsx ┘
```

### Output format decision

**Excel-first.** The audit's core job is bulk triage of potentially hundreds of granular
findings; Excel is sortable, filterable, annotatable, and is the natural substrate for
Phase 2 (consultant edits status columns in place). HTML/PDF is deferred — added only if a
formal executive summary is ever requested (YAGNI).

---

## 4. Step A — Snapshot + the app-map bridge

### Extraction strategy — per-object, NOT bulk (critical)

`execute_query` **silently truncates at ~100 rows with no error and no signal** (validated:
`ImportDetail` `COUNT(*)`=10,137 but an unbounded select returned ~100 rows mid-record).
A bulk pull would silently drop >99% of detail rows and produce a clean-looking,
confidently-wrong audit — release-blocking. Therefore:

1. **Pull detail tables per resolved object**, driven off the confirmed app-map — loop over
   the IFs/EFs/tables it names: `SELECT … FROM ImportDetail WHERE Name='<if>' ORDER BY Row`,
   etc. This is complete (validated: `I_T_BANKS_BRANCHES` → all 23 rows) and naturally scopes
   the snapshot to what's actually audited.
2. **Assert completeness after every pull**: compare returned row count to
   `SELECT COUNT(*) FROM <table> WHERE Name='<obj>'`; **fail loud** on mismatch. Apply the
   same guard to any per-table pull that could exceed the cap (`DataDictionary`,
   `DatabaseDetail`).
3. **Do not hardcode the cap** (~100 is environment-dependent) — always assert against
   `COUNT(*)`.

### Snapshot JSON (`baselines/applaud/applaud_snapshot.json`, gitignored)

Five indexed collections, populated by the per-object pulls above. Includes extraction
metadata (`system`, `mdb_path`, `extracted_at`, `extractor_version`).

- `data_dictionary` — `{name → {data_type, size, dec_places, req_opt, table_id}}`
- `tables` — `{table_name → {prefix, description, key_seqs, columns:[{ddid, bare, size, dec_places, odbc_name, row}]}}` (from `DatabaseTable` + `DatabaseDetail`)
- `imports` — `{if_name → [{row, ddid, bare, pic, input_type}]}` (from `Import` + `ImportDetail`)
- `exports` — `{ef_name → [{row, ddid, bare, pic, column_header}]}` (from `Export` + `ExportDetail`)
- `applications` — `{app_name → {dbid, description, steps:[{order, func_type, func_name}]}}` (from `Application`; steps resolved via `get_application`, which cleanly labels each step `IF` / `EF` / `CS`)

### The bridge: deriving the table↔IF/EF map

The `Application` table is the source of truth Brad confirmed exists on the front end. Each
application names a **primary table** (`Application.DBID`) and lists the IFs/EFs it uses (its
execution steps). Derivation, per target table:

1. Find `Application` rows where `DBID = <table>`.
2. Classify by name prefix: `I_*` → import application, `X_*` → export application
   (`CQ_*` = clear/CTQ application — relevant to deferred dim 6a, see §6).
3. Read each application's steps (`get_application`) to list the **IFs/EFs in execution
   order**.

**Canonical worked example — `T_BANKS_BRANCHES` (validated end-to-end):**

| Link | Value | Source (verified) |
|---|---|---|
| FBDI template . tab | `RapidImplementationForCashManagement` . `Bank Account` | mapping workbook |
| Applaud target table | `T_BANKS_BRANCHES` (prefix `T32`) | mapping workbook + `get_table_definition` |
| Bridge | `Application.DBID='T_BANKS_BRANCHES'` → `CQ_T_BANKS_BRANCHES`, `I_T_BANKS_BRANCHES`, `X_T_BANKS_BRANCHES` | `Application` query |
| Import file (IF) | `I_T_BANKS_BRANCHES` → step `I_T_BANKS_BRANCHES (IF)` | `get_application` |
| Export app | `X_T_BANKS_BRANCHES` → steps `T_BANKS_BRANCHES (EF)`, `X_T_BANKS_BRANCHES_VAL (EF)` | `get_application` |

**EF naming asymmetry — resolve EFs by reading `get_application` steps, never by assuming an
`X_` filename.** The export *application* is `X_T_BANKS_BRANCHES`, but its first EF *step* is
named `T_BANKS_BRANCHES` (no `X_`), plus a second `_VAL` validation EF.

### The app-map workbook (`FBDI_to_Applaud_AppMap.xlsx`, git-tracked)

Step A emits derived rows: `target_table | import_files | export_files | source_application
| origin(derived|confirmed)`. `import_files` / `export_files` are **lists**
(semicolon-delimited) so a table with multiple IFs/EFs is one row; the many-tables-per-FBDI-tab
fan-out is carried by the multiple FBDI-mapping rows that point at those tables. **This
workbook is the source of truth for audit scope.**

- On re-run, **confirmed rows win** over freshly-derived ones (NEW-derived fills only gaps;
  human confirmations/edits survive — same NEW-wins-then-OLD-fallback pattern as
  `populate_module`).
- Brad validates/edits this map before audit results are trusted (Open Question 1, resolved).
- **Audit scope = the confirmed app-map.** Whatever IFs/EFs a table's row lists are exactly
  what gets audited for that table. Removing a staging IF from a row excludes it — editorial
  control, no hardcoded prefix filter. The full chain:
  `FBDI tab → (FBDI mapping) → Applaud table → (confirmed app-map) → IF(s)/EF(s) → audited`.

### Snapshot freshness

Reuse `audit.py`'s 30-day staleness check: warn (do not block) when the snapshot is older
than `SNAPSHOT_MAX_AGE_DAYS`.

---

## 5. Step B — The dimension checks

Each check emits zero or more `Finding` records (§7). Severity in brackets.

**Dim 1 — Data element sizing.** For each mapped Oracle field: `parse_data_type` →
`applaud_type_for` → expected (`char 50` / `numeric 18,4`). Resolve the actual Applaud
element (target-table column `DDID` → `data_dictionary`). Compare:
- actual char `Size` < expected → **[HIGH] undersized** (truncation risk)
- actual numeric precision/scale < expected → **[HIGH] precision loss**
- actual type *class* ≠ expected class (Oracle `NUMBER` vs Applaud `X`) → **[HIGH] type-class mismatch**
- actual ≥ expected → pass (oversize is **[INFO]**)

**Dim 2 — IF coverage & ordering.** For each IF resolved via the app-map: build the ordered
Oracle field list (catalog tab order) and the ordered IF field list (`ImportDetail.Row`,
bare names). Then:
- Oracle field absent from IF → **[HIGH] missing field**
- relative order differs from FBDI tab order → **[MED] ordering violation** (report the
  `align.py` LCS displacement — *what* moved)
- IF field with no Oracle counterpart → **[INFO] extra field**

**Dim 3 — EF coverage & ordering.** Identical logic against the FBDI export(s) resolved via
`get_application` steps (§4 — not an `X_`-filename assumption). Framed as: *does our export
reproduce every Oracle FBDI column, in order, for a clean round-trip?* **Derive the
Oracle-comparison name from the bare `DDID`, not `ColumnHeader`** — `ColumnHeader` is empty on
real EFs (validated: every `T_BANKS_BRANCHES` EF row has `ColumnHeader=""`).

**Dim 4 — Target-table field coverage.** Every mapped Oracle field should have a column in
the target table's `DatabaseDetail` (match bare name / `ODBCName` ↔ Oracle technical name):
- Oracle field with no table column → **[HIGH] missing column**
- present but not mappable to the Oracle technical name → **[MED] name divergence**
- extra table column → **[INFO]**

**Dim 5 — Data element ↔ target-table consistency.** Within the `T_*` family the IF, EF, and
table share one prefix, so this is an **exact `DDID` match** (no bare-name reconciliation).
Every `DDID` used in a resolved IF/EF should also exist as a column in the target table:
- IF/EF `DDID` absent from the target table's `DatabaseDetail` → **[MED] orphaned data element**
  (loads into nothing)

This fires only on genuine intra-Applaud orphans; with a single shared prefix it must not
over-fire on cross-prefix noise (which is out of scope) nor degrade to a no-op. Validate the
trigger against the single-prefix reality during implementation.

---

## 6. Additional dimensions

**6a — Required-field conformance. DEFERRED (future stage).** Definian does **not** use
`DataDictionary.ReqOpt` for required checks. Required-field validation is implemented as
**CTQ (Critical-To-Quality) code-section checks inside `CQ_*` applications** — i.e. code
logic, not a metadata flag. Auditing it means parsing those code sections (`CodeSection` /
`CodeDetail`), which is materially more complex and out of scope here. The snapshot already
captures `CQ_*` applications, so the future path is: parse the CTQ application's required
checks and confirm each Oracle-required field is covered. Recorded for a later stage.

**6b — Release-delta cross-check. [HIGH — headline differentiator, in scope.]** Join against
the existing release comparison (catalog `Drift` / `align`): Oracle fields **added** in
`<new>` that are absent from the Applaud IF/EF/table (implementation is behind), and fields
Oracle **removed** that still linger in Applaud (cleanup). This is the one check no generic
schema-differ could perform — it ties the Applaud audit directly to *what Oracle just
changed*, leveraging the engine this repo already has.

**6c — Unmapped-but-present. [INFO, in scope.]** Applaud target tables with no FBDI mapping
at all — surfaces drift in the mapping workbook itself.

---

## 7. Findings model & Excel output

### The `Finding` record

One structured, addressable delta. This shape is what makes Phase 3 mechanical.

| Field | Example | Purpose |
|---|---|---|
| `finding_id` | stable hash of (object, field, attribute) | triage continuity across re-runs |
| `dimension` | `1-SIZING`, `2-IF`, `6b-RELEASE` | which check |
| `severity` | `HIGH` / `MED` / `INFO` | triage order |
| `fbdi_template`, `fbdi_tab`, `oracle_field`, `oracle_type` | `…/BANK_NAME`, `VARCHAR2(100)` | Oracle side |
| `applaud_object_type` | `DATA_ELEMENT` / `IMPORT` / `EXPORT` / `TABLE` | Phase-3 target kind |
| `applaud_object_name` | `I_T_BANKS_BRANCHES` / `T_BANKS_BRANCHES` | Phase-3 target object |
| `applaud_field` (DDID) | `T32BANK_NAME` | Phase-3 target field |
| `attribute` | `SIZE` / `SCALE` / `TYPE_CLASS` / `PRESENCE` / `ORDER` | what to change |
| `current_value` → `expected_value` | `char 30` → `char 100` | the delta to apply |
| `message` | "Undersized: Applaud char 30 < Oracle VARCHAR2(100)" | consultant-readable |
| `status`, `notes` | *(blank)* | Phase-2 columns the consultant edits |

The `(object_type, object_name, field, attribute, current→expected)` tuple is precisely the
instruction Phase 3 replays through the MCP write tools.

### Workbook sheets (reuse `audit.py` fills/header styling)

1. **Summary** — counts by dimension × severity; snapshot metadata (system, snapshot
   timestamp, release); app-map coverage %.
2. **Findings** — master list, one row per finding, severity-colored, with empty
   `Status`/`Notes` columns. This *is* the Phase-2 working surface.
3. **High Priority** — HIGH-severity subset, sorted — the consultant's worklist.
4. **Coverage** — mapped tables where no IF/EF could be resolved, plus 6c unmapped tables.
   Tells the consultant *what wasn't checked and why*, so silence is never mistaken for a
   pass.

Phase-2 status vocabulary: `ACCEPTED` / `DEFERRED` / `ACTIONED` (confirmed with Brad).

---

## 8. Phased plan

**Phase 1 — Report-only (this build).** `snapshot-applaud` (extract + derive app-map) →
human confirms `FBDI_to_Applaud_AppMap.xlsx` → `audit-applaud` (dims 1–5 + 6b + 6c) → Excel
findings workbook. Done when the workbook is correct and the app-map is confirmed for
ORACLE_MASTER.

**Phase 2 — Interactive review.** Consultant fills `Status`/`Notes` in the **Findings**
sheet. A re-run reconciles by `finding_id`: preserves prior status, marks vanished findings
`RESOLVED`, flags new ones. No new UI — the workbook round-trips. Natural home for the
deferred 6a CTQ check once added.

**Phase 3 — Write-to-MDB (out of scope; designed toward).** Read back `ACCEPTED` findings;
replay each delta through the MCP write tools (`push_data_elements`, `push_database_table`,
`push_code_to_section`, `create_import` / `create_export`).

### Phase-1 decisions that keep Phase 3 mechanical (guardrails)

- **Addressable deltas** — every finding carries `(object_type, object_name, field,
  attribute, current→expected)`; no re-derivation at write time.
- **Idempotent by `finding_id`** — re-applying an actioned delta is a no-op.
- **Snapshot is the pre-image** — Phase 3 verifies `current_value` still matches the
  snapshot before writing; refuse if the MDB drifted since extraction.
- **Avoid** prose-only findings and row-position keys — both make Phase 3 harder.

---

## 9. Open questions (all resolved, none Phase-1 blockers)

1. **App-map first run.** Brad will validate the Applaud-table→IF/EF map before audit
   results are trusted. Ship the auditor and confirm in the same session. *(Resolved.)*
2. **Snapshot freshness.** Reuse `audit.py`'s 30-day staleness warning (warn, not block).
   *(Resolved.)*
3. **Multi-IF/EF tables.** Scope is whatever the confirmed app-map lists per table — no
   hardcoded "FBDI-facing only" filter. The map is the editorial control. *(Resolved.)*

---

## 10. Testing approach

Mirror the repo convention: synthetic fixtures built inline per test (no shared fixture
files), `py -m pytest tests/`. Cover:

- **Snapshot/app-map derivation** — synthetic `Application`/`get_application` shapes →
  correct IF/EF resolution and ordering; confirmed-rows-win merge behavior.
- **Each dimension check** — undersized vs oversized vs type-class (dim 1); missing /
  reordered / extra (dims 2/3); missing column / name divergence (dim 4); orphaned element
  (dim 5); added/removed release deltas (6b); unmapped table (6c).
- **Bare-name matching at the Oracle↔Applaud boundary** (Oracle `BANK_NAME` ↔ Applaud
  `T32BANK_NAME` after stripping prefix `T32`); and **exact-`DDID` matching intra-Applaud**
  (IF/EF `DDID` ↔ table `DDID`, same prefix) for Dim 5.
- **Row-count-assertion guard** — a per-object pull whose returned count differs from
  `COUNT(*)` must **fail loud**, not silently proceed (guards the truncation bug).
- **Prefix-derivation fallback** — parenthetical present → parsed; parenthetical absent →
  DDID-derived **and logged**; never silent.
- **EF resolution via `get_application` steps** — including the `X_`-app / no-`X_`-EF-filename
  asymmetry; do not assume an `X_` filename.
- **Finding reconciliation** by `finding_id` (Phase-2 readiness): stable IDs across re-runs.
- **Excel writer** — sheet structure, severity fills, Status/Notes columns present.

Note: `applaud-mcp` queries are not exercised in unit tests — Step A's extraction is mocked
at the data boundary so Step B audits run fully offline against synthetic snapshots.

---

## 11. Environment prerequisite (plan task)

The `applaud-mcp` server's default `MDB_FILE_PATH` points at a now-missing file
(`…/MDB_for_ApplaudMCP/AP0STE.mdb`); the real databases moved into the `ORACLE_MASTER/` and
`AWC_MASTER/` subdirectories. A bare call with no `system`/`file_path` errors out.

**Environment discrepancy to resolve.** The external audit (§2 of `AUDIT_RESULTS_*.md`) reports
`list_systems` returning two configured aliases (`ORACLE_MASTER`, `AWC_MASTER`). In **this
Claude Code session, `list_systems` returns "No named systems configured"** — the auditor's
MCP environment had `MDB_SYSTEMS` configured; this one does not. Both observations are real;
they're different environments. All verified calls in this session passed an explicit
`file_path` (which works); `system: 'ORACLE_MASTER'` is unverified here because no aliases are
configured.

**Plan task:** configure `MDB_SYSTEMS` aliases in *this* environment via the `/update-config`
skill so `ORACLE_MASTER` / `AWC_MASTER` resolve as MCP aliases — matching the auditor's setup
and letting `--system` (§3) delegate name→path resolution to the MCP server instead of
duplicating a map in `config.py`. This is the path that makes Step A portable across both
environments.

**Blocker status:** not a hard blocker — Step A works today by passing an explicit `file_path`
(resolved from `config.py`, §3). Configuring `MDB_SYSTEMS` is the preferred, portable fix and
should be done early, but implementation can proceed on `file_path` if needed.
