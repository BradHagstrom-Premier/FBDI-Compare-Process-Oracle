# Applaud Compliance Audit — Design

**Date:** 2026-06-02
**Status:** Approved (brainstorming) — ready for implementation planning
**Branch:** `feat/applaud-compliance-audit`

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
- **Data elements are namespaced by a TableId prefix.** The same logical field appears as
  `T31BANK_NAME` in target table `T_BANKS` (prefix `T31`) and `O33BANK_NAME` in import
  `I_O_BANKS` (prefix `O33`). Linking IF/EF fields to table columns and to Oracle technical
  names must therefore match on the **bare** name (prefix stripped), not the raw `DDID`.
  Prefixes are authoritative — read from the table description `"T_BANKS (T31)"` /
  `DatabaseTable`, never guessed. Reuse `audit.py`'s `derive_bare_name` / `extract_prefix`
  / `_label_to_technical` machinery.
- **`execute_query` has no JOINs and no aggregates beyond `COUNT(*)`.** All joins happen
  client-side in Python after bulk-pulling each detail table. This is why the design
  snapshots (§4) rather than querying live per object.
- **The Application table is the table↔IF/EF bridge** (see §4).

### The one structural gap

The mapping workbook (`FBDI_to_ApplaudTables_Mapping.xlsx`) maps FBDI tab → Applaud
**target table** only. It says nothing about which Import/Export *File* serves a table, and
the names do not transform cleanly (target `T_BANKS` is fed by import `I_O_BANKS`). That
bridge is solved in §4 via the `Application` metadata.

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
| `fbdi/applaud_snapshot.py` | **Step A.** Bulk-extracts the MDB via `applaud-mcp` into `baselines/applaud/applaud_snapshot.json` (gitignored). Also derives the candidate app-map. |
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

### Snapshot JSON (`baselines/applaud/applaud_snapshot.json`, gitignored)

Five indexed collections, each a one-shot bulk `execute_query` pull (no per-object round
trips). Includes extraction metadata (`system`, `mdb_path`, `extracted_at`,
`extractor_version`).

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

Worked example (validated live): `get_application("X_T_BANKS")` (desc *"FBDI Fields for
T_BANKS"*) returns steps `T_BANKS (EF)` → `X_T_BANKS_VAL (EF)`. So target table `T_BANKS`'s
FBDI export resolves to those two EFs, in that order.

### The app-map workbook (`FBDI_to_Applaud_AppMap.xlsx`, git-tracked)

Step A emits derived rows: `target_table | import_files | export_files | source_application
| origin(derived|confirmed)`. **This workbook is the source of truth for audit scope.**

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

**Dim 3 — EF coverage & ordering.** Identical logic against the FBDI export(s) (the `X_T_*`
"FBDI Fields for …" exports). Framed as: *does our export reproduce every Oracle FBDI
column, in order, for a clean round-trip?*

**Dim 4 — Target-table field coverage.** Every mapped Oracle field should have a column in
the target table's `DatabaseDetail` (match bare name / `ODBCName` ↔ Oracle technical name):
- Oracle field with no table column → **[HIGH] missing column**
- present but not mappable to the Oracle technical name → **[MED] name divergence**
- extra table column → **[INFO]**

**Dim 5 — Data element ↔ target-table consistency.** Every `DDID` (bare name) used in a
resolved IF/EF should also exist as a column in the target table:
- IF/EF field absent from target table → **[MED] orphaned data element** (loads into nothing)

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
| `applaud_object_name` | `I_O_BANKS` / `T_BANKS` | Phase-3 target object |
| `applaud_field` (DDID) | `O33BANK_NAME` | Phase-3 target field |
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
- **Bare-name matching** across TableId-prefix namespacing (`O33BANK_NAME` ↔ `T31BANK_NAME`
  ↔ Oracle `BANK_NAME`).
- **Finding reconciliation** by `finding_id` (Phase-2 readiness): stable IDs across re-runs.
- **Excel writer** — sheet structure, severity fills, Status/Notes columns present.

Note: `applaud-mcp` queries are not exercised in unit tests — Step A's extraction is mocked
at the data boundary so Step B audits run fully offline against synthetic snapshots.

---

## 11. Environment prerequisite (plan task)

The `applaud-mcp` server's default `MDB_FILE_PATH` points at a now-missing file
(`…/MDB_for_ApplaudMCP/AP0STE.mdb`); the real databases moved into the `ORACLE_MASTER/` and
`AWC_MASTER/` subdirectories. This is an environment/config issue independent of the audit
design, but Step A relies on reaching the right MDB.

**Plan task:** as part of implementation, fix the stale MCP config via the `/update-config`
skill — either repoint `MDB_FILE_PATH` to a valid default (`ORACLE_MASTER/AP0STE.mdb`) or,
preferably, configure named systems (`MDB_SYSTEMS`) so `ORACLE_MASTER` / `AWC_MASTER`
resolve as MCP aliases. The latter also lets our `--system` flag (§3) delegate name→path
resolution to the MCP server instead of duplicating the map in `config.py`. Decide between
the two at plan time; either unblocks Step A.
