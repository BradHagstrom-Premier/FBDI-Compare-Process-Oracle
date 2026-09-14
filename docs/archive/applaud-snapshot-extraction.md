# Step A — Applaud snapshot extraction (agent-driven)

The agent runs this sequence with `applaud-mcp`, passing `system: 'ORACLE_MASTER'`
(or `file_path` fallback), feeding raw results to `fbdi/applaud_snapshot.py` and
`fbdi/applaud_appmap.py` helpers. EVERY per-object pull is validated with
`assert_complete()` against its `COUNT(*)` (applaud-mcp silently truncates ~100 rows).

1. Load the FBDI mapping (pure Python): `report.load_mapping(...)` → the set of
   MAPPED/NEEDS_REVIEW `T_*` target tables (the audit scope).

2. For each target table T:
   a. `get_table_definition(T)` → description (prefix parenthetical) + key sequences.
   b. `SELECT Name,Row,DDID,ODBCName FROM DatabaseDetail WHERE Name='T' ORDER BY Row`
      → assert against `SELECT COUNT(*) FROM DatabaseDetail WHERE Name='T'`.
      **DatabaseDetail carries Row order + DDID only — its DataType/Size/DecPlaces/
      ODBCName columns are EMPTY on real data. Do NOT read type/size from it.**
   c. `derive_prefix(description, [col DDIDs])` → prefix P (3-char TableId code).
   d. **`SELECT Name,DataType,Size,DecPlaces FROM DataDictionary WHERE Name LIKE 'P%'`**
      → assert vs `SELECT COUNT(*) FROM DataDictionary WHERE Name LIKE 'P%'`. This is
      the real type/size source. (`LIKE 'P%'` naturally excludes `@`-audit fields, which
      start with `@`.) Build `dd_by_ddid = {row.Name: row}`.
   e. `build_table(T, P, fallback, description, key_seqs, raw_columns, dd_by_ddid=dd_by_ddid)`
      — joins DD type/size onto each column; drops `@`-audit fields.
   f. `SELECT Name,Description,DBID FROM Application WHERE DBID='T'` → the I_/X_/CQ_ apps.
   g. For each I_/X_ app: `get_application(app)` → steps (IF/EF func_type + func_name + order).

3. For each resolved IF: `SELECT Name,Row,DDID,InputType,Pic FROM ImportDetail
   WHERE Name='if' ORDER BY Row` → assert vs `COUNT(*)`; `build_file_fields(..., kind='IF')`
   (drops `@`-audit fields).

4. For each resolved EF: `SELECT Name,Row,DDID,Pic,ColumnHeader FROM ExportDetail
   WHERE Name='ef' ORDER BY Row` → assert vs `COUNT(*)`; `build_file_fields(..., kind='EF')`.

5. `derive_appmap(applications, target_tables)` → merge with any confirmed
   `FBDI_to_Applaud_AppMap.xlsx` via `merge_appmap` → `write_appmap_workbook`.

6. Assemble `ApplaudSnapshot(...)` and `.write(applaud_snapshot_path(system))`.

DataDictionary IS pulled in Phase 1 — sizing comes from DataDictionary, NOT DatabaseDetail
(which has no type data). `@`-prefixed fields are excluded at assembly. The orchestrator
skill (Candidate C) automates this with HITL checkpoints.

## First-run confirmation (per pass-3 audit §3 — already pre-validated 2026-06-02)

Before trusting output, run the audit on `T_BANKS_BRANCHES` / `Bank Account` and confirm the
business fields match. Expected: **~22/23 clean matches + one genuine divergence** — Oracle
"EDI ID Number" (→`EDI_ID_NUMBER`) vs Applaud `EFT_ID_NUMBER` (a real reviewable finding, not
noise). 23 spurious PRESENCE findings would mean `_label_to_technical` / `oracle_match_key`
normalization is broken. (Pre-validated live this session: 22 clean + that single divergence.)
