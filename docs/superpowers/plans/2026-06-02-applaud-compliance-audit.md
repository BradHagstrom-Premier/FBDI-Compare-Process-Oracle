# Applaud Compliance Audit — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

> **HANDBACK GATE:** Pass-1 audit (`AUDIT_RESULTS_applaud-compliance-audit.md`) and pass-2 audit
> (`AUDIT_RESULTS_plan_pass2.md`) are both incorporated. Pass-2 fixes — all re-verified live this
> session — are: **(Blocker 1)** Dim 1 sizing now sources type/size from `DataDictionary`, not
> the empty `DatabaseDetail` (Tasks 2, 7, 16); **(Blocker 2)** `@`-prefixed audit fields excluded
> at assembly + from the LCP prefix fallback (Tasks 2, 3); **(§3.2, promoted to required)** Oracle
> `technical` is `None` on thin tabs — `oracle_match_key` normalizes the label via
> `_label_to_technical` across all matching dims, with an end-to-end integration test (Tasks 7, 15);
> **(§3.1)** `ODBCName` empty → bare-name is the effective Dim 4 key (noted, Task 10);
> **(§3.3)** date-vs-char type-class test added (Task 7).
>
> **Pass-3 (`AUDIT_RESULTS_plan_pass3.md`): PASS — cleared to implement.** No further audit gate.
> The one residual the auditor couldn't check (that `_label_to_technical` yields UPPER_SNAKE) was
> **pre-validated live this session**: normalizing the real `Bank Account` catalog labels against
> the live `I_T_BANKS_BRANCHES` bares gives **22/23 clean matches** plus the single predicted
> divergence — Oracle "EDI ID Number" (→`EDI_ID_NUMBER`) vs Applaud `EFT_ID_NUMBER`. So the §3.2
> normalization is confirmed correct; no `oracle_match_key` adjustment needed. The Task 15
> integration test still stands as the permanent regression guard.

**Goal:** Build an offline audit engine that compares an Applaud system (`.mdb`, via `applaud-mcp`) against the Oracle FBDI release it targets, emitting a consultant-readable Excel findings workbook of field-level misalignments.

**Architecture:** Two steps. **Step A (agent-driven):** the agent calls `applaud-mcp` per-object, feeding raw query results to pure-Python assembly helpers (`applaud_snapshot.py`, `applaud_appmap.py`) that validate (row-count assertion) and write `applaud_snapshot.json` + a confirmable `FBDI_to_Applaud_AppMap.xlsx`. **Step B (pure CLI):** `audit_applaud.py` loads the snapshot + Oracle catalog + FBDI mapping + app-map, runs seven dimension checks, and writes `Applaud_Compliance_Report_<rel>_<sys>.xlsx`. Scope is `T_*` target tables; within that family IF/EF/table share one TableId prefix, so intra-Applaud matching is exact-DDID and bare-name matching is only needed at the Oracle↔Applaud boundary.

**Tech Stack:** Python 3.14+, openpyxl, pytest. Reuses `fbdi/applaud_type.py`, `fbdi/type_parser.py`, `fbdi/align.py`, `fbdi/report.py` (`load_catalog_release`, `load_mapping`), and `fbdi/audit.py` styling helpers. On Brad's Windows setup use the `py` launcher (`py -m pytest`).

**Spec:** `docs/superpowers/specs/2026-06-02-applaud-compliance-audit-design.md` (authoritative). Where the audit-results file conflicts with the spec, the audit-results file won and the spec was already revised to match.

---

## File Structure

| File | Responsibility |
|---|---|
| `fbdi/applaud_snapshot.py` | Snapshot dataclasses; pure-Python assembly helpers that take **raw MCP query results** (lists of dicts) → typed objects; the `assert_complete()` row-count guard; JSON write/load. No MCP I/O. |
| `fbdi/applaud_appmap.py` | Prefix derivation (parenthetical + logged LCP fallback); app-map derivation from `Application`/`get_application` results; app-map workbook write/load; confirmed-wins merge. |
| `fbdi/audit_applaud.py` | `Finding` dataclass + stable `finding_id`; the seven dimension checks; `run_audit()` orchestration; Excel findings writer. |
| `fbdi/cli.py` | Add the `audit-applaud` subcommand + `_run_audit_applaud()` dispatch. (Edit.) |
| `fbdi/config.py` | Add `APPLAUD_SYSTEMS` name→path map, `DEFAULT_APPLAUD_SYSTEM`, snapshot path helper. (Edit.) |
| `docs/superpowers/references/applaud-snapshot-extraction.md` | The exact per-object MCP query sequence the agent (and later the skill) follows for Step A. |
| `tests/test_applaud_snapshot.py` | Assembly + row-count guard tests. |
| `tests/test_applaud_appmap.py` | Prefix derivation + app-map derivation + merge tests. |
| `tests/test_audit_applaud.py` | Per-dimension check tests + finding_id + Excel writer smoke. |

---

## Task 1: Snapshot data model + JSON I/O

**Files:**
- Create: `fbdi/applaud_snapshot.py`
- Test: `tests/test_applaud_snapshot.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_applaud_snapshot.py
from fbdi.applaud_snapshot import (
    DataColumn, FileField, SnapshotTable, ApplaudSnapshot,
)


def test_snapshot_roundtrips_through_json(tmp_path):
    snap = ApplaudSnapshot(
        system="ORACLE_MASTER",
        mdb_path="X:/AP0STE.mdb",
        extracted_at="2026-06-02T00:00:00+00:00",
        extractor_version="1",
        tables={
            "T_BANKS_BRANCHES": SnapshotTable(
                name="T_BANKS_BRANCHES", prefix="T32", prefix_fallback=False,
                description="T_BANKS_BRANCHES (T32)", key_seqs=[["T32COUNTRY"]],
                columns=[DataColumn(ddid="T32BANK_NAME", bare="BANK_NAME",
                                    data_type="X", size=100, dec_places=None,
                                    odbc_name="BANK_NAME", row=2)],
            )
        },
        imports={"I_T_BANKS_BRANCHES": [FileField(row=1, ddid="T32COUNTRY",
                  bare="COUNTRY", pic="X(60)", input_type="C", column_header=None)]},
        exports={"T_BANKS_BRANCHES": [FileField(row=1, ddid="T32COUNTRY",
                  bare="COUNTRY", pic="X(60)", input_type=None, column_header="")]},
        applications={"I_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES",
                  "description": "", "steps": [{"order": 1, "func_type": "IF",
                  "func_name": "I_T_BANKS_BRANCHES"}]}},
    )
    path = tmp_path / "snap.json"
    snap.write(path)
    loaded = ApplaudSnapshot.load(path)
    assert loaded == snap
    assert loaded.tables["T_BANKS_BRANCHES"].columns[0].bare == "BANK_NAME"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_applaud_snapshot.py::test_snapshot_roundtrips_through_json -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'fbdi.applaud_snapshot'`

- [ ] **Step 3: Write minimal implementation**

```python
# fbdi/applaud_snapshot.py
"""Applaud MDB snapshot — typed model, assembly helpers, JSON I/O.

Step A (extraction) is agent-driven: the agent calls applaud-mcp per-object and
feeds raw query-result rows (lists of dicts) to the assembly helpers here, which
validate (row-count guard) and produce a typed ApplaudSnapshot. No MCP I/O lives
in this module, so every function is unit-testable with synthetic inputs.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, asdict, field
from pathlib import Path


@dataclass
class DataColumn:
    ddid: str
    bare: str
    data_type: str            # Access DataType code: "X" char, "N" numeric, ...
    size: int | None
    dec_places: int | None
    odbc_name: str | None
    row: int


@dataclass
class FileField:
    row: int
    ddid: str
    bare: str
    pic: str | None
    input_type: str | None        # ImportDetail.InputType (IF only)
    column_header: str | None     # ExportDetail.ColumnHeader (EF only; often "")


@dataclass
class SnapshotTable:
    name: str
    prefix: str | None
    prefix_fallback: bool
    description: str
    key_seqs: list[list[str]]
    columns: list[DataColumn]


@dataclass
class ApplaudSnapshot:
    system: str
    mdb_path: str
    extracted_at: str
    extractor_version: str
    tables: dict[str, SnapshotTable] = field(default_factory=dict)
    imports: dict[str, list[FileField]] = field(default_factory=dict)
    exports: dict[str, list[FileField]] = field(default_factory=dict)
    applications: dict[str, dict] = field(default_factory=dict)

    def write(self, path: Path) -> None:
        Path(path).write_text(
            json.dumps(asdict(self), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    @classmethod
    def load(cls, path: Path) -> "ApplaudSnapshot":
        d = json.loads(Path(path).read_text(encoding="utf-8"))
        tables = {
            name: SnapshotTable(
                name=t["name"], prefix=t["prefix"],
                prefix_fallback=t["prefix_fallback"], description=t["description"],
                key_seqs=[list(k) for k in t["key_seqs"]],
                columns=[DataColumn(**c) for c in t["columns"]],
            )
            for name, t in d["tables"].items()
        }
        imports = {n: [FileField(**f) for f in rows] for n, rows in d["imports"].items()}
        exports = {n: [FileField(**f) for f in rows] for n, rows in d["exports"].items()}
        return cls(
            system=d["system"], mdb_path=d["mdb_path"],
            extracted_at=d["extracted_at"], extractor_version=d["extractor_version"],
            tables=tables, imports=imports, exports=exports,
            applications=d["applications"],
        )
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_applaud_snapshot.py::test_snapshot_roundtrips_through_json -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/applaud_snapshot.py tests/test_applaud_snapshot.py
git commit -m "feat(applaud-audit): snapshot data model + JSON I/O"
```

---

## Task 2: Row-count guard + IF/EF/table assembly (with DataDictionary type join + `@`-field exclusion)

**Files:**
- Modify: `fbdi/applaud_snapshot.py`
- Test: `tests/test_applaud_snapshot.py`

The `execute_query` API silently truncates at ~100 rows. Every per-object pull must be checked against its `COUNT(*)`. Assembly helpers take raw rows (as `applaud-mcp` returns them — `list[dict]`) plus the expected count.

**Two data-layer facts confirmed live against `T_BANKS_BRANCHES` (drive this task):**
- **`DatabaseDetail` carries NO type data** — every column returns `DataType=""`, `Size=0`, `DecPlaces=0`, `ODBCName=""`. It supplies only `Row` order + `DDID`. The real type/size lives on **`DataDictionary`** (`T32BANK_NAME` → `DataType='X', Size=100`). So `build_table` joins each column's `DDID` to a `DataDictionary` slice (`dd_by_ddid`) to populate `data_type`/`size`/`dec_places`.
- **`@`-prefixed fields are internal Definian audit/tracking columns** (26 of `T_BANKS_BRANCHES`'s 49: `@T32DO_NOT_LOAD`, `@T32LEGACY_HEADER1..10`, `@T32LEGACY_FIELD1..10`, …) and must be **excluded from all Dim 1–6 matching**. `_strip_prefix("@T32LEGACY_HEADER1","T32")` would leave them mangled, so they are dropped at assembly via `is_audit_field()`.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_applaud_snapshot.py  (append)
import pytest
from fbdi.applaud_snapshot import (
    assert_complete, build_file_fields, build_table, is_audit_field,
    SnapshotIncompleteError,
)


def test_assert_complete_raises_on_truncation():
    rows = [{"Row": i} for i in range(100)]
    with pytest.raises(SnapshotIncompleteError) as exc:
        assert_complete("ImportDetail", "I_X", rows, expected_count=137)
    assert "I_X" in str(exc.value) and "100" in str(exc.value) and "137" in str(exc.value)


def test_assert_complete_passes_when_counts_match():
    rows = [{"Row": i} for i in range(23)]
    assert_complete("ImportDetail", "I_T_BANKS_BRANCHES", rows, expected_count=23)


def test_is_audit_field_detects_at_prefix():
    assert is_audit_field("@T32LEGACY_HEADER1") is True
    assert is_audit_field("T32BANK_NAME") is False


def test_build_file_fields_strips_prefix_orders_and_drops_audit_fields():
    raw = [
        {"Row": 2, "DDID": "T32BANK_NAME", "Pic": "X(100)", "InputType": "C"},
        {"Row": 1, "DDID": "T32COUNTRY", "Pic": "X(60)", "InputType": "C"},
        {"Row": 3, "DDID": "@T32DO_NOT_LOAD", "Pic": "X(1)", "InputType": "C"},  # audit field
    ]
    fields = build_file_fields(raw, prefix="T32", kind="IF")
    assert [f.bare for f in fields] == ["COUNTRY", "BANK_NAME"]   # @ field dropped
    assert fields[0].row == 1 and fields[0].input_type == "C"
    assert fields[0].column_header is None


def test_build_table_joins_datadictionary_type_and_drops_audit_fields():
    # DatabaseDetail carries blank type (real-data shape); DataDictionary has the type.
    raw_cols = [
        {"Row": 1, "DDID": "T32COUNTRY", "DataType": "", "Size": 0,
         "DecPlaces": 0, "ODBCName": ""},
        {"Row": 2, "DDID": "@T32DO_NOT_LOAD", "DataType": "", "Size": 0,
         "DecPlaces": 0, "ODBCName": ""},                          # audit field
    ]
    dd_by_ddid = {"T32COUNTRY": {"DataType": "X", "Size": 60, "DecPlaces": 0}}
    table = build_table("T_BANKS_BRANCHES", prefix="T32", prefix_fallback=False,
                        description="T_BANKS_BRANCHES (T32)", key_seqs=[["T32COUNTRY"]],
                        raw_columns=raw_cols, dd_by_ddid=dd_by_ddid)
    assert [c.bare for c in table.columns] == ["COUNTRY"]          # @ field dropped
    # type/size come from DataDictionary, NOT the blank DatabaseDetail row
    assert table.columns[0].data_type == "X" and table.columns[0].size == 60
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_applaud_snapshot.py -k "assert_complete or build_" -v`
Expected: FAIL with `ImportError: cannot import name 'assert_complete'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/applaud_snapshot.py  (append)

class SnapshotIncompleteError(RuntimeError):
    """Raised when a per-object pull returned fewer rows than COUNT(*) — the
    applaud-mcp ~100-row silent truncation. Fail loud; never proceed."""


def assert_complete(table: str, obj_name: str, rows: list, expected_count: int) -> None:
    if len(rows) != expected_count:
        raise SnapshotIncompleteError(
            f"{table} WHERE Name='{obj_name}': got {len(rows)} rows but "
            f"COUNT(*)={expected_count}. Likely the ~100-row execute_query cap. "
            "Re-pull per-object; do not proceed with a partial snapshot."
        )


def is_audit_field(ddid: str) -> bool:
    """`@`-prefixed DDIDs are internal Definian audit/tracking columns
    (@…DO_NOT_LOAD, @…LEGACY_*). Excluded from all Dim 1-6 matching."""
    return ddid.lstrip().startswith("@")


def _strip_prefix(ddid: str, prefix: str | None) -> str:
    if prefix and ddid.upper().startswith(prefix.upper()):
        return ddid[len(prefix):]
    return ddid


def build_file_fields(raw_rows: list[dict], prefix: str | None, kind: str) -> list[FileField]:
    """kind is 'IF' or 'EF'. Orders by Row; strips the TableId prefix to bare name;
    drops `@`-audit fields."""
    out: list[FileField] = []
    for r in sorted(raw_rows, key=lambda x: x["Row"]):
        ddid = str(r["DDID"])
        if is_audit_field(ddid):
            continue
        out.append(FileField(
            row=int(r["Row"]),
            ddid=ddid,
            bare=_strip_prefix(ddid, prefix),
            pic=(str(r["Pic"]) if r.get("Pic") is not None else None),
            input_type=(str(r["InputType"]) if kind == "IF" and r.get("InputType") is not None else None),
            column_header=(str(r.get("ColumnHeader") or "") if kind == "EF" else None),
        ))
    return out


def build_table(name: str, prefix: str | None, prefix_fallback: bool,
                description: str, key_seqs: list[list[str]],
                raw_columns: list[dict],
                dd_by_ddid: dict[str, dict]) -> SnapshotTable:
    """Columns: Row/DDID from DatabaseDetail (the only data it reliably carries);
    data_type/size/dec_places JOINED from DataDictionary (DatabaseDetail's type
    columns are empty on real data). `@`-audit fields are dropped."""
    cols: list[DataColumn] = []
    for r in sorted(raw_columns, key=lambda x: x["Row"]):
        ddid = str(r["DDID"])
        if is_audit_field(ddid):
            continue
        dd = dd_by_ddid.get(ddid, {})
        size = dd.get("Size")
        dec = dd.get("DecPlaces")
        cols.append(DataColumn(
            ddid=ddid,
            bare=_strip_prefix(ddid, prefix),
            data_type=(str(dd["DataType"]).strip() if dd.get("DataType") is not None else ""),
            size=(int(size) if size not in (None, "") else None),
            dec_places=(int(dec) if dec not in (None, "") else None),
            # ODBCName is empty in ORACLE_MASTER; kept for completeness only.
            odbc_name=(str(r["ODBCName"]) if r.get("ODBCName") else None),
            row=int(r["Row"]),
        ))
    return SnapshotTable(name=name, prefix=prefix, prefix_fallback=prefix_fallback,
                         description=description, key_seqs=key_seqs, columns=cols)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_applaud_snapshot.py -v`
Expected: PASS (all snapshot tests)

- [ ] **Step 5: Commit**

```bash
git add fbdi/applaud_snapshot.py tests/test_applaud_snapshot.py
git commit -m "feat(applaud-audit): row-count guard + IF/EF/table assembly helpers"
```

---

## Task 3: Prefix derivation (parenthetical + logged LCP fallback)

**Files:**
- Create: `fbdi/applaud_appmap.py`
- Test: `tests/test_applaud_appmap.py`

The Applaud-side prefix is read from the table description parenthetical (`"… (T32)"`). When absent (e.g. `O_BANKS` → no parenthetical), derive the 3-char TableId code (`^[A-Z][A-Z0-9]{2}`, e.g. `O33`) from the first business DDID and **log** that a fallback was used. (Implementation note: a longest-common-prefix fallback is **wrong** — two field names sharing a leading letter, e.g. `BANK_NAME`/`BRANCH_NUMBER`, extend the LCP past the 3-char code to `O33B`. Use the TableId-code regex.) Oracle/mapping-side prefix comes from the mapping workbook's `Prefix` column, not here.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_applaud_appmap.py
import logging
from fbdi.applaud_appmap import derive_prefix


def test_derive_prefix_from_parenthetical():
    prefix, fallback = derive_prefix("T_BANKS_BRANCHES (T32)", ["T32COUNTRY", "T32BANK_NAME"])
    assert prefix == "T32" and fallback is False


def test_derive_prefix_falls_back_to_lcp_and_logs(caplog):
    with caplog.at_level(logging.WARNING):
        prefix, fallback = derive_prefix("O_BANKS", ["O33BANK_NAME", "O33BRANCH_NUMBER"])
    assert prefix == "O33" and fallback is True
    assert any("fallback" in r.message.lower() for r in caplog.records)


def test_derive_prefix_none_when_no_columns_and_no_parenthetical():
    prefix, fallback = derive_prefix("WEIRD_TABLE", [])
    assert prefix is None and fallback is True


def test_derive_prefix_fallback_ignores_audit_fields():
    # @-audit fields must not skew the longest-common-prefix derivation.
    prefix, fallback = derive_prefix(
        "O_BANKS", ["O33BANK_NAME", "@O33LEGACY_FIELD1", "O33BRANCH_NUMBER"])
    assert prefix == "O33" and fallback is True
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_applaud_appmap.py -k derive_prefix -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'fbdi.applaud_appmap'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/applaud_appmap.py
"""Applaud application-map bridge: prefix derivation + table<->IF/EF derivation.

Pure-Python. Fed raw Application / get_application results (and DatabaseDetail
DDIDs for the prefix fallback). No MCP I/O.
"""
from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass, field

_log = logging.getLogger(__name__)

# Matches a trailing "(T32)" / "(O33)" style prefix tag in a table description.
_PAREN_PREFIX_RE = re.compile(r"\(([A-Z0-9]+)\)\s*$")


def _longest_common_prefix(strings: list[str]) -> str:
    if not strings:
        return ""
    s1, s2 = min(strings), max(strings)
    i = 0
    while i < len(s1) and i < len(s2) and s1[i] == s2[i]:
        i += 1
    return s1[:i]


def derive_prefix(description: str, column_ddids: list[str]) -> tuple[str | None, bool]:
    """Return (prefix, used_fallback).

    Parenthetical first (authoritative). Otherwise the longest common prefix of
    the table's column DDIDs (all share the TableId prefix), logged as a fallback.
    """
    m = _PAREN_PREFIX_RE.search((description or "").strip())
    if m:
        return m.group(1), False
    # Exclude @-audit fields; derive the 3-char TableId code from the first
    # business DDID. (LCP is wrong: shared leading field letters over-extend it.)
    business = [d.upper() for d in column_ddids if not d.lstrip().startswith("@")]
    lcp = _longest_common_prefix(business)  # NOTE: replaced by _TABLEID_RE — see fbdi/applaud_appmap.py
    if lcp:
        _log.warning(
            "Prefix fallback for %r: no description parenthetical; derived %r "
            "from common DDID prefix.", description, lcp,
        )
        return lcp, True
    _log.warning("Prefix fallback for %r: no parenthetical and no columns; prefix is None.",
                 description)
    return None, True
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_applaud_appmap.py -k derive_prefix -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/applaud_appmap.py tests/test_applaud_appmap.py
git commit -m "feat(applaud-audit): prefix derivation with logged LCP fallback"
```

---

## Task 4: App-map derivation from Application/get_application results

**Files:**
- Modify: `fbdi/applaud_appmap.py`
- Test: `tests/test_applaud_appmap.py`

For each target table, find apps whose `DBID` equals the table, classify by name prefix (`I_`→import, `X_`→export), and collect the IF/EF file names from each app's steps **in execution order**. EFs are resolved from `get_application` steps (func_type `EF`) — never by assuming an `X_` filename.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_applaud_appmap.py  (append)
from fbdi.applaud_appmap import AppMapRow, derive_appmap


def test_derive_appmap_resolves_if_and_ef_in_order():
    applications = {
        "I_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES", "description": "",
            "steps": [{"order": 1, "func_type": "IF", "func_name": "I_T_BANKS_BRANCHES"}]},
        "X_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES", "description": "FBDI Fields",
            "steps": [{"order": 1, "func_type": "EF", "func_name": "T_BANKS_BRANCHES"},
                      {"order": 2, "func_type": "EF", "func_name": "X_T_BANKS_BRANCHES_VAL"}]},
        "CQ_T_BANKS_BRANCHES": {"dbid": "T_BANKS_BRANCHES", "description": "",
            "steps": [{"order": 1, "func_type": "CS", "func_name": "CS_REQ"}]},
        "X_T_OTHER": {"dbid": "T_OTHER", "description": "",
            "steps": [{"order": 1, "func_type": "EF", "func_name": "T_OTHER"}]},
    }
    rows = derive_appmap(applications, {"T_BANKS_BRANCHES"})
    assert len(rows) == 1
    row = rows[0]
    assert row.target_table == "T_BANKS_BRANCHES"
    assert row.import_files == ["I_T_BANKS_BRANCHES"]
    assert row.export_files == ["T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES_VAL"]
    assert set(row.source_applications) == {"I_T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES"}
    assert row.origin == "derived"


def test_derive_appmap_table_with_no_apps_yields_empty_row():
    rows = derive_appmap({}, {"T_LONELY"})
    assert rows[0].target_table == "T_LONELY"
    assert rows[0].import_files == [] and rows[0].export_files == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_applaud_appmap.py -k derive_appmap -v`
Expected: FAIL with `ImportError: cannot import name 'derive_appmap'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/applaud_appmap.py  (append)

@dataclass
class AppMapRow:
    target_table: str
    import_files: list[str] = field(default_factory=list)
    export_files: list[str] = field(default_factory=list)
    source_applications: list[str] = field(default_factory=list)
    origin: str = "derived"          # "derived" | "confirmed"


def _steps_of_type(app: dict, func_type: str) -> list[str]:
    steps = sorted(app.get("steps", []), key=lambda s: s.get("order", 0))
    return [s["func_name"] for s in steps if s.get("func_type") == func_type]


def derive_appmap(applications: dict, target_tables: set[str]) -> list[AppMapRow]:
    """One AppMapRow per target table. Apps are matched by DBID; IF/EF file names
    come from the apps' get_application steps in execution order."""
    rows: list[AppMapRow] = []
    for table in sorted(target_tables):
        imports: list[str] = []
        exports: list[str] = []
        sources: list[str] = []
        for app_name in sorted(applications):
            app = applications[app_name]
            if app.get("dbid") != table:
                continue
            ifs = _steps_of_type(app, "IF")
            efs = _steps_of_type(app, "EF")
            if ifs or efs:
                sources.append(app_name)
            for f in ifs:
                if f not in imports:
                    imports.append(f)
            for f in efs:
                if f not in exports:
                    exports.append(f)
        rows.append(AppMapRow(target_table=table, import_files=imports,
                              export_files=exports, source_applications=sources,
                              origin="derived"))
    return rows
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_applaud_appmap.py -k derive_appmap -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/applaud_appmap.py tests/test_applaud_appmap.py
git commit -m "feat(applaud-audit): derive table<->IF/EF app-map from Application results"
```

---

## Task 5: App-map workbook write / load / confirmed-wins merge

**Files:**
- Modify: `fbdi/applaud_appmap.py`
- Test: `tests/test_applaud_appmap.py`

The workbook is the source of truth for audit scope. Lists are semicolon-delimited. On re-run, confirmed rows win; derived rows only fill gaps (NEW-derived-fills, OLD-confirmed-wins — same spirit as `populate_module`).

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_applaud_appmap.py  (append)
from fbdi.applaud_appmap import write_appmap_workbook, load_appmap_workbook, merge_appmap


def test_appmap_workbook_roundtrip(tmp_path):
    rows = [AppMapRow("T_BANKS_BRANCHES", ["I_T_BANKS_BRANCHES"],
                      ["T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES_VAL"],
                      ["I_T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES"], "derived")]
    path = tmp_path / "appmap.xlsx"
    write_appmap_workbook(rows, path)
    loaded = load_appmap_workbook(path)
    assert loaded["T_BANKS_BRANCHES"].import_files == ["I_T_BANKS_BRANCHES"]
    assert loaded["T_BANKS_BRANCHES"].export_files == ["T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES_VAL"]


def test_merge_keeps_confirmed_and_adds_new_derived():
    confirmed = {"T_A": AppMapRow("T_A", ["I_HAND_EDITED"], [], ["X"], "confirmed")}
    derived = [
        AppMapRow("T_A", ["I_T_A_AUTO"], ["E_T_A"], ["X"], "derived"),  # must NOT override confirmed
        AppMapRow("T_B", ["I_T_B"], [], ["X"], "derived"),              # new -> added
    ]
    merged = merge_appmap(derived, confirmed)
    by = {r.target_table: r for r in merged}
    assert by["T_A"].import_files == ["I_HAND_EDITED"] and by["T_A"].origin == "confirmed"
    assert by["T_B"].import_files == ["I_T_B"] and by["T_B"].origin == "derived"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_applaud_appmap.py -k "appmap_workbook or merge" -v`
Expected: FAIL with `ImportError: cannot import name 'write_appmap_workbook'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/applaud_appmap.py  (append)
from pathlib import Path
from openpyxl import Workbook, load_workbook

_APPMAP_HEADERS = ["Target Table", "Import Files", "Export Files",
                   "Source Applications", "Origin"]


def _join(items: list[str]) -> str:
    return "; ".join(items)


def _split(cell) -> list[str]:
    if cell is None:
        return []
    return [p.strip() for p in str(cell).split(";") if p.strip()]


def write_appmap_workbook(rows: list[AppMapRow], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "App Map"
    ws.append(_APPMAP_HEADERS)
    for r in rows:
        ws.append([r.target_table, _join(r.import_files), _join(r.export_files),
                   _join(r.source_applications), r.origin])
    ws.freeze_panes = "A2"
    wb.save(path)


def load_appmap_workbook(path: Path) -> dict[str, AppMapRow]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["App Map"] if "App Map" in wb.sheetnames else wb.active
    out: dict[str, AppMapRow] = {}
    rows = ws.iter_rows(min_row=2, values_only=True)
    for row in rows:
        table, imports, exports, sources, origin = (list(row) + [None] * 5)[:5]
        if not table:
            continue
        out[str(table)] = AppMapRow(
            target_table=str(table), import_files=_split(imports),
            export_files=_split(exports), source_applications=_split(sources),
            origin=(str(origin) if origin else "derived"),
        )
    wb.close()
    return out


def merge_appmap(derived: list[AppMapRow],
                 confirmed: dict[str, AppMapRow]) -> list[AppMapRow]:
    """Confirmed rows win; derived rows fill only tables not already confirmed."""
    out: dict[str, AppMapRow] = dict(confirmed)
    for r in derived:
        if r.target_table not in out:
            out[r.target_table] = r
    return [out[k] for k in sorted(out)]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_applaud_appmap.py -v`
Expected: PASS (all app-map tests)

- [ ] **Step 5: Commit**

```bash
git add fbdi/applaud_appmap.py tests/test_applaud_appmap.py
git commit -m "feat(applaud-audit): app-map workbook I/O + confirmed-wins merge"
```

---

## Task 6: Finding model + stable finding_id

**Files:**
- Create: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py
from fbdi.audit_applaud import Finding, make_finding_id


def test_finding_id_is_stable_and_attribute_sensitive():
    base = dict(dimension="1-SIZING", applaud_object_type="TABLE",
                applaud_object_name="T_BANKS_BRANCHES", applaud_field="T32BANK_NAME")
    id_size = make_finding_id(attribute="SIZE", **base)
    id_size_again = make_finding_id(attribute="SIZE", **base)
    id_scale = make_finding_id(attribute="SCALE", **base)
    assert id_size == id_size_again        # stable across runs
    assert id_size != id_scale             # attribute-sensitive
    assert len(id_size) == 12


def test_finding_defaults_status_and_notes_blank():
    f = Finding(finding_id="abc", dimension="1-SIZING", severity="HIGH",
                fbdi_template="t", fbdi_tab="tab", oracle_field="BANK_NAME",
                oracle_type="VARCHAR2(100)", applaud_object_type="TABLE",
                applaud_object_name="T_BANKS_BRANCHES", applaud_field="T32BANK_NAME",
                attribute="SIZE", current_value="char 30", expected_value="char 100",
                message="Undersized")
    assert f.status == "" and f.notes == ""
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k finding -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'fbdi.audit_applaud'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py
"""Applaud compliance audit engine (Step B).

Pure-Python over an ApplaudSnapshot + Oracle catalog + FBDI mapping + confirmed
app-map. Runs the dimension checks and writes an Excel findings workbook. Every
finding is an addressable delta: (object_type, object_name, field, attribute,
current -> expected), so a future write phase can replay accepted findings.
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass


@dataclass
class Finding:
    finding_id: str
    dimension: str
    severity: str                 # HIGH | MED | INFO
    fbdi_template: str
    fbdi_tab: str
    oracle_field: str
    oracle_type: str
    applaud_object_type: str      # DATA_ELEMENT | IMPORT | EXPORT | TABLE
    applaud_object_name: str
    applaud_field: str
    attribute: str                # SIZE | SCALE | TYPE_CLASS | PRESENCE | ORDER
    current_value: str
    expected_value: str
    message: str
    status: str = ""              # Phase-2: ACCEPTED | DEFERRED | ACTIONED
    notes: str = ""


def make_finding_id(*, dimension: str, applaud_object_type: str,
                    applaud_object_name: str, applaud_field: str,
                    attribute: str) -> str:
    key = "|".join([dimension, applaud_object_type, applaud_object_name,
                    applaud_field, attribute])
    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:12]
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k finding -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Finding model + stable finding_id"
```

---

## Task 7: Dim 1 — data element sizing

**Files:**
- Modify: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

Compare the Oracle expected shape (from the catalog `AlignedField` via `applaud_type_for`) against the Applaud column's actual shape — `data_type`/`size`/`dec_places` now sourced from **DataDictionary** (joined in `build_table`, Task 2), since `DatabaseDetail` carries no type data. Flags: undersized, precision-loss, type-class mismatch. Oversize is INFO. This task also adds `oracle_match_key()` — the **label→technical normalization** (§3.2) used by every matching dimension: the catalog's `technical` is `None` on thin tabs (e.g. the canonical `Bank Account` tab exposes only labels like "Bank Name"), so the Oracle identity is `technical` when present else `_label_to_technical(label)` (→ `BANK_NAME`), matching the Applaud bare name.

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_audit_applaud.py  (append)
from fbdi.align import AlignedField
from fbdi.applaud_snapshot import DataColumn, build_table
from fbdi.audit_applaud import expected_shape, actual_shape, check_sizing, oracle_match_key


def test_oracle_match_key_normalizes_label_when_technical_missing():
    # Thin tab: technical is None, only a label is present.
    thin = AlignedField(2, "Bank Name", None, None, None, None, None)
    assert oracle_match_key(thin) == "BANK_NAME"
    # Technical present: used as-is (uppercased).
    rich = AlignedField(5, "Bank Name", "BANK_NAME", "VARCHAR2", 60, None, True)
    assert oracle_match_key(rich) == "BANK_NAME"


def test_shapes_char_and_numeric():
    of = AlignedField(position=1, label="Bank Name", technical="BANK_NAME",
                      data_type="VARCHAR2", length=100, scale=None, required=True)
    assert expected_shape(of) == ("char", 100, None)
    col = DataColumn(ddid="T32BANK_NAME", bare="BANK_NAME", data_type="X",
                     size=30, dec_places=None, odbc_name="BANK_NAME", row=1)
    assert actual_shape(col) == ("char", 30, None)


def test_actual_shape_reflects_datadictionary_not_blank_databasedetail():
    # Regression for Blocker 1: DatabaseDetail type cols are blank on real data;
    # build_table must source type/size from DataDictionary so actual_shape is correct.
    raw_cols = [{"Row": 1, "DDID": "T32BANK_NAME", "DataType": "", "Size": 0,
                 "DecPlaces": 0, "ODBCName": ""}]
    dd = {"T32BANK_NAME": {"DataType": "X", "Size": 100, "DecPlaces": 0}}
    table = build_table("T_BANKS_BRANCHES", "T32", False, "T_BANKS_BRANCHES (T32)",
                        [["T32COUNTRY"]], raw_cols, dd_by_ddid=dd)
    assert actual_shape(table.columns[0]) == ("char", 100, None)   # NOT ("", 0, None)


def test_check_sizing_flags_undersized():
    of = AlignedField(1, "Bank Name", "BANK_NAME", "VARCHAR2", 100, None, True)
    col = DataColumn("T32BANK_NAME", "BANK_NAME", "X", 30, None, "BANK_NAME", 1)
    findings = check_sizing("Tmpl", "Bank Account", "T_BANKS_BRANCHES",
                            {"BANK_NAME": of}, [col])
    assert len(findings) == 1
    f = findings[0]
    assert f.attribute == "SIZE" and f.severity == "HIGH"
    assert f.current_value == "char 30" and f.expected_value == "char 100"


def test_check_sizing_flags_type_class_mismatch():
    of = AlignedField(1, "Amount", "AMOUNT", "NUMBER", 18, 4, False)
    col = DataColumn("T32AMOUNT", "AMOUNT", "X", 50, None, "AMOUNT", 1)
    findings = check_sizing("Tmpl", "Tab", "T_X", {"AMOUNT": of}, [col])
    assert findings[0].attribute == "TYPE_CLASS" and findings[0].severity == "HIGH"


def test_check_sizing_oversize_is_info_not_high():
    of = AlignedField(1, "Code", "CODE", "VARCHAR2", 10, None, False)
    col = DataColumn("T32CODE", "CODE", "X", 50, None, "CODE", 1)
    findings = check_sizing("Tmpl", "Tab", "T_X", {"CODE": of}, [col])
    assert findings == [] or all(f.severity == "INFO" for f in findings)


def test_check_sizing_date_stored_as_char_is_type_class_finding():
    # §3.3: Oracle DATE expected ("date"); Applaud stores it as char -> real TYPE_CLASS,
    # not silently swallowed by the exp_cls=="" guard.
    of = AlignedField(1, "Effective Date", "EFFECTIVE_DATE", "DATE", None, None, False)
    col = DataColumn("T32EFFECTIVE_DATE", "EFFECTIVE_DATE", "X", 30, None, "", 1)
    findings = check_sizing("Tmpl", "Tab", "T_X", {"EFFECTIVE_DATE": of}, [col])
    assert len(findings) == 1 and findings[0].attribute == "TYPE_CLASS"
    assert findings[0].severity == "HIGH"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_audit_applaud.py -k "shapes or sizing" -v`
Expected: FAIL with `ImportError: cannot import name 'expected_shape'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py  (append)
from fbdi.align import AlignedField
from fbdi.applaud_snapshot import DataColumn
from fbdi.applaud_type import applaud_type_for
from fbdi.audit import _label_to_technical
from fbdi.type_parser import ParsedType

Shape = tuple[str, int | None, int | None]   # (class, size, scale)


def oracle_match_key(of: AlignedField) -> str:
    """Normalized Oracle identity used by every matching dimension (§3.2).

    The FBDI catalog leaves `technical` = None on thin tabs (e.g. the canonical
    RapidImplementationForCashManagement / 'Bank Account' tab, which exposes only
    labels like 'Bank Name'). Use technical when present; otherwise normalize the
    label via audit._label_to_technical ('Bank Name' -> 'BANK_NAME'), matching the
    Applaud bare name. Returns UPPER_SNAKE_CASE (or '' if neither is present)."""
    if of.technical:
        return of.technical.upper()
    return _label_to_technical(of.label or "").upper()


def _shape_from_applaud_str(s: str) -> Shape:
    """Parse 'char 100' / 'numeric 18,4' / 'date' into (class, size, scale)."""
    parts = s.split()
    cls = parts[0] if parts else ""
    size = scale = None
    if len(parts) == 2:
        nums = parts[1].split(",")
        size = int(nums[0]) if nums[0].isdigit() else None
        if len(nums) == 2 and nums[1].isdigit():
            scale = int(nums[1])
    return (cls, size, scale)


def expected_shape(of: AlignedField) -> Shape:
    pt = ParsedType(data_type=(of.data_type or ""), length=of.length,
                    scale=of.scale, parse_warning=False)
    return _shape_from_applaud_str(applaud_type_for(pt))


def actual_shape(col: DataColumn) -> Shape:
    dt = (col.data_type or "").strip().upper()
    if dt == "X":
        return ("char", col.size, None)
    if dt == "N":
        return ("numeric", col.size, col.dec_places)
    return (dt.lower(), col.size, col.dec_places)


def check_sizing(fbdi_template: str, fbdi_tab: str, table_name: str,
                 oracle_by_bare: dict[str, AlignedField],
                 columns: list[DataColumn]) -> list["Finding"]:
    col_by_bare = {c.bare.upper(): c for c in columns}
    findings: list[Finding] = []
    for bare, of in oracle_by_bare.items():
        col = col_by_bare.get(bare.upper())
        if col is None:
            continue   # presence is Dim 4's job, not Dim 1
        exp_cls, exp_size, exp_scale = expected_shape(of)
        act_cls, act_size, act_scale = actual_shape(col)
        oracle_type = (applaud_type_for(ParsedType(of.data_type or "", of.length,
                                                   of.scale, False)))
        common = dict(fbdi_template=fbdi_template, fbdi_tab=fbdi_tab,
                      oracle_field=bare, oracle_type=oracle_type,
                      applaud_object_type="TABLE", applaud_object_name=table_name,
                      applaud_field=col.ddid)

        def mk(attribute, severity, current, expected, message):
            return Finding(finding_id=make_finding_id(
                dimension="1-SIZING", applaud_object_type="TABLE",
                applaud_object_name=table_name, applaud_field=col.ddid,
                attribute=attribute), dimension="1-SIZING", severity=severity,
                attribute=attribute, current_value=current, expected_value=expected,
                message=message, **common)

        if exp_cls != act_cls and exp_cls not in ("", act_cls):
            findings.append(mk("TYPE_CLASS", "HIGH",
                f"{act_cls} {act_size}", f"{exp_cls} {exp_size}",
                f"Type-class mismatch: Applaud {act_cls} vs Oracle {exp_cls}"))
            continue
        if exp_cls == "char" and exp_size and act_size is not None and act_size < exp_size:
            findings.append(mk("SIZE", "HIGH", f"char {act_size}", f"char {exp_size}",
                f"Undersized: Applaud char {act_size} < Oracle char {exp_size}"))
        elif exp_cls == "numeric":
            if exp_size and act_size is not None and act_size < exp_size:
                findings.append(mk("SIZE", "HIGH",
                    f"numeric {act_size},{act_scale or 0}",
                    f"numeric {exp_size},{exp_scale or 0}",
                    f"Precision loss: Applaud {act_size} < Oracle {exp_size} digits"))
            elif (exp_scale or 0) > (act_scale or 0):
                findings.append(mk("SCALE", "HIGH",
                    f"numeric {act_size},{act_scale or 0}",
                    f"numeric {exp_size},{exp_scale or 0}",
                    f"Scale loss: Applaud scale {act_scale or 0} < Oracle {exp_scale}"))
    return findings
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_audit_applaud.py -k "shapes or sizing" -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Dim 1 data element sizing check"
```

---

## Task 8: Dim 2 — IF coverage & ordering (via align_tabs)

**Files:**
- Modify: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

Presence is a set difference (Oracle bare names vs IF bare names). Ordering is checked only among fields **common to both**, by comparing their relative order via a local LCS — fields outside the LCS are out of place. **`align_tabs` is deliberately NOT reused here:** it classifies `SHIFTED` by absolute position equality, which is correct for two releases of the same tab but wrong here — Oracle catalog positions and IF row numbers live in different numbering spaces (one missing field offsets every downstream row), so it would emit spurious ORDER findings and surface reordered common fields as REMOVED+ADDED. (`align_tabs` *is* used for Dim 6b, Task 12, where both sides are Oracle catalog rows in the same numbering space.)

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py  (append)
from fbdi.applaud_snapshot import FileField
from fbdi.audit_applaud import check_file_coverage


def _of(pos, tech):
    return AlignedField(pos, tech, tech, "VARCHAR2", 50, None, False)


def test_check_if_flags_missing_extra_and_order():
    oracle = [_of(1, "COUNTRY"), _of(2, "BANK_NAME"), _of(3, "BANK_CODE")]
    if_fields = [
        FileField(1, "T32BANK_NAME", "BANK_NAME", "X(100)", "C", None),  # out of order
        FileField(2, "T32COUNTRY", "COUNTRY", "X(60)", "C", None),
        FileField(3, "T32EXTRA", "EXTRA", "X(10)", "C", None),           # extra
        # BANK_CODE missing
    ]
    findings = check_file_coverage("Tmpl", "Bank Account", "I_T_BANKS_BRANCHES",
                                   "IMPORT", "2-IF", oracle, if_fields)
    kinds = {(f.attribute, f.applaud_field): f for f in findings}
    assert any(f.attribute == "PRESENCE" and f.oracle_field == "BANK_CODE"
               and f.severity == "HIGH" for f in findings)
    assert any(f.attribute == "PRESENCE" and "EXTRA" in f.applaud_field
               and f.severity == "INFO" for f in findings)
    assert any(f.attribute == "ORDER" and f.severity == "MED" for f in findings)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k check_if -v`
Expected: FAIL with `ImportError: cannot import name 'check_file_coverage'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py  (append)

def _lcs_sequence(a: list[str], b: list[str]) -> list[str]:
    """Longest common subsequence of two string lists (the in-order overlap)."""
    m, n = len(a), len(b)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(m - 1, -1, -1):
        for j in range(n - 1, -1, -1):
            dp[i][j] = (dp[i + 1][j + 1] + 1) if a[i] == b[j] else max(dp[i + 1][j], dp[i][j + 1])
    out: list[str] = []
    i = j = 0
    while i < m and j < n:
        if a[i] == b[j]:
            out.append(a[i]); i += 1; j += 1
        elif dp[i + 1][j] >= dp[i][j + 1]:
            i += 1
        else:
            j += 1
    return out


def _presence_finding(fbdi_template, fbdi_tab, object_type, object_name, dimension,
                      *, oracle_field, applaud_field, severity, message) -> "Finding":
    id_field = applaud_field if applaud_field != "(missing)" else oracle_field
    return Finding(
        finding_id=make_finding_id(dimension=dimension, applaud_object_type=object_type,
            applaud_object_name=object_name, applaud_field=id_field, attribute="PRESENCE"),
        dimension=dimension, severity=severity, fbdi_template=fbdi_template,
        fbdi_tab=fbdi_tab, oracle_field=oracle_field, oracle_type="",
        applaud_object_type=object_type, applaud_object_name=object_name,
        applaud_field=applaud_field, attribute="PRESENCE",
        current_value="absent" if applaud_field == "(missing)" else "present",
        expected_value="present" if applaud_field == "(missing)" else "(advisory)",
        message=message)


def check_file_coverage(fbdi_template: str, fbdi_tab: str, object_name: str,
                        object_type: str, dimension: str,
                        oracle_fields: list[AlignedField],
                        file_fields: list[FileField]) -> list["Finding"]:
    """Dim 2 (IF) / Dim 3 (EF): coverage (set difference) + ordering (LCS over
    fields common to both). Identity is the bare DDID (Applaud side) vs the
    normalized Oracle key; never ColumnHeader (empty on real EFs)."""
    oracle_order = [oracle_match_key(f) for f in oracle_fields if oracle_match_key(f)]
    oracle_set = set(oracle_order)
    file_sorted = sorted(file_fields, key=lambda x: x.row)
    file_order = [f.bare.upper() for f in file_sorted]
    file_set = set(file_order)
    findings: list[Finding] = []

    # PRESENCE — Oracle field missing from the file (HIGH)
    for f in oracle_fields:
        name = oracle_match_key(f)
        if name and name not in file_set:
            findings.append(_presence_finding(
                fbdi_template, fbdi_tab, object_type, object_name, dimension,
                oracle_field=name, applaud_field="(missing)", severity="HIGH",
                message=f"Missing field: Oracle {name} not in {object_name}"))

    # PRESENCE — extra file field with no Oracle counterpart (INFO)
    for ff in file_sorted:
        if ff.bare.upper() not in oracle_set:
            findings.append(_presence_finding(
                fbdi_template, fbdi_tab, object_type, object_name, dimension,
                oracle_field="", applaud_field=ff.ddid, severity="INFO",
                message=f"Extra field: {object_name} has {ff.ddid} with no Oracle counterpart"))

    # ORDER — among fields common to both, is relative order preserved? (MED)
    common_in_oracle = [n for n in oracle_order if n in file_set]
    common_in_file = [n for n in file_order if n in oracle_set]
    if common_in_oracle != common_in_file:
        keep = set(_lcs_sequence(common_in_oracle, common_in_file))
        ddid_by_bare = {f.bare.upper(): f.ddid for f in file_sorted}
        for idx, n in enumerate(common_in_file):
            if n in keep:
                continue
            field_id = ddid_by_bare.get(n, n)
            findings.append(Finding(
                finding_id=make_finding_id(dimension=dimension, applaud_object_type=object_type,
                    applaud_object_name=object_name, applaud_field=field_id, attribute="ORDER"),
                dimension=dimension, severity="MED", fbdi_template=fbdi_template,
                fbdi_tab=fbdi_tab, oracle_field=n, oracle_type="",
                applaud_object_type=object_type, applaud_object_name=object_name,
                applaud_field=field_id, attribute="ORDER",
                current_value=f"file pos {idx + 1}", expected_value="Oracle relative order",
                message=f"Ordering violation: {n} is out of Oracle field order in {object_name}"))
    return findings
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k check_if -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Dim 2 IF coverage + ordering via align_tabs"
```

---

## Task 9: Dim 3 — EF coverage & ordering

**Files:**
- Modify: `tests/test_audit_applaud.py` (no new production code — `check_file_coverage` already serves EFs via `object_type="EXPORT"`, `dimension="3-EF"`)

Confirms the shared helper handles EFs and that the Oracle-comparison name is the bare DDID (EF `ColumnHeader` is empty on real data).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py  (append)
def test_check_ef_uses_bare_ddid_not_empty_column_header():
    oracle = [_of(1, "COUNTRY"), _of(2, "BANK_NAME")]
    ef_fields = [
        FileField(1, "T32COUNTRY", "COUNTRY", "X(60)", None, ""),     # ColumnHeader empty
        FileField(2, "T32BANK_NAME", "BANK_NAME", "X(100)", None, ""),
    ]
    findings = check_file_coverage("Tmpl", "Bank Account", "T_BANKS_BRANCHES",
                                   "EXPORT", "3-EF", oracle, ef_fields)
    # Full coverage, correct order -> no PRESENCE/ORDER findings
    assert findings == []
```

- [ ] **Step 2: Run test to verify it fails or passes**

Run: `py -m pytest tests/test_audit_applaud.py -k check_ef -v`
Expected: PASS immediately (helper already supports EFs). If it fails, fix `check_file_coverage` so a fully-covered, correctly-ordered EF yields zero findings.

- [ ] **Step 3: (only if Step 2 failed) adjust `check_file_coverage`**

No change expected; the bare DDID is already the identity used by `_file_fields_as_aligned`.

- [ ] **Step 4: Re-run to confirm**

Run: `py -m pytest tests/test_audit_applaud.py -k "check_if or check_ef" -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/test_audit_applaud.py
git commit -m "test(applaud-audit): Dim 3 EF coverage uses bare DDID (empty ColumnHeader)"
```

---

## Task 10: Dim 4 — target-table field coverage

**Files:**
- Modify: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

Every mapped Oracle field should have a column in the target table (match the normalized Oracle key ↔ column bare name / `ODBCName`). Missing column = HIGH. **Note:** `ODBCName` is empty across `ORACLE_MASTER` (confirmed live), so **bare-name is the effective match key**; the `ODBCName` branch is defensive only — a future maintainer should not treat it as load-bearing.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py  (append)
from fbdi.audit_applaud import check_table_coverage


def test_check_table_coverage_flags_missing_column():
    oracle = [_of(1, "COUNTRY"), _of(2, "BANK_NAME")]
    cols = [DataColumn("T32COUNTRY", "COUNTRY", "X", 60, None, "COUNTRY", 1)]
    findings = check_table_coverage("Tmpl", "Bank Account", "T_BANKS_BRANCHES",
                                    oracle, cols)
    assert len(findings) == 1
    assert findings[0].oracle_field == "BANK_NAME"
    assert findings[0].attribute == "PRESENCE" and findings[0].severity == "HIGH"


def test_check_table_coverage_matches_on_odbcname():
    oracle = [_of(1, "BANK_NAME")]
    cols = [DataColumn("T32BNK", "BNK", "X", 60, None, "BANK_NAME", 1)]  # bare differs, ODBC matches
    findings = check_table_coverage("Tmpl", "Tab", "T_X", oracle, cols)
    assert findings == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k check_table_coverage -v`
Expected: FAIL with `ImportError: cannot import name 'check_table_coverage'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py  (append)

def check_table_coverage(fbdi_template: str, fbdi_tab: str, table_name: str,
                         oracle_fields: list[AlignedField],
                         columns: list[DataColumn]) -> list["Finding"]:
    present = set()
    for c in columns:
        present.add(c.bare.upper())
        if c.odbc_name:
            present.add(c.odbc_name.upper())
    findings: list[Finding] = []
    for of in oracle_fields:
        tech = oracle_match_key(of)
        if not tech or tech in present:
            continue
        findings.append(Finding(
            finding_id=make_finding_id(dimension="4-TABLE", applaud_object_type="TABLE",
                applaud_object_name=table_name, applaud_field=tech, attribute="PRESENCE"),
            dimension="4-TABLE", severity="HIGH", fbdi_template=fbdi_template,
            fbdi_tab=fbdi_tab, oracle_field=tech,
            oracle_type="", applaud_object_type="TABLE", applaud_object_name=table_name,
            applaud_field="(missing)", attribute="PRESENCE",
            current_value="absent", expected_value="present",
            message=f"Missing column: Oracle {tech} has no column in {table_name}"))
    return findings
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k check_table_coverage -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Dim 4 target-table field coverage"
```

---

## Task 11: Dim 5 — data element ↔ target-table consistency (orphans)

**Files:**
- Modify: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

Within the `T_*` family IF/EF/table share one prefix, so this is an exact-`DDID` match. An IF/EF `DDID` absent from the target table's columns = orphaned data element (MED).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py  (append)
from fbdi.audit_applaud import check_orphans


def test_check_orphans_flags_if_field_absent_from_table():
    table_cols = [DataColumn("T32COUNTRY", "COUNTRY", "X", 60, None, "COUNTRY", 1)]
    if_fields = [
        FileField(1, "T32COUNTRY", "COUNTRY", "X(60)", "C", None),
        FileField(2, "T32GHOST", "GHOST", "X(10)", "C", None),   # not a table column
    ]
    findings = check_orphans("Tmpl", "Bank Account", "T_BANKS_BRANCHES",
                             "I_T_BANKS_BRANCHES", "IMPORT", table_cols, if_fields)
    assert len(findings) == 1
    assert findings[0].applaud_field == "T32GHOST"
    assert findings[0].attribute == "PRESENCE" and findings[0].severity == "MED"


def test_check_orphans_silent_when_all_match():
    table_cols = [DataColumn("T32COUNTRY", "COUNTRY", "X", 60, None, "COUNTRY", 1)]
    if_fields = [FileField(1, "T32COUNTRY", "COUNTRY", "X(60)", "C", None)]
    assert check_orphans("T", "tab", "T_X", "I_X", "IMPORT", table_cols, if_fields) == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k check_orphans -v`
Expected: FAIL with `ImportError: cannot import name 'check_orphans'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py  (append)

def check_orphans(fbdi_template: str, fbdi_tab: str, table_name: str,
                  object_name: str, object_type: str,
                  table_columns: list[DataColumn],
                  file_fields: list[FileField]) -> list["Finding"]:
    table_ddids = {c.ddid.upper() for c in table_columns}
    findings: list[Finding] = []
    for f in file_fields:
        if f.ddid.upper() in table_ddids:
            continue
        findings.append(Finding(
            finding_id=make_finding_id(dimension="5-ORPHAN", applaud_object_type=object_type,
                applaud_object_name=object_name, applaud_field=f.ddid, attribute="PRESENCE"),
            dimension="5-ORPHAN", severity="MED", fbdi_template=fbdi_template,
            fbdi_tab=fbdi_tab, oracle_field="", oracle_type="",
            applaud_object_type=object_type, applaud_object_name=object_name,
            applaud_field=f.ddid, attribute="PRESENCE",
            current_value=f"in {object_name}", expected_value=f"column in {table_name}",
            message=f"Orphaned data element: {f.ddid} used in {object_name} but absent "
                    f"from table {table_name}"))
    return findings
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k check_orphans -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Dim 5 orphaned data element check"
```

---

## Task 12: Dim 6b — release-delta cross-check

**Files:**
- Modify: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

Given the catalog alignment between old and new release for a tab (`align_tabs(old_oracle, new_oracle) -> list[Change]`), an Oracle field **ADDED** in the new release that is absent from the Applaud table/IF/EF = behind (HIGH); a field **REMOVED** by Oracle that still exists in Applaud = stale (MED).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py  (append)
from fbdi.align import Change
from fbdi.audit_applaud import check_release_delta


def test_release_delta_flags_added_missing_and_removed_lingering():
    # Oracle added NEW_FIELD in new release; removed OLD_FIELD.
    changes = [
        Change("ADDED", None, 3, None, _of(3, "NEW_FIELD")),
        Change("REMOVED", 2, None, _of(2, "OLD_FIELD"), None),
    ]
    applaud_bares = {"COUNTRY", "OLD_FIELD"}   # NEW_FIELD missing; OLD_FIELD lingering
    findings = check_release_delta("Tmpl", "Bank Account", "T_BANKS_BRANCHES",
                                   changes, applaud_bares, old_release="26A", new_release="26B")
    by_field = {f.oracle_field: f for f in findings}
    assert by_field["NEW_FIELD"].severity == "HIGH"
    assert "added" in by_field["NEW_FIELD"].message.lower()
    assert by_field["OLD_FIELD"].severity == "MED"
    assert "removed" in by_field["OLD_FIELD"].message.lower()


def test_release_delta_silent_when_applaud_in_sync():
    changes = [Change("ADDED", None, 3, None, _of(3, "NEW_FIELD"))]
    findings = check_release_delta("T", "tab", "T_X", changes,
                                   {"NEW_FIELD"}, "26A", "26B")  # already present
    assert findings == []


def test_build_release_changes_aligns_per_tab():
    from fbdi.audit_applaud import build_release_changes
    old = {("Tmpl", "Bank Account"): [_of(1, "COUNTRY"), _of(2, "OLD_FIELD")]}
    new = {("Tmpl", "Bank Account"): [_of(1, "COUNTRY"), _of(2, "NEW_FIELD")]}
    changes = build_release_changes(old, new)
    kinds = {c.change_type for c in changes[("Tmpl", "Bank Account")]}
    assert "ADDED" in kinds and "REMOVED" in kinds
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k release_delta -v`
Expected: FAIL with `ImportError: cannot import name 'check_release_delta'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py  (append)
from fbdi.align import Change, align_tabs


def build_release_changes(
    old_catalog: dict[tuple[str, str], list[AlignedField]],
    new_catalog: dict[tuple[str, str], list[AlignedField]],
) -> dict[tuple[str, str], list[Change]]:
    """Per-tab align(old, new) so Dim 6b can see Oracle's added/removed fields.
    Both sides are catalog rows (same numbering space) — the correct use of align_tabs."""
    out: dict[tuple[str, str], list[Change]] = {}
    for key, new_rows in new_catalog.items():
        old_rows = old_catalog.get(key)
        if old_rows:
            out[key] = align_tabs(old_rows, new_rows)
    return out


def check_release_delta(fbdi_template: str, fbdi_tab: str, table_name: str,
                        changes: list[Change], applaud_bares: set[str],
                        old_release: str, new_release: str) -> list["Finding"]:
    present = {b.upper() for b in applaud_bares}
    findings: list[Finding] = []
    for ch in changes:
        if ch.change_type == "ADDED" and ch.new_field is not None:
            name = oracle_match_key(ch.new_field)
            if name and name not in present:
                findings.append(_release_finding(fbdi_template, fbdi_tab, table_name,
                    name, "HIGH", "ADDED",
                    f"Behind release: Oracle added {name} in {new_release}; "
                    f"absent from Applaud {table_name}"))
        elif ch.change_type == "REMOVED" and ch.old_field is not None:
            name = oracle_match_key(ch.old_field)
            if name and name in present:
                findings.append(_release_finding(fbdi_template, fbdi_tab, table_name,
                    name, "MED", "REMOVED",
                    f"Stale field: Oracle removed {name} in {new_release}; "
                    f"still present in Applaud {table_name}"))
    return findings


def _release_finding(fbdi_template, fbdi_tab, table_name, name, severity,
                     change_type, message) -> "Finding":
    return Finding(
        finding_id=make_finding_id(dimension="6b-RELEASE", applaud_object_type="TABLE",
            applaud_object_name=table_name, applaud_field=name, attribute=change_type),
        dimension="6b-RELEASE", severity=severity, fbdi_template=fbdi_template,
        fbdi_tab=fbdi_tab, oracle_field=name, oracle_type="",
        applaud_object_type="TABLE", applaud_object_name=table_name,
        applaud_field=name, attribute=change_type,
        current_value=("absent" if change_type == "ADDED" else "present"),
        expected_value=("present" if change_type == "ADDED" else "removed"),
        message=message)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k release_delta -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Dim 6b release-delta cross-check"
```

---

## Task 13: Dim 6c — unmapped-but-present + coverage rows

**Files:**
- Modify: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

`check_unmapped` flags Applaud `T_*` tables in the snapshot that have no FBDI mapping (INFO). `coverage_gaps` records mapped tables whose confirmed app-map resolved no IF/EF (so the consultant knows what wasn't checked).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py  (append)
from fbdi.audit_applaud import check_unmapped, coverage_gaps


def test_check_unmapped_flags_snapshot_table_without_mapping():
    snapshot_tables = {"T_BANKS_BRANCHES", "T_ORPHAN_TABLE"}
    mapped_tables = {"T_BANKS_BRANCHES"}
    findings = check_unmapped(snapshot_tables, mapped_tables)
    assert len(findings) == 1
    assert findings[0].applaud_object_name == "T_ORPHAN_TABLE"
    assert findings[0].severity == "INFO" and findings[0].dimension == "6c-UNMAPPED"


def test_coverage_gaps_lists_mapped_tables_with_no_if_ef():
    gaps = coverage_gaps(
        mapped_tables={"T_A", "T_B"},
        appmap={"T_A": (["I_T_A"], ["E_T_A"]), "T_B": ([], [])},
    )
    assert gaps == [("T_B", "no IF/EF resolved in app-map")]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k "unmapped or coverage_gaps" -v`
Expected: FAIL with `ImportError: cannot import name 'check_unmapped'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py  (append)

def check_unmapped(snapshot_tables: set[str], mapped_tables: set[str]) -> list["Finding"]:
    findings: list[Finding] = []
    for t in sorted(snapshot_tables):
        if not t.upper().startswith("T_") or t in mapped_tables:
            continue
        findings.append(Finding(
            finding_id=make_finding_id(dimension="6c-UNMAPPED", applaud_object_type="TABLE",
                applaud_object_name=t, applaud_field="", attribute="PRESENCE"),
            dimension="6c-UNMAPPED", severity="INFO", fbdi_template="", fbdi_tab="",
            oracle_field="", oracle_type="", applaud_object_type="TABLE",
            applaud_object_name=t, applaud_field="", attribute="PRESENCE",
            current_value="present", expected_value="(no FBDI mapping)",
            message=f"Unmapped Applaud table: {t} has no FBDI mapping row"))
    return findings


def coverage_gaps(mapped_tables: set[str],
                  appmap: dict[str, tuple[list[str], list[str]]]) -> list[tuple[str, str]]:
    gaps: list[tuple[str, str]] = []
    for t in sorted(mapped_tables):
        ifs, efs = appmap.get(t, ([], []))
        if not ifs and not efs:
            gaps.append((t, "no IF/EF resolved in app-map"))
    return gaps
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k "unmapped or coverage_gaps" -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Dim 6c unmapped + coverage-gap detection"
```

---

## Task 14: Excel findings writer (4 sheets)

**Files:**
- Modify: `fbdi/audit_applaud.py`
- Test: `tests/test_audit_applaud.py`

Reuse `audit.py` styling. Sheets: **Summary**, **Findings** (master, with empty `Status`/`Notes`), **High Priority** (HIGH only), **Coverage** (gaps).

- [ ] **Step 1: Write the failing test**

```python
# tests/test_audit_applaud.py  (append)
from openpyxl import load_workbook
from fbdi.audit_applaud import write_findings_workbook


def _sample_finding(sev="HIGH"):
    return Finding(finding_id="abc123", dimension="1-SIZING", severity=sev,
                   fbdi_template="Tmpl", fbdi_tab="Bank Account", oracle_field="BANK_NAME",
                   oracle_type="VARCHAR2(100)", applaud_object_type="TABLE",
                   applaud_object_name="T_BANKS_BRANCHES", applaud_field="T32BANK_NAME",
                   attribute="SIZE", current_value="char 30", expected_value="char 100",
                   message="Undersized")


def test_write_findings_workbook_has_four_sheets_and_status_columns(tmp_path):
    path = tmp_path / "report.xlsx"
    write_findings_workbook(
        findings=[_sample_finding("HIGH"), _sample_finding("INFO")],
        coverage=[("T_B", "no IF/EF resolved in app-map")],
        meta={"system": "ORACLE_MASTER", "release": "26B",
              "extracted_at": "2026-06-02T00:00:00+00:00"},
        path=path)
    wb = load_workbook(path)
    assert wb.sheetnames == ["Summary", "Findings", "High Priority", "Coverage"]
    findings_headers = [c.value for c in wb["Findings"][1]]
    assert "Status" in findings_headers and "Notes" in findings_headers
    # High Priority excludes the INFO finding
    assert wb["High Priority"].max_row == 2   # header + 1 HIGH row
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k findings_workbook -v`
Expected: FAIL with `ImportError: cannot import name 'write_findings_workbook'`

- [ ] **Step 3: Write the implementation**

```python
# fbdi/audit_applaud.py  (append)
from pathlib import Path
from openpyxl import Workbook
from fbdi.audit import _HEADER_FILL, _HEADER_FONT, _style_header_row
from openpyxl.styles import PatternFill

_SEVERITY_FILLS = {
    "HIGH": PatternFill("solid", fgColor="FCE4D6"),
    "MED":  PatternFill("solid", fgColor="FFF2CC"),
    "INFO": PatternFill("solid", fgColor="E2EFDA"),
}

_FINDINGS_HEADERS = [
    "Finding ID", "Dimension", "Severity", "FBDI Template", "FBDI Tab",
    "Oracle Field", "Oracle Type", "Applaud Object Type", "Applaud Object",
    "Applaud Field", "Attribute", "Current", "Expected", "Message", "Status", "Notes",
]


def _finding_row(f: "Finding") -> list:
    return [f.finding_id, f.dimension, f.severity, f.fbdi_template, f.fbdi_tab,
            f.oracle_field, f.oracle_type, f.applaud_object_type, f.applaud_object_name,
            f.applaud_field, f.attribute, f.current_value, f.expected_value,
            f.message, f.status, f.notes]


def _write_findings_sheet(ws, findings: list["Finding"]) -> None:
    ws.append(_FINDINGS_HEADERS)
    _style_header_row(ws, len(_FINDINGS_HEADERS))
    for f in findings:
        ws.append(_finding_row(f))
        fill = _SEVERITY_FILLS.get(f.severity)
        if fill:
            for col in range(1, len(_FINDINGS_HEADERS) + 1):
                ws.cell(row=ws.max_row, column=col).fill = fill
    ws.freeze_panes = "A2"


def write_findings_workbook(findings: list["Finding"], coverage: list[tuple[str, str]],
                            meta: dict, path: Path) -> None:
    sev_order = {"HIGH": 0, "MED": 1, "INFO": 2}
    findings = sorted(findings, key=lambda f: (sev_order.get(f.severity, 9), f.dimension))

    wb = Workbook()

    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.append(["Applaud Compliance Report"])
    ws_sum.append(["System", meta.get("system", "")])
    ws_sum.append(["Release", meta.get("release", "")])
    ws_sum.append(["Snapshot extracted", meta.get("extracted_at", "")])
    ws_sum.append([])
    ws_sum.append(["Dimension", "HIGH", "MED", "INFO"])
    dims = sorted({f.dimension for f in findings})
    for d in dims:
        ws_sum.append([d,
            sum(1 for f in findings if f.dimension == d and f.severity == "HIGH"),
            sum(1 for f in findings if f.dimension == d and f.severity == "MED"),
            sum(1 for f in findings if f.dimension == d and f.severity == "INFO")])

    _write_findings_sheet(wb.create_sheet("Findings"), findings)
    _write_findings_sheet(wb.create_sheet("High Priority"),
                          [f for f in findings if f.severity == "HIGH"])

    ws_cov = wb.create_sheet("Coverage")
    ws_cov.append(["Table", "Coverage Note"])
    _style_header_row(ws_cov, 2)
    for table, note in coverage:
        ws_cov.append([table, note])
    ws_cov.freeze_panes = "A2"

    wb.save(path)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k findings_workbook -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): Excel findings writer (Summary/Findings/High/Coverage)"
```

---

## Task 15: `run_audit` orchestration + config + `audit-applaud` CLI

**Files:**
- Modify: `fbdi/config.py`, `fbdi/audit_applaud.py`, `fbdi/cli.py`
- Test: `tests/test_audit_applaud.py`

`run_audit` wires it together: load snapshot + catalog + mapping + confirmed app-map, iterate mapped (template, tab)→table, gather that table's columns/IFs/EFs from the snapshot per the app-map, run each check, then 6c/coverage, and write the workbook.

- [ ] **Step 1: Add config constants**

```python
# fbdi/config.py  (append)

# Applaud system aliases -> .mdb path. Mirrors MDB_SYSTEMS in the applaud-mcp env;
# kept here so Step B can name-qualify snapshot files / output without the MCP up.
APPLAUD_SYSTEMS = {
    "ORACLE_MASTER": "C:/Users/10193/Definian/MDB_for_ApplaudMCP/ORACLE_MASTER/AP0STE.mdb",
    "AWC_MASTER":    "C:/Users/10193/Definian/MDB_for_ApplaudMCP/AWC_MASTER/AP0STE.mdb",
}
DEFAULT_APPLAUD_SYSTEM = "ORACLE_MASTER"


def applaud_snapshot_path(system: str):
    from pathlib import Path
    return Path("baselines") / "applaud" / f"applaud_snapshot_{system}.json"
```

- [ ] **Step 2: Write the failing orchestration test**

```python
# tests/test_audit_applaud.py  (append)
from fbdi.applaud_snapshot import ApplaudSnapshot, SnapshotTable, DataColumn, FileField
from fbdi.audit_applaud import run_audit


def _build_snapshot():
    return ApplaudSnapshot(
        system="ORACLE_MASTER", mdb_path="X", extracted_at="2026-06-02T00:00:00+00:00",
        extractor_version="1",
        tables={"T_BANKS_BRANCHES": SnapshotTable(
            name="T_BANKS_BRANCHES", prefix="T32", prefix_fallback=False,
            description="T_BANKS_BRANCHES (T32)", key_seqs=[["T32COUNTRY"]],
            columns=[DataColumn("T32COUNTRY", "COUNTRY", "X", 60, None, "COUNTRY", 1),
                     DataColumn("T32BANK_NAME", "BANK_NAME", "X", 30, None, "BANK_NAME", 2)])},
        imports={"I_T_BANKS_BRANCHES": [
            FileField(1, "T32COUNTRY", "COUNTRY", "X(60)", "C", None),
            FileField(2, "T32BANK_NAME", "BANK_NAME", "X(30)", "C", None)]},
        exports={"T_BANKS_BRANCHES": [
            FileField(1, "T32COUNTRY", "COUNTRY", "X(60)", None, ""),
            FileField(2, "T32BANK_NAME", "BANK_NAME", "X(30)", None, "")]},
        applications={})


def test_run_audit_produces_workbook_and_sizing_finding(tmp_path):
    snap = _build_snapshot()
    # Oracle catalog grouped by (template, tab): BANK_NAME is VARCHAR2(100) (snapshot has 30)
    catalog = {("Tmpl", "Bank Account"): [
        AlignedField(1, "Country", "COUNTRY", "VARCHAR2", 60, None, False),
        AlignedField(2, "Bank Name", "BANK_NAME", "VARCHAR2", 100, None, True)]}
    mapping = {("Tmpl", "Bank Account"): {"applaud_table": "T_BANKS_BRANCHES",
               "prefix": "T32", "module": "Fin", "status": "MAPPED", "in_base": ""}}
    appmap = {"T_BANKS_BRANCHES": AppMapRow("T_BANKS_BRANCHES",
              ["I_T_BANKS_BRANCHES"], ["T_BANKS_BRANCHES"],
              ["I_T_BANKS_BRANCHES", "X_T_BANKS_BRANCHES"], "confirmed")}
    out = tmp_path / "report.xlsx"
    findings = run_audit(snap, catalog, mapping, appmap, release="26B",
                         release_changes={}, out_path=out)
    assert out.exists()
    assert any(f.dimension == "1-SIZING" and f.oracle_field == "BANK_NAME"
               and f.severity == "HIGH" for f in findings)


def test_run_audit_thin_tab_label_only_no_spurious_presence(tmp_path):
    """§3.2 integration check: the canonical Bank Account tab has technical=None
    (labels only). With label->technical normalization wired, the IF's known-good
    fields must produce ZERO spurious 2-IF PRESENCE findings."""
    # Applaud IF + table use T32 bares; Oracle catalog rows are label-only (technical=None).
    bares = ["COUNTRY", "BANK_NAME", "BANK_CODE", "ALTERNATE_BANK_NAME"]
    labels = ["Country", "Bank Name", "Bank Code", "Alternate Bank Name"]
    snap = ApplaudSnapshot(
        system="ORACLE_MASTER", mdb_path="X", extracted_at="2026-06-02T00:00:00+00:00",
        extractor_version="1",
        tables={"T_BANKS_BRANCHES": SnapshotTable(
            name="T_BANKS_BRANCHES", prefix="T32", prefix_fallback=False,
            description="T_BANKS_BRANCHES (T32)", key_seqs=[["T32COUNTRY"]],
            columns=[DataColumn(f"T32{b}", b, "X", 100, None, "", i + 1)
                     for i, b in enumerate(bares)])},
        imports={"I_T_BANKS_BRANCHES": [
            FileField(i + 1, f"T32{b}", b, "X(100)", "C", None)
            for i, b in enumerate(bares)]},
        exports={}, applications={})
    catalog = {("RapidImplementationForCashManagement", "Bank Account"): [
        AlignedField(i + 1, lbl, None, None, None, None, None)   # technical=None
        for i, lbl in enumerate(labels)]}
    mapping = {("RapidImplementationForCashManagement", "Bank Account"): {
        "applaud_table": "T_BANKS_BRANCHES", "prefix": "T32", "module": "Fin",
        "status": "MAPPED", "in_base": ""}}
    appmap = {"T_BANKS_BRANCHES": AppMapRow("T_BANKS_BRANCHES",
              ["I_T_BANKS_BRANCHES"], [], ["I_T_BANKS_BRANCHES"], "confirmed")}
    findings = run_audit(snap, catalog, mapping, appmap, release="26B",
                         release_changes={}, out_path=tmp_path / "r.xlsx")
    if_presence = [f for f in findings
                   if f.dimension == "2-IF" and f.attribute == "PRESENCE"]
    assert if_presence == []   # normalization matched all four label-only fields
```

- [ ] **Step 3: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py -k run_audit -v`
Expected: FAIL with `ImportError: cannot import name 'run_audit'`

- [ ] **Step 4: Write the orchestration implementation**

```python
# fbdi/audit_applaud.py  (append)
from fbdi.applaud_snapshot import ApplaudSnapshot
from fbdi.applaud_appmap import AppMapRow


def run_audit(snapshot: ApplaudSnapshot,
              catalog: dict[tuple[str, str], list[AlignedField]],
              mapping: dict[tuple[str, str], dict],
              appmap: dict[str, AppMapRow],
              release: str,
              release_changes: dict[tuple[str, str], list[Change]],
              out_path: Path) -> list["Finding"]:
    findings: list[Finding] = []
    mapped_tables: set[str] = set()
    appmap_pairs: dict[str, tuple[list[str], list[str]]] = {}

    for (template, tab), info in mapping.items():
        table_name = info.get("applaud_table")
        if not table_name:
            continue
        oracle_fields = catalog.get((template, tab), [])
        if not oracle_fields:
            continue
        table = snapshot.tables.get(table_name)
        row = appmap.get(table_name)
        mapped_tables.add(table_name)
        ifs = row.import_files if row else []
        efs = row.export_files if row else []
        appmap_pairs[table_name] = (ifs, efs)

        oracle_by_bare = {oracle_match_key(f): f
                          for f in oracle_fields if oracle_match_key(f)}

        if table is not None:
            findings += check_sizing(template, tab, table_name, oracle_by_bare, table.columns)
            findings += check_table_coverage(template, tab, table_name, oracle_fields, table.columns)

        for if_name in ifs:
            if_fields = snapshot.imports.get(if_name, [])
            findings += check_file_coverage(template, tab, if_name, "IMPORT", "2-IF",
                                            oracle_fields, if_fields)
            if table is not None:
                findings += check_orphans(template, tab, table_name, if_name, "IMPORT",
                                          table.columns, if_fields)
        for ef_name in efs:
            ef_fields = snapshot.exports.get(ef_name, [])
            findings += check_file_coverage(template, tab, ef_name, "EXPORT", "3-EF",
                                            oracle_fields, ef_fields)
            if table is not None:
                findings += check_orphans(template, tab, table_name, ef_name, "EXPORT",
                                          table.columns, ef_fields)

        changes = release_changes.get((template, tab))
        if changes and table is not None:
            applaud_bares = {c.bare.upper() for c in table.columns}
            findings += check_release_delta(template, tab, table_name, changes,
                                            applaud_bares, old_release="(prior)",
                                            new_release=release)

    findings += check_unmapped(set(snapshot.tables), mapped_tables)
    coverage = coverage_gaps(mapped_tables, appmap_pairs)

    write_findings_workbook(findings, coverage,
        {"system": snapshot.system, "release": release,
         "extracted_at": snapshot.extracted_at}, out_path)
    return findings
```

- [ ] **Step 5: Run test to verify it passes**

Run: `py -m pytest tests/test_audit_applaud.py -k run_audit -v`
Expected: PASS

- [ ] **Step 6: Wire the CLI subcommand**

In `fbdi/cli.py`, after the `report` subparser block (around line 151) add:

```python
    audit_applaud_parser = subparsers.add_parser(
        "audit-applaud",
        help="Audit an Applaud system against the Oracle FBDI release it targets",
    )
    audit_applaud_parser.add_argument("--release", required=True, help="Release tag, e.g. 26B")
    audit_applaud_parser.add_argument("--old-release", default=None,
                                      help="Prior release tag for Dim 6b (e.g. 26A); aligns the "
                                           "catalog's old sheet against --release. Omit to skip 6b.")
    audit_applaud_parser.add_argument("--system", default="ORACLE_MASTER",
                                      help="Applaud system alias (default: ORACLE_MASTER)")
    audit_applaud_parser.add_argument("--catalog", type=Path,
                                      default=Path("FBDI_Master_Catalog.xlsx"))
    audit_applaud_parser.add_argument("--mapping", type=Path,
                                      default=Path("FBDI_to_ApplaudTables_Mapping.xlsx"))
    audit_applaud_parser.add_argument("--appmap", type=Path,
                                      default=Path("FBDI_to_Applaud_AppMap.xlsx"))
    audit_applaud_parser.add_argument("--output", type=Path, default=None)
```

In the dispatch block (around line 167) add:

```python
    elif args.command == "audit-applaud":
        _run_audit_applaud(args)
```

Add the handler function near the other `_run_*` functions:

```python
def _run_audit_applaud(args: argparse.Namespace) -> None:
    from fbdi.applaud_snapshot import ApplaudSnapshot
    from fbdi.applaud_appmap import load_appmap_workbook
    from fbdi.audit_applaud import run_audit, build_release_changes
    from fbdi.report import load_catalog_release, load_mapping
    from fbdi.config import applaud_snapshot_path

    snap_path = applaud_snapshot_path(args.system)
    if not snap_path.exists():
        print(f"Error: snapshot not found: {snap_path}. Run Step A (agent-driven extraction) first.")
        sys.exit(1)
    snapshot = ApplaudSnapshot.load(snap_path)
    catalog = load_catalog_release(args.catalog, args.release)
    mapping = load_mapping(args.mapping)
    appmap = load_appmap_workbook(args.appmap) if args.appmap.exists() else {}

    release_changes = {}
    if args.old_release:
        old_catalog = load_catalog_release(args.catalog, args.old_release)
        release_changes = build_release_changes(old_catalog, catalog)

    out = args.output or Path(f"Applaud_Compliance_Report_{args.release}_{args.system}.xlsx")

    findings = run_audit(snapshot, catalog, mapping, appmap, release=args.release,
                         release_changes=release_changes, out_path=out)
    print(f"Findings: {len(findings)}  (HIGH={sum(1 for f in findings if f.severity=='HIGH')})")
    print(f"Output written to: {out}")
```

- [ ] **Step 7: Run the full suite**

Run: `py -m pytest tests/ -q`
Expected: PASS (all existing + new tests)

- [ ] **Step 8: Commit**

```bash
git add fbdi/config.py fbdi/audit_applaud.py fbdi/cli.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): run_audit orchestration + audit-applaud CLI + config"
```

---

## Task 16: Step A extraction reference doc + gitignore + .keep

**Files:**
- Create: `docs/superpowers/references/applaud-snapshot-extraction.md`
- Modify: `.gitignore`

Step A has no CLI (it's agent-driven), so the exact query sequence must be documented so the agent — and later the orchestrator skill (Candidate C) — runs it consistently with the row-count guard.

- [ ] **Step 1: Write the reference doc**

```markdown
# Step A — Applaud snapshot extraction (agent-driven)

The agent runs this sequence with applaud-mcp, passing `system: 'ORACLE_MASTER'`
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
   c. `derive_prefix(description, [col DDIDs])` → prefix P.
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
```

- [ ] **Step 2: Add gitignore entries + snapshot dir keeper**

Append to `.gitignore`:

```
# Applaud audit — gitignored snapshot (the app-map workbook IS tracked)
baselines/applaud/applaud_snapshot_*.json
Applaud_Compliance_Report_*.xlsx
```

Create `baselines/applaud/.keep` (empty file) so the directory exists.

- [ ] **Step 3: Verify the doc renders and gitignore works**

Run: `git check-ignore baselines/applaud/applaud_snapshot_ORACLE_MASTER.json`
Expected: prints the path (ignored). `FBDI_to_Applaud_AppMap.xlsx` must NOT be ignored.

Run: `git check-ignore FBDI_to_Applaud_AppMap.xlsx; echo "exit=$?"`
Expected: `exit=1` (not ignored — it is tracked source of truth)

- [ ] **Step 4: Commit**

```bash
git add docs/superpowers/references/applaud-snapshot-extraction.md .gitignore baselines/applaud/.keep
git commit -m "docs(applaud-audit): Step A extraction reference + gitignore rules"
```

---

## Self-Review

**1. Spec coverage:**
- §2 scope (T_* only, single-prefix, prefix fallback) → Tasks 3, 11 (exact-DDID orphans), and the matching logic throughout. ✓
- §4 per-object extraction + COUNT guard → Tasks 2, 16. ✓
- §4 app-map derivation + confirmed-wins + lists → Tasks 4, 5. ✓
- §4 EF via get_application steps (no X_ assumption) → Task 4 (`_steps_of_type`), Task 9. ✓
- §5 Dims 1–5 → Tasks 7, 8, 9, 10, 11. ✓
- §6 6b release-delta, 6c unmapped → Tasks 12, 13. ✓
- §6 6a (required-field) deferred → not implemented (correct). ✓
- §7 Finding model + finding_id + 4-sheet workbook + Status/Notes → Tasks 6, 14. ✓
- §3 audit-applaud CLI, config systems → Task 15. ✓
- §11 MCP config already applied (out of plan scope). ✓
- §10 tests: row-count guard (Task 2), prefix fallback (Task 3), each dimension, EF bare-DDID (Task 9), Excel writer (Task 14). finding_id reconciliation is Phase 2 — finding_id is generated (Task 6) but reconciliation is out of Phase-1 scope per spec. ✓

**Pass-2 audit coverage (`AUDIT_RESULTS_plan_pass2.md` §4 acceptance):**
1. Dim 1 sources type/size from `DataDictionary`; `build_table` joins DD by DDID; regression test feeds blank-`DatabaseDetail` + populated-DD and asserts `actual_shape` (Tasks 2, 7). ✓
2. `@`-prefixed fields excluded at assembly (`is_audit_field`, Task 2) + from LCP prefix fallback (Task 3), with tests. ✓
3. Task 16 reference doc corrected: "DataDictionary IS pulled; DatabaseDetail has no type data." ✓
4. `ODBCName` empty in ORACLE_MASTER → bare-name is the effective Dim 4 key (noted, Task 10). ✓
5. Integration check on the `T_BANKS_BRANCHES` / "Bank Account" thin tab confirms label-only Oracle fields produce zero spurious 2-IF PRESENCE findings; `oracle_match_key` wires `_label_to_technical` into every matching dim (Tasks 7, 15). ✓
6. §0 "keep as-is" items preserved unchanged. ✓
- §3.3 date-vs-char type-class test added (Task 7). ✓

**2. Placeholder scan:** No TBD/TODO; every code step has complete code; commands have expected output. ✓

**3. Type consistency:** `Finding`, `DataColumn`, `FileField`, `SnapshotTable`, `ApplaudSnapshot`, `AppMapRow`, `make_finding_id` signatures are identical across all tasks that use them. `check_file_coverage` is defined once (Task 8) and reused for EFs (Task 9) and in `run_audit` (Task 15). Dims 2/3 use a local set-difference + `_lcs_sequence`; only Dim 6b (Task 12) uses `align_tabs(old, new) -> list[Change]`, whose `ADDED`/`REMOVED` carry `new_field`/`old_field` respectively (verified against `fbdi/align.py:135-168`). ✓

**Notes for the third-pass spot-check:**
- **§3.2 normalization (the key thing to re-check):** `oracle_match_key` maps the catalog's
  label→technical via `_label_to_technical` when `technical` is None. On the live `Bank Account`
  tab this matched 22/23 fields; the one near-miss is catalog **"EDI ID Number"** (→`EDI_ID_NUMBER`)
  vs Applaud **`EFT_ID_NUMBER`** — a genuine naming divergence the report *should* surface (one
  HIGH "missing" + one INFO "extra"), not noise to suppress. Confirm the implementer's integration
  run reproduces ~22 clean matches + that single reviewable divergence, not 23 false positives.
- **`align_tabs` numbering-space caveat (already handled):** `align_tabs` classifies `SHIFTED` by absolute position equality (`fbdi/align.py:107`), so it is used **only** for Dim 6b (old vs new Oracle catalog rows, same numbering space). Dims 2/3 deliberately do their own presence (set) + ordering (LCS) so Oracle-position-vs-IF-row offsets don't produce spurious findings. Confirm this reasoning holds.
- **Dim 6b is wired via `--old-release`:** when provided, the CLI loads the prior release's catalog sheet and `build_release_changes` aligns it per tab (Task 12), so 6b is live (not a silent no-op). When `--old-release` is omitted, 6b is intentionally skipped. Confirm the catalog reliably carries both release sheets at audit time (it does for 26A/26B); if a release is missing its sheet, `load_catalog_release` raises `ValueError` — the auditor should decide whether to soften that to a warning.
