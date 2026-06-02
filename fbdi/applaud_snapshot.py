"""Applaud MDB snapshot — typed model, assembly helpers, JSON I/O.

Step A (extraction) is agent-driven: the agent calls applaud-mcp per-object and
feeds raw query-result rows (lists of dicts) to the assembly helpers here, which
validate (row-count guard) and produce a typed ApplaudSnapshot. No MCP I/O lives
in this module, so every function is unit-testable with synthetic inputs.

Two data-layer facts (verified live against ORACLE_MASTER/AP0STE.mdb) shape the
helpers:
  - DatabaseDetail carries NO type data (DataType/Size/DecPlaces/ODBCName are
    empty on real data); the canonical type/size lives on DataDictionary. So
    build_table joins each column's DDID to a DataDictionary slice.
  - "@"-prefixed DDIDs are internal Definian audit/tracking columns and are
    excluded from the snapshot (and thus from all Dim 1-6 matching).
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
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)  # baselines/applaud/ is gitignored, created at runtime
        path.write_text(
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


# ---------------------------------------------------------------------------
# Assembly helpers (Task 2) — fed raw MCP query results; no MCP I/O.
# ---------------------------------------------------------------------------

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
    (@...DO_NOT_LOAD, @...LEGACY_*). Excluded from all Dim 1-6 matching."""
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
