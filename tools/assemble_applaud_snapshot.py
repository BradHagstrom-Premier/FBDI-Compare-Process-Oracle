"""Assemble an ApplaudSnapshot from the raw JSON produced by
tools/extract_applaud_snapshot.mjs (Step A, assembly half).

The Node extractor writes byte-exact raw rows to baselines/applaud/raw/extract.json;
this reads that and runs the pure-Python assembly helpers in fbdi.applaud_snapshot
(which drop @-audit fields and fail loud on any non-@ column missing a
DataDictionary entry), then writes the snapshot to the configured path.

Pilot scope: the 10 confirmed tables. (The orchestrator would generalise this
config from the confirmed app-map + a prefix source; embedded here for the pilot.)
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from fbdi.applaud_snapshot import ApplaudSnapshot, build_table, build_file_fields
from fbdi.config import applaud_snapshot_path

RAW = Path("baselines/applaud/raw/extract.json")

# table | prefix | import file | export file(s)
PILOT = [
    ("T_AP_INVOICE_INT",           "TA1", "I_T_AP_INVOICE_INT",           ["X_T_AP_INVOICE_INT", "X_T_AP_INVOICE_INT_TXT"]),
    ("T_AP_INVOICE_LINES",         "T99", "I_T_AP_INVOICE_LINES",         ["X_T_AP_INVOICE_LINES"]),
    ("T_BANKS_BRANCHES",           "T32", "I_T_BANKS_BRANCHES",           ["T_BANKS_BRANCHES"]),
    ("T_BPA_PO_LINES_INTERFACE",   "T64", "I_T_BPA_PO_LINES_INTERFACE",   ["X_T_BPA_PO_LINES_INTERFACE"]),
    ("T_EGP_COMPONENTS_INTERFACE", "T91", "I_T_EGP_COMPONENTS_INTERFACE", ["X_T_EGP_COMPONENTS_INTERFACE"]),
    ("T_EGP_ITEM_CATEGORIES_INT",  "T87", "I_T_EGP_ITEM_CATEGORIES_INT",  ["X_T_EGP_ITEM_CATEGORIES_INT"]),
    ("T_EGO_ITEM_INTF_EFF_B",      "T86", "I_T_EGO_ITEM_INTF_EFF_B",      ["X_T_EGO_ITEM_INTF_EFF_B"]),
    ("T_MSC_ST_ASSIGNMENT_SETS",   "T04", "I_T_MSC_ST_ASSIGNMENT_SETS",   ["X_T_MSC_ST_ASSIGNMENT_SETS"]),
    ("T_POZ_SUPPLIERS_INT",        "T07", "I_T_POZ_SUPPLIERS_INT",        ["X_T_POZ_SUPPLIERS"]),
    ("T_POZ_SUPPLIER_SITES_INT",   "T09", "I_T_POZ_SUPPLIER_SITES_INT",   ["X_T_POZ_SUPPLIER_SITES"]),
]


def main() -> None:
    """Read the raw extract JSON, assemble the pilot ApplaudSnapshot, and write it."""
    raw = json.loads(RAW.read_text(encoding="utf-8"))
    db = raw["database_detail"]
    dd = raw["data_dictionary"]
    imp = raw["import_detail"]
    exp = raw["export_detail"]

    tables, imports, exports = {}, {}, {}
    for table, prefix, if_name, ef_names in PILOT:
        dd_by_ddid = {r["Name"]: r for r in dd[prefix]}
        tables[table] = build_table(
            name=table, prefix=prefix, prefix_fallback=False,
            description=f"{table} ({prefix})", key_seqs=[],
            raw_columns=db[table], dd_by_ddid=dd_by_ddid,
        )
        imports[if_name] = build_file_fields(imp[if_name], prefix, "IF")
        for ef in ef_names:
            exports[ef] = build_file_fields(exp[ef], prefix, "EF")

    snap = ApplaudSnapshot(
        system=raw["_meta"]["system"],
        mdb_path=raw["_meta"]["mdb_path"],
        extracted_at=datetime.now(timezone.utc).isoformat(timespec="seconds"),
        extractor_version="first-run-pilot-1-mdbreader",
        tables=tables, imports=imports, exports=exports, applications={},
    )
    out = applaud_snapshot_path(snap.system)
    snap.write(out)

    print(f"Wrote snapshot: {out}")
    print(f"  tables={len(snap.tables)}  imports={len(snap.imports)}  exports={len(snap.exports)}")
    for name, t in sorted(snap.tables.items()):
        print(f"  {name:<30} cols={len(t.columns)}")


if __name__ == "__main__":
    main()
