"""Helper script to process MCP responses and build the snapshot incrementally.

Usage patterns:
  python _extract_helper.py parse_def <table_name> <def_text_file>
  python _extract_helper.py parse_fields <table_name> <prefix> <saved_json_path>
  python _extract_helper.py add_missing <table_name>
  python _extract_helper.py add_table <table_name> <def_json_str> <fields_json_str>
  python _extract_helper.py status
  python _extract_helper.py finalize
"""
import json
import re
import sys
from pathlib import Path
from datetime import datetime, timezone

PARTIAL = Path("_snapshot_partial.json")
MDB_PATH = "C:/Users/10193/Definian/MDB_for_ApplaudMCP/AP0STE.mdb"

def load_partial():
    return json.loads(PARTIAL.read_text())

def save_partial(p):
    PARTIAL.write_text(json.dumps(p, indent=2))

def parse_definition(def_text):
    """Parse get_table_definition output text into dict with prefix, description, type, key_sequences."""
    # Description line
    desc_m = re.search(r"^Description:\s*(.+?)\s*$", def_text, re.M)
    description = desc_m.group(1).strip() if desc_m else ""
    prefix_m = re.search(r"\(([A-Z0-9]+)\)\s*$", description)
    prefix = prefix_m.group(1) if prefix_m else ""

    type_m = re.search(r"^Type:\s*(\S+)", def_text, re.M)
    ttype = type_m.group(1).strip() if type_m else ""

    # Parse sequences
    key_sequences = []
    # Find each "Sequence 'N'" block and its keys
    seq_blocks = re.findall(
        r"Sequence '(\d+)'[^\n]*\n((?:\s+Key \d+:[^\n]+\n?)+)",
        def_text,
    )
    for seq_num, body in seq_blocks:
        keys = []
        for km in re.finditer(r"Key \d+:\s*([^\s(]+)", body):
            keys.append(km.group(1))
        key_sequences.append({"seq": seq_num, "keys": keys})

    return {
        "prefix": prefix,
        "description": description,
        "type": ttype,
        "key_sequences": key_sequences,
    }

def parse_fields(records, prefix):
    """Transform DataDictionary records into field dicts, filtered to those matching prefix."""
    fields = []
    warnings = []
    for r in records:
        full = r.get("Name", "")
        legacy = full.startswith("@")
        clean = full.lstrip("@")
        # Only keep if clean starts with prefix
        if not clean.upper().startswith(prefix.upper()):
            continue
        bare = clean[len(prefix):]
        fields.append({
            "name": full,
            "bare_name": bare,
            "is_legacy_tracking": legacy,
            "data_type": r.get("DataType", ""),
            "length": r.get("Size", 0),
        })
    return fields, warnings

def add_table(name, definition, fields):
    p = load_partial()
    if name in p["completed_names"]:
        return
    entry = {
        "name": name,
        "prefix": definition["prefix"],
        "description": definition["description"],
        "type": definition["type"],
        "key_sequences": definition["key_sequences"],
        "fields": fields,
    }
    p["tables"].append(entry)
    p["completed_names"].append(name)
    save_partial(p)

def add_missing(name, reason="not found via get_table_definition"):
    p = load_partial()
    if name in p["completed_names"]:
        return
    p["missing_tables"].append({"name": name, "reason": reason})
    p["completed_names"].append(name)
    save_partial(p)

def status():
    p = load_partial()
    with open("_working_set.json") as f:
        ws = json.load(f)
    done = set(p["completed_names"])
    pending = [item["name"] for item in ws if item["name"] not in done]
    print(f"completed: {len(done)}, pending: {len(pending)}, tables: {len(p['tables'])}, missing: {len(p['missing_tables'])}")
    if pending:
        print(f"next 10 pending: {pending[:10]}")
    return pending

def next_pending(n=15):
    p = load_partial()
    with open("_working_set.json") as f:
        ws = json.load(f)
    done = set(p["completed_names"])
    pending = [item["name"] for item in ws if item["name"] not in done]
    for name in pending[:n]:
        print(name)

def finalize():
    p = load_partial()
    snapshot = {
        "mdb_path": MDB_PATH,
        "extracted_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "extractor_version": "1",
        "tables": p["tables"],
        "missing_tables": p["missing_tables"],
    }
    Path("applaud_snapshot.json").write_text(json.dumps(snapshot, indent=2))
    print(f"Wrote applaud_snapshot.json: {len(snapshot['tables'])} tables, {len(snapshot['missing_tables'])} missing")

if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "status":
        status()
    elif cmd == "next":
        n = int(sys.argv[2]) if len(sys.argv) > 2 else 15
        next_pending(n)
    elif cmd == "add_missing":
        add_missing(sys.argv[2])
    elif cmd == "add_table":
        name = sys.argv[2]
        definition = json.loads(sys.argv[3])
        fields = json.loads(sys.argv[4])
        add_table(name, definition, fields)
    elif cmd == "finalize":
        finalize()
    elif cmd == "parse_def":
        # Read def text from file, emit JSON
        text = Path(sys.argv[2]).read_text()
        print(json.dumps(parse_definition(text)))
    elif cmd == "parse_fields":
        prefix = sys.argv[2]
        recs_path = sys.argv[3]
        records = json.loads(Path(recs_path).read_text())
        fields, warnings = parse_fields(records, prefix)
        print(json.dumps({"fields": fields, "warnings": warnings}))
    else:
        print(f"unknown command: {cmd}")
        sys.exit(2)
