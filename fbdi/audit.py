"""
audit.py — FBDI ↔ Applaud mapping audit engine.

Consumes applaud_snapshot.json + FBDI_Master_Catalog.xlsx (26B) +
fbdi_applaud_mapping.xlsx and produces Claude_fbdi_applaud_mapping.xlsx
(3 sheets) + Claude_fbdi_applaud_mapping_audit.md.

Run: python -m fbdi.audit
"""
from __future__ import annotations

import json
import re
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


REPO_ROOT = Path(__file__).parent.parent
SNAPSHOT_PATH = REPO_ROOT / "applaud_snapshot.json"
CATALOG_PATH = REPO_ROOT / "FBDI_Master_Catalog.xlsx"
PRIOR_MAPPING_PATH = REPO_ROOT / "fbdi_applaud_mapping.xlsx"
OUTPUT_MAPPING_PATH = REPO_ROOT / "Claude_fbdi_applaud_mapping.xlsx"
OUTPUT_AUDIT_PATH = REPO_ROOT / "Claude_fbdi_applaud_mapping_audit.md"
CATALOG_RELEASE = "26B"
SNAPSHOT_MAX_AGE_DAYS = 30

_STRIP_SUFFIXES = ("_ALL", "_INT", "_INTERFACE")


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class SnapshotField:
    name: str
    bare_name: str
    is_legacy_tracking: bool
    data_type: str
    length: int


@dataclass
class SnapshotKeySeq:
    seq: str
    keys: list[str]


@dataclass
class SnapshotTable:
    name: str
    prefix: str | None
    description: str
    type: str
    key_sequences: list[SnapshotKeySeq]
    fields: list[SnapshotField]

    def business_fields(self) -> list[SnapshotField]:
        return [f for f in self.fields if not f.is_legacy_tracking]

    def key_bare_names(self) -> set[str]:
        bare: set[str] = set()
        for seq in self.key_sequences:
            for k in seq.keys:
                # Keys are stored as full prefixed names; strip prefix
                if self.prefix and k.upper().startswith(self.prefix.upper()):
                    bare.add(k[len(self.prefix):].upper())
                else:
                    bare.add(k.upper())
        return bare


@dataclass
class ApplaudSnapshot:
    mdb_path: str
    extracted_at: str
    extractor_version: str
    tables: list[SnapshotTable]
    missing_tables: list[dict]

    def table_by_name(self) -> dict[str, SnapshotTable]:
        return {t.name: t for t in self.tables}

    def missing_set(self) -> set[str]:
        return {m["name"] for m in self.missing_tables}


@dataclass
class Candidate:
    fbdi_file: str
    fbdi_tab: str
    name_alignment: str           # EXACT | PARTIAL | NONE
    key_coverage: float
    column_overlap: float
    prefix_conformance: bool
    applaud_key_fields_matched: list[str]
    applaud_fields_matched: list[str]
    applaud_fields_missing: list[str]


@dataclass
class EvidenceBundle:
    candidates_evaluated: list[Candidate] = field(default_factory=list)
    rejected_alternatives: list[Candidate] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass
class PriorRow:
    applaud_table: str
    prior_status: str
    prefix: str
    mapping_text: str
    module: str
    notes: str


@dataclass
class AuditRow:
    applaud_table: str
    prefix: str
    verdict: str                   # YES | UNMAPPED | NEEDS_REVIEW | FILE_TOO_LARGE | FILE_ERROR
    fbdi_mapping: str
    confidence: str                # H | M | L | ""
    rationale: str
    prior_verdict: str
    changed: bool
    needs_deep_rationale: bool
    evidence: EvidenceBundle


# Type aliases
CatalogIndex = dict[tuple[str, str], set[str]]   # {(file_name, tab_name): set[column_technical]}
CandidateIndex = dict[str, list[Candidate]]       # {applaud_table_name: sorted candidates}


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_snapshot(path: Path = SNAPSHOT_PATH) -> ApplaudSnapshot:
    if not path.exists():
        raise FileNotFoundError(f"Snapshot missing — run Step A first: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    tables = []
    for t in data["tables"]:
        fields = [
            SnapshotField(
                name=f["name"],
                bare_name=f["bare_name"],
                is_legacy_tracking=f["is_legacy_tracking"],
                data_type=f["data_type"],
                length=f["length"],
            )
            for f in t["fields"]
        ]
        key_seqs = [
            SnapshotKeySeq(seq=k["seq"], keys=k["keys"])
            for k in t["key_sequences"]
        ]
        tables.append(SnapshotTable(
            name=t["name"],
            prefix=t.get("prefix"),
            description=t.get("description", ""),
            type=t.get("type", ""),
            key_sequences=key_seqs,
            fields=fields,
        ))
    return ApplaudSnapshot(
        mdb_path=data["mdb_path"],
        extracted_at=data["extracted_at"],
        extractor_version=data["extractor_version"],
        tables=tables,
        missing_tables=data.get("missing_tables", []),
    )


def load_catalog(
    path: Path = CATALOG_PATH, release: str = CATALOG_RELEASE
) -> CatalogIndex:
    if not path.exists():
        raise FileNotFoundError(f"Catalog missing: {path}")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if release not in wb.sheetnames:
            raise ValueError(f"No '{release}' tab in catalog. Available: {wb.sheetnames}")
        ws = wb[release]
        rows_iter = ws.iter_rows(values_only=True)
        raw_headers = next(rows_iter)
        headers = [str(h).strip().lower() if h else "" for h in raw_headers]
        try:
            file_col = headers.index("file_name")
            tab_col = headers.index("tab_name")
            tech_col = headers.index("column_technical")
        except ValueError as exc:
            raise ValueError(f"Catalog missing expected header: {exc}. Got: {headers}")
        index: CatalogIndex = {}
        for row in rows_iter:
            fname = str(row[file_col]).strip() if row[file_col] else ""
            tab = str(row[tab_col]).strip() if row[tab_col] else ""
            tech = str(row[tech_col]).strip() if row[tech_col] else ""
            if fname and tab:
                key = (fname, tab)
                index.setdefault(key, set())
                if tech:
                    index[key].add(tech.upper())
        return index
    finally:
        wb.close()


def load_prior_mapping(path: Path = PRIOR_MAPPING_PATH) -> dict[str, PriorRow]:
    if not path.exists():
        raise FileNotFoundError(f"Prior mapping missing: {path}")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        # Find Applaud Tables sheet by name (case-insensitive), fall back to index 1
        ws = None
        for name in wb.sheetnames:
            if "applaud" in name.lower():
                ws = wb[name]
                break
        if ws is None:
            if len(wb.sheetnames) >= 2:
                ws = wb.worksheets[1]
            else:
                raise ValueError(
                    f"No 'Applaud Tables' sheet found. Sheets: {wb.sheetnames}"
                )
        rows_iter = ws.iter_rows(values_only=True)
        raw_headers = next(rows_iter)
        headers = [
            str(h).strip().lower().replace(" ", "_") if h else ""
            for h in raw_headers
        ]

        def _col(name: str) -> int:
            return headers.index(name) if name in headers else -1

        col_table = _col("applaud_table")
        col_status = _col("status")
        col_prefix = _col("prefix")
        col_mapping = _col("fbdi_template_mappings")
        col_module = _col("module")
        col_notes = _col("notes")

        if col_table == -1:
            raise ValueError(
                f"Sheet2 missing 'applaud_table' column. Headers found: {headers}"
            )

        result: dict[str, PriorRow] = {}
        for row in rows_iter:
            def _val(idx: int) -> str:
                if idx == -1 or idx >= len(row):
                    return ""
                v = row[idx]
                return str(v).strip() if v is not None else ""

            table_name = _val(col_table)
            if not table_name or table_name.startswith("#"):
                continue
            result[table_name] = PriorRow(
                applaud_table=table_name,
                prior_status=_val(col_status),
                prefix=_val(col_prefix),
                mapping_text=_val(col_mapping),
                module=_val(col_module),
                notes=_val(col_notes),
            )
        return result
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Prefix + bare_name utilities
# ---------------------------------------------------------------------------

_PREFIX_RE = re.compile(r'\(([A-Z0-9]+)\)\s*$')


def extract_prefix(description: str) -> str | None:
    m = _PREFIX_RE.search(description.strip())
    return m.group(1) if m else None


def derive_bare_name(field_name: str, prefix: str) -> tuple[str, bool]:
    """Return (bare_name, is_legacy_tracking)."""
    name = field_name
    is_legacy = False
    if name.startswith("@"):
        is_legacy = True
        name = name[1:]  # strip @
    upper_prefix = prefix.upper()
    if name.upper().startswith(upper_prefix):
        return name[len(prefix):], is_legacy
    return name, is_legacy
