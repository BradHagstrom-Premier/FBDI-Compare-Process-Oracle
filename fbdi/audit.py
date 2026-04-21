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
