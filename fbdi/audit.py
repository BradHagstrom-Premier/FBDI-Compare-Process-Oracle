"""
audit.py — FBDI ↔ Applaud mapping audit engine.

Consumes applaud_snapshot.json + FBDI_Master_Catalog.xlsx (26B) +
fbdi_applaud_mapping.xlsx and produces Claude_fbdi_applaud_mapping.xlsx
(3 sheets) + Claude_fbdi_applaud_mapping_audit.md.

Run: python -m fbdi.audit
"""
from __future__ import annotations

import json
import logging
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
_log = logging.getLogger(__name__)


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


# ---------------------------------------------------------------------------
# Signal computation
# ---------------------------------------------------------------------------

def compute_name_alignment(applaud_table: str, fbdi_tab: str) -> str:
    """Compare Applaud table name (strip T_) against FBDI tab name."""
    stripped = applaud_table.upper().removeprefix("T_")
    tab_upper = fbdi_tab.upper()

    if stripped == tab_upper:
        return "EXACT"

    # Try stripping suffixes from both sides
    def _base(s: str) -> str:
        for suffix in _STRIP_SUFFIXES:
            if s.endswith(suffix):
                return s[: -len(suffix)]
        return s

    if _base(stripped) == _base(tab_upper):
        return "PARTIAL"
    if _base(stripped) == tab_upper or stripped == _base(tab_upper):
        return "PARTIAL"

    return "NONE"


def compute_key_coverage(
    applaud_key_bare_names: set[str], fbdi_columns: set[str]
) -> float:
    if not applaud_key_bare_names:
        return 0.0
    fbdi_upper = {c.upper() for c in fbdi_columns}
    matched = sum(1 for k in applaud_key_bare_names if k.upper() in fbdi_upper)
    return matched / len(applaud_key_bare_names)


def compute_column_overlap(
    applaud_fields: list[SnapshotField], fbdi_columns: set[str]
) -> float:
    biz_fields = [f for f in applaud_fields if not f.is_legacy_tracking]
    if not biz_fields:
        return 0.0
    fbdi_upper = {c.upper() for c in fbdi_columns}
    matched = sum(1 for f in biz_fields if f.bare_name.upper() in fbdi_upper)
    return matched / len(biz_fields)


def check_prefix_conformance(
    applaud_table: str, prefix: str, fbdi_tab: str
) -> bool:
    """True when Applaud table name minus T_ exactly equals the FBDI tab name."""
    return applaud_table.upper().removeprefix("T_") == fbdi_tab.upper()


# ---------------------------------------------------------------------------
# Pass 1 — candidate index
# ---------------------------------------------------------------------------

_PASS1_MIN_NAME_ALIGNMENT = {"EXACT", "PARTIAL"}
_PASS1_MIN_KEY_COVERAGE = 0.5
_PASS1_MIN_COLUMN_OVERLAP = 0.3


def _sort_key(c: Candidate) -> tuple:
    align_order = {"EXACT": 0, "PARTIAL": 1, "NONE": 2}
    return (align_order[c.name_alignment], -c.key_coverage, -c.column_overlap)


def build_candidate_index(
    snapshot: ApplaudSnapshot, catalog: CatalogIndex
) -> CandidateIndex:
    index: CandidateIndex = {}
    table_by_name = snapshot.table_by_name()

    for applaud_table_name, snap_table in table_by_name.items():
        candidates: list[Candidate] = []
        key_bare = snap_table.key_bare_names()

        for (fbdi_file, fbdi_tab), fbdi_cols in catalog.items():
            name_align = compute_name_alignment(applaud_table_name, fbdi_tab)
            key_cov = compute_key_coverage(key_bare, fbdi_cols)
            col_ovlp = compute_column_overlap(snap_table.fields, fbdi_cols)
            prefix_ok = check_prefix_conformance(
                applaud_table_name, snap_table.prefix or "", fbdi_tab
            )

            # Pass-1 threshold: keep if any signal clears its floor
            if (
                name_align in _PASS1_MIN_NAME_ALIGNMENT
                or key_cov >= _PASS1_MIN_KEY_COVERAGE
                or col_ovlp >= _PASS1_MIN_COLUMN_OVERLAP
            ):
                fbdi_upper = {c.upper() for c in fbdi_cols}
                biz_fields = snap_table.business_fields()
                matched = [f.bare_name for f in biz_fields if f.bare_name.upper() in fbdi_upper]
                missing = [f.bare_name for f in biz_fields if f.bare_name.upper() not in fbdi_upper]
                key_matched = [k for k in key_bare if k.upper() in fbdi_upper]

                candidates.append(Candidate(
                    fbdi_file=fbdi_file,
                    fbdi_tab=fbdi_tab,
                    name_alignment=name_align,
                    key_coverage=key_cov,
                    column_overlap=col_ovlp,
                    prefix_conformance=prefix_ok,
                    applaud_key_fields_matched=key_matched,
                    applaud_fields_matched=matched,
                    applaud_fields_missing=missing,
                ))

        candidates.sort(key=_sort_key)
        index[applaud_table_name] = candidates

    return index


# ---------------------------------------------------------------------------
# Prior-mapping text parser
# ---------------------------------------------------------------------------

def parse_prior_mapping(mapping_text: str) -> list[tuple[str, str]]:
    """Parse "Template / Tab[; Template / Tab]" → [(file, tab), ...]."""
    result: list[tuple[str, str]] = []
    if not mapping_text or not mapping_text.strip():
        return result
    for segment in mapping_text.split(";"):
        segment = segment.strip()
        if not segment:
            continue
        parts = segment.split(" / ", maxsplit=1)
        if len(parts) != 2 or not parts[0].strip() or not parts[1].strip():
            _log.warning("Malformed prior mapping segment (skipping): %r", segment)
            continue
        result.append((parts[0].strip(), parts[1].strip()))
    return result


# ---------------------------------------------------------------------------
# Confidence tier evaluator
# ---------------------------------------------------------------------------

def evaluate_confidence(candidate: Candidate) -> str:
    """Return H, M, or L per spec §6.2. Evaluated in order; first match wins."""
    if (
        candidate.name_alignment == "EXACT"
        and (candidate.key_coverage == 1.0 or candidate.column_overlap >= 0.7)
    ):
        return "H"
    if candidate.name_alignment == "PARTIAL" or (
        0 < candidate.key_coverage < 1.0 and candidate.column_overlap >= 0.4
    ):
        return "M"
    return "L"


# ---------------------------------------------------------------------------
# Pass 2 — adjudication
# ---------------------------------------------------------------------------

_CARRYTHROUGH_VERDICTS = {"FILE_TOO_LARGE", "FILE_ERROR"}


def _find_candidate(
    candidates: list[Candidate], fbdi_file: str, fbdi_tab: str
) -> Candidate | None:
    for c in candidates:
        if c.fbdi_file == fbdi_file and c.fbdi_tab == fbdi_tab:
            return c
    return None


def adjudicate_table(
    applaud_table: str,
    snap_table: SnapshotTable | None,
    candidates: list[Candidate],
    prior: PriorRow,
) -> AuditRow:
    evidence = EvidenceBundle(candidates_evaluated=list(candidates))
    prefix = snap_table.prefix or prior.prefix if snap_table else prior.prefix

    # ── PREFLIGHT ────────────────────────────────────────────────────────────
    if snap_table is None and prior.prior_status not in _CARRYTHROUGH_VERDICTS:
        return AuditRow(
            applaud_table=applaud_table, prefix=prefix,
            verdict="UNMAPPED", fbdi_mapping="",
            confidence="H", rationale="Applaud table not present in MDB snapshot",
            prior_verdict=prior.prior_status, changed=False,
            needs_deep_rationale=False, evidence=evidence,
        )

    if prior.prior_status in _CARRYTHROUGH_VERDICTS:
        return AuditRow(
            applaud_table=applaud_table, prefix=prefix,
            verdict=prior.prior_status, fbdi_mapping=prior.mapping_text,
            confidence="", rationale="Sized out / unreadable in 26B — unchanged from prior",
            prior_verdict=prior.prior_status, changed=False,
            needs_deep_rationale=False, evidence=evidence,
        )

    # ── PRIOR MAPPING PARSE ──────────────────────────────────────────────────
    prior_claims = parse_prior_mapping(prior.mapping_text)
    best_candidate = candidates[0] if candidates else None

    verdict: str
    fbdi_mapping: str
    confidence: str
    rationale: str

    # ── UNMAPPED / blank ─────────────────────────────────────────────────────
    if prior.prior_status in ("UNMAPPED", "") or (
        prior.prior_status == "YES" and not prior.mapping_text.strip()
    ):
        if best_candidate:
            conf = evaluate_confidence(best_candidate)
            if conf == "H":
                verdict = "YES"
                fbdi_mapping = f"{best_candidate.fbdi_file} / {best_candidate.fbdi_tab}"
                confidence = "H"
                rationale = (
                    f"Promoted from UNMAPPED — EXACT name match, "
                    f"key={best_candidate.key_coverage:.0%}, overlap={best_candidate.column_overlap:.0%}"
                )
            elif conf == "M":
                verdict = "NEEDS_REVIEW"
                fbdi_mapping = f"{best_candidate.fbdi_file} / {best_candidate.fbdi_tab}"
                confidence = "M"
                rationale = "Potential new mapping — Medium confidence; verify with Brad"
            else:
                verdict = "UNMAPPED"
                fbdi_mapping = ""
                confidence = "H"
                rationale = "No FBDI tab in 26B catalog scores above threshold"
        else:
            verdict = "UNMAPPED"
            fbdi_mapping = ""
            confidence = "H"
            rationale = "No FBDI tab in 26B catalog scores above threshold"

    # ── SINGLE prior claim ───────────────────────────────────────────────────
    elif len(prior_claims) == 1:
        file, tab = prior_claims[0]
        matched_c = _find_candidate(candidates, file, tab)
        if matched_c:
            conf = evaluate_confidence(matched_c)
            if conf in ("H", "M"):
                verdict = "YES"
                fbdi_mapping = f"{file} / {tab}"
                confidence = conf
                rationale = (
                    f"name={matched_c.name_alignment}, "
                    f"key={matched_c.key_coverage:.0%}, "
                    f"overlap={matched_c.column_overlap:.0%}"
                )
            else:
                verdict = "NEEDS_REVIEW"
                fbdi_mapping = f"{file} / {tab}"
                confidence = "L"
                rationale = "Prior claim scores Low against 26B catalog — verify"
        else:
            verdict = "NEEDS_REVIEW"
            fbdi_mapping = f"{file} / {tab}"
            confidence = "L" if best_candidate else "H"
            rationale = (
                "Prior references file/tab not found in 26B catalog or below all thresholds"
            )

    # ── MULTI prior claims ───────────────────────────────────────────────────
    else:
        high_or_med: list[tuple[str, str, Candidate, str]] = []
        low_or_absent: list[tuple[str, str]] = []
        for file, tab in prior_claims:
            c = _find_candidate(candidates, file, tab)
            if c:
                conf = evaluate_confidence(c)
                if conf in ("H", "M"):
                    high_or_med.append((file, tab, c, conf))
                else:
                    low_or_absent.append((file, tab))
                    evidence.rejected_alternatives.append(c)
            else:
                low_or_absent.append((file, tab))

        if len(high_or_med) == len(prior_claims):
            # All claims score High or Medium → keep multi
            verdict = "YES"
            fbdi_mapping = "; ".join(f"{f} / {t}" for f, t, _, _ in high_or_med)
            confidence = "H" if all(conf == "H" for _, _, _, conf in high_or_med) else "M"
            rationale = f"Multi-mapping retained — {len(high_or_med)} legs verified"
        elif len(high_or_med) == 1:
            # One good leg — collapse to single
            file, tab, c, conf = high_or_med[0]
            verdict = "YES"
            fbdi_mapping = f"{file} / {tab}"
            confidence = conf
            rationale = (
                f"Collapsed from multi — 1/{len(prior_claims)} legs scored {conf}; "
                f"rest below threshold"
            )
        else:
            verdict = "NEEDS_REVIEW"
            fbdi_mapping = "; ".join(f"{f} / {t}" for f, t in prior_claims)
            confidence = "M" if high_or_med else "L"
            rationale = "Multi-mapping contested — see audit.md for per-leg evidence"

    # ── PREFIX AUDIT (all verdicts) ──────────────────────────────────────────
    if verdict == "YES" and fbdi_mapping:
        # Check prefix_conformance on the first/primary chosen candidate
        first_claim = parse_prior_mapping(fbdi_mapping)
        if first_claim:
            chosen_c = _find_candidate(candidates, first_claim[0][0], first_claim[0][1])
            if chosen_c and not chosen_c.prefix_conformance:
                evidence.notes.append(
                    f"Prefix mismatch — expected T_<tab> convention, "
                    f"got prefix={prefix} for tab={first_claim[0][1]}"
                )

    changed = verdict != prior.prior_status or fbdi_mapping != prior.mapping_text.strip()
    needs_deep = (
        verdict == "NEEDS_REVIEW"
        or changed
        or confidence == "L"
        or bool(evidence.notes)
    )

    return AuditRow(
        applaud_table=applaud_table, prefix=prefix,
        verdict=verdict, fbdi_mapping=fbdi_mapping,
        confidence=confidence, rationale=rationale,
        prior_verdict=prior.prior_status, changed=changed,
        needs_deep_rationale=needs_deep, evidence=evidence,
    )
