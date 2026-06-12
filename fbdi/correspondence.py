"""Oracle <-> Applaud field-name correspondence layer.

Pure-Python, no MCP/live I/O (same discipline as applaud_appmap.py). Proposes
candidate field correspondences for HITL confirmation, persists confirmed pairs
in a committed workbook, and resolves them into a {applaud_bare: oracle_key}
alias the audit applies before set-intersection.

Verified live against ORACLE_MASTER; see
docs/superpowers/AUDIT_RESULTS_field-correspondence.md for the data behind the
truncation/abbreviation/type rules.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

# Applaud's application-level name cap (NOT the DataDictionary.Name TEXT(60)
# schema). Longest observed bare is 27 with a 3-char prefix (audit §2.1).
APPLAUD_NAME_CAP = 30

# Longest suffix Applaud appends-then-truncates (NUMBER=6). Bounds how much a
# normalized name may exceed its counterpart on the "appended suffix" path so a
# genuinely short coincidental prefix is not read as a truncation hit (audit §2.2).
MAX_SUFFIX_SLACK = 6

# Data-grounded abbreviation seed (audit §8.3). Abbrev -> full expansion; expanded
# on both sides before comparison. Seed ONLY from post-exact residual divergences;
# do NOT add Oracle's own spellings (they already match in the exact pre-pass).
ABBREVIATIONS: dict[str, str] = {
    "BU": "BUSINESSUNIT",
    "BUS": "BUSINESS",
    "DISC": "DISCOUNT",
    "NUM": "NUMBER",
    "DESCR": "DESCRIPTION",
    "DESC": "DESCRIPTION",
    "AMT": "AMOUNT",
    "INV": "INVOICE",
    "COMP": "COMPONENT",
    "REFER": "REFERENCE",
}

_BOOL_SUFFIX_RE = re.compile(r"_(FLAG|FLG|F)$")
_TRAILING_DIGITS_RE = re.compile(r"(\d+)$")


def _split_trailing_digits(s: str) -> tuple[str, str]:
    """('TIMESTAMP10') -> ('TIMESTAMP', '10'); ('VENDOR_NAME') -> ('VENDOR_NAME', '')."""
    m = _TRAILING_DIGITS_RE.search(s)
    if not m:
        return s, ""
    return s[: m.start()], m.group(1)


def expand_abbreviations(name: str) -> str:
    """Token-wise abbreviation expansion on the underscore-delimited form.
    Unknown tokens pass through; already-expanded input is stable (idempotent)."""
    return "_".join(ABBREVIATIONS.get(tok, tok) for tok in name.split("_"))


def normalize_name(name: str) -> str:
    """Canonical comparison form: upper, strip '*', strip a full boolean suffix,
    expand abbreviations, then FULL underscore squash (audit §2.3 — collapse
    position carries no information). Returns a contiguous A-Z0-9 string."""
    s = (name or "").strip().upper().strip("*")
    s = _BOOL_SUFFIX_RE.sub("", s)
    s = expand_abbreviations(s)
    return s.replace("_", "")


def truncation_window(prefix: str | None) -> int:
    """Per-table truncation window = 30 - len(prefix) (audit §2.1)."""
    return APPLAUD_NAME_CAP - len(prefix or "")
