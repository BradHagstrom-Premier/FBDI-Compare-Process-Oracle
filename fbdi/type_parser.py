"""Parse Oracle data-type strings from FBDI templates into structured fields.

FBDI templates store types in a 'Data Type' row as strings like:
  VARCHAR2(5 CHAR), VARCHAR2(2048 CHAR), VARCHAR2(80), Varchar2(250),
  NUMBER(18), NUMBER(18,4), DATE, CLOB, BLOB

This module parses those strings once so downstream comparison to Applaud
doesn't re-parse on every run.
"""

import re
from dataclasses import dataclass


@dataclass
class ParsedType:
    """Result of parsing a data-type string.

    data_type is uppercase ('VARCHAR2', 'NUMBER', 'DATE'). Empty string
    means the input was blank/None. length and scale are None when
    absent. parse_warning is True only for non-empty inputs that couldn't
    be decoded; blank inputs are not warnings.
    """
    data_type: str
    length: int | None
    scale: int | None
    parse_warning: bool


# Shape:
#   TYPENAME
#   TYPENAME(length)
#   TYPENAME(length CHAR|BYTE)
#   TYPENAME(length,scale)
# Case-insensitive; leading/trailing whitespace tolerated.
_TYPE_RE = re.compile(
    r"^\s*"
    r"([A-Za-z][A-Za-z0-9]*)"              # type name
    r"\s*"
    r"(?:"
        r"\(\s*"
        r"(\d+)"                           # length / precision
        r"(?:\s*,\s*(\d+))?"               # optional scale
        r"(?:\s+(?:CHAR|BYTE))?"           # optional CHAR|BYTE suffix
        r"\s*\)"
    r")?"
    r"\s*$",
    re.IGNORECASE,
)


def parse_data_type(raw: str | None) -> ParsedType:
    """Parse an Oracle data-type string into (data_type, length, scale).

    Returns ParsedType with parse_warning=True when raw is non-empty but
    doesn't match any known shape. Blank/None returns an empty ParsedType
    with parse_warning=False (blank is legitimate, not a failure).
    """
    if raw is None or not str(raw).strip():
        return ParsedType("", None, None, False)

    m = _TYPE_RE.match(str(raw))
    if not m:
        return ParsedType("", None, None, True)

    dtype = m.group(1).upper()
    length = int(m.group(2)) if m.group(2) else None
    scale = int(m.group(3)) if m.group(3) else None
    return ParsedType(dtype, length, scale, False)
