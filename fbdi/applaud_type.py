"""Oracle → Applaud type translator.

Pure function. Consumes the ParsedType dataclass from fbdi.type_parser and
emits the Applaud-side type string used in the compliance report.

Mapping (per the design spec):
  VARCHAR2(N)   → "char N"
  NUMBER(p, s)  → "numeric p,s"
  NUMBER(p)     → "numeric p"
  NUMBER        → "numeric"          (no defaults invented)
  DATE          → "date"
  TIMESTAMP     → "date"
  CLOB/BLOB/RAW → "<type>" (lowercase passthrough)
  unknown       → "<type>" (lowercase passthrough)
  blank/parse_warning → ""           (don't fabricate a type)
"""

from __future__ import annotations

from fbdi.type_parser import ParsedType


def applaud_type_for(t: ParsedType) -> str:
    if not t.data_type or t.parse_warning:
        return ""

    name = t.data_type.upper()

    if name == "VARCHAR2":
        return f"char {t.length}" if t.length is not None else "char"

    if name == "NUMBER":
        if t.length is not None and t.scale is not None:
            return f"numeric {t.length},{t.scale}"
        if t.length is not None:
            return f"numeric {t.length}"
        return "numeric"

    if name in ("DATE", "TIMESTAMP"):
        return "date"

    return name.lower()
