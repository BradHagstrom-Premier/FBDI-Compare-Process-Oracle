# Applaud Field-Correspondence Layer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make confirmed Oracle↔Applaud field-name correspondences count as matches in the audit, so only genuinely unmatched fields become "missing field" findings — collapsing the pilot's ~957 false HIGH positives toward the real residual.

**Architecture:** A new pure-Python module `fbdi/correspondence.py` proposes candidate field correspondences (normalize → abbreviation-expand → truncation-aware match → type-class veto → greedy bijection → confidence tiers). A human confirms/overrides them in a review workbook; confirmed pairs persist in a committed `FBDI_to_Applaud_FieldMap.xlsx`. At audit time, `run_audit` aliases the **Applaud `bare`** side before set-intersection so the four existing dimension checks are untouched. This is the field-grain clone of the existing app-map `derived → confirmed` pattern (`fbdi/applaud_appmap.py`).

**Tech Stack:** Python 3.14+, openpyxl, pytest. Pure functions over `ApplaudSnapshot` / `AlignedField` / `DataColumn` — no MCP or live I/O in the module (same discipline as `applaud_appmap.py`).

**Authoritative inputs:** spec `docs/superpowers/specs/2026-06-10-applaud-field-correspondence-design.md` as amended by the audit `docs/superpowers/AUDIT_RESULTS_field-correspondence.md`. Where they conflict, the audit wins.

---

## Assumptions (REQUIRED items from the audit — every one has a task + test below)

1. **Digit-run truncation (audit §1.1).** Truncation is *not* always right-truncation: when a name ends in an ordinal, Oracle drops a letter from the middle to keep the digits (`GLOBAL_ATTRIBUTE_TIMESTAMP10` → `…TIMESTAM10`). Rule: when both sides end in a digit run, the runs must be **equal**; strip them and stem-match the remainders. → Task 3, `test_digit_run_truncation`.
2. **DataType code `U` → character class (audit §1.2).** `U` (Unicode text) is live inside the pilot (`T07VENDOR_NAME` is `U(100)`). `actual_shape` must bucket `U` with `X`. The type-class veto stays **strictly char-vs-numeric, never date-vs-char** (TIMESTAMP→`X(150)`, DATE→`D(8)`). → Task 1 (engine fix) + Task 3 (`test_u_column_not_vetoed`, `test_date_vs_char_does_not_veto`).
3. **Derivation input = the filtered column set (audit §1.3, §1.4).** Correspondence derivation must exclude `@`-prefixed audit fields and non-prefix working columns (e.g. `X_PHANTOM`) from the candidate pool — the same set the four audit checks see. → Task 3 (`test_derivation_excludes_audit_and_nonprefix`) + Task 8 (snapshot lock-in).
4. **`X_PHANTOM` is already handled at the snapshot layer — VERIFIED, not a new bug (audit §1.4, corrected).** `applaud_snapshot.build_table` (lines 163-179) already excludes non-prefix phantom columns, and `_strip_prefix` (line 126) is prefix-aware (it returns `X_PHANTOM` unchanged, never `HANTOM`). A live scan of the PR #3 workbook (`Applaud_Compliance_Report_26B_ORACLE_MASTER.xlsx`) found **zero** `HANTOM` findings. Task 8 therefore writes a *lock-in regression test* for the existing behavior — it does **not** "fix" a non-existent mis-strip.

**Confirmed assumptions (no code needed):** `bare` is the resolution key, ODBCName is empty in ORACLE_MASTER (audit §5/§8.4); one-to-one per table holds (§8.5); `Oracle Key == oracle_match_key(of)` (§8.6).

---

## File Structure

| File | Responsibility | Created/Modified |
|---|---|---|
| `fbdi/correspondence.py` | All correspondence logic: normalization, abbreviation table, scoring/tiers, derivation, fieldmap + review workbook I/O, merge, `apply_review_decisions`, `build_alias`. Pure-Python. | **Create** |
| `fbdi/audit_applaud.py` | `actual_shape` gains `U→char` (Task 1). `run_audit` gains `fieldmap`/`accept_confidence` params + aliasing + rejected-provenance annotation (Task 7). The four check functions are untouched. | Modify |
| `fbdi/cli.py` | New `correspondence-derive` / `correspondence-confirm` subcommands; `audit-applaud` gains `--fieldmap` / `--accept-confidence`. | Modify |
| `tests/test_correspondence.py` | All unit tests for the new module. | **Create** |
| `tests/test_audit_applaud.py` | The load-bearing aliasing regression (Task 7). | Modify (append) |
| `tests/test_applaud_snapshot.py` | X_PHANTOM / @-field lock-in (Task 8). | Modify (append) |
| `.gitignore` | Ignore disposable `Applaud_FieldMap_Review_*.xlsx`; keep `FBDI_to_Applaud_FieldMap.xlsx` tracked. | Modify |

**Naming locked across tasks** (a function called one name in Task 3 must be that name in Task 7):
`FieldCorrespondence(applaud_table, oracle_key, applaud_bare, applaud_ddid, confidence, origin="derived", score=0.0, signals="", notes="")`;
`normalize_name`, `expand_abbreviations`, `truncation_window`, `names_correspond`, `score_candidate`,
`derive_table_correspondences`, `derive_correspondences`,
`write_fieldmap_workbook`, `load_fieldmap_workbook`, `merge_fieldmap`,
`write_review_workbook`, `load_review_workbook`, `apply_review_decisions`, `InvalidCorrectedBareError`,
`build_alias`. Constants: `ABBREVIATIONS`, `APPLAUD_NAME_CAP=30`, `MAX_SUFFIX_SLACK=6`, `TIER_BANDS`.

---

## Task 1: Engine fix — `actual_shape` maps `U → char` (audit §1.2)

**Files:**
- Modify: `fbdi/audit_applaud.py:134-142`
- Test: `tests/test_audit_applaud.py`

This is small, isolated, and unblocks the type-veto correctness the later tasks depend on (correspondence reuses `actual_shape`).

- [ ] **Step 1: Write the failing test**

Append to `tests/test_audit_applaud.py`:

```python
from fbdi.applaud_snapshot import DataColumn
from fbdi.audit_applaud import actual_shape


def test_actual_shape_maps_u_to_char():
    # Audit §1.2: DataType 'U' (Unicode text) is character class, same bucket as 'X'.
    col = DataColumn(ddid="T07VENDOR_NAME", bare="VENDOR_NAME", data_type="U",
                     size=100, dec_places=None, odbc_name=None, row=1)
    assert actual_shape(col) == ("char", 100, None)


def test_actual_shape_keeps_x_and_n():
    x = DataColumn("T_X", "X1", "X", 50, None, None, 1)
    n = DataColumn("T_N", "N1", "N", 18, 4, None, 2)
    assert actual_shape(x) == ("char", 50, None)
    assert actual_shape(n) == ("numeric", 18, 4)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_audit_applaud.py::test_actual_shape_maps_u_to_char -v`
Expected: FAIL — current code returns `("u", 100, None)` (the `else` lowercases the code).

- [ ] **Step 3: Add the `U` branch**

In `fbdi/audit_applaud.py`, change `actual_shape` (currently lines 134-142):

```python
def actual_shape(col: DataColumn) -> Shape:
    """Applaud column's actual shape (class, size, scale) from its DataDictionary type.
    X and U (Unicode text, audit §1.2) -> char; N -> numeric; else the lowercased code
    (e.g. D -> 'd' for DATE columns)."""
    dt = (col.data_type or "").strip().upper()
    if dt in ("X", "U"):
        return ("char", col.size, None)
    if dt == "N":
        return ("numeric", col.size, col.dec_places)
    return (dt.lower(), col.size, col.dec_places)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_audit_applaud.py -k actual_shape -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "fix(applaud-audit): map DataType U to character class (audit §1.2)"
```

---

## Task 2: `correspondence.py` — normalization primitives + abbreviation table

**Files:**
- Create: `fbdi/correspondence.py`
- Test: `tests/test_correspondence.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_correspondence.py`:

```python
from fbdi.correspondence import (
    ABBREVIATIONS, APPLAUD_NAME_CAP, MAX_SUFFIX_SLACK,
    expand_abbreviations, normalize_name, truncation_window, _split_trailing_digits,
)


def test_squash_collapses_all_underscores():
    # Audit §2.3: underscore position carries no information -> full squash both sides.
    assert normalize_name("REMIT_ADVICEDELIVERY_METHOD") == "REMITADVICEDELIVERYMETHOD"
    assert normalize_name("REMIT_ADVICEDELIVERYMETHOD") == "REMITADVICEDELIVERYMETHOD"


def test_normalize_strips_star_and_uppercases():
    assert normalize_name("Supplier_Name*") == "SUPPLIERNAME"


def test_strip_bool_suffix_full_forms():
    # Full _FLAG/_FLG/_F are stripped; truncated forms are left to names_correspond (Task 3).
    assert normalize_name("ALWAYS_TAKE_DISCOUNT_FLAG") == normalize_name("ALWAYS_TAKE_DISCOUNT")


def test_expand_abbreviations_is_token_wise_and_bidirectional_safe():
    # BU -> BUSINESSUNIT expansion on the Oracle side.
    assert expand_abbreviations("PROCUREMENT_BU") == "PROCUREMENT_BUSINESSUNIT"
    # A token not in the table passes through unchanged (idempotent).
    assert expand_abbreviations("ITEM_NUMBER") == "ITEM_NUMBER"
    # Already-expanded input is stable.
    assert expand_abbreviations("PROCUREMENT_BUSINESSUNIT") == "PROCUREMENT_BUSINESSUNIT"


def test_abbreviation_table_has_data_grounded_seed():
    # Audit §8.3 seed entries.
    for k in ("BU", "BUS", "DISC", "NUM", "DESCR", "DESC", "AMT", "INV", "COMP", "REFER"):
        assert k in ABBREVIATIONS


def test_truncation_window_is_30_minus_prefix():
    assert truncation_window("T09") == APPLAUD_NAME_CAP - 3
    assert truncation_window("TA1") == 27


def test_split_trailing_digits():
    assert _split_trailing_digits("TIMESTAMP10") == ("TIMESTAMP", "10")
    assert _split_trailing_digits("VENDOR_NAME") == ("VENDOR_NAME", "")
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_correspondence.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'fbdi.correspondence'`.

- [ ] **Step 3: Write the module's primitives**

Create `fbdi/correspondence.py`:

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_correspondence.py -v`
Expected: PASS (7 tests).

- [ ] **Step 5: Commit**

```bash
git add fbdi/correspondence.py tests/test_correspondence.py
git commit -m "feat(correspondence): normalization primitives + abbreviation seed"
```

---

## Task 3: Matching, scoring, tiers, and per-table derivation

**Files:**
- Modify: `fbdi/correspondence.py`
- Test: `tests/test_correspondence.py`

This is the correctness core. `names_correspond` implements the truncation/digit-run rules; `score_candidate` weights name 0.6 / type 0.25 / position 0.15; `derive_table_correspondences` runs the exact pre-pass, vetoes char-vs-numeric, and greedily assigns a one-to-one bijection.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_correspondence.py`:

```python
from fbdi.align import AlignedField
from fbdi.applaud_snapshot import DataColumn
from fbdi.correspondence import (
    FieldCorrespondence, names_correspond, score_candidate, TIER_BANDS,
    derive_table_correspondences,
)


def _col(ddid, bare, dt="X", size=100, dec=None, row=1):
    return DataColumn(ddid=ddid, bare=bare, data_type=dt, size=size,
                      dec_places=dec, odbc_name=None, row=row)


def _of(technical, dt="VARCHAR2", length=100, scale=None, position=1):
    return AlignedField(position=position, label=None, technical=technical,
                        data_type=dt, length=length, scale=scale, required=None)


# --- names_correspond ---

def test_right_truncation_within_window():
    # CONSUMPTION_ADVICE_LINE_NUMBER -> Applaud lost the final R (truncated NUMBER).
    win = 27
    assert names_correspond("CONSUMPTIONADVICELINENUMBER",
                            "CONSUMPTIONADVICELINENUMBE", applaud_bare_len=26, window=win)


def test_appended_then_truncated_suffix():
    # PROCUREMENT_BU -> expand -> PROCUREMENTBUSINESSUNIT; Applaud appended NAME, truncated to NAM.
    assert names_correspond("PROCUREMENTBUSINESSUNIT",
                            "PROCUREMENTBUSINESSUNITNAM", applaud_bare_len=27, window=27)


def test_digit_run_truncation():
    # Audit §1.1 named test case. Digits (10) must be equal; stems differ by the dropped P.
    assert names_correspond("GLOBALATTRIBUTETIMESTAMP10",
                            "GLOBALATTRIBUTETIMESTAM10", applaud_bare_len=25, window=27)


def test_digit_run_must_be_equal():
    # TIMESTAMP10 must NOT match TIMESTAMP1 just because stems share a prefix.
    assert not names_correspond("GLOBALATTRIBUTETIMESTAMP10",
                                "GLOBALATTRIBUTETIMESTAMP1", applaud_bare_len=25, window=27)


def test_coincidental_short_prefix_does_not_match():
    # 'BANK' is a prefix of 'BANKACCOUNTNUMBER' but the delta is far past MAX_SUFFIX_SLACK
    # and Applaud was not truncated at the cap -> reject.
    assert not names_correspond("BANKACCOUNTNUMBER", "BANK", applaud_bare_len=4, window=27)


# --- type veto ---

def test_char_vs_numeric_vetoes_candidate():
    # Name matches but Oracle char vs Applaud numeric -> no candidate emitted.
    oracle = {"AMOUNT": _of("AMOUNT", dt="VARCHAR2", length=50)}
    cols = [_col("T01AMOUNT", "AMOUNT", dt="N", size=18, dec=2)]
    out = derive_table_correspondences("T_X", "T01", oracle, cols, decided=set())
    assert out == []


def test_u_column_not_vetoed():
    # Audit §1.2: U is char; an Oracle char field still matches a U column on a name divergence.
    oracle = {"VENDOR_NAME_NEW": _of("VENDOR_NAME_NEW", dt="VARCHAR2", length=100)}
    cols = [_col("T07VENDOR_NAMENEW", "VENDOR_NAMENEW", dt="U", size=100)]
    out = derive_table_correspondences("T_POZ", "T07", oracle, cols, decided=set())
    assert len(out) == 1 and out[0].applaud_bare == "VENDOR_NAMENEW"


def test_date_vs_char_does_not_veto():
    # Audit §1.2: Applaud stores TIMESTAMP as X (char); Oracle TIMESTAMP -> 'date'. No veto.
    oracle = {"GLOBAL_ATTRIBUTE_TIMESTAMP10": _of("GLOBAL_ATTRIBUTE_TIMESTAMP10",
                                                  dt="TIMESTAMP", length=None)}
    cols = [_col("T09GLOBAL_ATTRIBUTE_TIMESTAM10", "GLOBAL_ATTRIBUTE_TIMESTAM10",
                 dt="X", size=150)]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols, decided=set())
    assert len(out) == 1


# --- exact pre-pass + exclusions ---

def test_exact_matches_are_not_proposed():
    oracle = {"ITEM_NUMBER": _of("ITEM_NUMBER")}
    cols = [_col("T01ITEM_NUMBER", "ITEM_NUMBER")]
    assert derive_table_correspondences("T_X", "T01", oracle, cols, decided=set()) == []


def test_derivation_excludes_audit_and_nonprefix():
    # Audit §1.3/§1.4: @-fields and non-prefix working columns never enter the candidate pool.
    oracle = {"PROCUREMENT_BU": _of("PROCUREMENT_BU")}
    cols = [
        _col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM", size=25),
        _col("@T09LEGACY_AUDIT", "@T09LEGACY_AUDIT"),   # @-field (defensive)
        _col("X_PHANTOM", "X_PHANTOM"),                 # non-prefix working column
    ]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols, decided=set())
    bares = {c.applaud_bare for c in out}
    assert "PROCUREMENT_BUSINESSUNITNAM" in bares
    assert all(not b.startswith("@") and b != "X_PHANTOM" for b in bares)


def test_decided_pairs_are_skipped():
    oracle = {"PROCUREMENT_BU": _of("PROCUREMENT_BU")}
    cols = [_col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM", size=25)]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols,
                                       decided={("T_POZ", "PROCUREMENT_BU")})
    assert out == []


def test_bijection_one_to_one_per_table():
    # Two Oracle keys both plausibly match one Applaud column; only the best wins.
    oracle = {"PROCUREMENT_BU": _of("PROCUREMENT_BU", position=1),
              "PROCUREMENT_BUS": _of("PROCUREMENT_BUS", position=2)}
    cols = [_col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM", size=25)]
    out = derive_table_correspondences("T_POZ", "T09", oracle, cols, decided=set())
    assert len(out) == 1   # the single Applaud column is assigned once


def test_tiers_are_ordered_high_probable_weak():
    names = [b for b, _ in TIER_BANDS]
    assert names == ["HIGH", "PROBABLE", "WEAK"]
    pc = score_candidate(_of("PROCUREMENT_BU"),
                         _col("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM"),
                         window=27, position_score=1.0)
    assert pc[0] > 0.0 and pc[1]  # (score, signals) — signals non-empty
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_correspondence.py -k "names_correspond or veto or exact or bijection or tiers or derivation or decided or right_trunc or appended or digit or coincid or u_column or date_vs" -v`
Expected: FAIL — `ImportError: cannot import name 'FieldCorrespondence'`.

- [ ] **Step 3: Implement matching, scoring, tiers, and derivation**

Append to `fbdi/correspondence.py`:

```python
from fbdi.align import AlignedField, _lcs_match
from fbdi.applaud_snapshot import DataColumn
from fbdi.audit_applaud import expected_shape, actual_shape

# Score weights (audit §2.4 keeps position weak — DO NOT raise the 0.15).
_W_NAME, _W_TYPE, _W_POSITION = 0.6, 0.25, 0.15

# Confidence bands on the single weighted score, highest first. EXACT is handled
# in the pre-pass and never persisted, so it is not a band here.
TIER_BANDS: list[tuple[str, float]] = [("HIGH", 0.85), ("PROBABLE", 0.55), ("WEAK", 0.0)]


@dataclass
class FieldCorrespondence:
    applaud_table: str
    oracle_key: str
    applaud_bare: str
    applaud_ddid: str
    confidence: str               # HIGH | PROBABLE | WEAK  (or "confirmed"/"rejected" origin)
    origin: str = "derived"       # derived | confirmed | rejected
    score: float = 0.0
    signals: str = ""
    notes: str = ""


def names_correspond(oracle_norm: str, applaud_norm: str, applaud_bare_len: int,
                     window: int) -> bool:
    """True if the two normalized names plausibly denote the same field.

    Implements audit §1.1 (digit-run preservation), §2.1 (derived window) and
    §2.2 (truncation-aware prefix-of-other). `applaud_bare_len` is the RAW stored
    bare length (truncation happened on the stored name)."""
    # Digit-run preservation (audit §1.1): equal trailing digits, stem-match the rest.
    o_stem, o_dig = _split_trailing_digits(oracle_norm)
    a_stem, a_dig = _split_trailing_digits(applaud_norm)
    if o_dig or a_dig:
        if o_dig != a_dig:
            return False
        oracle_norm, applaud_norm = o_stem, a_stem

    if oracle_norm == applaud_norm:
        return True

    # Applaud right-truncated from a longer Oracle logical name.
    if oracle_norm.startswith(applaud_norm):
        delta = len(oracle_norm) - len(applaud_norm)
        # Valid if only a short suffix was lost, OR Applaud was cut at the cap
        # (a hard truncation legitimately drops many trailing chars).
        return delta <= MAX_SUFFIX_SLACK or applaud_bare_len >= window - 1

    # Applaud appended a suffix (NAME/FLAG/...) then possibly truncated it.
    if applaud_norm.startswith(oracle_norm):
        return len(applaud_norm) - len(oracle_norm) <= MAX_SUFFIX_SLACK

    return False


def _name_score(oracle_key: str, applaud_bare: str, applaud_bare_len: int,
                window: int) -> float:
    """1.0 if normalized-equal (a non-truncation divergence, e.g. abbreviation or
    underscore-only), 0.8 if a truncation/suffix match, 0.0 if no correspondence."""
    o, a = normalize_name(oracle_key), normalize_name(applaud_bare)
    if o == a:
        return 1.0
    if names_correspond(o, a, applaud_bare_len, window):
        return 0.8
    return 0.0


def _type_class_conflict(of: AlignedField, col: DataColumn) -> bool:
    """Char-vs-numeric clash only (audit §1.2) — never date-vs-char."""
    return {expected_shape(of)[0], actual_shape(col)[0]} == {"char", "numeric"}


def _type_score(of: AlignedField, col: DataColumn) -> float:
    exp_cls, act_cls = expected_shape(of)[0], actual_shape(col)[0]
    if exp_cls and act_cls and exp_cls == act_cls:
        return 1.0
    return 0.5


def _tier(score: float) -> str:
    for name, floor in TIER_BANDS:
        if score >= floor:
            return name
    return "WEAK"


def score_candidate(of: AlignedField, col: DataColumn, window: int,
                    position_score: float) -> tuple[float, str]:
    """Weighted score + a human-readable signals string. Caller has already
    confirmed a non-zero name correspondence and a passing type veto."""
    ns = _name_score(of.technical or "", col.bare, len(col.bare), window)
    ts = _type_score(of, col)
    score = _W_NAME * ns + _W_TYPE * ts + _W_POSITION * position_score
    signals = f"name={ns:.2f} type={ts:.2f} pos={position_score:.2f}"
    return round(score, 4), signals


def _candidate_excluded(col: DataColumn, prefix: str | None) -> bool:
    """Audit §1.3/§1.4: @-audit fields and non-prefix working columns (X_PHANTOM)
    never enter the candidate pool. Defensive — build_table already drops them."""
    if col.bare.lstrip().startswith("@") or col.ddid.lstrip().startswith("@"):
        return True
    if prefix and not col.ddid.upper().startswith(prefix.upper()):
        return True
    return False


def derive_table_correspondences(
    applaud_table: str, prefix: str | None,
    oracle_by_key: dict[str, AlignedField],
    applaud_columns: list[DataColumn],
    decided: set[tuple[str, str]],
) -> list[FieldCorrespondence]:
    """Propose one-to-one correspondences for the residual after the exact pre-pass.

    `oracle_by_key` maps oracle_match_key -> AlignedField. `decided` is the set of
    (table, oracle_key) pairs already in the committed map (confirmed/rejected) —
    never re-proposed."""
    window = truncation_window(prefix)
    cols = [c for c in applaud_columns if not _candidate_excluded(c, prefix)]

    # 1. Exact pre-pass: keys present verbatim on both sides need no map entry.
    applaud_bares = {c.bare.upper() for c in cols}
    residual_oracle = {k: of for k, of in oracle_by_key.items()
                       if k.upper() not in applaud_bares
                       and (applaud_table, k) not in decided}
    residual_cols = [c for c in cols if c.bare.upper() not in oracle_by_key]

    # Position support: LCS over the residual order (Oracle position vs Applaud row).
    o_order = [k for k in residual_oracle]
    a_order = [c.bare for c in sorted(residual_cols, key=lambda c: c.row)]

    # 2. Build candidate pairs (name-match + passing type veto), scored.
    candidates: list[tuple[float, str, str, DataColumn, str]] = []  # (score, tier, okey, col, signals)
    for o_idx, (okey, of) in enumerate(residual_oracle.items()):
        for a_idx, col in enumerate(residual_cols):
            if _name_score(okey, col.bare, len(col.bare), window) == 0.0:
                continue
            if _type_class_conflict(of, col):
                continue   # char-vs-numeric veto (audit §1.2)
            pos = _position_score(o_idx, a_idx, len(o_order), len(a_order))
            score, signals = score_candidate(of, col, window, pos)
            candidates.append((score, _tier(score), okey, col, signals))

    # 3. Greedy one-to-one bijection: best score first, both sides must be free.
    candidates.sort(key=lambda t: t[0], reverse=True)
    used_oracle: set[str] = set()
    used_bare: set[str] = set()
    out: list[FieldCorrespondence] = []
    for score, tier, okey, col, signals in candidates:
        if okey in used_oracle or col.bare.upper() in used_bare:
            continue
        used_oracle.add(okey)
        used_bare.add(col.bare.upper())
        out.append(FieldCorrespondence(
            applaud_table=applaud_table, oracle_key=okey, applaud_bare=col.bare,
            applaud_ddid=col.ddid, confidence=tier, origin="derived",
            score=score, signals=signals))
    return out


def _position_score(o_idx: int, a_idx: int, n_oracle: int, n_applaud: int) -> float:
    """Gentle order-agreement gradient in [0,1]; a weak tiebreak only (audit §2.4)."""
    span = max(n_oracle, n_applaud, 1)
    return max(0.0, 1.0 - abs(o_idx - a_idx) / span)


def derive_correspondences(
    tables: dict[str, tuple[str | None, dict[str, AlignedField], list[DataColumn]]],
    decided: set[tuple[str, str]],
) -> list[FieldCorrespondence]:
    """Derive across many tables. `tables` maps applaud_table ->
    (prefix, oracle_by_key, applaud_columns). Output sorted table -> tier -> score."""
    out: list[FieldCorrespondence] = []
    for table in sorted(tables):
        prefix, oracle_by_key, cols = tables[table]
        out.extend(derive_table_correspondences(table, prefix, oracle_by_key, cols, decided))
    tier_rank = {name: i for i, (name, _) in enumerate(TIER_BANDS)}
    out.sort(key=lambda fc: (fc.applaud_table, tier_rank.get(fc.confidence, 9), -fc.score))
    return out
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_correspondence.py -v`
Expected: PASS (all tests in the file).

- [ ] **Step 5: Commit**

```bash
git add fbdi/correspondence.py tests/test_correspondence.py
git commit -m "feat(correspondence): truncation/digit-run matching, type veto, bijection, tiers"
```

---

## Task 4: Fieldmap workbook I/O + `merge_fieldmap`

**Files:**
- Modify: `fbdi/correspondence.py`
- Test: `tests/test_correspondence.py`

Clones `applaud_appmap`'s workbook helpers. Committed map sheet `"Field Map"`, columns
`Applaud Table | Oracle Key | Applaud Bare | Applaud DDID | Confidence | Origin | Notes`.
Merge keys on `(table, oracle_key)`: confirmed/rejected win; a fresh derive fills only undecided pairs.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_correspondence.py`:

```python
from fbdi.correspondence import (
    write_fieldmap_workbook, load_fieldmap_workbook, merge_fieldmap,
)


def _fc(table, okey, bare, origin="derived", conf="HIGH"):
    return FieldCorrespondence(applaud_table=table, oracle_key=okey, applaud_bare=bare,
                               applaud_ddid=table[:3] + bare, confidence=conf, origin=origin)


def test_fieldmap_workbook_roundtrip(tmp_path):
    rows = [_fc("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM",
                origin="confirmed", conf="HIGH")]
    path = tmp_path / "fieldmap.xlsx"
    write_fieldmap_workbook(rows, path)
    loaded = load_fieldmap_workbook(path)
    assert loaded["T_POZ"][0].oracle_key == "PROCUREMENT_BU"
    assert loaded["T_POZ"][0].applaud_bare == "PROCUREMENT_BUSINESSUNITNAM"
    assert loaded["T_POZ"][0].origin == "confirmed"


def test_merge_confirmed_wins_over_rederive():
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "HAND_PICKED_BARE",
                               origin="confirmed")]}
    derived = [_fc("T_POZ", "PROCUREMENT_BU", "AUTO_BARE", origin="derived"),  # must NOT win
               _fc("T_POZ", "NEW_KEY", "NEW_BARE", origin="derived")]          # undecided -> added
    merged = merge_fieldmap(derived, committed)
    by = {(fc.oracle_key): fc for fc in merged["T_POZ"]}
    assert by["PROCUREMENT_BU"].applaud_bare == "HAND_PICKED_BARE"
    assert by["PROCUREMENT_BU"].origin == "confirmed"
    assert by["NEW_KEY"].origin == "derived"


def test_merge_rejected_also_wins_and_suppresses_reproposal():
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "", origin="rejected")]}
    derived = [_fc("T_POZ", "PROCUREMENT_BU", "AUTO_BARE", origin="derived")]
    merged = merge_fieldmap(derived, committed)
    assert merged["T_POZ"][0].origin == "rejected"


def test_rederive_idempotence_across_releases():
    # 26B decision survives a fresh 26B->next derive that re-proposes the same key.
    committed = {"T_POZ": [_fc("T_POZ", "PROCUREMENT_BU", "CONFIRMED_BARE",
                               origin="confirmed")]}
    rederived = [_fc("T_POZ", "PROCUREMENT_BU", "AUTO_BARE", origin="derived")]
    once = merge_fieldmap(rederived, committed)
    twice = merge_fieldmap(rederived, once)
    assert twice["T_POZ"][0].applaud_bare == "CONFIRMED_BARE"
    assert len(twice["T_POZ"]) == 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_correspondence.py -k "fieldmap or merge or rederive" -v`
Expected: FAIL — `ImportError: cannot import name 'write_fieldmap_workbook'`.

- [ ] **Step 3: Implement workbook I/O + merge**

Append to `fbdi/correspondence.py` (add `from pathlib import Path` and
`from openpyxl import Workbook, load_workbook` to the import block at the top of the file):

```python
_FIELDMAP_HEADERS = ["Applaud Table", "Oracle Key", "Applaud Bare", "Applaud DDID",
                     "Confidence", "Origin", "Notes"]


def write_fieldmap_workbook(rows: list[FieldCorrespondence], path: Path) -> None:
    """Write the committed Oracle<->Applaud field map (sheet 'Field Map')."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Field Map"
    ws.append(_FIELDMAP_HEADERS)
    for r in sorted(rows, key=lambda fc: (fc.applaud_table, fc.oracle_key)):
        ws.append([r.applaud_table, r.oracle_key, r.applaud_bare, r.applaud_ddid,
                   r.confidence, r.origin, r.notes])
    ws.freeze_panes = "A2"
    wb.save(path)


def load_fieldmap_workbook(path: Path) -> dict[str, list[FieldCorrespondence]]:
    """Load the committed field map into {applaud_table: [FieldCorrespondence, ...]}."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Field Map"] if "Field Map" in wb.sheetnames else wb.active
    out: dict[str, list[FieldCorrespondence]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        table, okey, bare, ddid, conf, origin, notes = (list(row) + [None] * 7)[:7]
        if not table or not okey:
            continue
        out.setdefault(str(table), []).append(FieldCorrespondence(
            applaud_table=str(table), oracle_key=str(okey),
            applaud_bare=(str(bare) if bare else ""), applaud_ddid=(str(ddid) if ddid else ""),
            confidence=(str(conf) if conf else ""), origin=(str(origin) if origin else "derived"),
            notes=(str(notes) if notes else "")))
    wb.close()
    return out


def merge_fieldmap(
    derived: list[FieldCorrespondence],
    committed: dict[str, list[FieldCorrespondence]],
) -> dict[str, list[FieldCorrespondence]]:
    """Confirmed/rejected rows win; derived rows fill only undecided (table, oracle_key).
    Clones merge_appmap semantics (applaud_appmap.py:165). Idempotent across re-derives."""
    out: dict[str, dict[str, FieldCorrespondence]] = {}
    for table, rows in committed.items():
        out[table] = {r.oracle_key: r for r in rows}
    for r in derived:
        bucket = out.setdefault(r.applaud_table, {})
        if r.oracle_key not in bucket:
            bucket[r.oracle_key] = r
    return {table: [bucket[k] for k in sorted(bucket)] for table, bucket in sorted(out.items())}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_correspondence.py -k "fieldmap or merge or rederive" -v`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add fbdi/correspondence.py tests/test_correspondence.py
git commit -m "feat(correspondence): committed fieldmap workbook I/O + merge precedence"
```

---

## Task 5: Review workbook emit + load + `apply_review_decisions` (Corrected Bare fail-loud)

**Files:**
- Modify: `fbdi/correspondence.py`
- Test: `tests/test_correspondence.py`

Review workbook columns:
`Applaud Table | Oracle Key | Oracle Type | Candidate Applaud Bare | Applaud DDID | Applaud Type | Confidence | Score | Signals | Conflicts/Alternatives | Confirm? | Corrected Bare`.
`apply_review_decisions` turns reviewer input into `FieldCorrespondence` rows: `Y`→confirmed; non-empty `Corrected Bare`→confirmed with the substitute (VALIDATED against the table's bare set, audit §4.1); `N`→rejected.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_correspondence.py`:

```python
import pytest
from fbdi.correspondence import (
    ReviewRow, write_review_workbook, load_review_workbook,
    apply_review_decisions, InvalidCorrectedBareError,
)


def _review(table, okey, cand, confirm="", corrected=""):
    return ReviewRow(applaud_table=table, oracle_key=okey, oracle_type="char 100",
                     candidate_bare=cand, applaud_ddid=table[:3] + cand,
                     applaud_type="char 25", confidence="HIGH", score=0.88,
                     signals="name=0.80", alternatives="", confirm=confirm,
                     corrected_bare=corrected)


def test_review_workbook_roundtrip(tmp_path):
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM")]
    path = tmp_path / "review.xlsx"
    write_review_workbook(rows, path, exact_counts={"T_POZ": (212, 226)})
    loaded = load_review_workbook(path)
    assert loaded[0].oracle_key == "PROCUREMENT_BU"
    assert loaded[0].candidate_bare == "PROCUREMENT_BUSINESSUNITNAM"


def test_apply_confirm_yes_becomes_confirmed():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM", confirm="Y")]
    valid = {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}}
    out = apply_review_decisions(rows, valid)
    assert out[0].origin == "confirmed"
    assert out[0].applaud_bare == "PROCUREMENT_BUSINESSUNITNAM"


def test_apply_confirm_no_becomes_rejected():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM", confirm="N")]
    out = apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})
    assert out[0].origin == "rejected"


def test_apply_corrected_bare_overrides_candidate():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "WRONG_GUESS", corrected="REAL_BARE")]
    out = apply_review_decisions(rows, {"T_POZ": {"REAL_BARE", "WRONG_GUESS"}})
    assert out[0].origin == "confirmed" and out[0].applaud_bare == "REAL_BARE"


def test_apply_corrected_bare_not_in_table_fails_loud():
    # Audit §4.1: a typo'd Corrected Bare must abort the merge, not commit a dead alias.
    rows = [_review("T_POZ", "PROCUREMENT_BU", "WRONG_GUESS", corrected="TYPOO_BARE")]
    with pytest.raises(InvalidCorrectedBareError):
        apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})


def test_apply_skips_undecided_rows():
    rows = [_review("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM")]  # no Y/N
    out = apply_review_decisions(rows, {"T_POZ": {"PROCUREMENT_BUSINESSUNITNAM"}})
    assert out == []
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_correspondence.py -k "review or apply or corrected" -v`
Expected: FAIL — `ImportError: cannot import name 'ReviewRow'`.

- [ ] **Step 3: Implement review workbook + decisions**

Append to `fbdi/correspondence.py`:

```python
_REVIEW_HEADERS = ["Applaud Table", "Oracle Key", "Oracle Type", "Candidate Applaud Bare",
                   "Applaud DDID", "Applaud Type", "Confidence", "Score", "Signals",
                   "Conflicts/Alternatives", "Confirm?", "Corrected Bare"]


@dataclass
class ReviewRow:
    applaud_table: str
    oracle_key: str
    oracle_type: str
    candidate_bare: str
    applaud_ddid: str
    applaud_type: str
    confidence: str
    score: float
    signals: str
    alternatives: str
    confirm: str = ""          # reviewer input: 'Y' | 'N' | ''
    corrected_bare: str = ""   # reviewer input: substitute bare


class InvalidCorrectedBareError(ValueError):
    """Raised when a reviewer-entered Corrected Bare is not an actual bare in the
    table (audit §4.1) — fail loud rather than commit an alias that maps to nothing."""


def write_review_workbook(rows: list[ReviewRow], path: Path,
                          exact_counts: dict[str, tuple[int, int]] | None = None) -> None:
    """Write the disposable HITL review workbook. `exact_counts` maps table ->
    (exact_matched, total) so the reviewer sees denominator context (audit §6)."""
    exact_counts = exact_counts or {}
    wb = Workbook()
    ws = wb.active
    ws.title = "Review"
    ws.append(_REVIEW_HEADERS)
    current_table = None
    for r in rows:
        if r.applaud_table != current_table:
            current_table = r.applaud_table
            matched, total = exact_counts.get(current_table, (0, 0))
            ws.append([f"--- {current_table}: {matched} of {total} matched exactly; "
                       f"deciding the residual {max(total - matched, 0)} ---"])
        ws.append([r.applaud_table, r.oracle_key, r.oracle_type, r.candidate_bare,
                   r.applaud_ddid, r.applaud_type, r.confidence, r.score, r.signals,
                   r.alternatives, r.confirm, r.corrected_bare])
    ws.freeze_panes = "A2"
    wb.save(path)


def load_review_workbook(path: Path) -> list[ReviewRow]:
    """Load reviewer decisions. Header-separator rows (a single '--- ...' cell) are skipped."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Review"] if "Review" in wb.sheetnames else wb.active
    out: list[ReviewRow] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        cells = (list(row) + [None] * 12)[:12]
        table, okey = cells[0], cells[1]
        if not table or not okey:
            continue   # blank row or a "--- table header ---" separator
        out.append(ReviewRow(
            applaud_table=str(table), oracle_key=str(okey),
            oracle_type=str(cells[2] or ""), candidate_bare=str(cells[3] or ""),
            applaud_ddid=str(cells[4] or ""), applaud_type=str(cells[5] or ""),
            confidence=str(cells[6] or ""), score=float(cells[7] or 0.0),
            signals=str(cells[8] or ""), alternatives=str(cells[9] or ""),
            confirm=str(cells[10] or "").strip(), corrected_bare=str(cells[11] or "").strip()))
    wb.close()
    return out


def apply_review_decisions(
    rows: list[ReviewRow],
    valid_bares_by_table: dict[str, set[str]],
) -> list[FieldCorrespondence]:
    """Turn reviewer input into FieldCorrespondence rows.

    Corrected Bare (if present) wins -> confirmed with the substitute, VALIDATED
    against the table's bare set (audit §4.1, fail loud). Else Confirm? 'Y' ->
    confirmed; 'N' -> rejected; blank -> skipped (undecided)."""
    out: list[FieldCorrespondence] = []
    for r in rows:
        valid = {b.upper() for b in valid_bares_by_table.get(r.applaud_table, set())}
        if r.corrected_bare:
            if r.corrected_bare.upper() not in valid:
                raise InvalidCorrectedBareError(
                    f"{r.applaud_table}: Corrected Bare {r.corrected_bare!r} for Oracle "
                    f"key {r.oracle_key!r} is not a column in that table. Fix the typo "
                    "or clear the cell; refusing to commit an alias that maps to nothing.")
            out.append(FieldCorrespondence(
                applaud_table=r.applaud_table, oracle_key=r.oracle_key,
                applaud_bare=r.corrected_bare, applaud_ddid=r.applaud_ddid,
                confidence="HIGH", origin="confirmed", notes="reviewer-corrected"))
        elif r.confirm.upper() == "Y":
            out.append(FieldCorrespondence(
                applaud_table=r.applaud_table, oracle_key=r.oracle_key,
                applaud_bare=r.candidate_bare, applaud_ddid=r.applaud_ddid,
                confidence=r.confidence, origin="confirmed", score=r.score,
                signals=r.signals))
        elif r.confirm.upper() == "N":
            out.append(FieldCorrespondence(
                applaud_table=r.applaud_table, oracle_key=r.oracle_key,
                applaud_bare="", applaud_ddid="", confidence=r.confidence,
                origin="rejected", notes="reviewer-rejected"))
        # blank Confirm? + no Corrected Bare -> undecided, skip
    return out
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_correspondence.py -k "review or apply or corrected" -v`
Expected: PASS (6 tests).

- [ ] **Step 5: Commit**

```bash
git add fbdi/correspondence.py tests/test_correspondence.py
git commit -m "feat(correspondence): review workbook + fail-loud Corrected Bare validation (audit §4.1)"
```

---

## Task 6: `build_alias` resolver + confidence gate

**Files:**
- Modify: `fbdi/correspondence.py`
- Test: `tests/test_correspondence.py`

`build_alias(fieldmap_for_table, accept_confidence)` → `{applaud_bare_upper: oracle_key_upper}`.
Default gate `confirmed` (only `origin=confirmed`). A tier name (`HIGH`/`PROBABLE`/`WEAK`) additionally admits `origin=derived` rows at or above that tier — a pre-review noise-reduction pass. `rejected` rows are never aliased.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_correspondence.py`:

```python
from fbdi.correspondence import build_alias


def test_build_alias_confirmed_only_by_default():
    rows = [_fc("T_POZ", "PROCUREMENT_BU", "PROCUREMENT_BUSINESSUNITNAM",
                origin="confirmed"),
            _fc("T_POZ", "OTHER_KEY", "OTHER_BARE", origin="derived", conf="HIGH")]
    alias = build_alias(rows, accept_confidence="confirmed")
    assert alias == {"PROCUREMENT_BUSINESSUNITNAM": "PROCUREMENT_BU"}


def test_build_alias_admits_derived_at_or_above_tier():
    rows = [_fc("T_POZ", "K1", "BARE_HIGH", origin="derived", conf="HIGH"),
            _fc("T_POZ", "K2", "BARE_WEAK", origin="derived", conf="WEAK")]
    alias = build_alias(rows, accept_confidence="HIGH")
    assert alias == {"BARE_HIGH": "K1"}   # WEAK excluded


def test_build_alias_never_aliases_rejected():
    rows = [_fc("T_POZ", "K1", "BARE", origin="rejected", conf="HIGH")]
    assert build_alias(rows, accept_confidence="WEAK") == {}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_correspondence.py -k build_alias -v`
Expected: FAIL — `ImportError: cannot import name 'build_alias'`.

- [ ] **Step 3: Implement the resolver**

Append to `fbdi/correspondence.py`:

```python
def build_alias(fieldmap_for_table: list[FieldCorrespondence],
                accept_confidence: str = "confirmed") -> dict[str, str]:
    """Resolve a table's field map into {applaud_bare_upper: oracle_key_upper}.

    'confirmed' (default): only origin=confirmed rows. A tier name ('HIGH' /
    'PROBABLE' / 'WEAK') additionally admits origin=derived rows at or above that
    tier (pre-review pass). origin=rejected is never aliased."""
    tier_rank = {name: i for i, (name, _) in enumerate(TIER_BANDS)}  # HIGH=0 best
    gate = (accept_confidence or "confirmed").strip()
    alias: dict[str, str] = {}
    for fc in fieldmap_for_table:
        if fc.origin == "rejected" or not fc.applaud_bare:
            continue
        if fc.origin == "confirmed":
            admit = True
        elif gate == "confirmed":
            admit = False
        else:
            admit = tier_rank.get(fc.confidence, 99) <= tier_rank.get(gate, -1)
        if admit:
            alias[fc.applaud_bare.upper()] = fc.oracle_key.upper()
    return alias
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_correspondence.py -k build_alias -v`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add fbdi/correspondence.py tests/test_correspondence.py
git commit -m "feat(correspondence): build_alias resolver + confidence gate"
```

---

## Task 7: Wire `run_audit` aliasing + rejected-provenance + load-bearing regression

**Files:**
- Modify: `fbdi/audit_applaud.py:513-579`
- Test: `tests/test_audit_applaud.py`

`run_audit` gains `fieldmap` + `accept_confidence` params. Inside the per-table loop, after
`table = snapshot.tables.get(table_name)`, build the alias and pass **aliased copies** of the
table columns and IF/EF fields (via `dataclasses.replace(c, bare=...)`) into the four checks.
DDID is left untouched (Dim 5). `rejected` keys annotate matching missing-field findings with a
provenance note (audit §4.2) — severity unchanged.

- [ ] **Step 1: Write the failing regression test**

Append to `tests/test_audit_applaud.py`:

```python
import dataclasses
from pathlib import Path

from fbdi.align import AlignedField
from fbdi.applaud_snapshot import ApplaudSnapshot, SnapshotTable, DataColumn
from fbdi.correspondence import FieldCorrespondence
from fbdi.audit_applaud import run_audit


def _snapshot_with_procurement(size):
    col = DataColumn(ddid="T09PROCUREMENT_BUSINESSUNITNAM",
                     bare="PROCUREMENT_BUSINESSUNITNAM", data_type="X", size=size,
                     dec_places=None, odbc_name=None, row=1)
    table = SnapshotTable(name="T_POZ_SUPPLIER_SITES_INT", prefix="T09",
                          prefix_fallback=False, description="(T09)", key_seqs=[],
                          columns=[col])
    return ApplaudSnapshot(system="ORACLE_MASTER", mdb_path="x", extracted_at="2026-06-11",
                           extractor_version="t", tables={"T_POZ_SUPPLIER_SITES_INT": table})


def _procurement_catalog():
    # Oracle field PROCUREMENT_BU, char 40 -> bigger than the Applaud char 25 column.
    of = AlignedField(position=1, label=None, technical="PROCUREMENT_BU",
                      data_type="VARCHAR2", length=40, scale=None, required=None)
    return {("PO_TPL", "Suppliers"): [of]}


_MAPPING = {("PO_TPL", "Suppliers"): {"applaud_table": "T_POZ_SUPPLIER_SITES_INT"}}


def test_alias_collapses_missing_field_and_fires_sizing(tmp_path):
    snap = _snapshot_with_procurement(size=25)
    catalog = _procurement_catalog()

    # No fieldmap: PROCUREMENT_BU reads as a missing field (Dim 4 HIGH).
    out_a = tmp_path / "a.xlsx"
    base = run_audit(snap, catalog, _MAPPING, appmap={}, release="26B",
                     release_changes={}, out_path=out_a)
    assert any(f.dimension == "4-TABLE" and f.oracle_field == "PROCUREMENT_BU"
               and f.current_value == "absent" for f in base)

    # With a confirmed alias: the missing-field finding vanishes AND Dim 1 sizing fires
    # (Applaud char 25 < Oracle char 40), surfacing the previously-skipped resize gap.
    fieldmap = {"T_POZ_SUPPLIER_SITES_INT": [FieldCorrespondence(
        applaud_table="T_POZ_SUPPLIER_SITES_INT", oracle_key="PROCUREMENT_BU",
        applaud_bare="PROCUREMENT_BUSINESSUNITNAM",
        applaud_ddid="T09PROCUREMENT_BUSINESSUNITNAM", confidence="HIGH",
        origin="confirmed")]}
    out_b = tmp_path / "b.xlsx"
    aliased = run_audit(snap, catalog, _MAPPING, appmap={}, release="26B",
                        release_changes={}, out_path=out_b, fieldmap=fieldmap)
    assert not any(f.dimension == "4-TABLE" and f.oracle_field == "PROCUREMENT_BU"
                   and f.current_value == "absent" for f in aliased)
    assert any(f.dimension == "1-SIZING" and f.attribute == "SIZE"
               and f.applaud_object_name == "T_POZ_SUPPLIER_SITES_INT" for f in aliased)


def test_rejected_key_annotates_finding_without_changing_severity(tmp_path):
    snap = _snapshot_with_procurement(size=25)
    catalog = _procurement_catalog()
    fieldmap = {"T_POZ_SUPPLIER_SITES_INT": [FieldCorrespondence(
        applaud_table="T_POZ_SUPPLIER_SITES_INT", oracle_key="PROCUREMENT_BU",
        applaud_bare="", applaud_ddid="", confidence="HIGH", origin="rejected")]}
    out = tmp_path / "c.xlsx"
    findings = run_audit(snap, catalog, _MAPPING, appmap={}, release="26B",
                         release_changes={}, out_path=out, fieldmap=fieldmap)
    miss = [f for f in findings if f.dimension == "4-TABLE"
            and f.oracle_field == "PROCUREMENT_BU" and f.current_value == "absent"]
    assert len(miss) == 1
    assert miss[0].severity == "HIGH"                  # unchanged
    assert "Reviewed" in miss[0].notes                 # provenance note added
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `py -m pytest tests/test_audit_applaud.py -k "alias_collapses or rejected_key" -v`
Expected: FAIL — `run_audit() got an unexpected keyword argument 'fieldmap'`.

- [ ] **Step 3: Add the params, aliasing, and annotation to `run_audit`**

In `fbdi/audit_applaud.py`, add `import dataclasses` near the top imports. Then update the
`run_audit` signature and the per-table loop. Change the signature (currently lines 513-520):

```python
def run_audit(snapshot: ApplaudSnapshot,
              catalog: dict[tuple[str, str], list[AlignedField]],
              mapping: dict[tuple[str, str], dict],
              appmap: dict[str, AppMapRow],
              release: str,
              release_changes: dict[tuple[str, str], list[Change]],
              out_path: Path,
              old_release: str | None = None,
              fieldmap: dict[str, list] | None = None,
              accept_confidence: str = "confirmed") -> list[Finding]:
```

Immediately inside the per-table loop, after `table = snapshot.tables.get(table_name)`
(current line 534), insert the aliasing block:

```python
        table = snapshot.tables.get(table_name)

        # --- Field-correspondence aliasing (spec §7) -------------------------
        # Alias the Applaud *bare* side so the four checks below match renamed
        # fields. DDID is left untouched (Dim 5 orphans match on DDID).
        from fbdi.correspondence import build_alias
        fm_rows = (fieldmap or {}).get(table_name, [])
        alias = build_alias(fm_rows, accept_confidence) if fm_rows else {}
        rejected_keys = {r.oracle_key.upper() for r in fm_rows
                         if getattr(r, "origin", "") == "rejected"}

        def _aliased(seq):
            return [dataclasses.replace(c, bare=alias.get(c.bare.upper(), c.bare))
                    for c in seq]

        if table is not None and alias:
            table = dataclasses.replace(table, columns=_aliased(table.columns))
        # ---------------------------------------------------------------------

        n_before = len(findings)
```

Then alias the IF/EF fields where they are fetched. Replace the IF loop body
(`if_fields = snapshot.imports.get(if_name, [])`) so the fetched list is aliased before use:

```python
        for if_name in ifs:
            if_fields = _aliased(snapshot.imports.get(if_name, [])) if alias \
                else snapshot.imports.get(if_name, [])
            findings += check_file_coverage(template, tab, if_name, "IMPORT", "2-IF",
                                            oracle_fields, if_fields)
            if table is not None:
                findings += check_orphans(template, tab, table_name, if_name, "IMPORT",
                                          table.columns, if_fields)
        for ef_name in efs:
            ef_fields = _aliased(snapshot.exports.get(ef_name, [])) if alias \
                else snapshot.exports.get(ef_name, [])
            findings += check_file_coverage(template, tab, ef_name, "EXPORT", "3-EF",
                                            oracle_fields, ef_fields)
            if table is not None:
                findings += check_orphans(template, tab, table_name, ef_name, "EXPORT",
                                          table.columns, ef_fields)
```

> NOTE: `check_orphans` matches on DDID, and `_aliased` only replaces `bare` (DDID
> is preserved by `dataclasses.replace`), so Dim 5 is unaffected — exactly as the spec requires.

Finally, after the `check_release_delta` block that closes the per-table work (just before the
loop ends, after the Dim 6b `findings += check_release_delta(...)` call), annotate rejected
findings added during this table's iteration:

```python
        # Provenance for reviewer-rejected keys (audit §4.2): note, don't suppress.
        if rejected_keys:
            for f in findings[n_before:]:
                if (f.applaud_object_name == table_name and f.current_value == "absent"
                        and f.oracle_field.upper() in rejected_keys):
                    f.notes = ("Reviewed — confirmed no Applaud counterpart"
                               if not f.notes else f.notes)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `py -m pytest tests/test_audit_applaud.py -k "alias_collapses or rejected_key" -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Run the full audit suite to confirm no regression**

Run: `py -m pytest tests/test_audit_applaud.py -v`
Expected: PASS (all — existing tests still green; default `fieldmap=None` is a no-op).

- [ ] **Step 6: Commit**

```bash
git add fbdi/audit_applaud.py tests/test_audit_applaud.py
git commit -m "feat(applaud-audit): alias Applaud bare via fieldmap; annotate rejected findings (spec §7, audit §4.2)"
```

---

## Task 8: Lock-in regression — snapshot excludes `X_PHANTOM` and `@`-fields (audit §1.4, corrected)

**Files:**
- Test: `tests/test_applaud_snapshot.py`

No production code changes — `build_table` (lines 163-179) and `_strip_prefix` (line 126) already
do the right thing, and the PR #3 workbook has zero `HANTOM` findings. This task pins that behavior
so a future refactor can't reintroduce a `HANTOM` mis-strip, and confirms the correspondence
candidate pool inherits a clean column set.

- [ ] **Step 1: Write the lock-in tests**

Append to `tests/test_applaud_snapshot.py` (match the existing import style in that file):

```python
from fbdi.applaud_snapshot import build_table, build_file_fields, _strip_prefix


def test_strip_prefix_is_prefix_aware_never_mangles_nonprefix():
    # Audit §1.4 (corrected): a non-prefix name is returned unchanged, NOT 'HANTOM'.
    assert _strip_prefix("X_PHANTOM", "T91") == "X_PHANTOM"
    assert _strip_prefix("T91COMP_NAME", "T91") == "COMP_NAME"


def test_build_table_excludes_nonprefix_phantom_column(caplog):
    raw_columns = [
        {"Row": 1, "DDID": "T91COMP_NAME", "ODBCName": None},
        {"Row": 2, "DDID": "X_PHANTOM", "ODBCName": None},   # non-prefix working column
    ]
    dd = {"T91COMP_NAME": {"DataType": "X", "Size": 100, "DecPlaces": None}}
    table = build_table("T_EGP_COMPONENTS_INTERFACE", "T91", False,
                        "(T91)", [], raw_columns, dd)
    bares = {c.bare for c in table.columns}
    assert bares == {"COMP_NAME"}                 # X_PHANTOM dropped, no 'HANTOM'
    assert all("HANTOM" not in c.bare for c in table.columns)


def test_build_table_excludes_audit_fields():
    raw_columns = [
        {"Row": 1, "DDID": "T91COMP_NAME", "ODBCName": None},
        {"Row": 2, "DDID": "@T91LEGACY_AUDIT", "ODBCName": None},
    ]
    dd = {"T91COMP_NAME": {"DataType": "X", "Size": 100, "DecPlaces": None}}
    table = build_table("T_EGP_COMPONENTS_INTERFACE", "T91", False,
                        "(T91)", [], raw_columns, dd)
    assert [c.ddid for c in table.columns] == ["T91COMP_NAME"]
```

- [ ] **Step 2: Run tests to verify they pass immediately**

Run: `py -m pytest tests/test_applaud_snapshot.py -k "strip_prefix or phantom or audit_fields" -v`
Expected: PASS (3 tests) — these document already-correct behavior. If any FAIL, the snapshot
layer regressed and must be fixed there before proceeding (do not patch around it in
`correspondence.py`).

- [ ] **Step 3: Commit**

```bash
git add tests/test_applaud_snapshot.py
git commit -m "test(applaud-snapshot): lock in X_PHANTOM / @-field exclusion (audit §1.4)"
```

---

## Task 9: CLI — `correspondence-derive`, `correspondence-confirm`, `audit-applaud --fieldmap`

**Files:**
- Modify: `fbdi/cli.py`
- Test: `tests/test_correspondence.py` (a thin assembly test) + manual smoke

The derive/confirm commands load snapshot + catalog + mapping exactly like `_run_audit_applaud`
(`cli.py:453-518`). `audit-applaud` gains `--fieldmap` (default `FBDI_to_Applaud_FieldMap.xlsx`,
loaded if present) and `--accept-confidence` (default `confirmed`).

- [ ] **Step 1: Write a failing assembly test for the derive helper**

The CLI assembles a `{table: (prefix, oracle_by_key, columns)}` dict for `derive_correspondences`.
Factor that assembly into a pure helper so it is testable without argparse. Append to
`tests/test_correspondence.py`:

```python
from fbdi.correspondence import assemble_derivation_inputs


def test_assemble_derivation_inputs_groups_by_table():
    snap = _snapshot_with_procurement_sites()   # helper below
    catalog = {("PO_TPL", "Suppliers"): [_of("PROCUREMENT_BU")]}
    mapping = {("PO_TPL", "Suppliers"): {"applaud_table": "T_POZ_SUPPLIER_SITES_INT"}}
    inputs = assemble_derivation_inputs(snap, catalog, mapping)
    prefix, oracle_by_key, cols = inputs["T_POZ_SUPPLIER_SITES_INT"]
    assert prefix == "T09"
    assert "PROCUREMENT_BU" in oracle_by_key
    assert cols and cols[0].bare == "PROCUREMENT_BUSINESSUNITNAM"


def _snapshot_with_procurement_sites():
    from fbdi.applaud_snapshot import ApplaudSnapshot, SnapshotTable, DataColumn
    col = DataColumn("T09PROCUREMENT_BUSINESSUNITNAM", "PROCUREMENT_BUSINESSUNITNAM",
                     "X", 25, None, None, 1)
    t = SnapshotTable("T_POZ_SUPPLIER_SITES_INT", "T09", False, "(T09)", [], [col])
    return ApplaudSnapshot("ORACLE_MASTER", "x", "2026-06-11", "t",
                           tables={"T_POZ_SUPPLIER_SITES_INT": t})
```

(`_of` is already defined in the test file from Task 3.)

- [ ] **Step 2: Run test to verify it fails**

Run: `py -m pytest tests/test_correspondence.py -k assemble -v`
Expected: FAIL — `ImportError: cannot import name 'assemble_derivation_inputs'`.

- [ ] **Step 3: Implement the assembly helper**

Append to `fbdi/correspondence.py`:

```python
def assemble_derivation_inputs(
    snapshot, catalog: dict[tuple[str, str], list[AlignedField]],
    mapping: dict[tuple[str, str], dict],
) -> dict[str, tuple[str | None, dict[str, AlignedField], list[DataColumn]]]:
    """Group the audit's (template, tab)->table chain into per-table derivation inputs:
    {applaud_table: (prefix, {oracle_match_key: AlignedField}, [DataColumn, ...])}.
    Mirrors run_audit's loop so derivation sees exactly the audit's column set."""
    from fbdi.audit_applaud import oracle_match_key
    out: dict[str, tuple[str | None, dict[str, AlignedField], list[DataColumn]]] = {}
    for (template, tab), info in mapping.items():
        table_name = info.get("applaud_table")
        if not table_name:
            continue
        oracle_fields = catalog.get((template, tab), [])
        table = snapshot.tables.get(table_name)
        if not oracle_fields or table is None:
            continue
        oracle_by_key = {oracle_match_key(f): f for f in oracle_fields if oracle_match_key(f)}
        prefix = table.prefix
        out[table_name] = (prefix, oracle_by_key, list(table.columns))
    return out
```

- [ ] **Step 4: Run test to verify it passes**

Run: `py -m pytest tests/test_correspondence.py -k assemble -v`
Expected: PASS.

- [ ] **Step 5: Add the two subparsers and the `audit-applaud` flags**

In `fbdi/cli.py`, after the `audit_applaud_parser` block (ends at line 176), add:

```python
    corr_derive_parser = subparsers.add_parser(
        "correspondence-derive",
        help="Propose Oracle<->Applaud field correspondences into a review workbook")
    corr_derive_parser.add_argument("--release", required=True, help="Release tag, e.g. 26B")
    corr_derive_parser.add_argument("--system", default="ORACLE_MASTER")
    corr_derive_parser.add_argument("--catalog", type=Path,
                                    default=Path("FBDI_Master_Catalog.xlsx"))
    corr_derive_parser.add_argument("--mapping", type=Path,
                                    default=Path("FBDI_to_ApplaudTables_Mapping.xlsx"))
    corr_derive_parser.add_argument("--map", dest="fieldmap", type=Path,
                                    default=Path("FBDI_to_Applaud_FieldMap.xlsx"),
                                    help="Committed field map; already-decided pairs are skipped")
    corr_derive_parser.add_argument("--tables", default=None,
                                    help="Comma-separated Applaud target tables to scope to")
    corr_derive_parser.add_argument("--output", type=Path, default=None)

    corr_confirm_parser = subparsers.add_parser(
        "correspondence-confirm",
        help="Merge reviewer decisions from a review workbook into the committed field map")
    corr_confirm_parser.add_argument("--review", type=Path, required=True)
    corr_confirm_parser.add_argument("--system", default="ORACLE_MASTER")
    corr_confirm_parser.add_argument("--map", dest="fieldmap", type=Path,
                                     default=Path("FBDI_to_Applaud_FieldMap.xlsx"))
```

Add the `audit-applaud` flags after line 171 (`--output`):

```python
    audit_applaud_parser.add_argument("--fieldmap", type=Path,
                                      default=Path("FBDI_to_Applaud_FieldMap.xlsx"),
                                      help="Committed Oracle<->Applaud field map (loaded if present)")
    audit_applaud_parser.add_argument("--accept-confidence", default="confirmed",
                                      choices=["confirmed", "HIGH", "PROBABLE", "WEAK"],
                                      help="Minimum acceptance for aliasing (default: confirmed)")
```

Add the dispatch arms after the `audit-applaud` arm (line 195):

```python
    elif args.command == "correspondence-derive":
        _run_correspondence_derive(args)
    elif args.command == "correspondence-confirm":
        _run_correspondence_confirm(args)
```

- [ ] **Step 6: Wire the fieldmap into `_run_audit_applaud`**

In `fbdi/cli.py`, in `_run_audit_applaud`, after the `appmap = ...` line (currently line 501) add:

```python
    from fbdi.correspondence import load_fieldmap_workbook
    fieldmap = load_fieldmap_workbook(args.fieldmap) if args.fieldmap.exists() else None
```

and change the `run_audit(...)` call (lines 514-515) to pass it through:

```python
    findings = run_audit(snapshot, catalog, mapping, appmap, release=release,
                         release_changes=release_changes, out_path=out, old_release=old_release,
                         fieldmap=fieldmap, accept_confidence=args.accept_confidence)
```

- [ ] **Step 7: Implement the two command handlers**

Add to `fbdi/cli.py` (near `_run_audit_applaud`):

```python
def _run_correspondence_derive(args: argparse.Namespace) -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(name)s: %(message)s")
    from fbdi.applaud_snapshot import ApplaudSnapshot
    from fbdi.report import load_catalog_release, load_mapping
    from fbdi.config import applaud_snapshot_path
    from fbdi.correspondence import (
        assemble_derivation_inputs, derive_correspondences, load_fieldmap_workbook,
        write_review_workbook, ReviewRow, normalize_name,
    )
    from fbdi.audit_applaud import expected_shape, actual_shape

    snap_path = applaud_snapshot_path(args.system)
    if not snap_path.exists():
        print(f"Error: snapshot not found: {snap_path}. Run Step A extraction first.")
        sys.exit(1)
    snapshot = ApplaudSnapshot.load(snap_path)
    release = args.release.upper()
    try:
        catalog = load_catalog_release(args.catalog, release)
    except ValueError as exc:
        print(f"Error: {exc}"); sys.exit(1)
    mapping = load_mapping(args.mapping)
    if args.tables:
        from fbdi.audit_applaud import filter_mapping_to_tables, UnknownTableError
        names = [t for t in args.tables.split(",") if t.strip()]
        try:
            mapping = filter_mapping_to_tables(mapping, names)
        except UnknownTableError as exc:
            print(f"Error: {exc}"); sys.exit(1)

    committed = load_fieldmap_workbook(args.fieldmap) if args.fieldmap.exists() else {}
    decided = {(t, fc.oracle_key) for t, rows in committed.items() for fc in rows}

    inputs = assemble_derivation_inputs(snapshot, catalog, mapping)
    derived = derive_correspondences(inputs, decided)

    # exact_counts for the reviewer's denominator context (audit §6).
    exact_counts: dict[str, tuple[int, int]] = {}
    for table, (prefix, oracle_by_key, cols) in inputs.items():
        bares = {c.bare.upper() for c in cols}
        exact = sum(1 for k in oracle_by_key if k.upper() in bares)
        exact_counts[table] = (exact, len(oracle_by_key))

    col_by_table_bare = {(t, c.bare): c for t, (_, _, cols) in inputs.items() for c in cols}
    of_by_table_key = {(t, k): of for t, (_, obk, _) in inputs.items()
                       for k, of in obk.items()}

    rows: list[ReviewRow] = []
    for fc in derived:
        of = of_by_table_key.get((fc.applaud_table, fc.oracle_key))
        col = col_by_table_bare.get((fc.applaud_table, fc.applaud_bare))
        rows.append(ReviewRow(
            applaud_table=fc.applaud_table, oracle_key=fc.oracle_key,
            oracle_type=" ".join(str(x) for x in expected_shape(of) if x not in (None, "")) if of else "",
            candidate_bare=fc.applaud_bare, applaud_ddid=fc.applaud_ddid,
            applaud_type=" ".join(str(x) for x in actual_shape(col) if x not in (None, "")) if col else "",
            confidence=fc.confidence, score=fc.score, signals=fc.signals, alternatives=""))

    out = args.output or Path(f"Applaud_FieldMap_Review_{release}_{args.system}.xlsx")
    write_review_workbook(rows, out, exact_counts=exact_counts)
    print(f"Derived {len(rows)} candidate correspondence(s) across "
          f"{len(inputs)} table(s). Review workbook: {out}")


def _run_correspondence_confirm(args: argparse.Namespace) -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(name)s: %(message)s")
    from fbdi.applaud_snapshot import ApplaudSnapshot
    from fbdi.config import applaud_snapshot_path
    from fbdi.correspondence import (
        load_review_workbook, apply_review_decisions, InvalidCorrectedBareError,
        load_fieldmap_workbook, merge_fieldmap, write_fieldmap_workbook,
    )
    if not args.review.is_file():
        print(f"Error: review workbook not found: {args.review}"); sys.exit(1)
    snap_path = applaud_snapshot_path(args.system)
    if not snap_path.exists():
        print(f"Error: snapshot not found: {snap_path}."); sys.exit(1)
    snapshot = ApplaudSnapshot.load(snap_path)
    valid_bares = {name: {c.bare for c in t.columns}
                   for name, t in snapshot.tables.items()}

    review_rows = load_review_workbook(args.review)
    try:
        decisions = apply_review_decisions(review_rows, valid_bares)
    except InvalidCorrectedBareError as exc:
        print(f"Error: {exc}"); sys.exit(1)

    committed = load_fieldmap_workbook(args.fieldmap) if args.fieldmap.exists() else {}
    merged = merge_fieldmap(decisions, committed)
    flat = [fc for rows in merged.values() for fc in rows]
    write_fieldmap_workbook(flat, args.fieldmap)
    n_conf = sum(1 for fc in flat if fc.origin == "confirmed")
    n_rej = sum(1 for fc in flat if fc.origin == "rejected")
    print(f"Merged {len(decisions)} decision(s). Field map now: {n_conf} confirmed, "
          f"{n_rej} rejected -> {args.fieldmap}")
```

- [ ] **Step 8: Run the full suite + a CLI help smoke check**

Run: `py -m pytest tests/ -q`
Expected: PASS (all tests; ~415 collected).
Run: `py -m fbdi correspondence-derive --help`
Expected: prints usage with `--release`, `--map`, `--tables`, `--output` (exit 0).

- [ ] **Step 9: Commit**

```bash
git add fbdi/cli.py fbdi/correspondence.py tests/test_correspondence.py
git commit -m "feat(cli): correspondence-derive / -confirm + audit-applaud --fieldmap/--accept-confidence"
```

---

## Task 10: `.gitignore` entries + final verification

**Files:**
- Modify: `.gitignore`

- [ ] **Step 1: Add the review-workbook ignore (keep the field map tracked)**

In `.gitignore`, after line 20 (`Applaud_Compliance_Report_*.xlsx`), add:

```
# Disposable field-correspondence review workbook (the committed map
# FBDI_to_Applaud_FieldMap.xlsx IS tracked, like FBDI_to_Applaud_AppMap.xlsx)
Applaud_FieldMap_Review_*.xlsx
```

- [ ] **Step 2: Verify the field map would be tracked and the review workbook ignored**

Run: `git check-ignore -v Applaud_FieldMap_Review_26B_ORACLE_MASTER.xlsx; git check-ignore -v FBDI_to_Applaud_FieldMap.xlsx`
Expected: the first prints a matching `.gitignore` rule; the second prints nothing and exits 1
(i.e., it is NOT ignored — it will be tracked when created).

- [ ] **Step 3: Run the full suite one final time**

Run: `py -m pytest tests/ -q`
Expected: PASS (all tests).

- [ ] **Step 4: Commit**

```bash
git add .gitignore
git commit -m "chore: ignore disposable Applaud_FieldMap_Review_*.xlsx; keep field map tracked"
```

---

## Operational follow-up (post-merge, NOT part of the TDD plan — requires live MCP)

These steps run against the live ORACLE_MASTER snapshot and produce the committed map. They are
listed for the operator, not the implementing agent:

1. Ensure `baselines/applaud/applaud_snapshot.json` is current (Step A extraction).
2. `py -m fbdi correspondence-derive --release 26B --tables T_AP_INVOICE_INT,T_AP_INVOICE_LINES,T_BANKS_BRANCHES,T_BPA_PO_LINES_INTERFACE,T_EGP_COMPONENTS_INTERFACE,T_EGP_ITEM_CATEGORIES_INT,T_EGO_ITEM_INTF_EFF_B,T_MSC_ST_ASSIGNMENT_SETS,T_POZ_SUPPLIERS_INT,T_POZ_SUPPLIER_SITES_INT`
3. Brad reviews `Applaud_FieldMap_Review_26B_ORACLE_MASTER.xlsx` — extends `ABBREVIATIONS` from the
   WEAK tier (the missing-abbreviation worklist, audit §3), settles the `LINE_ATTRIBUTE16-20`
   canary (audit §8.3), marks `Confirm?`/`Corrected Bare`.
4. `py -m fbdi correspondence-confirm --review Applaud_FieldMap_Review_26B_ORACLE_MASTER.xlsx`
5. `git add FBDI_to_Applaud_FieldMap.xlsx && git commit` the curated map.
6. Re-run `py -m fbdi audit-applaud --release 26B --old-release 26A` and confirm HIGH
   "missing field" counts collapse from 957 toward the genuine residual (spec §11), comparing
   against the first-run notes distribution.

---

## Self-Review

**Spec coverage (spec §§1-11 as amended by the audit):**
- §3 three-phase command model → Task 9 (`correspondence-derive` / `-confirm`; `audit-applaud --fieldmap`). ✓
- §4 module + reuse of `expected_shape`/`actual_shape`/`_lcs_match`; no top-level import cycle (build_alias imported inside `run_audit`) → Tasks 2-6, 7. ✓
- §5 derivation ladder: exact pre-pass, full squash (§2.3), truncation window `30-len(prefix)` (§2.1), digit-run (§1.1), abbreviation seed (§8.3), char-vs-numeric veto + `U→char` (§1.2), weak position (§2.4), tiers, bijection (§8.5) → Tasks 1-3. ✓
- §6 committed map + review workbook + merge precedence + Corrected-Bare fail-loud (§4.1) → Tasks 4-5. ✓
- §7 audit integration: alias bare only, DDID untouched, default gate confirmed, sizing side-benefit, rejected provenance (§4.2) → Task 7. ✓
- §1.3/§1.4 @-field + non-prefix exclusion → Task 3 (`_candidate_excluded`) + Task 8 (snapshot lock-in). ✓
- §9 build sequence test list (digit-run named case, U column, coincidental short prefix, bijection, tier bands, roundtrip, merge precedence, re-derive idempotence, Y/N/Corrected, load-bearing regression) → Tasks 2-7. ✓
- §10 fixture realism (the six edge cases are distributed across Task 3 + Task 8) and denominator header (Task 5 `exact_counts` / Task 9). ✓
- §11 verification → Operational follow-up. ✓

**Placeholder scan:** No TBD/TODO; every code step shows complete code; every test step shows full test bodies and exact run commands with expected output. ✓

**Type consistency:** `FieldCorrespondence`, `ReviewRow`, `InvalidCorrectedBareError`, and every function name match across Tasks 2-9 (see "Naming locked across tasks"). `run_audit`'s new params (`fieldmap`, `accept_confidence`) are consumed identically in Task 7 (definition) and Task 9 (CLI call). `build_alias` signature is identical in Task 6 (def) and Task 7 (use). ✓

**One deviation from the spec, deliberately recorded:** the spec/audit §1.4 assumed the snapshot
mis-strips `X_PHANTOM` to `HANTOM` and "likely already pollutes the pilot findings." Live
verification (PR #3 workbook scan + reading `applaud_snapshot.py:126,163-179`) shows the snapshot
already excludes it correctly — zero `HANTOM` findings. Task 8 therefore locks in the existing
behavior rather than fixing a non-existent bug, and `derive_table_correspondences` adds a
defensive `_candidate_excluded` guard so the candidate pool is clean even if fed raw columns.
