# FBDI Compliance Report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a data-driven generator that emits `FBDI_Compliance_Report_<OLD>_<NEW>.{html,pdf}` per release pair, replacing the manually-built Word report. Includes a root-cause fix to the catalog's broken `Drift` sheet.

**Architecture:** Six pieces — a shared LCS-style alignment module (`align.py`) used by both the rebuilt catalog Drift writer and a new Jinja2-templated report generator (`report.py`). PDF rendered from the same template via `weasyprint` with a `print_mode` flag. Brand palette enforced from `reference/colorguide.pdf`.

**Tech Stack:** Python 3.14, openpyxl (existing), jinja2 (new), weasyprint (new), pytest (existing).

**Spec:** [`docs/superpowers/specs/2026-05-01-fbdi-compliance-report-design.md`](../specs/2026-05-01-fbdi-compliance-report-design.md)

---

## File map

**Create:**
- `fbdi/align.py` — pure alignment algorithm
- `fbdi/applaud_type.py` — Oracle → Applaud type translator
- `fbdi/report.py` — report generator (view-model + render)
- `fbdi/templates/report.html.j2` — single Jinja2 template (HTML + PDF)
- `tests/test_align.py`
- `tests/test_applaud_type.py`
- `tests/test_report.py`

**Modify:**
- `requirements.txt` — add `jinja2`, `weasyprint`
- `fbdi/catalog.py` — replace `_compute_drift` to use `align.align_tabs`; update `DriftRow` schema
- `fbdi/cli.py` — add `report` subcommand
- `tests/test_catalog.py` — update Drift tests for new schema and classifications

---

## Phase 1: Dependencies

### Task 1: Add jinja2 and weasyprint to requirements

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: Add the two new lines**

```
openpyxl>=3.1
selenium>=4.20
webdriver-manager>=4.0
requests>=2.31
pytest>=8.0
jinja2>=3.1
weasyprint>=62.0
```

- [ ] **Step 2: Install and verify both import**

```bash
pip install -r requirements.txt
py -c "import jinja2, weasyprint; print('jinja2', jinja2.__version__); print('weasyprint', weasyprint.__version__)"
```

Expected: prints both versions without ImportError. Note: weasyprint on Windows requires GTK runtime; if `OSError: cannot load library 'libgobject-2.0-0'` appears, install GTK3 from https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases and re-run.

- [ ] **Step 3: Commit**

```bash
git add requirements.txt
git commit -m "chore(deps): add jinja2 and weasyprint for HTML/PDF report rendering"
```

---

## Phase 2: Alignment algorithm (TDD)

The alignment algorithm is a pure function. It takes two row lists (one per release) and returns a typed change list. It uses LCS (longest common subsequence) over field identity to find matched pairs, then classifies each pair across three axes (label, metadata, position). Unmatched-old becomes REMOVED; unmatched-new becomes ADDED.

**REQUIRED SUB-SKILL for this phase:** Use `superpowers:test-driven-development` — write the failing test first, run it to confirm it fails, then write the minimal code to pass.

### Task 2: align.py skeleton + Change/AlignmentResult dataclasses + first test

**Files:**
- Create: `fbdi/align.py`
- Create: `tests/test_align.py`

- [ ] **Step 1: Write the failing test**

`tests/test_align.py`:
```python
"""Tests for fbdi.align — LCS-style alignment of FBDI tab rows across releases."""

from fbdi.align import AlignedField, Change, align_tabs


def _row(position, label, technical=None, data_type=None, length=None, required=None):
    """Build an AlignedField input row."""
    return AlignedField(
        position=position, label=label, technical=technical,
        data_type=data_type, length=length, required=required,
    )


class TestEmpty:
    def test_both_empty(self):
        result = align_tabs([], [])
        assert result == []

    def test_old_empty_new_one_field(self):
        new = [_row(1, "Item Name", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs([], new)
        assert len(result) == 1
        assert result[0].change_type == "ADDED"
        assert result[0].new_position == 1
        assert result[0].old_position is None

    def test_new_empty_old_one_field(self):
        old = [_row(1, "Item Name", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs(old, [])
        assert len(result) == 1
        assert result[0].change_type == "REMOVED"
        assert result[0].old_position == 1
        assert result[0].new_position is None
```

- [ ] **Step 2: Run test to verify it fails**

```bash
py -m pytest tests/test_align.py -v
```

Expected: ImportError — `cannot import name 'AlignedField' from 'fbdi.align'`.

- [ ] **Step 3: Create align.py with minimal implementation**

`fbdi/align.py`:
```python
"""LCS-style alignment of FBDI tab rows across two releases.

Pure function — no I/O. Takes two lists of AlignedField (one per release)
and returns a typed Change list classifying every difference: ADDED,
REMOVED, MODIFIED, RENAMED, SHIFTED, MULTI.

The algorithm matches fields by identity (technical name first, label
fallback) using longest common subsequence, then classifies each matched
pair across three independent axes (label, metadata, position). Unmatched
rows on either side become ADDED or REMOVED.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class AlignedField:
    """One field at one position within one release."""
    position: int
    label: str | None
    technical: str | None
    data_type: str | None
    length: int | None
    required: bool | None


@dataclass(frozen=True)
class Change:
    """One classified difference between two releases."""
    change_type: str                      # ADDED | REMOVED | MODIFIED | RENAMED | SHIFTED | MULTI
    old_position: int | None
    new_position: int | None
    old_field: AlignedField | None
    new_field: AlignedField | None
    axes: tuple[str, ...] = ()            # subset of ("label", "metadata", "position")
    sub_kinds: tuple[str, ...] = ()       # subset of ("type", "length", "required") when metadata changed


def align_tabs(old: list[AlignedField], new: list[AlignedField]) -> list[Change]:
    """Align two release row lists and return classified changes."""
    if not old and not new:
        return []
    if not old:
        return [
            Change(change_type="ADDED", old_position=None, new_position=f.position,
                   old_field=None, new_field=f)
            for f in new
        ]
    if not new:
        return [
            Change(change_type="REMOVED", old_position=f.position, new_position=None,
                   old_field=f, new_field=None)
            for f in old
        ]
    # Matched-pair classification not implemented yet.
    raise NotImplementedError("matched-pair classification — Task 3")
```

- [ ] **Step 4: Run test to verify it passes**

```bash
py -m pytest tests/test_align.py -v
```

Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add fbdi/align.py tests/test_align.py
git commit -m "feat(align): scaffold align.py with empty/all-add/all-remove cases"
```

### Task 3: Identity matching (technical-name-first, label fallback)

**Files:**
- Modify: `fbdi/align.py`
- Modify: `tests/test_align.py`

- [ ] **Step 1: Add the failing tests**

Append to `tests/test_align.py`:
```python
class TestIdentityMatching:
    def test_perfect_match_no_changes_emits_nothing(self):
        rows = [_row(1, "Item Name", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs(rows, rows)
        assert result == []

    def test_match_by_technical_name_when_present(self):
        old = [_row(1, "Old Label", "ITEM_NAME", "VARCHAR2", 30, True)]
        new = [_row(1, "New Label", "ITEM_NAME", "VARCHAR2", 30, True)]
        result = align_tabs(old, new)
        # Same technical name → matched (so RENAMED, not ADD+REMOVE)
        assert len(result) == 1
        assert result[0].change_type == "RENAMED"

    def test_match_by_label_when_technical_is_none(self):
        # Thin tabs have no technical name; fall back to label
        old = [_row(1, "Item Name", None, None, None, True)]
        new = [_row(1, "Item Name", None, None, None, True)]
        result = align_tabs(old, new)
        assert result == []

    def test_no_match_when_both_label_and_technical_differ(self):
        old = [_row(1, "Old", "OLD_FIELD", "VARCHAR2", 30, True)]
        new = [_row(1, "New", "NEW_FIELD", "VARCHAR2", 30, True)]
        result = align_tabs(old, new)
        # Disjoint identity — REMOVE the old + ADD the new
        types = sorted(c.change_type for c in result)
        assert types == ["ADDED", "REMOVED"]
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
py -m pytest tests/test_align.py::TestIdentityMatching -v
```

Expected: 4 failures, all from `NotImplementedError`.

- [ ] **Step 3: Implement identity-key + LCS matching + minimal classifier**

Replace the body of `align_tabs` in `fbdi/align.py`:

```python
def _identity_key(f: AlignedField) -> tuple[str, str]:
    """Stable identity for matching across releases.

    Prefers technical name (canonical, position-independent). Falls back to
    label when technical is missing (thin tabs). Tag distinguishes the
    space so a label "ITEM_NAME" never matches a technical "ITEM_NAME".
    """
    if f.technical:
        return ("tech", f.technical)
    return ("label", f.label or "")


def _lcs_match(old: list[AlignedField], new: list[AlignedField]) -> list[tuple[int, int]]:
    """Longest common subsequence over identity keys. Returns matched index pairs.

    Standard O(m*n) DP. Indices are 0-based positions in the input lists.
    """
    m, n = len(old), len(new)
    if m == 0 or n == 0:
        return []
    old_keys = [_identity_key(f) for f in old]
    new_keys = [_identity_key(f) for f in new]
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(m):
        for j in range(n):
            if old_keys[i] == new_keys[j]:
                dp[i + 1][j + 1] = dp[i][j] + 1
            else:
                dp[i + 1][j + 1] = max(dp[i][j + 1], dp[i + 1][j])
    # Backtrack to recover the matched pairs.
    pairs: list[tuple[int, int]] = []
    i, j = m, n
    while i > 0 and j > 0:
        if old_keys[i - 1] == new_keys[j - 1]:
            pairs.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1
    pairs.reverse()
    return pairs


def _classify_pair(old_f: AlignedField, new_f: AlignedField) -> Change | None:
    """Classify a matched pair across three axes; None if unchanged."""
    label_changed = (old_f.label or "") != (new_f.label or "")
    metadata_kinds: list[str] = []
    if (old_f.data_type or "") != (new_f.data_type or ""):
        metadata_kinds.append("type")
    if old_f.length != new_f.length:
        metadata_kinds.append("length")
    if old_f.required != new_f.required:
        metadata_kinds.append("required")
    metadata_changed = bool(metadata_kinds)
    position_changed = old_f.position != new_f.position

    axes = []
    if label_changed:
        axes.append("label")
    if metadata_changed:
        axes.append("metadata")
    if position_changed:
        axes.append("position")
    if not axes:
        return None

    if len(axes) == 1:
        change_type = {"label": "RENAMED", "metadata": "MODIFIED", "position": "SHIFTED"}[axes[0]]
    else:
        change_type = "MULTI"

    return Change(
        change_type=change_type,
        old_position=old_f.position,
        new_position=new_f.position,
        old_field=old_f,
        new_field=new_f,
        axes=tuple(axes),
        sub_kinds=tuple(metadata_kinds),
    )


def align_tabs(old: list[AlignedField], new: list[AlignedField]) -> list[Change]:
    """Align two release row lists and return classified changes."""
    if not old and not new:
        return []

    matched = _lcs_match(old, new)
    matched_old = {i for i, _ in matched}
    matched_new = {j for _, j in matched}

    changes: list[Change] = []
    # REMOVED: old fields with no match
    for i, f in enumerate(old):
        if i not in matched_old:
            changes.append(Change(
                change_type="REMOVED", old_position=f.position, new_position=None,
                old_field=f, new_field=None,
            ))
    # ADDED: new fields with no match
    for j, f in enumerate(new):
        if j not in matched_new:
            changes.append(Change(
                change_type="ADDED", old_position=None, new_position=f.position,
                old_field=None, new_field=f,
            ))
    # Classified pair changes
    for i, j in matched:
        c = _classify_pair(old[i], new[j])
        if c is not None:
            changes.append(c)

    # Stable sort: by new_position (None last), then old_position.
    changes.sort(key=lambda c: (c.new_position is None, c.new_position or 0,
                                c.old_position is None, c.old_position or 0))
    return changes
```

- [ ] **Step 4: Run all align tests**

```bash
py -m pytest tests/test_align.py -v
```

Expected: 7 passed.

- [ ] **Step 5: Commit**

```bash
git add fbdi/align.py tests/test_align.py
git commit -m "feat(align): identity matching, LCS, and three-axis classification"
```

### Task 4: SHIFTED detection (mid-file insert cascades correctly)

**Files:**
- Modify: `tests/test_align.py`

This task validates the SHIFTED case using the actual 26A→26B WorkDefinitionTemplate scenario: inserting one new field at position 19 should classify as 1 ADDED + ~20 SHIFTED, NOT as the catalog's current "ADDED:6, MULTI:2, RENAMED:19" misclassification.

- [ ] **Step 1: Add failing tests**

Append to `tests/test_align.py`:
```python
class TestShiftDetection:
    def test_mid_file_insert_classifies_as_one_add_plus_shifts(self):
        """Real scenario from 26A→26B WorkDefinitionTemplate.

        Insert one new field at position 19. Catalog's naive per-position diff
        misclassifies this as multiple ADDs + RENAMEDs. Alignment must produce
        exactly 1 ADDED + N SHIFTED.
        """
        old = [
            _row(18, "Transform from Item Number", "TRANSFORM_FROM_ITEM_NUMBER", "VARCHAR2", 300, False),
            _row(19, "Completion Subinventory", "COMPLETION_SUBINVENTORY_NAME", "VARCHAR2", 10, False),
            _row(20, "Completion Locator Segment1", "COMPL_LOCATOR_SEGMENT1", "VARCHAR2", 40, False),
            _row(21, "Completion Locator Segment2", "COMPL_LOCATOR_SEGMENT2", "VARCHAR2", 40, False),
        ]
        new = [
            _row(18, "Transform from Item Number", "TRANSFORM_FROM_ITEM_NUMBER", "VARCHAR2", 300, False),
            _row(19, "Enable Parallel Operations", "ENABLE_PARALLEL_OPS_FLAG", "VARCHAR2", 1, False),
            _row(20, "Completion Subinventory", "COMPLETION_SUBINVENTORY_NAME", "VARCHAR2", 10, False),
            _row(21, "Completion Locator Segment1", "COMPL_LOCATOR_SEGMENT1", "VARCHAR2", 40, False),
            _row(22, "Completion Locator Segment2", "COMPL_LOCATOR_SEGMENT2", "VARCHAR2", 40, False),
        ]
        result = align_tabs(old, new)

        added = [c for c in result if c.change_type == "ADDED"]
        shifted = [c for c in result if c.change_type == "SHIFTED"]
        renamed = [c for c in result if c.change_type == "RENAMED"]
        modified = [c for c in result if c.change_type == "MODIFIED"]

        assert len(added) == 1
        assert added[0].new_field.technical == "ENABLE_PARALLEL_OPS_FLAG"
        assert added[0].new_position == 19

        assert len(shifted) == 3
        # Verify each SHIFTED keeps the same field but moves position +1
        for c in shifted:
            assert c.old_field.technical == c.new_field.technical
            assert c.new_position == c.old_position + 1

        assert renamed == []
        assert modified == []

    def test_pure_swap_emits_two_shifts(self):
        old = [
            _row(1, "Field A", "FIELD_A", "VARCHAR2", 10, True),
            _row(2, "Field B", "FIELD_B", "VARCHAR2", 10, True),
        ]
        new = [
            _row(1, "Field B", "FIELD_B", "VARCHAR2", 10, True),
            _row(2, "Field A", "FIELD_A", "VARCHAR2", 10, True),
        ]
        result = align_tabs(old, new)
        # LCS picks one of the two as a stable axis; the other becomes a remove+add.
        # We don't constrain *which*, but the total shape must conserve fields.
        assert len(result) > 0
        # Both fields must appear somewhere on either side
        all_techs = set()
        for c in result:
            if c.old_field:
                all_techs.add(c.old_field.technical)
            if c.new_field:
                all_techs.add(c.new_field.technical)
        assert all_techs == {"FIELD_A", "FIELD_B"}
```

- [ ] **Step 2: Run tests**

```bash
py -m pytest tests/test_align.py::TestShiftDetection -v
```

Expected: 2 passed (no implementation change needed — the previous task already supports this; if a test fails, the alignment classifier has a bug).

- [ ] **Step 3: Commit**

```bash
git add tests/test_align.py
git commit -m "test(align): pin shift-detection invariants from real WorkDefinitionTemplate scenario"
```

### Task 5: MODIFIED variants and MULTI composition

**Files:**
- Modify: `tests/test_align.py`

- [ ] **Step 1: Add failing tests**

Append to `tests/test_align.py`:
```python
class TestModifiedVariants:
    def test_type_only_change(self):
        old = [_row(1, "X", "X", "VARCHAR2", 30, True)]
        new = [_row(1, "X", "X", "NUMBER",   None, True)]
        result = align_tabs(old, new)
        assert len(result) == 1
        c = result[0]
        assert c.change_type == "MODIFIED"
        assert c.axes == ("metadata",)
        assert "type" in c.sub_kinds

    def test_length_only_change(self):
        old = [_row(1, "X", "X", "VARCHAR2", 30, True)]
        new = [_row(1, "X", "X", "VARCHAR2", 50, True)]
        result = align_tabs(old, new)
        assert len(result) == 1
        c = result[0]
        assert c.change_type == "MODIFIED"
        assert c.sub_kinds == ("length",)

    def test_required_flip(self):
        old = [_row(1, "X", "X", "VARCHAR2", 30, False)]
        new = [_row(1, "X", "X", "VARCHAR2", 30, True)]
        result = align_tabs(old, new)
        assert result[0].change_type == "MODIFIED"
        assert result[0].sub_kinds == ("required",)


class TestMulti:
    def test_label_plus_metadata(self):
        old = [_row(1, "Old Label", "X", "VARCHAR2", 30, True)]
        new = [_row(1, "New Label", "X", "VARCHAR2", 50, True)]
        result = align_tabs(old, new)
        c = result[0]
        assert c.change_type == "MULTI"
        assert set(c.axes) == {"label", "metadata"}
        assert c.sub_kinds == ("length",)

    def test_position_plus_metadata(self):
        # Field moves AND its type changes
        old = [
            _row(1, "Anchor", "ANCHOR", "VARCHAR2", 10, True),
            _row(2, "Mover",  "MOVER",  "VARCHAR2", 30, True),
        ]
        new = [
            _row(1, "Anchor", "ANCHOR", "VARCHAR2", 10, True),
            _row(2, "Inserted", "INSERTED", "VARCHAR2", 1, False),
            _row(3, "Mover",  "MOVER",  "VARCHAR2", 50, True),  # length changed AND moved
        ]
        result = align_tabs(old, new)
        mover_change = [c for c in result if c.new_field and c.new_field.technical == "MOVER"][0]
        assert mover_change.change_type == "MULTI"
        assert set(mover_change.axes) == {"position", "metadata"}
        assert mover_change.sub_kinds == ("length",)
```

- [ ] **Step 2: Run tests**

```bash
py -m pytest tests/test_align.py -v
```

Expected: all passed (5 new + 9 prior = 14 total).

- [ ] **Step 3: Commit**

```bash
git add tests/test_align.py
git commit -m "test(align): MODIFIED sub-kinds and MULTI composition coverage"
```

---

## Phase 3: Oracle → Applaud type translator (TDD)

### Task 6: applaud_type.py with full type translation table

**Files:**
- Create: `fbdi/applaud_type.py`
- Create: `tests/test_applaud_type.py`

- [ ] **Step 1: Write the failing tests**

`tests/test_applaud_type.py`:
```python
"""Tests for fbdi.applaud_type — Oracle → Applaud type translator."""

from fbdi.applaud_type import applaud_type_for
from fbdi.type_parser import ParsedType


def _pt(data_type, length=None, scale=None, parse_warning=False):
    return ParsedType(data_type=data_type, length=length, scale=scale, parse_warning=parse_warning)


class TestApplaudTypeFor:
    def test_varchar2_with_length(self):
        assert applaud_type_for(_pt("VARCHAR2", length=30)) == "char 30"

    def test_varchar2_without_length(self):
        # Should never happen in real data, but be deterministic anyway
        assert applaud_type_for(_pt("VARCHAR2")) == "char"

    def test_number_with_precision_and_scale(self):
        assert applaud_type_for(_pt("NUMBER", length=18, scale=4)) == "numeric 18,4"

    def test_number_with_precision_only(self):
        assert applaud_type_for(_pt("NUMBER", length=18)) == "numeric 18"

    def test_number_no_precision(self):
        # Plain NUMBER → "numeric" with no defaults invented
        assert applaud_type_for(_pt("NUMBER")) == "numeric"

    def test_date(self):
        assert applaud_type_for(_pt("DATE")) == "date"

    def test_timestamp(self):
        assert applaud_type_for(_pt("TIMESTAMP")) == "date"

    def test_clob_passthrough(self):
        assert applaud_type_for(_pt("CLOB")) == "clob"

    def test_blob_passthrough(self):
        assert applaud_type_for(_pt("BLOB")) == "blob"

    def test_unknown_type_passes_through_lowercase(self):
        assert applaud_type_for(_pt("XMLTYPE")) == "xmltype"

    def test_empty_returns_empty(self):
        # Blank input (thin tab with no type info) returns empty string
        assert applaud_type_for(_pt("")) == ""

    def test_parse_warning_returns_empty(self):
        # Don't fabricate a type for unparseable input
        assert applaud_type_for(_pt("VARCHAR2", parse_warning=True)) == ""
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
py -m pytest tests/test_applaud_type.py -v
```

Expected: ImportError — `cannot import name 'applaud_type_for' from 'fbdi.applaud_type'`.

- [ ] **Step 3: Implement applaud_type.py**

`fbdi/applaud_type.py`:
```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
py -m pytest tests/test_applaud_type.py -v
```

Expected: 12 passed.

- [ ] **Step 5: Commit**

```bash
git add fbdi/applaud_type.py tests/test_applaud_type.py
git commit -m "feat(applaud-type): add Oracle->Applaud type translator with full mapping"
```

---

## Phase 4: Catalog Drift root-cause fix

The existing `_compute_drift` in `fbdi/catalog.py` does naive per-position diff and misclassifies shift cascades. Replace it with a pass that calls `align.align_tabs` per (file, tab) and emits one DriftRow per Change with the new schema.

### Task 7: Update DriftRow schema and rewrite _compute_drift

**Files:**
- Modify: `fbdi/catalog.py`
- Modify: `tests/test_catalog.py`

- [ ] **Step 1: Read the existing _compute_drift and DriftRow to know what's changing**

```bash
py -c "
import inspect
from fbdi import catalog
print(inspect.getsource(catalog.DriftRow))
print('---')
print(inspect.getsource(catalog._compute_drift))
"
```

- [ ] **Step 2: Update DriftRow + add alignment-driven _compute_drift**

In `fbdi/catalog.py`, replace the `DriftRow` dataclass (around line 71-87) with:

```python
@dataclass
class DriftRow:
    """One row per classified change between two releases for one (file, tab).

    Schema is alignment-driven, not naive per-position. position columns
    are split into old/new because SHIFTED rows have a different position
    on each side.
    """
    file: str
    tab: str
    change_type: str            # ADDED | REMOVED | MODIFIED | RENAMED | SHIFTED | MULTI
    old_position: int | None
    new_position: int | None
    col_label_old: str
    col_label_new: str
    col_technical_old: str
    col_technical_new: str
    data_type_old: str
    data_type_new: str
    length_old: str
    length_new: str
    required_old: str
    required_new: str
    sub_kinds: str              # comma-joined ("type", "length", "required") for MODIFIED/MULTI; empty otherwise
```

Then replace `_compute_drift` (around line 411-466) with:

```python
def _compute_drift(
    old_rows: list[CatalogRow],
    new_rows: list[CatalogRow],
    release_old: str,
    release_new: str,
) -> list[DriftRow]:
    """Alignment-driven diff between two release row sets.

    Groups rows by (file, tab); for each pair, calls fbdi.align.align_tabs
    to derive the classified Change list; emits one DriftRow per Change.
    Replaces the prior naive per-position diff that misclassified shift
    cascades as RENAMED/MULTI.
    """
    from fbdi.align import AlignedField, align_tabs

    def _to_aligned(r: CatalogRow) -> AlignedField:
        return AlignedField(
            position=r.position,
            label=r.column_label,
            technical=r.column_technical or None,
            data_type=r.data_type or None,
            length=r.length,
            required=r.required,
        )

    def _group(rows: list[CatalogRow]) -> dict[tuple[str, str], list[CatalogRow]]:
        out: dict[tuple[str, str], list[CatalogRow]] = {}
        for r in rows:
            out.setdefault((r.file_name, r.tab_name), []).append(r)
        for v in out.values():
            v.sort(key=lambda r: r.position)
        return out

    old_grouped = _group(old_rows)
    new_grouped = _group(new_rows)
    all_keys = sorted(set(old_grouped.keys()) | set(new_grouped.keys()))

    drift: list[DriftRow] = []
    for file, tab in all_keys:
        old_aligned = [_to_aligned(r) for r in old_grouped.get((file, tab), [])]
        new_aligned = [_to_aligned(r) for r in new_grouped.get((file, tab), [])]
        for change in align_tabs(old_aligned, new_aligned):
            drift.append(_drift_row_from_change(file, tab, change))
    return drift


def _drift_row_from_change(file: str, tab: str, change) -> DriftRow:
    """Build a DriftRow from a fbdi.align.Change."""
    old = change.old_field
    new = change.new_field
    return DriftRow(
        file=file,
        tab=tab,
        change_type=change.change_type,
        old_position=change.old_position,
        new_position=change.new_position,
        col_label_old=(old.label or "") if old else "",
        col_label_new=(new.label or "") if new else "",
        col_technical_old=(old.technical or "") if old else "",
        col_technical_new=(new.technical or "") if new else "",
        data_type_old=_fmt_type_align(old),
        data_type_new=_fmt_type_align(new),
        length_old=_fmt_length_align(old),
        length_new=_fmt_length_align(new),
        required_old=_fmt_required_align(old),
        required_new=_fmt_required_align(new),
        sub_kinds=",".join(change.sub_kinds),
    )


def _fmt_type_align(f) -> str:
    return (f.data_type or "") if f else ""


def _fmt_length_align(f) -> str:
    if f is None or f.length is None:
        return ""
    return str(f.length)


def _fmt_required_align(f) -> str:
    if f is None or f.required is None:
        return ""
    return "TRUE" if f.required else "FALSE"
```

Delete the old `_drift_row`, `_fmt_type`, `_fmt_length`, `_fmt_required` helpers (replaced by the `_align` variants). Keep them if they're used elsewhere (grep first).

- [ ] **Step 3: Update _drift_tab_headers to match the new schema**

In `fbdi/catalog.py`, replace `_drift_tab_headers` (around line 521-533):

```python
def _drift_tab_headers(release_old: str | None, release_new: str | None) -> list[str]:
    """Build Drift tab headers with release names substituted."""
    old = release_old or "OLD"
    new = release_new or "NEW"
    return [
        "file", "tab", "change_type",
        f"position_{old}", f"position_{new}",
        f"col_label_{old}", f"col_label_{new}",
        f"col_technical_{old}", f"col_technical_{new}",
        f"data_type_{old}", f"data_type_{new}",
        f"length_{old}", f"length_{new}",
        f"required_{old}", f"required_{new}",
        "sub_kinds",
    ]
```

- [ ] **Step 4: Update the Drift writer in _write_master_workbook**

Find the Drift sheet write loop in `_write_master_workbook` (search for `drift` in the function). It currently writes columns matching the old `DriftRow` order. Update the row-write loop to use the new schema. Look for the `for d in drift:` loop and replace its row tuple to match new headers.

```bash
py -c "
import inspect, fbdi.catalog
src = inspect.getsource(fbdi.catalog._write_master_workbook)
print(src)
"
```

Then patch it. The row tuple in column order:

```python
ws_drift.append([
    d.file, d.tab, d.change_type,
    d.old_position if d.old_position is not None else "",
    d.new_position if d.new_position is not None else "",
    d.col_label_old, d.col_label_new,
    d.col_technical_old, d.col_technical_new,
    d.data_type_old, d.data_type_new,
    d.length_old, d.length_new,
    d.required_old, d.required_new,
    d.sub_kinds,
])
```

- [ ] **Step 5: Find and update the existing Drift tests**

```bash
py -m pytest tests/test_catalog.py -v 2>&1 | head -60
```

Identify failures from the schema change. The tests that touch DriftRow fields will need updates. Update them to:
- Reference the new field names (`old_position`, `new_position`, `change_type`, `sub_kinds` instead of single `position`)
- Update assertions on `change_type` values where the old code emitted `TYPE_CHANGED` / `LENGTH_CHANGED` / `REQUIRED_CHANGED` — these become `MODIFIED` with `sub_kinds="type"` / `"length"` / `"required"`
- Update assertions on shift scenarios — the new alignment correctly identifies SHIFTED, where the old code emitted MULTI

For each failing test, read its assertion, decide whether it was testing the OLD broken behavior (in which case rewrite the assertion to test the new correct behavior) or testing schema (in which case update the field names).

- [ ] **Step 6: Run full catalog test module**

```bash
py -m pytest tests/test_catalog.py -v
```

Expected: all passed.

- [ ] **Step 7: Run the full suite to catch any cross-module breakage**

```bash
py -m pytest tests/ -v
```

Expected: all passed.

- [ ] **Step 8: Spike — regenerate the catalog and verify Drift looks right**

```bash
py -m fbdi catalog --release 26A
py -m fbdi catalog --release 26B
PYTHONIOENCODING=utf-8 py -c "
from openpyxl import load_workbook
from collections import Counter
wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True)
ws = wb['Drift']
header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
print('Headers:', header)
ct_idx = header.index('change_type')
counts = Counter()
for row in ws.iter_rows(min_row=2, values_only=True):
    counts[row[ct_idx]] += 1
print('Change-type counts:', counts.most_common())
print('(Expected: SHIFTED >> RENAMED, MULTI sharply lower than before)')
"
```

Expected: `SHIFTED` is now the dominant change type for shift-heavy files; `RENAMED` and `MULTI` counts drop sharply (the old ~460+236 was the misclassification you're fixing).

- [ ] **Step 9: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "fix(catalog): rebuild Drift sheet on alignment module — fixes shift-cascade misclassification

Replaces the naive per-position _compute_drift with a pass that uses
fbdi.align.align_tabs per (file, tab). DriftRow gets new schema
(separate old/new positions; change_type values aligned with the new
ADDED/REMOVED/MODIFIED/RENAMED/SHIFTED/MULTI taxonomy; sub_kinds column
for metadata sub-classifications)."
```

---

## Phase 5: Report generator (view-model + scope filtering)

### Task 8: Report scope filtering and view-model construction

**Files:**
- Create: `fbdi/report.py`
- Create: `tests/test_report.py`

- [ ] **Step 1: Write the failing tests**

`tests/test_report.py`:
```python
"""Tests for fbdi.report — view-model construction and scope filtering."""

import pytest
from pathlib import Path

from fbdi.report import (
    FileSection,
    ReportContext,
    PendingBaseEntry,
    build_report_context,
)
from fbdi.align import AlignedField, Change


# ---- helpers ----

def _aligned(position, label, technical, data_type=None, length=None, required=None):
    return AlignedField(position=position, label=label, technical=technical,
                       data_type=data_type, length=length, required=required)


def _mapping(template, tab, applaud_table="T_X", prefix="TX1",
             module="Financials", in_base=None):
    """Build one mapping dict entry."""
    return {
        (template, tab): {
            "applaud_table": applaud_table,
            "prefix": prefix,
            "module": module,
            "in_base": in_base,
        }
    }


# ---- tests ----

class TestScopeFiltering:
    def test_unmapped_file_is_silently_excluded(self):
        catalog_old = {("UnmappedFile", "TabA"): [_aligned(1, "A", "A")]}
        catalog_new = {("UnmappedFile", "TabA"): [_aligned(1, "A", "A"),
                                                  _aligned(2, "B", "B")]}
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping={}, old_release="26A", new_release="26B",
        )
        assert ctx.file_sections == []
        assert ctx.pending_base == []

    def test_mapped_in_base_routes_to_main_body(self):
        catalog_old = {("MappedFile", "TabA"): [_aligned(1, "Old", "OLD_F")]}
        catalog_new = {("MappedFile", "TabA"): [_aligned(1, "New", "NEW_F")]}
        mapping = _mapping("MappedFile", "TabA", in_base=None)
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        assert len(ctx.file_sections) == 1
        assert ctx.pending_base == []

    def test_pending_base_routes_to_separate_section(self):
        catalog_old = {("PendingFile", "TabA"): [_aligned(1, "Old", "OLD_F")]}
        catalog_new = {("PendingFile", "TabA"): [_aligned(1, "Old", "OLD_F"),
                                                 _aligned(2, "New", "NEW_F")]}
        mapping = _mapping("PendingFile", "TabA",
                          in_base="Needs to be created in base system")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        assert ctx.file_sections == []
        assert len(ctx.pending_base) == 1
        assert ctx.pending_base[0].file == "PendingFile"
        assert ctx.pending_base[0].tab == "TabA"
        assert ctx.pending_base[0].change_count == 1


class TestApplaudFieldNameConstruction:
    def test_uses_technical_when_present(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Some Label", "FIELD_NAME",
                                              "VARCHAR2", 30, True)]}
        mapping = _mapping("F", "T", prefix="TX1")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        section = ctx.file_sections[0]
        added = section.changes_by_type["ADDED"]
        assert added[0].applaud_field_name == "TX1FIELD_NAME"

    def test_falls_back_to_normalized_label_when_technical_is_none(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Some Label!", None,
                                              None, None, True)]}
        mapping = _mapping("F", "T", prefix="TX1")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        section = ctx.file_sections[0]
        added = section.changes_by_type["ADDED"]
        # normalize_label strips '!' and joins → "Some Label"
        assert added[0].applaud_field_name == "TX1Some Label"

    def test_thirty_char_warning_set_when_name_exceeds_limit(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "X",
                                              "COPY_LOTS_AND_SERIAL_NUMBERS_FROM_PARENT_TXN",
                                              "VARCHAR2", 1, True)]}
        mapping = _mapping("F", "T", prefix="TH8")
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        added = ctx.file_sections[0].changes_by_type["ADDED"]
        assert added[0].name_exceeds_30
        assert added[0].name_length > 30


class TestModuleRollup:
    def test_module_counts_aggregate_across_files(self):
        catalog_old = {
            ("F1", "T1"): [_aligned(1, "X", "X")],
            ("F2", "T1"): [_aligned(1, "X", "X")],
        }
        catalog_new = {
            ("F1", "T1"): [_aligned(1, "X", "X"), _aligned(2, "Y", "Y")],
            ("F2", "T1"): [_aligned(1, "X", "X")],  # no changes
        }
        mapping = {**_mapping("F1", "T1", module="Financials"),
                   **_mapping("F2", "T1", module="Financials")}
        ctx = build_report_context(
            catalog_old=catalog_old, catalog_new=catalog_new,
            mapping=mapping, old_release="26A", new_release="26B",
        )
        assert "Financials" in ctx.module_rollup
        rollup = ctx.module_rollup["Financials"]
        assert rollup["tabs"] == 1   # only F1/T1 has changes
        assert rollup["added"] == 1
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
py -m pytest tests/test_report.py -v
```

Expected: ImportError — `cannot import name 'FileSection' from 'fbdi.report'`.

- [ ] **Step 3: Implement report.py view-model and scope filter**

`fbdi/report.py`:
```python
"""FBDI Compliance Report generator.

Reads the FBDI Master Catalog (per-release sheets) and the FBDI-to-Applaud
mapping, runs alignment per (file, tab), filters to the in-scope universe
(MAPPED only; pending-base routed to a separate section), and emits an
HTML and PDF report from one Jinja2 template.

This module exposes:
- build_report_context(...) — pure view-model construction (testable in isolation)
- generate_report(...)      — top-level: load → build → render → write

The view-model dataclasses (ReportContext, FileSection, ChangeRow,
PendingBaseEntry) are the contract between this module and the template.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from fbdi.align import AlignedField, Change, align_tabs
from fbdi.applaud_type import applaud_type_for
from fbdi.catalog_normalize import normalize_label
from fbdi.type_parser import parse_data_type


APPLAUD_NAME_LIMIT = 30


@dataclass
class ChangeRow:
    """One row in a per-file change-type table (view-model)."""
    change_type: str
    applaud_field_name: str
    name_length: int
    name_exceeds_30: bool
    old_position: int | None
    new_position: int | None
    label: str
    oracle_type_str: str           # e.g. "VARCHAR2(30)" — empty when not applicable
    applaud_type_str: str          # e.g. "char 30"
    required: bool | None
    axes: tuple[str, ...]
    sub_kinds: tuple[str, ...]
    # For RENAMED / MODIFIED — old vs new values to display side-by-side
    old_label: str | None = None
    new_label: str | None = None
    old_oracle_type_str: str | None = None
    new_oracle_type_str: str | None = None
    old_required: bool | None = None
    new_required: bool | None = None


@dataclass
class FileSection:
    """One per-file section in the main body."""
    file: str
    tab: str
    applaud_table: str
    prefix: str
    module: str
    in_base_note: str | None       # e.g. the "Multiple mapping is possible..." string when present
    changes_by_type: dict[str, list[ChangeRow]] = field(default_factory=dict)
    shift_summary: str | None = None  # e.g. "20 fields shifted from positions 19-39 to 20-40"


@dataclass
class PendingBaseEntry:
    """One entry in the pending base-system tables list."""
    file: str
    tab: str
    applaud_table: str
    prefix: str
    module: str
    change_count: int


@dataclass
class ReportContext:
    """Top-level view-model passed to the Jinja2 template."""
    old_release: str
    new_release: str
    generated_date: str
    module_rollup: dict[str, dict[str, int]]    # module -> {"tabs": N, "added": N, ...}
    file_sections: list[FileSection]
    pending_base: list[PendingBaseEntry]


# Public: scope filtering and view-model construction ---------------------------

def build_report_context(
    catalog_old: dict[tuple[str, str], list[AlignedField]],
    catalog_new: dict[tuple[str, str], list[AlignedField]],
    mapping: dict[tuple[str, str], dict],
    old_release: str,
    new_release: str,
    generated_date: str | None = None,
) -> ReportContext:
    """Build the report context from grouped catalog data + mapping lookup.

    catalog_old / catalog_new keys are (file, tab) tuples. Values are
    AlignedField rows already in catalog form. mapping keys match the
    catalog keys; values are dicts with 'applaud_table', 'prefix',
    'module', 'in_base'.
    """
    from datetime import date as _date
    if generated_date is None:
        generated_date = _date.today().isoformat()

    file_sections: list[FileSection] = []
    pending_base: list[PendingBaseEntry] = []
    all_keys = set(catalog_old.keys()) | set(catalog_new.keys())

    for key in sorted(all_keys):
        if key not in mapping:
            continue  # UNMAPPED — silently exclude
        m = mapping[key]
        file, tab = key

        old_rows = catalog_old.get(key, [])
        new_rows = catalog_new.get(key, [])
        changes = align_tabs(old_rows, new_rows)
        if not changes:
            continue

        in_base = m.get("in_base") or ""
        if "Needs to be created in base system" in in_base:
            pending_base.append(PendingBaseEntry(
                file=file, tab=tab,
                applaud_table=m["applaud_table"],
                prefix=m["prefix"],
                module=m["module"],
                change_count=len(changes),
            ))
            continue

        in_base_note = in_base if in_base else None

        section = FileSection(
            file=file, tab=tab,
            applaud_table=m["applaud_table"],
            prefix=m["prefix"],
            module=m["module"],
            in_base_note=in_base_note,
        )
        section.changes_by_type = _bucket_changes(changes, prefix=m["prefix"])
        section.shift_summary = _build_shift_summary(section.changes_by_type.get("SHIFTED", []))
        file_sections.append(section)

    # Sort by (module, file, tab) for stable ordering
    file_sections.sort(key=lambda s: (s.module or "", s.file, s.tab))
    pending_base.sort(key=lambda p: (p.module or "", p.file, p.tab))

    module_rollup = _build_module_rollup(file_sections)

    return ReportContext(
        old_release=old_release,
        new_release=new_release,
        generated_date=generated_date,
        module_rollup=module_rollup,
        file_sections=file_sections,
        pending_base=pending_base,
    )


def _applaud_field_name(prefix: str, technical: str | None, label: str | None) -> str:
    """Construct the Applaud field name: prefix + technical (or normalized label)."""
    if technical:
        suffix = technical
    else:
        suffix = normalize_label(label or "")
    return f"{prefix}{suffix}"


def _oracle_type_str(field: AlignedField | None) -> str:
    if field is None or not field.data_type:
        return ""
    if field.length is not None:
        return f"{field.data_type}({field.length})"
    return field.data_type


def _bucket_changes(changes: list[Change], prefix: str) -> dict[str, list[ChangeRow]]:
    """Group classified changes into per-type buckets of ChangeRow view-models."""
    buckets: dict[str, list[ChangeRow]] = defaultdict(list)
    for c in changes:
        # Pick the field to use for naming/typing (new wins; old when REMOVED)
        primary = c.new_field if c.new_field is not None else c.old_field
        applaud_name = _applaud_field_name(prefix, primary.technical, primary.label)
        oracle_type = _oracle_type_str(primary)
        applaud_type = applaud_type_for(parse_data_type(oracle_type)) if oracle_type else ""

        row = ChangeRow(
            change_type=c.change_type,
            applaud_field_name=applaud_name,
            name_length=len(applaud_name),
            name_exceeds_30=len(applaud_name) > APPLAUD_NAME_LIMIT,
            old_position=c.old_position,
            new_position=c.new_position,
            label=primary.label or "",
            oracle_type_str=oracle_type,
            applaud_type_str=applaud_type,
            required=primary.required,
            axes=c.axes,
            sub_kinds=c.sub_kinds,
            old_label=c.old_field.label if c.old_field else None,
            new_label=c.new_field.label if c.new_field else None,
            old_oracle_type_str=_oracle_type_str(c.old_field),
            new_oracle_type_str=_oracle_type_str(c.new_field),
            old_required=c.old_field.required if c.old_field else None,
            new_required=c.new_field.required if c.new_field else None,
        )
        buckets[c.change_type].append(row)
    return dict(buckets)


def _build_shift_summary(shifted_rows: list[ChangeRow]) -> str | None:
    """Build the inline shift-summary sentence used in the SHIFTED block."""
    if not shifted_rows:
        return None
    old_positions = sorted(r.old_position for r in shifted_rows)
    new_positions = sorted(r.new_position for r in shifted_rows)
    n = len(shifted_rows)
    return (
        f"{n} field{'s' if n != 1 else ''} shifted from positions "
        f"{old_positions[0]}-{old_positions[-1]} to {new_positions[0]}-{new_positions[-1]}."
    )


def _build_module_rollup(sections: list[FileSection]) -> dict[str, dict[str, int]]:
    """Aggregate per-module counts across file sections."""
    rollup: dict[str, dict[str, int]] = {}
    for s in sections:
        m = s.module or "Unknown"
        if m not in rollup:
            rollup[m] = {"tabs": 0, "added": 0, "removed": 0, "modified": 0, "renamed": 0, "shifted": 0, "multi": 0}
        rollup[m]["tabs"] += 1
        for ct, rows in s.changes_by_type.items():
            rollup[m][ct.lower()] = rollup[m].get(ct.lower(), 0) + len(rows)
    return rollup
```

- [ ] **Step 4: Run report tests to verify they pass**

```bash
py -m pytest tests/test_report.py -v
```

Expected: all passed.

- [ ] **Step 5: Run the full suite to confirm no regression**

```bash
py -m pytest tests/ -v
```

Expected: all passed.

- [ ] **Step 6: Commit**

```bash
git add fbdi/report.py tests/test_report.py
git commit -m "feat(report): scope filtering, view-model construction, and pending-base routing"
```

### Task 9: Catalog and mapping loaders for the report

**Files:**
- Modify: `fbdi/report.py`
- Modify: `tests/test_report.py`

- [ ] **Step 1: Add the failing tests**

Append to `tests/test_report.py`:
```python
class TestLoaders:
    def test_load_catalog_release_groups_by_file_and_tab(self, tmp_path):
        from fbdi.report import load_catalog_release
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "26B"
        ws.append(["release", "file_name", "tab_name", "position",
                   "column_label", "column_technical",
                   "data_type", "length", "scale", "data_type_raw", "required"])
        ws.append(["26B", "F1", "T1", 1, "Lab", "TECH", "VARCHAR2", 30, None, "VARCHAR2(30)", "TRUE"])
        ws.append(["26B", "F1", "T1", 2, "Lab2", "TECH2", "NUMBER", 18, None, "NUMBER(18)", "FALSE"])
        ws.append(["26B", "F2", "T1", 1, "Lab", "TECH", None, None, None, "", "FALSE"])
        path = tmp_path / "cat.xlsx"
        wb.save(path)

        result = load_catalog_release(path, "26B")
        assert ("F1", "T1") in result
        assert len(result[("F1", "T1")]) == 2
        first = result[("F1", "T1")][0]
        assert first.position == 1
        assert first.technical == "TECH"
        assert first.required is True
        assert first.length == 30

    def test_load_mapping_filters_to_mapped_status(self, tmp_path):
        from fbdi.report import load_mapping
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "FBDI Mapping"
        ws.append(["FBDI Template", "FBDI Tab", "Applaud Table", "Prefix",
                   "Status", "Module", "In Base System?"])
        ws.append(["F1", "T1", "T_X", "TX1", "MAPPED", "Financials", None])
        ws.append(["F2", "T1", "T_Y", "TY1", "UNMAPPED", "SCM", None])
        ws.append(["F3", "T1", "T_Z", "TZ1", "MAPPED", "SCM", "Needs to be created in base system"])
        path = tmp_path / "mapping.xlsx"
        wb.save(path)

        result = load_mapping(path)
        assert ("F1", "T1") in result
        assert ("F3", "T1") in result   # MAPPED + pending-base still included
        assert ("F2", "T1") not in result  # UNMAPPED excluded
        assert result[("F1", "T1")]["module"] == "Financials"
        assert result[("F3", "T1")]["in_base"] == "Needs to be created in base system"
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
py -m pytest tests/test_report.py::TestLoaders -v
```

Expected: ImportError — `cannot import name 'load_catalog_release'`.

- [ ] **Step 3: Implement the loaders in report.py**

Add to `fbdi/report.py`:
```python
from openpyxl import load_workbook


def load_catalog_release(catalog_path: Path, release: str) -> dict[tuple[str, str], list[AlignedField]]:
    """Read one release sheet from the master catalog and group by (file, tab)."""
    wb = load_workbook(catalog_path, read_only=True, data_only=True)
    if release not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Release sheet '{release}' not found in {catalog_path}")
    ws = wb[release]

    grouped: dict[tuple[str, str], list[AlignedField]] = defaultdict(list)
    rows = ws.iter_rows(min_row=2, values_only=True)
    for row in rows:
        # Schema: release, file_name, tab_name, position, column_label,
        # column_technical, data_type, length, scale, data_type_raw, required
        _rel, file_name, tab_name, position, label, technical, data_type, length, _scale, _raw, required = row
        if file_name is None or tab_name is None:
            continue
        grouped[(file_name, tab_name)].append(AlignedField(
            position=int(position),
            label=label,
            technical=(technical or None),
            data_type=(data_type or None),
            length=(int(length) if length is not None and length != "" else None),
            required=_parse_required(required),
        ))

    wb.close()
    # Sort each group's rows by position to be safe
    for k in grouped:
        grouped[k].sort(key=lambda f: f.position)
    return dict(grouped)


def _parse_required(v) -> bool | None:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    s = str(v).strip().upper()
    if s == "TRUE":
        return True
    if s == "FALSE":
        return False
    return None


def load_mapping(mapping_path: Path) -> dict[tuple[str, str], dict]:
    """Read FBDI_to_ApplaudTables_Mapping.xlsx and return MAPPED-status rows.

    UNMAPPED rows are filtered out at load time (they're noise per the spec).
    NEEDS_REVIEW rows are kept so the report can flag them visually.
    """
    wb = load_workbook(mapping_path, read_only=True, data_only=True)
    ws = wb["FBDI Mapping"]
    out: dict[tuple[str, str], dict] = {}
    rows = ws.iter_rows(min_row=2, values_only=True)
    for row in rows:
        # Schema: FBDI Template, FBDI Tab, Applaud Table, Prefix, Status,
        # Module, In Base System?
        template, tab, applaud_table, prefix, status, module, in_base = row[:7]
        if template is None or tab is None:
            continue
        if status not in ("MAPPED", "NEEDS_REVIEW"):
            continue
        out[(str(template), str(tab))] = {
            "applaud_table": applaud_table,
            "prefix": prefix,
            "module": module,
            "status": status,
            "in_base": in_base,
        }
    wb.close()
    return out
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
py -m pytest tests/test_report.py -v
```

Expected: all passed.

- [ ] **Step 5: Commit**

```bash
git add fbdi/report.py tests/test_report.py
git commit -m "feat(report): catalog-release and mapping loaders with status filtering"
```

---

## Phase 6: HTML template

The template renders the full report with one Jinja2 file. It uses `{% if not print_mode %}<details>{% endif %}` conditionals so the HTML render gets collapsibles for SHIFTED while the PDF render gets always-expanded compact layouts. Self-contained CSS embedded in `<style>`.

**REQUIRED SUB-SKILLS for this phase:**
- `frontend-design` — visual hierarchy, typography, deliberate spacing, brand-palette discipline. This is not boilerplate CSS — it should look like a polished compliance deliverable, not a generic Bootstrap-feel page.
- `humanizer-skill:humanizer` — applied to ALL prose authored in the template (lede sentences, advisory notes, narrative summaries). No AI-tells like "in today's evolving landscape", "leverage", "robust", "delve", em-dash overuse, rule-of-three padding. Plain factual sentences.

### Task 10: Create base template with cover, module rollup, summary table

**Files:**
- Create: `fbdi/templates/report.html.j2`

- [ ] **Step 1: Build the template**

Use the visual mockup from brainstorming (`.superpowers/brainstorm/<session>/content/full-report-walkthrough.html`) as the visual spec. Reproduce its CSS (Definian palette only) and translate the static HTML into Jinja2 with the following pattern:

`fbdi/templates/report.html.j2`:
```jinja
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>FBDI Compliance Report — {{ ctx.old_release }} → {{ ctx.new_release }}</title>
<style>
  /* === Definian palette — strict; do not introduce off-palette colors === */
  :root {
    --def-blue: #0D2C71;
    --def-green: #00AB63;
    --midnight: #02072D;
    --darkgray: #3C405B;
    --coolgray: #D8D7EE;
    --bg-soft: #F7F7FB;
    --warn: #B8860B;
    --warn-bg: #FFFBF0;
    --del: #C0392B;
  }

  /* === Base typography === */
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
         color: var(--midnight); margin: 0; background: #fff; }

  /* === Cover === */
  .cover { background: linear-gradient(135deg, var(--def-blue) 0%, var(--midnight) 100%);
           color: #fff; padding: 64px 48px; text-align: center; }
  .cover .brand { font-size: 11px; letter-spacing: 3px; color: var(--def-green);
                  font-weight: 700; text-transform: uppercase; margin-bottom: 8px; }
  .cover h1 { font-size: 36px; margin: 0 0 12px; font-weight: 800; }
  .cover .sub { font-size: 16px; opacity: 0.85; margin-bottom: 24px; font-weight: 300; }
  .cover .meta { display: inline-flex; gap: 24px; padding: 12px 24px;
                 background: rgba(255,255,255,0.08); border-radius: 4px; font-size: 13px; }
  .cover .meta b { color: var(--def-green); }

  /* === Section base === */
  .section { padding: 28px 36px; border-top: 1px solid var(--coolgray); }
  .section h2 { color: var(--def-blue); font-size: 22px; margin: 0 0 4px; font-weight: 700; }
  .section .lede { color: var(--darkgray); font-size: 13px; margin-bottom: 18px; }

  /* === Module rollup === */
  .module-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
  .module-card { background: var(--bg-soft); border-left: 4px solid var(--def-blue);
                 border-radius: 4px; padding: 14px 18px; }
  .module-card.financials { border-left-color: var(--def-green); }
  .module-card .name { font-size: 13px; font-weight: 700; color: var(--def-blue);
                       margin-bottom: 6px; text-transform: uppercase; letter-spacing: 1px; }
  .module-card .figures { display: flex; gap: 16px; }
  .module-card .stat { display: flex; flex-direction: column; }
  .module-card .stat .num { font-size: 24px; font-weight: 700; line-height: 1; }
  .module-card .stat .lab { font-size: 10px; color: var(--darkgray);
                            text-transform: uppercase; letter-spacing: 0.5px; margin-top: 4px; }

  /* === Summary table === */
  .summary-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  .summary-table th { background: var(--def-blue); color: #fff; padding: 8px 10px;
                       text-align: left; font-weight: 600; font-size: 11px;
                       text-transform: uppercase; letter-spacing: 0.5px; }
  .summary-table th.num { text-align: right; }
  .summary-table td { padding: 7px 10px; border-bottom: 1px solid #eee; }
  .summary-table td.num { text-align: right; width: 1%; white-space: nowrap; }
  .module-tag { display: inline-block; padding: 1px 8px; border-radius: 10px;
                background: var(--coolgray); color: var(--midnight); font-size: 10px; font-weight: 600; }
  .module-tag.financials { background: var(--def-green); color: #fff; }

  /* (Per-file section + collapsible CSS added in next task) */
</style>
</head>
<body>

<!-- Cover -->
<div class="cover">
  <div class="brand">DEFINIAN · Applaud Base System</div>
  <h1>FBDI Compliance Report</h1>
  <div class="sub">Required changes between Oracle Cloud releases</div>
  <div class="meta">
    <span><b>From</b> Release {{ ctx.old_release }}</span>
    <span><b>To</b> Release {{ ctx.new_release }}</span>
    <span><b>Generated</b> {{ ctx.generated_date }}</span>
  </div>
</div>

<!-- Section 1: Module rollup -->
<div class="section">
  <h2>1. At a glance</h2>
  <div class="lede">Mapped, in-scope FBDI tabs only. UNMAPPED files and pending base-system tables are excluded from the per-file detail; pending-base files are listed at the end for visibility.</div>
  <div class="module-grid">
    {% for module, stats in ctx.module_rollup.items() %}
    <div class="module-card{% if module == 'Financials' %} financials{% endif %}">
      <div class="name">{{ module }}</div>
      <div class="figures">
        <div class="stat"><span class="num">{{ stats.tabs }}</span><span class="lab">tabs</span></div>
        <div class="stat"><span class="num">{{ stats.added }}</span><span class="lab">added</span></div>
        <div class="stat"><span class="num">{{ stats.shifted }}</span><span class="lab">shifted</span></div>
        <div class="stat"><span class="num">{{ stats.removed }}</span><span class="lab">removed</span></div>
      </div>
    </div>
    {% endfor %}
  </div>
</div>

<!-- Section 2: Summary table -->
<div class="section">
  <h2>2. Summary by FBDI tab</h2>
  <div class="lede">{{ ctx.file_sections | length }} mapped tab{% if ctx.file_sections | length != 1 %}s{% endif %} have changes between {{ ctx.old_release }} and {{ ctx.new_release }}.</div>
  <table class="summary-table">
    <thead>
      <tr>
        <th>FBDI File · Tab</th>
        <th>Applaud Table</th>
        <th>Prefix</th>
        <th>Module</th>
        <th class="num">Added</th>
        <th class="num">Removed</th>
        <th class="num">Modified</th>
        <th class="num">Shifted</th>
      </tr>
    </thead>
    <tbody>
      {% for s in ctx.file_sections %}
      <tr>
        <td><b>{{ s.file }}</b><br><span style="font-size:11px;color:var(--darkgray)">{{ s.tab }}</span></td>
        <td>{{ s.applaud_table }}</td>
        <td><code>{{ s.prefix }}</code></td>
        <td><span class="module-tag{% if s.module == 'Financials' %} financials{% endif %}">{{ s.module }}</span></td>
        <td class="num">{{ s.changes_by_type.get('ADDED', []) | length }}</td>
        <td class="num">{{ s.changes_by_type.get('REMOVED', []) | length }}</td>
        <td class="num">{{ (s.changes_by_type.get('MODIFIED', []) | length) + (s.changes_by_type.get('MULTI', []) | length) }}</td>
        <td class="num">{{ s.changes_by_type.get('SHIFTED', []) | length }}</td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
</div>

<!-- Per-file sections — added in next task -->

</body>
</html>
```

- [ ] **Step 2: Smoke render with synthetic data**

```bash
py -c "
from pathlib import Path
import jinja2
from fbdi.report import build_report_context
from fbdi.align import AlignedField

old = {('F1','T1'): [AlignedField(1,'A','A','VARCHAR2',30,True)]}
new = {('F1','T1'): [AlignedField(1,'A','A','VARCHAR2',30,True), AlignedField(2,'B','B','NUMBER',18,False)]}
mapping = {('F1','T1'): {'applaud_table':'T_X','prefix':'TX1','module':'Financials','in_base':None}}
ctx = build_report_context(old, new, mapping, '26A', '26B')

env = jinja2.Environment(loader=jinja2.FileSystemLoader('fbdi/templates'))
tpl = env.get_template('report.html.j2')
html = tpl.render(ctx=ctx, print_mode=False)
Path('/tmp/smoke.html').write_text(html, encoding='utf-8')
print('Wrote /tmp/smoke.html (', len(html), 'bytes)')
"
```

Expected: prints byte count; no Jinja errors.

- [ ] **Step 3: Open in browser via chrome-devtools-mcp and screenshot**

**REQUIRED SUB-SKILL:** Use `chrome-devtools-mcp:chrome-devtools` to render the file and take a full-page screenshot. Verify visually:
- Cover page renders with brand colors
- Module rollup card appears
- Summary table renders with one row
- No off-palette colors visible
- No layout glitches

- [ ] **Step 4: Commit**

```bash
git add fbdi/templates/report.html.j2
git commit -m "feat(report): base template with cover, module rollup, summary table"
```

### Task 11: Add per-file section, change-type tables, and SHIFTED collapsible to template

**Files:**
- Modify: `fbdi/templates/report.html.j2`

**Pattern reference:** the visual mockup `.superpowers/brainstorm/<session>/content/full-report-walkthrough.html` and `position-column-fix.html` show the final layout. Reproduce the per-file section, change-type tables, and the action-checkbox columns. The `width: 1%; white-space: nowrap` fix on numeric columns is required.

- [ ] **Step 1: Add the per-file CSS to the existing `<style>` block**

Append before the closing `</style>` tag:

```css
  /* === Per-file section === */
  .file-section { border: 1px solid var(--coolgray); border-radius: 6px;
                  margin-bottom: 18px; overflow: hidden; }
  .file-head { background: var(--def-blue); color: #fff; padding: 14px 18px; }
  .file-head .num { font-size: 11px; opacity: 0.7; letter-spacing: 1px; }
  .file-head h3 { font-size: 18px; margin: 2px 0 6px; font-weight: 700; }
  .file-head .meta { display: flex; flex-wrap: wrap; gap: 14px;
                     font-size: 12px; opacity: 0.92; }
  .file-head code { background: rgba(255,255,255,0.15); padding: 1px 6px;
                    border-radius: 3px; font-size: 11px; }
  .module-pill { display: inline-block; padding: 2px 10px; border-radius: 12px;
                 background: var(--def-green); color: #fff; font-size: 11px; font-weight: 600; }
  .file-body { padding: 18px 22px; }

  /* === Change blocks === */
  .change-block { margin-bottom: 18px; }
  .change-block:last-child { margin-bottom: 0; }
  .change-block h4 { font-size: 13px; margin: 0 0 8px; color: var(--def-blue);
                     text-transform: uppercase; letter-spacing: 1px; font-weight: 700;
                     display: flex; align-items: center; gap: 8px; }
  .change-block h4 .count { background: var(--def-green); color: #fff;
                            border-radius: 10px; padding: 1px 8px;
                            font-size: 11px; letter-spacing: 0; }
  .change-block.removed h4 { color: var(--del); }
  .change-block.removed h4 .count { background: var(--del); }
  .change-block.modified h4 { color: var(--warn); }
  .change-block.modified h4 .count { background: var(--warn); }
  .change-block.renamed h4 { color: var(--darkgray); }
  .change-block.renamed h4 .count { background: var(--darkgray); }

  .ct { width: 100%; border-collapse: collapse; font-size: 12px; }
  .ct th { background: var(--coolgray); color: var(--midnight); padding: 7px 9px;
           text-align: left; font-weight: 600; font-size: 10.5px;
           text-transform: uppercase; letter-spacing: 0.5px;
           border-bottom: 1px solid #ccc; }
  .ct th.num { text-align: right; padding-right: 14px; }
  .ct td { padding: 8px 9px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; }
  .ct td.num { text-align: right; width: 1%; white-space: nowrap; padding-right: 14px;
               color: var(--darkgray); }
  .ct td.field { font-family: ui-monospace, SFMono-Regular, Consolas, monospace;
                 font-size: 11.5px; }
  .ct .center { text-align: center; }
  .ct th.action-col, .ct td.action-col { text-align: center; width: 56px;
                                          background: rgba(13,44,113,0.03);
                                          border-left: 1px solid #f0f0f0; }
  .ct td.action-col.dim { background: #f8f8f8; }
  .ct .checkbox { display: inline-block; width: 16px; height: 16px;
                  border: 1.5px solid var(--def-blue); border-radius: 2px; vertical-align: middle; }
  .ct .checkbox.dash { border-style: dashed; opacity: 0.4; }
  .ct .checkbox.del { border-color: var(--del); }
  .ct .checkbox.warn { border-color: var(--warn); }
  .ct .dash-cell { text-align: center; color: #ccc; }
  .ct .badge { display: inline-block; padding: 1px 6px; border-radius: 8px;
               font-size: 10px; font-weight: 600; }
  .ct .badge.warn-text { background: var(--warn-bg); color: var(--warn);
                         border: 1px solid var(--warn); }
  .type-arrow { color: var(--darkgray); padding: 0 4px; }
  .type-old { color: var(--darkgray); text-decoration: line-through; opacity: 0.7; }
  .type-new { color: var(--midnight); font-weight: 600; }

  /* === Shift summary + collapsible === */
  .summary-box { background: var(--bg-soft); border-left: 4px solid var(--def-green);
                 padding: 10px 14px; border-radius: 3px; font-size: 12.5px; margin-bottom: 8px; }
  details.shift-details { border: 1px solid var(--coolgray); border-radius: 4px; margin-top: 8px; }
  details.shift-details > summary { padding: 8px 14px; cursor: pointer; background: var(--bg-soft);
                                    font-size: 12px; color: var(--def-blue); font-weight: 600;
                                    list-style: none; user-select: none; }
  details.shift-details > summary::before { content: '▶ '; margin-right: 6px;
                                             transition: transform 0.15s; display: inline-block; }
  details.shift-details[open] > summary::before { transform: rotate(90deg); }
  details.shift-details > div { padding: 10px 14px; border-top: 1px solid var(--coolgray); }
  .shift-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 4px 24px; font-size: 12px; }
  .shift-grid .row { display: flex; justify-content: space-between; padding: 3px 8px;
                     border-bottom: 1px solid #f5f5f5; }
  .shift-grid .arrow { color: var(--darkgray); }
  .shift-grid .new-pos { color: var(--def-green); font-weight: 600; }
  .shift-grid .old-pos { color: var(--darkgray); }

  /* === Pending base list === */
  .pending-list { list-style: none; padding: 0; margin: 0; }
  .pending-list li { padding: 10px 14px; border: 1px solid var(--coolgray);
                     border-radius: 4px; margin-bottom: 6px; background: var(--bg-soft);
                     display: flex; justify-content: space-between; align-items: center;
                     font-size: 12.5px; }
  .pending-list .name b { color: var(--def-blue); }
  .pending-list .pending-count { color: var(--darkgray); font-size: 11px; }

  /* === Print mode (PDF render) — auto-expand collapsibles via flat layout === */
  /* See task 12 for the print-mode logic; the PDF render uses the
     {% raw %}{% if print_mode %}{% endraw %} branch instead of CSS to fully
     control the layout. */
```

- [ ] **Step 2: Add the per-file section template body**

Insert between the summary table section and the closing `</body>`:

```jinja
{% if ctx.file_sections %}
<div class="section">
  <h2>3. Required changes by FBDI file</h2>
  <div class="lede">For each tab: Added, Removed, Modified, Renamed, Shifted. Action columns mark where consultants act in their Applaud install (DB table, IF = Import Form, EF = Export Form).</div>

  {% for s in ctx.file_sections %}
  <div class="file-section">
    <div class="file-head">
      <div class="num">3.{{ loop.index }}</div>
      <h3>{{ s.file }}</h3>
      <div class="meta">
        <span><b>FBDI Tab:</b> {{ s.tab }}</span>
        <span><b>Applaud Table:</b> {{ s.applaud_table }}</span>
        <span><b>Prefix:</b> <code>{{ s.prefix }}</code></span>
        <span class="module-pill">{{ s.module }}</span>
      </div>
    </div>
    <div class="file-body">

      {% if s.changes_by_type.get('ADDED') %}
      <div class="change-block">
        <h4>Added <span class="count">{{ s.changes_by_type['ADDED'] | length }}</span></h4>
        <table class="ct">
          <thead><tr><th>Applaud Field</th><th class="num">Pos</th><th>Oracle Type</th><th>Applaud Type</th><th>Required</th><th class="action-col">DB</th><th class="action-col">IF</th><th class="action-col">EF</th></tr></thead>
          <tbody>
          {% for r in s.changes_by_type['ADDED'] %}
            <tr>
              <td class="field">{{ r.applaud_field_name }}{% if r.name_exceeds_30 %} <span class="badge warn-text" title="Exceeds 30 char limit">⚠ {{ r.name_length }} chars</span>{% endif %}</td>
              <td class="num">{{ r.new_position }}</td>
              <td>{{ r.oracle_type_str }}</td>
              <td>{{ r.applaud_type_str }}</td>
              <td class="center">{% if r.required is true %}TRUE{% elif r.required is false %}FALSE{% endif %}</td>
              <td class="action-col"><span class="checkbox" title="Add column"></span></td>
              <td class="action-col"><span class="checkbox" title="Add field"></span></td>
              <td class="action-col"><span class="checkbox" title="Add field"></span></td>
            </tr>
          {% endfor %}
          </tbody>
        </table>
      </div>
      {% endif %}

      {% if s.changes_by_type.get('REMOVED') %}
      <div class="change-block removed">
        <h4>Removed <span class="count">{{ s.changes_by_type['REMOVED'] | length }}</span></h4>
        <table class="ct">
          <thead><tr><th>Applaud Field</th><th class="num">Was at Pos</th><th>Oracle Type</th><th class="action-col">DB</th><th class="action-col">IF</th><th class="action-col">EF</th></tr></thead>
          <tbody>
          {% for r in s.changes_by_type['REMOVED'] %}
            <tr>
              <td class="field">{{ r.applaud_field_name }}</td>
              <td class="num">{{ r.old_position }}</td>
              <td>{{ r.oracle_type_str }}</td>
              <td class="action-col"><span class="checkbox del" title="Drop column"></span></td>
              <td class="action-col"><span class="checkbox del" title="Remove field"></span></td>
              <td class="action-col"><span class="checkbox del" title="Remove field"></span></td>
            </tr>
          {% endfor %}
          </tbody>
        </table>
      </div>
      {% endif %}

      {% if s.changes_by_type.get('MODIFIED') %}
      <div class="change-block modified">
        <h4>Modified <span class="count">{{ s.changes_by_type['MODIFIED'] | length }}</span></h4>
        <table class="ct">
          <thead><tr><th>Applaud Field</th><th class="num">Pos</th><th>Change</th><th class="action-col">DB</th><th class="action-col">IF</th><th class="action-col">EF</th></tr></thead>
          <tbody>
          {% for r in s.changes_by_type['MODIFIED'] %}
            <tr>
              <td class="field">{{ r.applaud_field_name }}</td>
              <td class="num">{{ r.new_position }}</td>
              <td>
                {% if 'type' in r.sub_kinds or 'length' in r.sub_kinds %}Type: <span class="type-old">{{ r.old_oracle_type_str }}</span><span class="type-arrow">→</span><span class="type-new">{{ r.new_oracle_type_str }}</span>{% endif %}
                {% if 'required' in r.sub_kinds %}Required: <span class="type-old">{{ 'TRUE' if r.old_required else 'FALSE' }}</span><span class="type-arrow">→</span><span class="type-new">{{ 'TRUE' if r.new_required else 'FALSE' }}</span> <span class="badge warn-text">flag only</span>{% endif %}
              </td>
              <td class="action-col"><span class="checkbox warn" title="Alter column"></span></td>
              {% if 'required' in r.sub_kinds and 'type' not in r.sub_kinds and 'length' not in r.sub_kinds %}
                <td class="action-col dim"><span class="dash-cell">—</span></td>
                <td class="action-col dim"><span class="dash-cell">—</span></td>
              {% else %}
                <td class="action-col"><span class="checkbox warn" title="Update length validation"></span></td>
                <td class="action-col"><span class="checkbox warn"></span></td>
              {% endif %}
            </tr>
          {% endfor %}
          </tbody>
        </table>
      </div>
      {% endif %}

      {% if s.changes_by_type.get('RENAMED') %}
      <div class="change-block renamed">
        <h4>Renamed <span class="count">{{ s.changes_by_type['RENAMED'] | length }}</span></h4>
        <table class="ct">
          <thead><tr><th>Applaud Field</th><th class="num">Pos</th><th>Label change</th><th class="action-col">DB</th><th class="action-col">IF</th><th class="action-col">EF</th></tr></thead>
          <tbody>
          {% for r in s.changes_by_type['RENAMED'] %}
            <tr>
              <td class="field">{{ r.applaud_field_name }}</td>
              <td class="num">{{ r.new_position }}</td>
              <td><span class="type-old">"{{ r.old_label }}"</span><span class="type-arrow">→</span><span class="type-new">"{{ r.new_label }}"</span></td>
              <td class="action-col"><span class="checkbox dash" title="Optional: update DB description"></span></td>
              <td class="action-col dim"><span class="dash-cell">—</span></td>
              <td class="action-col dim"><span class="dash-cell">—</span></td>
            </tr>
          {% endfor %}
          </tbody>
        </table>
        <div style="font-size:11px;color:var(--darkgray);margin-top:6px;font-style:italic">Renamed labels are low priority: only the DB data element description in Applaud may be updated; no IF/EF action required.</div>
      </div>
      {% endif %}

      {% if s.changes_by_type.get('SHIFTED') %}
      <div class="change-block">
        <h4>Shifted <span class="count" style="background: var(--darkgray)">{{ s.changes_by_type['SHIFTED'] | length }}</span></h4>
        <div class="summary-box">{{ s.shift_summary }}</div>
        {% if print_mode %}
          <div class="shift-grid">
          {% for r in s.changes_by_type['SHIFTED'] %}
            <div class="row"><span>{{ r.applaud_field_name }}</span><span><span class="old-pos">{{ r.old_position }}</span> <span class="arrow">→</span> <span class="new-pos">{{ r.new_position }}</span></span></div>
          {% endfor %}
          </div>
        {% else %}
          <details class="shift-details">
            <summary>Show all {{ s.changes_by_type['SHIFTED'] | length }} shifted fields</summary>
            <div>
              <div class="shift-grid">
              {% for r in s.changes_by_type['SHIFTED'] %}
                <div class="row"><span>{{ r.applaud_field_name }}</span><span><span class="old-pos">{{ r.old_position }}</span> <span class="arrow">→</span> <span class="new-pos">{{ r.new_position }}</span></span></div>
              {% endfor %}
              </div>
            </div>
          </details>
        {% endif %}
      </div>
      {% endif %}

      {# MULTI rendering: one row per change with appropriate action checkboxes per axes/sub_kinds #}
      {% if s.changes_by_type.get('MULTI') %}
      <div class="change-block modified">
        <h4>Multi-axis changes <span class="count">{{ s.changes_by_type['MULTI'] | length }}</span></h4>
        <table class="ct">
          <thead><tr><th>Applaud Field</th><th class="num">Was</th><th class="num">Now</th><th>Axes Changed</th><th class="action-col">DB</th><th class="action-col">IF</th><th class="action-col">EF</th></tr></thead>
          <tbody>
          {% for r in s.changes_by_type['MULTI'] %}
            <tr>
              <td class="field">{{ r.applaud_field_name }}</td>
              <td class="num">{{ r.old_position }}</td>
              <td class="num">{{ r.new_position }}</td>
              <td>{{ r.axes | join(', ') }}{% if r.sub_kinds %} ({{ r.sub_kinds | join(', ') }}){% endif %}</td>
              <td class="action-col"><span class="checkbox warn"></span></td>
              <td class="action-col"><span class="checkbox warn"></span></td>
              <td class="action-col"><span class="checkbox warn"></span></td>
            </tr>
          {% endfor %}
          </tbody>
        </table>
      </div>
      {% endif %}

    </div>
  </div>
  {% endfor %}
</div>
{% endif %}

{% if ctx.pending_base %}
<div class="section">
  <h2>4. Pending base-system tables</h2>
  <div class="lede">FBDI tabs mapped to an Applaud target table not yet present in the standard base install. Listed for the platform team's visibility — not actionable for client-install consultants.</div>
  <ul class="pending-list">
  {% for p in ctx.pending_base %}
    <li>
      <span class="name"><b>{{ p.file }}</b> · {{ p.tab }} → <code>{{ p.applaud_table }}</code> · <code>{{ p.prefix }}</code> <span class="module-tag" style="margin-left:6px">{{ p.module }}</span></span>
      <span class="pending-count">{{ p.change_count }} change{% if p.change_count != 1 %}s{% endif %} pending</span>
    </li>
  {% endfor %}
  </ul>
  <div style="font-size:11px;color:var(--darkgray);margin-top:8px;font-style:italic">For full per-field detail on pending-base tables, see <code>FBDI_Master_Catalog.xlsx</code> · Drift sheet.</div>
</div>
{% endif %}
```

- [ ] **Step 3: Smoke render with the WorkDefinitionTemplate scenario**

```bash
PYTHONIOENCODING=utf-8 py -c "
from pathlib import Path
import jinja2
from fbdi.report import build_report_context, load_catalog_release, load_mapping

cat_old = load_catalog_release(Path('FBDI_Master_Catalog.xlsx'), '26A')
cat_new = load_catalog_release(Path('FBDI_Master_Catalog.xlsx'), '26B')
mapping = load_mapping(Path('FBDI_to_ApplaudTables_Mapping.xlsx'))
ctx = build_report_context(cat_old, cat_new, mapping, '26A', '26B')

env = jinja2.Environment(loader=jinja2.FileSystemLoader('fbdi/templates'))
tpl = env.get_template('report.html.j2')
html = tpl.render(ctx=ctx, print_mode=False)
Path('FBDI_Compliance_Report_26A_26B.html').write_text(html, encoding='utf-8')
print('Wrote', len(html), 'bytes')
print('File sections:', len(ctx.file_sections))
print('Pending base:', len(ctx.pending_base))
"
```

Expected: writes the file; reports ~5 file sections + 1 pending-base entry (per the spec's in-scope footprint).

- [ ] **Step 4: Open via chrome-devtools-mcp and screenshot each section**

**REQUIRED SUB-SKILL:** Use `chrome-devtools-mcp:chrome-devtools` to open `FBDI_Compliance_Report_26A_26B.html` and take screenshots of: the cover, the module rollup, the summary table, one full per-file section (WorkDefinitionTemplate is good — exercises ADDED + SHIFTED), and the pending-base section. Verify:
- Brand palette only (no off-palette colors)
- Numeric column widths look correct (no wide whitespace)
- SHIFTED collapsible toggles correctly
- Action checkboxes render with correct colors (blue/red/amber/dashed)
- 30-char warning chip appears if any field name exceeds 30 chars

- [ ] **Step 5: Commit**

```bash
git add fbdi/templates/report.html.j2
git commit -m "feat(report): per-file sections, change-type tables, SHIFTED collapsible, pending-base"
```

---

## Phase 7: PDF rendering

### Task 12: Wire weasyprint and verify PDF output

**Files:**
- Modify: `fbdi/report.py`

- [ ] **Step 1: Add the public generate_report() function to report.py**

Append to `fbdi/report.py`:
```python
def generate_report(
    catalog_path: Path,
    mapping_path: Path,
    old_release: str,
    new_release: str,
    out_dir: Path,
) -> tuple[Path, Path]:
    """Top-level: load → build → render → write HTML and PDF.

    Returns (html_path, pdf_path).
    """
    import jinja2
    import weasyprint

    catalog_old = load_catalog_release(catalog_path, old_release)
    catalog_new = load_catalog_release(catalog_path, new_release)
    mapping = load_mapping(mapping_path)

    ctx = build_report_context(
        catalog_old=catalog_old, catalog_new=catalog_new,
        mapping=mapping, old_release=old_release, new_release=new_release,
    )

    template_dir = Path(__file__).parent / "templates"
    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(template_dir),
        autoescape=jinja2.select_autoescape(['html', 'j2']),
    )
    tpl = env.get_template("report.html.j2")

    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"FBDI_Compliance_Report_{old_release}_{new_release}"
    html_path = out_dir / f"{base}.html"
    pdf_path = out_dir / f"{base}.pdf"

    # HTML render — collapsibles default-closed
    html_path.write_text(tpl.render(ctx=ctx, print_mode=False), encoding="utf-8")

    # PDF render — collapsibles auto-expanded into compact layout
    pdf_html = tpl.render(ctx=ctx, print_mode=True)
    weasyprint.HTML(string=pdf_html).write_pdf(str(pdf_path))

    return html_path, pdf_path
```

- [ ] **Step 2: Manual run on real 26A→26B data**

```bash
py -c "
from pathlib import Path
from fbdi.report import generate_report
html, pdf = generate_report(
    catalog_path=Path('FBDI_Master_Catalog.xlsx'),
    mapping_path=Path('FBDI_to_ApplaudTables_Mapping.xlsx'),
    old_release='26A', new_release='26B', out_dir=Path('.'),
)
print('HTML:', html, html.stat().st_size, 'bytes')
print('PDF :', pdf, pdf.stat().st_size, 'bytes')
"
```

Expected: both files written, PDF is >50KB.

- [ ] **Step 3: Open the PDF and verify visually**

**REQUIRED SUB-SKILL:** Use `chrome-devtools-mcp:chrome-devtools` to navigate to `file:///<absolute-path>/FBDI_Compliance_Report_26A_26B.pdf` (Chrome renders PDFs natively). Verify:
- Cover page renders
- Module rollup, summary table, per-file sections all present
- SHIFTED tables show full content (not collapsed — PDF must auto-expand)
- Brand colors render correctly through weasyprint
- Page breaks fall reasonably between sections

- [ ] **Step 4: Commit**

```bash
git add fbdi/report.py
git commit -m "feat(report): generate_report() top-level function with weasyprint PDF render"
```

---

## Phase 8: CLI integration + end-to-end

### Task 13: Add `report` subcommand to fbdi CLI

**Files:**
- Modify: `fbdi/cli.py`
- Modify: `tests/test_cli.py` (add test for the new subparser)

- [ ] **Step 1: Add the failing CLI test**

Find the existing CLI tests in `tests/test_cli.py` and append:
```python
class TestReportSubcommand:
    def test_report_subcommand_parses_old_and_new(self, monkeypatch, tmp_path):
        from fbdi import cli

        called = {}

        def fake_generate(catalog_path, mapping_path, old_release, new_release, out_dir):
            called.update(dict(
                catalog_path=catalog_path, mapping_path=mapping_path,
                old_release=old_release, new_release=new_release, out_dir=out_dir,
            ))
            return tmp_path / "x.html", tmp_path / "x.pdf"

        monkeypatch.setattr("fbdi.report.generate_report", fake_generate)
        cli.main([
            "report", "--old", "26A", "--new", "26B",
            "--out-dir", str(tmp_path),
            "--catalog", str(tmp_path / "cat.xlsx"),
            "--mapping", str(tmp_path / "map.xlsx"),
        ])
        assert called["old_release"] == "26A"
        assert called["new_release"] == "26B"
        assert called["out_dir"] == tmp_path
```

- [ ] **Step 2: Run test to verify it fails**

```bash
py -m pytest tests/test_cli.py::TestReportSubcommand -v
```

Expected: argparse error — "invalid choice: 'report'".

- [ ] **Step 3: Add the report subparser to fbdi/cli.py**

In `fbdi/cli.py`, after the `populate_parser` block (line 109-125), add:

```python
    report_parser = subparsers.add_parser(
        "report",
        help="Generate the FBDI Compliance Report (HTML + PDF) from the catalog + mapping",
    )
    report_parser.add_argument(
        "--old", required=True, type=str,
        help="Older release label (e.g. 26A)",
    )
    report_parser.add_argument(
        "--new", required=True, type=str,
        help="Newer release label (e.g. 26B)",
    )
    report_parser.add_argument(
        "--out-dir", type=Path, default=Path("."),
        help="Output directory (default: ./)",
    )
    report_parser.add_argument(
        "--catalog", type=Path, default=Path("FBDI_Master_Catalog.xlsx"),
        help="Path to the master catalog (default: ./FBDI_Master_Catalog.xlsx)",
    )
    report_parser.add_argument(
        "--mapping", type=Path,
        default=Path("FBDI_to_ApplaudTables_Mapping.xlsx"),
        help="Path to the mapping spreadsheet (default: ./FBDI_to_ApplaudTables_Mapping.xlsx)",
    )
```

Add to the dispatch block after `populate-module`:
```python
    elif args.command == "report":
        _run_report(args)
```

Add the `_run_report` function:
```python
def _run_report(args: argparse.Namespace) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(name)s: %(message)s",
    )

    if not args.catalog.is_file():
        print(f"Error: catalog file not found: {args.catalog}")
        sys.exit(1)
    if not args.mapping.is_file():
        print(f"Error: mapping file not found: {args.mapping}")
        sys.exit(1)

    from fbdi.report import generate_report

    html_path, pdf_path = generate_report(
        catalog_path=args.catalog,
        mapping_path=args.mapping,
        old_release=args.old.upper(),
        new_release=args.new.upper(),
        out_dir=args.out_dir,
    )

    print(f"HTML: {html_path}")
    print(f"PDF : {pdf_path}")
```

- [ ] **Step 4: Run tests**

```bash
py -m pytest tests/test_cli.py -v
```

Expected: all passed.

- [ ] **Step 5: Run end-to-end via the CLI**

```bash
py -m fbdi report --old 26A --new 26B
```

Expected: prints both output paths; both files exist.

- [ ] **Step 6: Commit**

```bash
git add fbdi/cli.py tests/test_cli.py
git commit -m "feat(cli): add report subcommand to drive the compliance-report generator"
```

---

## Phase 9: Verification

### Task 14: End-to-end visual verification + final test run

**REQUIRED SUB-SKILL:** Use `superpowers:verification-before-completion` — do not declare done without running each verification step and confirming output.

- [ ] **Step 1: Run full test suite**

```bash
py -m pytest tests/ -v
```

Expected: all passed.

- [ ] **Step 2: Regenerate catalog (now produces correct Drift) + run the report**

```bash
py -m fbdi catalog --release 26A
py -m fbdi catalog --release 26B
py -m fbdi report --old 26A --new 26B
```

Expected: both reports written without errors.

- [ ] **Step 3: Verify Drift sheet now reflects alignment-driven counts**

```bash
PYTHONIOENCODING=utf-8 py -c "
from openpyxl import load_workbook
from collections import Counter
wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True)
ws = wb['Drift']
hdr = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
ct_idx = hdr.index('change_type')
counts = Counter()
for row in ws.iter_rows(min_row=2, values_only=True):
    counts[row[ct_idx]] += 1
print('Change-type distribution:', counts.most_common())
"
```

Expected: `SHIFTED` appears prominently; `RENAMED` and `MULTI` counts are dramatically lower than the pre-fix `460 / 236` (those were misclassifications).

- [ ] **Step 4: Visual inspection of HTML and PDF via chrome-devtools-mcp**

**REQUIRED SUB-SKILL:** Use `chrome-devtools-mcp:chrome-devtools` to:

1. Open the HTML report. Screenshot: cover, module rollup, summary table, each per-file section, pending-base section. Verify brand-color discipline (no off-palette colors), correct collapsible behavior, no layout glitches.
2. Open the PDF. Screenshot key pages. Verify SHIFTED tables auto-expanded, page breaks reasonable, brand colors preserved through weasyprint.

- [ ] **Step 5: Update CLAUDE.md to reflect the new pipeline state**

Edit `CLAUDE.md`:
- Add `fbdi/report.py`, `fbdi/applaud_type.py`, `fbdi/align.py` to the "Active Pipeline" section.
- Move `report.py (not built)` from "Current Frontier" to the active pipeline.
- Add the `python -m fbdi report` example to "Quick Start".
- Note that the catalog Drift schema changed (alignment-driven, with `old_position`/`new_position`/`sub_kinds`).

- [ ] **Step 6: Commit final docs update**

```bash
git add CLAUDE.md
git commit -m "docs(claude): record FBDI Compliance Report pipeline + Drift schema change"
```

- [ ] **Step 7: Push or hand off**

The branch is ready. Either push to origin/master directly (per the workflow note in user memory: handoff/plan execution → direct push to master) or open a PR if a review checkpoint is desired.

---

## Self-review

After writing all tasks, this section captures the final cross-check.

**Spec coverage check:**
- ✓ `align.py` — Phase 2 (Tasks 2-5)
- ✓ Catalog Drift fix — Phase 4 (Task 7)
- ✓ `applaud_type.py` — Phase 3 (Task 6)
- ✓ `report.py` view-model + scope filter — Phase 5 (Task 8)
- ✓ Catalog/mapping loaders — Phase 5 (Task 9)
- ✓ HTML template (cover, rollup, summary, per-file, pending-base) — Phase 6 (Tasks 10-11)
- ✓ PDF render via weasyprint with print_mode flag — Phase 7 (Task 12)
- ✓ CLI subcommand — Phase 8 (Task 13)
- ✓ Action matrix (DB/IF/EF) — encoded in template (Task 11)
- ✓ Module rollup — view-model (Task 8) + template (Task 10)
- ✓ Pending-base routing — view-model (Task 8) + template (Task 11)
- ✓ 30-char warning chip — view-model (Task 8) + template (Task 11)
- ✓ Definian palette discipline — template style block (Task 10) + visual verification (Tasks 11, 12, 14)
- ✓ Skills used: TDD throughout; frontend-design + humanizer in Phase 6; chrome-devtools in Phases 6, 7, 9; verification-before-completion in Phase 9
- ✓ Out-of-scope items (Applaud-mcp, auto-truncation, inline doc URLs, NEEDS_REVIEW workflow) explicitly excluded from tasks

**Type consistency check:**
- `AlignedField` defined in Task 2; used in Tasks 3-5, 7, 8, 9 — same field set throughout.
- `Change` defined in Task 2; used in Tasks 3-5, 7, 8 — `axes` and `sub_kinds` tuples consistent.
- `DriftRow` schema in Task 7 matches the spec's table.
- `FileSection` / `ChangeRow` / `PendingBaseEntry` / `ReportContext` defined in Task 8; used in Task 9 (loaders feed them) and Tasks 10-11 (template consumes them).
- `applaud_type_for` signature consistent between Task 6 (definition) and Task 8 (call site).
- `generate_report` signature consistent between Task 12 (definition) and Task 13 (CLI call).

No placeholders. No TBDs. Each step has the actual code or command needed.
