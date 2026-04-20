# Catalog Subprocess Deadlock + Type Parser Cleanse — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Eliminate bogus TIMEOUT issues in `FBDI_Master_Catalog.xlsx` for `ChangeOrderImportTemplate` and `ItemImportTemplate`, and collapse the 463 `TYPE_PARSE_WARNING` rows down to the ~11 genuinely malformed ones — leaving the Issues tab readable and the 26A/26B snapshots complete.

**Architecture:** Two root-cause fixes delivered together.

1. **Subprocess queue drain-before-join.** Both `fbdi/catalog.py` and `fbdi/compare.py` run per-file work in a subprocess and shuttle results through a `multiprocessing.Queue`. Both use the pattern `proc.join(timeout); queue.get_nowait()`, which deadlocks whenever the pickled payload exceeds the OS pipe buffer (~64 KB on Windows). We extract a shared `run_worker()` helper in a new `fbdi/_subprocess_util.py` module that polls the queue first and joins second. Unit-testable in isolation, and wires both call sites to the same fix.
2. **Type parser tolerates temporal format masks and trailing periods.** 444 of 463 warnings are `DATE/TimeStamp(YYYY/MM/DD …)` format-mask shapes; 8 are `VARCHAR2(1 CHAR).` with a stray period. A second regex branch for `DATE|TIMESTAMP` format masks plus an optional trailing `\.?` on the main regex handles both, leaving only the ~11 genuinely broken Oracle strings as warnings.

**Tech Stack:** Python 3.14, `multiprocessing.Queue`, `openpyxl` (streaming reads), `pytest`. No new dependencies.

---

## File Structure

### Create

- **`fbdi/_subprocess_util.py`** — new module. Exposes `WorkerOutcome` dataclass and `run_worker(target, args, timeout)` helper. Centralises drain-before-join. Polls the queue and process liveness together so crash-before-put is reported as `"crashed"` rather than masquerading as a timeout.
- **`tests/test_subprocess_util.py`** — new test module. Exercises `run_worker()` directly with module-level worker targets: small payload, large payload (regression for this bug), timeout, crash, error-sentinel passthrough.

### Modify

- **`fbdi/catalog.py`** — `_run_file_in_subprocess` rewritten to delegate to `run_worker` and translate `WorkerOutcome` into catalog-specific `IssueRow`s. Delete the inline `multiprocessing.Process` plumbing.
- **`fbdi/compare.py`** — inline subprocess loop in `compare_all` rewritten to call `run_worker`. Delete the inline plumbing.
- **`fbdi/type_parser.py`** — extend `_TYPE_RE` to tolerate an optional trailing `.`; add `_TEMPORAL_FORMAT_RE` for `DATE|TIMESTAMP` format masks; update `parse_data_type` to fall through to the new regex when the strict one misses.
- **`tests/test_catalog.py`** — add regression test that builds a synthetic .xlsm with ~1500 columns and runs `_run_file_in_subprocess` end-to-end, asserting the large payload flows through without timeout.
- **`tests/test_compare.py`** — add regression test with a synthetic file pair large enough to exceed the pipe buffer.
- **`tests/test_type_parser.py`** — add positive tests for every new accepted shape (format masks + trailing period) and negative tests confirming genuinely-malformed strings still warn.

### No-op

- **`fbdi/diagnose.py`** / **`fbdi/build_mapping.py`** — don't use subprocess, not affected.
- **`fbdi/config.py`** — `CATALOG_TIMEOUT` stays at 120s; the fix is correctness, not a budget change.

---

## Task 1: Shared subprocess runner with drain-before-join

**Files:**
- Create: `fbdi/_subprocess_util.py`
- Test: `tests/test_subprocess_util.py`

- [ ] **Step 1.1: Write the failing tests**

Create `tests/test_subprocess_util.py`:

```python
"""Tests for fbdi._subprocess_util.run_worker."""

import time

from fbdi._subprocess_util import run_worker


# Worker targets must be module-level so Windows spawn can pickle them.

def _target_small(queue):
    queue.put(("ok", 42))


def _target_large(queue):
    # ~200 KB serialized — comfortably above the Windows pipe buffer (~64 KB)
    # and above the macOS pipe buffer (~16-64 KB depending on kernel).
    queue.put(("ok", ["x" * 100] * 2000))


def _target_slow(queue):
    time.sleep(10)
    queue.put(("ok", "too late"))


def _target_crash(queue):
    raise RuntimeError("worker died before put")


def _target_error_sentinel(queue):
    queue.put("ERROR: synthetic")


class TestRunWorker:
    def test_small_payload_returns_completed(self):
        out = run_worker(_target_small, args=(), timeout=30)
        assert out.status == "completed"
        assert out.payload == ("ok", 42)
        assert out.exitcode == 0

    def test_large_payload_no_deadlock(self):
        # Regression: pre-fix this hung for the full timeout because the
        # feeder thread could not drain the pipe while parent was joining.
        t0 = time.perf_counter()
        out = run_worker(_target_large, args=(), timeout=30)
        elapsed = time.perf_counter() - t0
        assert out.status == "completed"
        assert out.payload[0] == "ok"
        assert len(out.payload[1]) == 2000
        assert elapsed < 15, f"returned in {elapsed:.1f}s (possible deadlock)"

    def test_timeout_when_worker_hangs(self):
        out = run_worker(_target_slow, args=(), timeout=2)
        assert out.status == "timeout"
        assert out.payload is None

    def test_crash_before_put_reported_as_crashed(self):
        out = run_worker(_target_crash, args=(), timeout=30)
        assert out.status == "crashed"
        assert out.payload is None
        assert out.exitcode not in (0, None)

    def test_error_sentinel_is_passed_through(self):
        # Helper does not interpret "ERROR:" — caller owns that convention.
        out = run_worker(_target_error_sentinel, args=(), timeout=30)
        assert out.status == "completed"
        assert out.payload == "ERROR: synthetic"
```

- [ ] **Step 1.2: Run tests to verify they fail**

Run: `python -m pytest tests/test_subprocess_util.py -v`
Expected: `ModuleNotFoundError: No module named 'fbdi._subprocess_util'`

- [ ] **Step 1.3: Implement the helper**

Create `fbdi/_subprocess_util.py`:

```python
"""Shared subprocess runner for per-file isolation.

Both catalog.py and compare.py run per-file work in a fresh subprocess so that
openpyxl resource leaks don't accumulate across sequential loads. This module
centralises the result-handling pattern.

Drain-before-join is required. multiprocessing.Queue.put() hands the payload
to a background feeder thread which writes to an OS pipe (~64 KB buffer on
Windows). If the payload exceeds the pipe buffer and the parent joins first,
the feeder blocks, the child cannot exit, and join() times out. Reading from
the queue before joining unblocks the feeder.
"""

import multiprocessing
import queue as queue_module
import time
from dataclasses import dataclass
from typing import Any, Callable


@dataclass
class WorkerOutcome:
    """Outcome of one subprocess run.

    status:
      "completed" — worker put a message; payload holds it.
      "timeout"   — deadline reached while worker still alive.
      "crashed"   — worker exited without putting anything on the queue.
    """
    status: str
    payload: Any
    exitcode: int | None


def run_worker(
    target: Callable,
    args: tuple,
    timeout: int,
    poll_interval: float = 0.5,
) -> WorkerOutcome:
    """Run target(*args, queue) in a subprocess; drain queue before joining.

    The target must accept `queue` as its final positional argument and put
    exactly one message on it then return. The helper polls the queue and
    process liveness together so that a crash before put() is reported as
    "crashed" rather than "timeout", and a large payload cannot deadlock
    the feeder thread behind a join().
    """
    mp_queue: multiprocessing.Queue = multiprocessing.Queue()
    proc = multiprocessing.Process(target=target, args=args + (mp_queue,))
    proc.start()

    deadline = time.monotonic() + timeout
    payload: Any = None
    status: str | None = None

    while True:
        try:
            payload = mp_queue.get(timeout=poll_interval)
            status = "completed"
            break
        except queue_module.Empty:
            pass

        if not proc.is_alive():
            try:
                payload = mp_queue.get_nowait()
                status = "completed"
            except queue_module.Empty:
                status = "crashed"
            break

        if time.monotonic() >= deadline:
            proc.terminate()
            proc.join(5)
            return WorkerOutcome("timeout", None, proc.exitcode)

    proc.join(5)
    if proc.is_alive():
        proc.terminate()
        proc.join(5)

    return WorkerOutcome(status, payload, proc.exitcode)
```

- [ ] **Step 1.4: Run tests to verify they pass**

Run: `python -m pytest tests/test_subprocess_util.py -v`
Expected: 5 passed.

- [ ] **Step 1.5: Commit**

```bash
git add fbdi/_subprocess_util.py tests/test_subprocess_util.py
git commit -m "feat(fbdi): add subprocess runner with drain-before-join"
```

---

## Task 2: Wire catalog.py to the shared helper

**Files:**
- Modify: `fbdi/catalog.py` (remove `_run_file_in_subprocess` body, re-implement via `run_worker`)
- Modify: `tests/test_catalog.py` (add large-payload regression test)

- [ ] **Step 2.1: Write the failing regression test**

Append to `tests/test_catalog.py`:

```python
import time

from openpyxl import Workbook

from fbdi.catalog import _run_file_in_subprocess


class TestRunFileInSubprocessLargePayload:
    def test_large_file_does_not_deadlock(self, tmp_path):
        # Regression for the Windows pipe-buffer deadlock that caused
        # ChangeOrderImportTemplate and ItemImportTemplate to report
        # bogus TIMEOUTs in the 26A/26B catalog. Build a rich tab with
        # enough columns (~1500) that the pickled CatalogRow payload
        # comfortably exceeds the ~64 KB pipe buffer.
        wb = Workbook()
        ws = wb.active
        ws.title = "BIG_TAB"
        n = 1500
        # header_row = 5; metadata rows above it
        ws.cell(row=2, column=1, value="Name")
        ws.cell(row=3, column=1, value="Data Type")
        ws.cell(row=4, column=1, value="Required or Optional")
        ws.cell(row=5, column=1, value="Column name of the Table BIG_TAB")
        for i in range(1, n + 1):
            ws.cell(row=2, column=i + 1, value=f"Label {i}")
            ws.cell(row=3, column=i + 1, value="VARCHAR2(80)")
            ws.cell(row=4, column=i + 1, value="Required" if i % 2 else "Optional")
            ws.cell(row=5, column=i + 1, value=f"COL_{i:04d}")
        path = tmp_path / "BigTemplate.xlsm"
        wb.save(path)

        t0 = time.perf_counter()
        rows, issues = _run_file_in_subprocess(path, release="26A", timeout=30)
        elapsed = time.perf_counter() - t0

        assert issues == []
        assert len(rows) == n
        assert rows[0].column_technical == "COL_0001"
        assert rows[-1].column_technical == f"COL_{n:04d}"
        assert elapsed < 20, f"returned in {elapsed:.1f}s (possible deadlock)"
```

- [ ] **Step 2.2: Run the test to verify it fails**

Run: `python -m pytest tests/test_catalog.py::TestRunFileInSubprocessLargePayload -v`
Expected: either a TIMEOUT issue (elapsed > 30s, rows == 0) or the new test file depends on `run_worker` wiring in catalog.py — either way, FAIL. If the legacy `queue.get_nowait()` path is still in place the test will show `issues == [IssueRow(... TIMEOUT ...)]`.

- [ ] **Step 2.3: Replace `_run_file_in_subprocess` with the helper call**

In `fbdi/catalog.py`:

Add the import near the top with the other `from fbdi.` imports:

```python
from fbdi._subprocess_util import run_worker
```

Replace the body of `_run_file_in_subprocess` (currently lines 356-400) with:

```python
def _run_file_in_subprocess(
    path: Path, release: str, timeout: int = CATALOG_TIMEOUT
) -> tuple[list[CatalogRow], list[IssueRow]]:
    """Run extract_file in a fresh subprocess with timeout.

    Translates WorkerOutcome into catalog-specific IssueRows on failure
    paths (TIMEOUT, SUBPROCESS_FAILED). Happy path unpacks the
    (row_tuples, issue_tuples) payload.
    """
    outcome = run_worker(_catalog_worker, args=(str(path), release), timeout=timeout)
    file_stem = path.stem

    if outcome.status == "timeout":
        return [], [IssueRow(
            release=release, file=file_stem, tab="",
            issue_type="TIMEOUT", detail=f"exceeded {timeout}s",
        )]

    if outcome.status == "crashed":
        return [], [IssueRow(
            release=release, file=file_stem, tab="",
            issue_type="SUBPROCESS_FAILED",
            detail=f"exit code {outcome.exitcode}",
        )]

    result = outcome.payload
    if isinstance(result, str) and result.startswith("ERROR:"):
        return [], [IssueRow(
            release=release, file=file_stem, tab="",
            issue_type="SUBPROCESS_FAILED",
            detail=result,
        )]

    row_tuples, issue_tuples = result
    return _tuples_to_rows(row_tuples), _tuples_to_issues(issue_tuples)
```

Also: remove the now-unused `import multiprocessing` from the top of `fbdi/catalog.py` if no other symbols in the file reference it. Verify with `grep -n "multiprocessing" fbdi/catalog.py` — if only the `_catalog_worker` signature still references `multiprocessing.Queue`, keep the import. (It does — `_catalog_worker(path_str: str, release: str, queue: multiprocessing.Queue) -> None` on the existing line in the file.)

- [ ] **Step 2.4: Run the full catalog test module**

Run: `python -m pytest tests/test_catalog.py -v`
Expected: all existing tests pass + the new regression test passes. Total should be previous count + 1.

- [ ] **Step 2.5: Commit**

```bash
git add fbdi/catalog.py tests/test_catalog.py
git commit -m "fix(catalog): drain subprocess queue before join to avoid pipe-buffer deadlock"
```

---

## Task 3: Wire compare.py to the shared helper

**Files:**
- Modify: `fbdi/compare.py` (inline subprocess block in `compare_all` → helper call)
- Modify: `tests/test_compare.py` (large-payload regression test)

- [ ] **Step 3.1: Write the failing regression test**

Append to `tests/test_compare.py`:

```python
import time

from openpyxl import Workbook

from fbdi.compare import compare_all


class TestCompareAllLargePayload:
    def test_large_pair_does_not_deadlock(self, tmp_path):
        # Regression for the same pipe-buffer deadlock as catalog.py.
        # Build an old/new pair with a wide tab so the serialized
        # ComparisonRow payload exceeds the OS pipe buffer.
        old_dir = tmp_path / "old"
        new_dir = tmp_path / "new"
        old_dir.mkdir()
        new_dir.mkdir()
        n = 2500

        def _build(path, tech_prefix):
            wb = Workbook()
            ws = wb.active
            ws.title = "WIDE_TAB"
            ws.cell(row=5, column=1, value="Column name of the Table WIDE_TAB")
            for i in range(1, n + 1):
                ws.cell(row=5, column=i + 1, value=f"{tech_prefix}_{i:04d}")
            wb.save(path)

        _build(old_dir / "BigTemplate.xlsm", "OLD")
        _build(new_dir / "BigTemplate.xlsm", "NEW")

        output = tmp_path / "Comparison_Report.xlsx"
        t0 = time.perf_counter()
        out_path, timed_out = compare_all(old_dir, new_dir, output, timeout=30)
        elapsed = time.perf_counter() - t0

        assert timed_out == []
        assert out_path == output
        assert output.exists()
        assert elapsed < 25, f"compare_all took {elapsed:.1f}s (possible deadlock)"
```

- [ ] **Step 3.2: Run the test to verify it fails**

Run: `python -m pytest tests/test_compare.py::TestCompareAllLargePayload -v`
Expected: test times out near 30s and reports `timed_out == ['BigTemplate']`, causing assertion failure.

- [ ] **Step 3.3: Replace the inline subprocess block with the helper**

In `fbdi/compare.py`:

Add near the other `from fbdi.` imports:

```python
from fbdi._subprocess_util import run_worker
```

Inside `compare_all`, replace the loop body from `queue: multiprocessing.Queue = multiprocessing.Queue()` through the `for row_tuple in result:` block (currently lines 217-251) with:

```python
    for i, (old_path, new_path) in enumerate(matched, 1):
        logger.info("[%d/%d] Comparing: %s", i, len(matched), old_path.stem)

        outcome = run_worker(
            _compare_worker,
            args=(str(old_path), str(new_path)),
            timeout=timeout,
        )

        if outcome.status == "timeout":
            timed_out.append(old_path.stem)
            logger.warning(
                "TIMEOUT after %ds: %s — skipped", timeout, old_path.stem,
            )
            continue

        if outcome.status == "crashed":
            logger.error(
                "Subprocess failed (exit %s): %s",
                outcome.exitcode, old_path.stem,
            )
            continue

        result = outcome.payload
        if isinstance(result, str) and result.startswith("ERROR:"):
            logger.error("Compare error for %s: %s", old_path.stem, result)
            continue

        for row_tuple in result:
            all_rows.append(ComparisonRow(*row_tuple))
```

Remove the `import multiprocessing` line at the top of `fbdi/compare.py` — it is no longer referenced. Double-check with `grep -n "multiprocessing" fbdi/compare.py`. If `_compare_worker(old_path: str, new_path: str, queue: multiprocessing.Queue)` still references it, keep the import. It does — keep it.

- [ ] **Step 3.4: Run the full compare test module**

Run: `python -m pytest tests/test_compare.py -v`
Expected: all existing tests pass + the new regression test passes.

- [ ] **Step 3.5: Commit**

```bash
git add fbdi/compare.py tests/test_compare.py
git commit -m "fix(compare): drain subprocess queue before join to avoid pipe-buffer deadlock"
```

---

## Task 4: Extend type_parser for temporal format masks and trailing periods

**Files:**
- Modify: `fbdi/type_parser.py`
- Modify: `tests/test_type_parser.py`

The distinct unparseable strings across 26A+26B catalogs (463 rows total) break down as follows:

| Count | Example | Category |
|------:|---|---|
| 212 | `DATE(YYYY/MM/DD)` | temporal format mask |
| 118 | `DATE (YYYY/MM/DD)` | temporal format mask (space before paren) |
| 40 | `Date((yyyy/mm/dd hh24:mm)` | temporal, stray leading `(` |
| 34 | `Date(YYYY/MM/DD)` | temporal format mask |
| 18 | `Date (YYYY/MM/DD)` | temporal format mask |
| 10 | `TimeStamp(yyyy/mm/dd hh24:mm)` | temporal format mask |
| 8 | `VARCHAR2(1 CHAR).` | trailing period |
| 6 | `TimeStamp(hh24:mm)` | temporal format mask |
| 4 | `TimeStamp(yyyy/mm/dd hh24:mm:ss)` | temporal format mask |
| 2 | `Date(yyyy/mm/dd)` | temporal format mask |
| 2 | `Date(yyyy/mm/dd hh24:mm)` | temporal format mask |
| 2 | `For desc Asset it is mandatory for create` | genuinely malformed |
| 2 | `(VARCHAR2(150)` | genuinely malformed (unbalanced paren) |
| 2 | `VARCHAR2(18R)` | genuinely malformed (non-digit length) |
| 2 | `varchar2(4` | genuinely malformed (no closing paren) |
| 1 | `Item Number` | genuinely malformed |

Target: 452 cases fixed (all temporal + trailing period), ~11 remaining as genuine Oracle-source quality issues that should stay flagged.

- [ ] **Step 4.1: Write the failing tests**

Append to `tests/test_type_parser.py`:

```python
class TestTrailingPeriod:
    def test_varchar2_with_trailing_period(self):
        result = parse_data_type("VARCHAR2(1 CHAR).")
        assert result == ParsedType("VARCHAR2", 1, None, False)

    def test_number_with_scale_and_trailing_period(self):
        result = parse_data_type("NUMBER(18,4).")
        assert result == ParsedType("NUMBER", 18, 4, False)

    def test_date_with_trailing_period(self):
        result = parse_data_type("DATE.")
        assert result == ParsedType("DATE", None, None, False)


class TestTemporalFormatMask:
    def test_date_upper_slash_format(self):
        result = parse_data_type("DATE(YYYY/MM/DD)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_upper_space_then_format(self):
        result = parse_data_type("DATE (YYYY/MM/DD)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_mixed_case_format(self):
        result = parse_data_type("Date(YYYY/MM/DD)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_lower_format(self):
        result = parse_data_type("Date(yyyy/mm/dd)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_datetime_format(self):
        result = parse_data_type("Date(yyyy/mm/dd hh24:mm)")
        assert result == ParsedType("DATE", None, None, False)

    def test_date_stray_leading_paren(self):
        # 40-row Oracle typo: 'Date((yyyy/mm/dd hh24:mm)'
        result = parse_data_type("Date((yyyy/mm/dd hh24:mm)")
        assert result == ParsedType("DATE", None, None, False)

    def test_timestamp_time_only(self):
        result = parse_data_type("TimeStamp(hh24:mm)")
        assert result == ParsedType("TIMESTAMP", None, None, False)

    def test_timestamp_full_datetime(self):
        result = parse_data_type("TimeStamp(yyyy/mm/dd hh24:mm:ss)")
        assert result == ParsedType("TIMESTAMP", None, None, False)


class TestStillWarnsOnTrulyMalformed:
    # These 11-or-so Oracle strings are genuinely broken. The parser should
    # NOT silently swallow them — Brad wants to see real data-quality issues.
    def test_unbalanced_leading_paren_warns(self):
        result = parse_data_type("(VARCHAR2(150)")
        assert result.parse_warning is True

    def test_missing_closing_paren_warns(self):
        result = parse_data_type("varchar2(4")
        assert result.parse_warning is True

    def test_non_digit_length_warns(self):
        result = parse_data_type("VARCHAR2(18R)")
        assert result.parse_warning is True

    def test_sentence_warns(self):
        result = parse_data_type("For desc Asset it is mandatory for create")
        assert result.parse_warning is True

    def test_label_without_paren_warns(self):
        # Single-word tokens that are not recognized types (VARCHAR2 etc.)
        # with no parens cannot be distinguished from type names like DATE
        # by the strict regex, so they match as a bare type name. That's
        # the correct behavior — the only way to know "Item Number" is
        # wrong is that it has a space and no parens. Current strict regex
        # requires [A-Za-z0-9]* after the leading letter, so "Item Number"
        # fails due to the space and sets the warning flag.
        result = parse_data_type("Item Number")
        assert result.parse_warning is True
```

- [ ] **Step 4.2: Run tests to verify they fail**

Run: `python -m pytest tests/test_type_parser.py::TestTrailingPeriod tests/test_type_parser.py::TestTemporalFormatMask -v`
Expected: all new tests in `TestTrailingPeriod` and `TestTemporalFormatMask` FAIL (parse_warning=True for inputs that should now parse cleanly). The `TestStillWarnsOnTrulyMalformed` tests should already pass under the current regex.

- [ ] **Step 4.3: Extend the parser**

Replace the contents of `fbdi/type_parser.py` with:

```python
"""Parse Oracle data-type strings from FBDI templates into structured fields.

FBDI templates store types in a 'Data Type' row as strings like:
  VARCHAR2(5 CHAR), VARCHAR2(2048 CHAR), VARCHAR2(80), Varchar2(250),
  NUMBER(18), NUMBER(18,4), DATE, CLOB, BLOB

Some Oracle templates also ship format-mask variants for temporal types:
  DATE(YYYY/MM/DD), Date (yyyy/mm/dd hh24:mm), TimeStamp(hh24:mm:ss)

And a handful ship with a stray trailing period: VARCHAR2(1 CHAR).

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


# Strict shape — the supported Oracle type forms with optional length/scale:
#   TYPENAME
#   TYPENAME(length)
#   TYPENAME(length CHAR|BYTE)
#   TYPENAME(length,scale)
# Optional trailing period tolerated (Oracle ships `VARCHAR2(1 CHAR).` in a
# handful of templates).
_TYPE_RE = re.compile(
    r"^\s*"
    r"([A-Za-z][A-Za-z0-9]*)"              # 1: type name
    r"\s*"
    r"(?:"
        r"\(\s*"
        r"(\d+)"                           # 2: length / precision
        r"(?:\s*,\s*(\d+))?"               # 3: optional scale
        r"(?:\s+(?:CHAR|BYTE))?"           # optional CHAR|BYTE suffix
        r"\s*\)"
    r")?"
    r"\s*\.?\s*$",
    re.IGNORECASE,
)


# Temporal format-mask shape — DATE and TIMESTAMP only:
#   DATE(YYYY/MM/DD), Date (yyyy/mm/dd), TimeStamp(hh24:mm:ss)
# Also tolerates Oracle's stray leading paren seen in ChangeOrderImportTemplate
# (`Date((yyyy/mm/dd hh24:mm)`). length/scale remain None — a format mask is
# not a SQL length. Constrained to DATE/TIMESTAMP so it cannot rescue
# genuinely-broken strings like `VARCHAR2(18R)`.
_TEMPORAL_FORMAT_RE = re.compile(
    r"^\s*"
    r"(DATE|TIMESTAMP)"                    # 1: type name
    r"\s*"
    r"\(+\s*"                              # one or more `(` (typo tolerance)
    r"[A-Za-z0-9/:\-\s]+"                  # format mask chars
    r"\s*\)\s*\.?\s*$",
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

    s = str(raw)

    m = _TYPE_RE.match(s)
    if m:
        dtype = m.group(1).upper()
        length = int(m.group(2)) if m.group(2) else None
        scale = int(m.group(3)) if m.group(3) else None
        return ParsedType(dtype, length, scale, False)

    m = _TEMPORAL_FORMAT_RE.match(s)
    if m:
        return ParsedType(m.group(1).upper(), None, None, False)

    return ParsedType("", None, None, True)
```

- [ ] **Step 4.4: Run tests to verify they pass**

Run: `python -m pytest tests/test_type_parser.py -v`
Expected: all existing `TestParseDataType` tests still pass + new `TestTrailingPeriod`, `TestTemporalFormatMask`, `TestStillWarnsOnTrulyMalformed` tests all pass.

- [ ] **Step 4.5: Commit**

```bash
git add fbdi/type_parser.py tests/test_type_parser.py
git commit -m "feat(type_parser): accept temporal format masks and trailing periods"
```

---

## Task 5: Full regression pass + catalog regeneration + verification

This task has no code changes — it verifies the three fixes together against the real 26A and 26B templates and updates the catalog artifact.

- [ ] **Step 5.1: Run the full test suite**

Run: `python -m pytest tests/ -v`
Expected: all tests pass. New count should be prior total (116) + ~20 new tests from Tasks 1–4.

- [ ] **Step 5.2: Regenerate the 26A snapshot**

Run: `python -m fbdi catalog --release 26A`
Expected: process completes without "TIMEOUT" log messages. Log lines `[N/total] Cataloging: ChangeOrderImportTemplate` and `ItemImportTemplate` now complete in seconds each. Final line logs `... N releases, M rows, K issues, J drift`.

- [ ] **Step 5.3: Regenerate the 26B snapshot**

Run: `python -m fbdi catalog --release 26B`
Expected: same as 5.2 for 26B.

- [ ] **Step 5.4: Verify Issues tab shrunk**

Run:

```bash
python -c "
from openpyxl import load_workbook
from collections import Counter
wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True, data_only=True)
ws = wb['Issues']
c = Counter()
for row in ws.iter_rows(min_row=2, values_only=True):
    if any(v is not None for v in row):
        c[row[3]] += 1
print('Issues by type:', dict(c))
wb.close()
"
```

Expected output (approximate):
- `TIMEOUT: 0`
- `SUBPROCESS_FAILED: 0`
- `TYPE_PARSE_WARNING: ~11` (down from 463; the genuinely-malformed Oracle strings)
- `NO_HEADER: 0`
- `FILE_ERROR: 0`

- [ ] **Step 5.5: Verify ChangeOrder and ItemImport rows appear in both snapshots**

Run:

```bash
python -c "
from openpyxl import load_workbook
wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True, data_only=True)
for release in ['26A', '26B']:
    ws = wb[release]
    counts = {'ChangeOrderImportTemplate': 0, 'ItemImportTemplate': 0}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] in counts:
            counts[row[1]] += 1
    print(f'{release}: {counts}')
wb.close()
"
```

Expected: non-zero counts for both files in both releases (around 1,400 rows each — previously 0).

- [ ] **Step 5.6: Inspect the remaining TYPE_PARSE_WARNING rows**

Run:

```bash
python -c "
from openpyxl import load_workbook
from collections import Counter
wb = load_workbook('FBDI_Master_Catalog.xlsx', read_only=True, data_only=True)
ws = wb['Issues']
c = Counter()
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[3] == 'TYPE_PARSE_WARNING':
        c[row[4]] += 1
print('Remaining TYPE_PARSE_WARNING strings:')
for s, n in c.most_common():
    print(f'  {n:4d}  {s!r}')
wb.close()
"
```

Expected: only the genuinely-malformed strings (`(VARCHAR2(150)`, `varchar2(4`, `VARCHAR2(18R)`, `For desc Asset ...`, `Item Number`). If any temporal format-mask strings appear, Task 4's regex missed a case — extend `_TEMPORAL_FORMAT_RE` and add a targeted test before proceeding.

- [ ] **Step 5.7: Commit the regenerated catalog**

```bash
git add FBDI_Master_Catalog.xlsx
git commit -m "chore(catalog): regenerate 26A/26B after subprocess and type-parser fixes"
```

- [ ] **Step 5.8: Push to master**

```bash
git push origin master
```

(Per Brad's workflow — direct push to master, not PR. Confirmed in feedback memory.)

---

## Self-Review

**Spec coverage:**
- TIMEOUT fix for ChangeOrder / ItemImport → Tasks 1 + 2 + verification in 5.5 ✓
- Latent bug in compare.py → Task 3 ✓
- TYPE_PARSE_WARNING regex fix → Task 4 ✓
- Verification of final Issues count reduction → Step 5.4 + 5.6 ✓
- Catalog artifact updated in repo → Step 5.7 ✓
- Testing discipline (TDD red → green → commit per task) → all tasks follow this cycle ✓

**Placeholder scan:** No "TBD", "implement later", "similar to Task N", or "add appropriate X" — every step has complete code or exact commands with expected output.

**Type consistency:**
- `WorkerOutcome` has fields `status: str`, `payload: Any`, `exitcode: int | None` — used consistently in Tasks 1, 2, 3.
- `run_worker(target, args, timeout, poll_interval=0.5)` signature consistent across Task 1 (definition), Task 2 (catalog call with 3 args), Task 3 (compare call with 3 args).
- `status` values `"completed"`, `"timeout"`, `"crashed"` — catalog.py translates `"timeout"` → TIMEOUT IssueRow, `"crashed"` → SUBPROCESS_FAILED IssueRow; compare.py logs the same outcomes. Consistent.
- `ParsedType(data_type, length, scale, parse_warning)` field order unchanged — existing tests keep passing.

All consistent. Proceeding.
