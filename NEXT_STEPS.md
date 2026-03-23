# Next Steps — Oracle FBDI Pulldown

Ranked by value delivered per effort. Phase 1 (comparison engine) is complete and passing 33 tests.

---

## 1. Fix Large-File Skipping (High Value, Low Effort)

**The problem:** 6 files over 5MB are currently skipped during comparison due to openpyxl load time. These files likely contain real changes that aren't appearing in the report.

**The fix:** Switch those files to `read_only=True` mode in openpyxl, which streams rows instead of loading the full workbook into memory. This is a targeted change in `compare.py` — estimated 1–2 hours.

**Why first:** The current report has known gaps. Closing them makes the output trustworthy before building anything on top of it.

---

## 2. Fix the ~8 Non-Standard Header Tabs (Medium Value, Medium Effort)

**The problem:** ~8 tabs fail dynamic header detection and are excluded from the comparison. The scoring heuristic doesn't handle these edge cases.

**The fix:** Run the detection failure list, inspect those tabs manually, and either (a) extend the scoring heuristic or (b) add a small override map for the specific files/tabs that fail. Option (b) is faster but less principled.

**Why second:** Once large files and header detection are fixed, the comparison output is complete. Everything downstream (Audrey report, pipeline) is built on top of a complete comparison.

---

## 3. Port the Downloader into the `fbdi/` Package (Medium Value, Medium Effort)

**The problem:** `reference/test.py` is Dan's original Selenium script and lives outside the package. Populating `baselines/` still requires running it manually and managing the output by hand.

**The fix:** Extract the download logic into `fbdi/download.py` with a CLI entry point (`python -m fbdi download --release 26b`). It should: accept a release name, create `baselines/<release>/`, and download all xlsm files into it.

**Why third:** Once this exists, you have a two-command workflow (download + compare). The full pipeline (`python -m fbdi run`) becomes a thin wrapper around both.

---

## 4. Build the Audrey Report (`report.py`) — Phase 2

**The problem:** The comparison output (`Comparison_Report_*.xlsx`) is a raw diff. The Audrey change-tracking format used for client deliverables has a different structure. Currently this reformatting is done manually.

**The fix:** `fbdi/report.py` — reads the comparison output and produces a formatted Audrey-compatible Excel file. This is primarily a formatting/layout problem, not a data problem.

**Why fourth:** This is high client-facing value but depends on the comparison output being complete and correct (items 1 and 2 first).

---

## 5. Full Pipeline Command — Phase 3

**The problem:** Three separate commands (download, compare, report) with manual handoffs between them.

**The fix:** `python -m fbdi run --old 25d --new 26b` — chains all three phases. Mostly orchestration logic once the components exist.

**Why last:** This is pure convenience. Build it after the components are solid.

---

## Not Recommended (Yet)

- **Parallelizing the comparison** — the 75-minute runtime is long but acceptable for a periodic task. Optimize only if it becomes a bottleneck.
- **Test coverage for `report.py`** — write tests after the format is stable, not before.
- **CI/CD** — not worth the setup until the pipeline is complete.
