# Next Steps — Oracle FBDI Pulldown

Ranked by value delivered per effort. Phase 1 (comparison engine) is complete and passing 45 tests.

---

## 1. ~~Fix Large-File Skipping~~ — RESOLVED

**Resolution (Phase 2):** `compare.py` now pre-checks file size before loading. Files exceeding
`MAX_FILE_SIZE_BYTES` (5MB) are explicitly skipped with a warning log and an unavoidable
printed summary block in the CLI output. `compare_all()` returns a `(Path, list[dict])` tuple
so callers have structured access to the skipped-file list.

**Skipped files (both releases):** CseInstalledBaseAssetImport, CseWarrantyCoverage,
MntMaintenanceProgramImport, PayablesCollectionDocuments, PayablesStandardInvoiceImportTemplate,
WorkOrderResourceTransactionTemplate. These 6 files require manual review against each release.

---

## 2. ~~Fix the 7 Non-Standard Header Tabs~~ — RESOLVED

**Resolution (Phase 3):** Three targeted fixes eliminated all 14 `NO_HEADER` rows (7 pairs × 2 releases).
Post-fix diagnostic confirms: **0 NO_HEADER rows** across 1692 tab entries.

- `fill_ratio` now computed against actual populated column extent per row, not `ws.max_column` (fixes phantom-wide ImportAwards tabs)
- `MIN_CELLS` lowered from 3 → 2 and centralized in `config.py` (fixes 4 two-column header tabs)
- `"Messages"` added to `SKIP_TABS` (correctly excludes Oracle error-code lookup tab)

---

**Original diagnostic results (Phase 2):** Running `python -m fbdi diagnose --old baselines/25d --new baselines/26a`
across 421 files and 1692 tab entries identified 7 unique file/tab pairs with `NO_HEADER`:

| File | Tab | Best Score | Root Cause |
|---|---|---|---|
| ImportAwards | Award Projects | 0.225 | Header at row 4, but phantom max_col=772 → fill_ratio≈0.012 drags score below 0.35 threshold |
| ImportAwards | Award Personnel | 0.275 | Header at row 4, but phantom max_col=1598 → fill_ratio≈0.016 drags score below 0.35 threshold |
| ImportAwards | Award Assistance Listing Number | 0.000 | Header at row 4 with only 2 columns — below MIN_CELLS=3 |
| ImportGrantsPersonnel | Grants Personnel Keywords | 0.000 | Header at row 4 with only 2 columns — below MIN_CELLS=3 |
| LeaseContractManageTemplate | Assets | 0.000 | Header at row 3 with only 2 columns — below MIN_CELLS=3 |
| RevenueLeaseContractManageTemplate | Assets | 0.000 | Header at row 3 with only 2 columns — below MIN_CELLS=3 |
| PONNegotiationLinesImportTemplate | Messages | 0.000 | Not a data tab — contains error message key/value pairs. Correctly excluded. Add to SKIP_TABS. |

**Three root cause categories:**

**Category B — Phantom wide columns causing low fill_ratio (2 tabs):**
ImportAwards tabs have real headers at row 4, but `max_column` is 772–1598 (phantom Excel metadata).
`_scan_rows()` caps at 500, but fill_ratio = ~6 headers / 500 = 0.012 — far below threshold.
**Recommended fix:** Compute fill_ratio against actual non-empty column count in that row,
not `max_column`. This is a one-line change in `_scan_rows()`.

**Category A1 — Below MIN_CELLS threshold (4 tabs):**
4 tabs have real but narrow headers (exactly 2 columns). MIN_CELLS=3 gates them out before scoring.
**Recommended fix:** Lower MIN_CELLS from 3 to 2. These are valid 2-column import tabs.

**Category A2 — Not a data tab (1 tab):**
PONNegotiationLinesImportTemplate / Messages is a lookup table of Oracle error codes, not an
import field definition tab.
**Recommended fix:** Add "Messages" to SKIP_TABS in `config.py`.

**Combined impact of fixes:** All 7 unique NO_HEADER pairs would be resolved. Current error count
across both releases: 14 rows (7 pairs × 2 releases). Zero would remain after the above fixes.

---

## 3. Fix 5 Corrupt-Stylesheet Files (Low Value, Unknown Effort)

**Diagnostic results (Phase 2):** 5 files in both releases fail with `FILE_ERROR` — corrupt XML
stylesheet that openpyxl cannot parse:

- ImportAssignmentLaborSchedules
- ImportPayrollCosts
- ResourceBreakdownStructureImportTemplate
- SupplierSiteImportTemplate
- SusActivityImportTemplate

These currently produce no comparison output at all. Whether this is acceptable depends on
whether these templates have changed between releases. Investigate manually.

**Possible fix:** These may open with `read_only=True` mode despite the stylesheet error
(the Phase 1 engine already uses `read_only=True` and may handle them). Try loading with
`read_only=True` and catch the specific stylesheet exception to continue with reduced
fidelity rather than failing entirely.

---

## 4. Port the Downloader into the `fbdi/` Package (Medium Value, Medium Effort)

**The problem:** `reference/test.py` is Dan's original Selenium script and lives outside the package. Populating `baselines/` still requires running it manually and managing the output by hand.

**The fix:** Extract the download logic into `fbdi/download.py` with a CLI entry point (`python -m fbdi download --release 26b`). It should: accept a release name, create `baselines/<release>/`, and download all xlsm files into it.

**Why fourth:** Once this exists, you have a two-command workflow (download + compare). The full pipeline (`python -m fbdi run`) becomes a thin wrapper around both.

---

## 5. Complete the FBDI-to-Applaud Mapping (Manual — Brad)

**Status:** `fbdi_applaud_mapping.xlsx` has been generated at repo root. Run with `python -m fbdi.build_mapping`.

**What exists:** 610 rows — 9 YES (known Applaud mappings pre-populated), 11 problem rows (FILE_ERROR / FILE_TOO_LARGE flagged orange at top), 590 TBD rows awaiting manual review.

**What Brad needs to do:** Open `fbdi_applaud_mapping.xlsx` and fill in `applaud_table`, `prefix`, `module`, and `in_scope` for each TBD row that Definian integrations touch. Set `in_scope="NO"` for any FBDI tab that is not part of a Definian integration.

**Why this comes before report.py:** The completed mapping file becomes the config input for `report.py`. Without it, `report.py` cannot map comparison diffs to their Applaud target tables.

**After Brad completes it:** Commit the filled-in `fbdi_applaud_mapping.xlsx` to master. That file is then a versioned config artifact — update it after each release cycle as mappings are confirmed or changed.

---

## 6. Build the Audrey Report (`report.py`) — Phase 2

**The problem:** The comparison output (`Comparison_Report_*.xlsx`) is a raw diff. The Audrey change-tracking format used for client deliverables has a different structure. Currently this reformatting is done manually.

**The fix:** `fbdi/report.py` — reads the comparison output and the completed `fbdi_applaud_mapping.xlsx`, and produces a formatted Audrey-compatible Excel file. This is primarily a formatting/layout problem, not a data problem.

**Why sixth:** Depends on item 5 (completed mapping) and the comparison output being correct.

---

## 7. Full Pipeline Command — Phase 3

**The problem:** Three separate commands (download, compare, report) with manual handoffs between them.

**The fix:** `python -m fbdi run --old 25d --new 26b` — chains all three phases. Mostly orchestration logic once the components exist.

**Why last:** This is pure convenience. Build it after the components are solid.

---

## Not Recommended (Yet)

- **Parallelizing the comparison** — the 75-minute runtime is long but acceptable for a periodic task. Optimize only if it becomes a bottleneck.
- **Test coverage for `report.py`** — write tests after the format is stable, not before.
- **CI/CD** — not worth the setup until the pipeline is complete.
