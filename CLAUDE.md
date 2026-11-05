# CLAUDE.md — Oracle FBDI Pulldown

This file gives Claude Code persistent context for this project. Read it at the start of every session.

**For human readers:** [`docs/operator-guide.md`](docs/operator-guide.md) walks through running the quarterly refresh; [`docs/developer-guide.md`](docs/developer-guide.md) orients developers to the codebase.

---

## Project Purpose

This repo automates comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) between Oracle Cloud release versions. The goal is to identify field-level changes (added, removed, modified columns) across releases so that Definian's Oracle integration team can keep client integrations current. Primary deliverables are a structured Excel comparison report, a per-release master catalog, and an HTML/PDF release change report.

---

## Quick Start

```bash
# Compare two releases (uses baselines/<ver>/originals/)
py -m fbdi compare --old 26A --new 26B --output Comparison_Report_26A_26B.xlsx

# Generate/update the FBDI master catalog for a release
py -m fbdi catalog --release 26B

# Diagnose header detection across releases
py -m fbdi diagnose --old baselines/26A/originals --new baselines/26B/originals --output Diagnostic_Report_26A_26B.xlsx

# Download + smart-clear FBDI templates for a new release
py tools/download_and_clear.py 26B                 # download + clear (exit 1 on any shortfall)
py tools/download_and_clear.py 26B --clear-only    # re-clear only (skip download)
py tools/download_and_clear.py 26B --discover-only # list what Oracle publishes vs originals/; writes nothing

# Verify blanks against originals (package survey + Dashboard paste anchor)
py tools/verify_blanks.py 26B

# Generate the FBDI Release Change Report (HTML default; --pdf adds the PDF)
py -m fbdi report --old 26B --new 26C               # writes FBDI_Change_Report_26B_26C.html
py -m fbdi report --old 26B --new 26C --pdf         # also writes FBDI_Change_Report_26B_26C.pdf

# Tests
py -m pytest tests/            # full suite
py -m pytest tests/test_clear.py -v
```

**Requirements:** Python 3.14+. Install deps via `pip install -r requirements.txt` (openpyxl, requests, pytest, jinja2, weasyprint, oletools; no browser — the downloader is plain HTTP). oletools reads VBA for the RapidImpl blank tests and `tools/build_vba_patch.py`, which also needs pywin32 + Excel. weasyprint is only needed for the opt-in `--pdf` path and requires MSYS2 mingw64 GTK (Pango ≥1.44) on Windows. On Brad's Windows setup, use the `py` launcher (`py -m fbdi report ...`, `py -m pytest ...`); bare `python` invokes the Microsoft Store launcher stub.

---

## Active Pipeline (Built and Working)

- **`fbdi/` package** — Python comparison engine
  - `detect_header.py` — dynamically identifies the header row in each FBDI tab using content scoring (no hardcoded filename map). Uses `iter_rows` for streaming scans.
  - `compare.py` — diffs two releases tab-by-tab, field-by-field. Each pair runs in a fresh subprocess (via `_subprocess_util.run_worker`) with a 120s timeout to isolate openpyxl resource leaks.
  - `clear.py` — smart clearing of FBDI templates using `detect_header_row` (preserves headers at any row — 4, 5, 8, etc.). Patches the original package byte-level via `ooxml.Package` (openpyxl is read-only, for detection only). Only the cleared sheets' `<sheetData>` cells change, so ActiveX buttons, x14 dropdowns, drawings and VBA ship as Oracle published them. Clears visible data sheets only: hidden LOV/reference sheets and `SKIP_CLEAR` settings sheets (ProjectImportTemplate::Options) are left alone. Drops `calcChain.xml` when a cleared cell held a formula.
  - `ooxml.py` — minimal byte-level OOXML package editor (`Package`: read/replace/delete parts, resolve worksheets, drop calcChain, atomic save) plus the shared sheet/row/cell regexes.
  - `blank_patches.py` — filename-keyed registry of post-clear fixes, applied by `clear_workbook` on every refresh. RapidImpl gets the Bank Account A7 "Dashboard paste anchor" (the Dashboard pastes at last-non-empty-col-A + 1) and a `vbaProject.bin` swap whose Generate macro saves `.xml`, not CSV (Oracle hard-coded `.FilterIndex = 5`). The swap is guarded by the original's sha256. The asset is `fbdi/assets/RapidImplementationForCashManagement.vbaProject.bin`, rebuilt by `tools/build_vba_patch.py`. A mismatch raises `BlankPatchError`, so that file fails rather than shipping unfixed.
  - `diagnose.py` — reports header-detection outcomes per tab (`DETECTED`, `NO_HEADER`, `SKIPPED_TAB`, `FILE_TOO_LARGE`, `FILE_ERROR`). Uses full (non-read_only) openpyxl mode.
  - `catalog.py` — generates `FBDI_Master_Catalog.xlsx` with per-release snapshots (file × tab × position × label × technical × type × length × scale × required) + `Issues` + `Drift` tabs. Shares `_subprocess_util.run_worker` with `compare.py`. The `Drift` sheet uses alignment-driven classification (`align.align_tabs`) — schema columns: `file`, `tab`, `change_type`, `old_position`, `new_position`, `old_label`, `new_label`, `old_technical`, `new_technical`, `old_data_type`, `new_data_type`, `old_length`, `new_length`, `old_scale`, `new_scale`, `old_required`, `new_required`, `sub_kinds`. Also mines Oracle's per-column cell-comment metadata (`TECHNICAL_NAME / VARCHAR2(N) / Description`) as a fallback via `_extract_metadata_from_comment` — anchored on `_ORACLE_TYPE_ALLOWLIST` to avoid false-positives on bare alpha tokens like "Required". **Row-label gutters:** many tabs stack `Name / Description / Data Type / Technical Name` (or `Column name of the Table X`) down column A with fields from column B (Oracle's Generate-CSV macro copies from B). `_find_gutter` detects this, and `_extract_rich` then reads fields from column B and picks up the metadata rows wherever they sit, above or below the header. A header-row `*` supplies `required` only when there is no Required row. A `*` on a separate label row is deliberately ignored: on the Scp* templates it marks a key column, not a required one.
  - `type_parser.py` — parses Oracle data-type strings (`VARCHAR2(N CHAR)`, `NUMBER(p,s)`, `DATE`, `DATE(YYYY/MM/DD)`, `TimeStamp(hh24:mm)`, trailing-period variants) into structured fields. Emits `TYPE_PARSE_WARNING` only for genuinely malformed strings.
  - `_subprocess_util.py` — shared `run_worker(target, args, timeout)` helper used by `catalog.py` and `compare.py`. Drains the result queue *before* joining the child process — required to avoid a pipe-buffer deadlock on Windows when payloads exceed ~64 KB.
  - `catalog_normalize.py` — normalizes FBDI column labels (strips punctuation/symbols, keeps alphanumerics/underscores/whitespace, collapses whitespace) to produce the catalog's `column_label`. Technical UPPER_SNAKE_CASE names are left untouched.
  - `align.py` — pure LCS-style alignment algorithm. `align_tabs(old_rows, new_rows) -> list[Change]`. Classifies changes across three axes (label, metadata, position); SHIFTED/RENAMED/MODIFIED/ADDED/REMOVED/MULTI. Shared by `catalog.py` (Drift writer) and `report.py`.
  - `report.py` — FBDI Release Change Report generator. `generate_report(catalog_path, old_release, new_release, out_dir, pdf=False) -> (html_path, pdf_path|None)`. Reads the master catalog's per-release sheets, aligns each `(file, tab)` across the two releases via `align_tabs`, groups the changed tabs by module from `baselines/<release>/file_modules.json` (NEW wins over OLD; files with no entry group under "Unclassified"), and renders a consultant-facing report from `fbdi/templates/report.html.j2` (+ `_report_macros.j2`). HTML is always written; PDF is opt-in via `pdf=True` (weasyprint imported lazily only then). No mapping workbook.
    - **Design system (Definian, Tier B):** the template implements `DESIGN.md`/`.impeccable/design.json` — navy cover, module→file/tab structure, color-coded change tables, sticky search/filter toolbar + collapsible sections (vanilla JS, HTML only). `generate_report` inlines a **vendored self-contained CSS** (`fbdi/templates/assets/definian.css` — Definian tokens + 5 base64 Aptos faces, ~1.3 MB) via the `design_css` template var, so the HTML is one portable file (no external CSS/font fetch, deterministic in weasyprint). Unit `_render_report` tests omit `design_css` and stay fast. Regenerate the asset with `py tools/build_report_font_css.py` (reads the design-system `fonts/*.ttf`; run only when the face set changes). Semantic accents are Definian status tokens (MODIFIED `#E5A100`, REMOVED `#D64545`); badges use AA-safe tint+dark-text, not white-on-fill.
  - `cli.py` / `__main__.py` — CLI entry point (`compare`, `diagnose`, `catalog`, `report`). `_resolve_dir()` makes `--old 26A` resolve to `baselines/26A/originals/`.
  - `config.py`, `utils.py` — shared configuration and helpers (`utils.quiet_openpyxl()` silences openpyxl's read-time "will be removed" UserWarnings; compare, catalog, diagnose and clear only read, so they are noise).
- **`tools/download_and_clear.py`** — standalone downloader + smart clearing entry point. It lives outside `fbdi/` so network I/O stays out of the comparison engine.
  - **Discovery is static HTML, no browser.** Each docs book's `toc.htm` (`BOOK_TOC_TEMPLATES`, in PPM → Fin → Proc → SCM order, where the first book wins a file's module) lists its topic pages, and each topic page links its `.xlsm`.
  - **It fails loudly.** Discovery runs first, and any TOC/page fetch error exits 1 *before* anything is deleted or `file_modules.json` is rewritten. After that it exits 1 on a failed or non-zip download, a file that vanished from disk after download (re-checked after a settle delay and again after the clear), or a clear failure. It names each missing file with a recovery URL. `MANUAL_FILES` (RapidImpl) survive the `originals/` wipe.
  - **Module map.** It writes `baselines/<VER>/file_modules.json` from discovery.
  - **Flags.** `--discover-only` lists what Oracle publishes and diffs it against `originals/` without writing anything. `--clear-only` skips blanks that already exist; delete `blanks/` to regenerate.
- **`tools/verify_blanks.py`** — read-only original-vs-blank gate, Stage 4.6 of the skill; `tests/test_blanks_integrity.py` calls the same functions.
  - **Package survey.** It checks every part, so no control, validation, drawing or conditional format is lost. Hidden, instructions and `SKIP_CLEAR` sheets stay byte-identical. Only `<sheetData>` of cleared sheets may change. VBA and `<dimension>` may differ only for files registered in `blank_patches.PATCHES`.
  - **Paste anchor.** On each data sheet, the last non-empty column-A row must be the header row.
- **`tools/build_vba_patch.py`** — one-time / on-hash-change Excel COM helper that rebuilds the RapidImpl patched `vbaProject.bin` (see `blank_patches.py`). Needs Excel with "Trust access to the VBA project object model" and ActiveX not set to "Disable all controls". It verifies the module set and `VB_Control` bindings survived, and refuses templates with auto-run VBA.
- **`.claude/skills/fbdi-compare-release/`** — orchestrator skill that chains the full download → clear → compare → catalog → report pipeline with human-in-the-loop checkpoints. Triggers on phrases like "Compare 26A to 26B" or "Oracle released 26C". Bundles Python helpers (`check_env.py`, `verify_download.py`, `summarize_report.py`, `verify_run.py`, `verify_rerun.py`) under `scripts/`; Stage 4.6 runs `tools/verify_blanks.py` and reference docs under `references/`. See the skill's `SKILL.md` for the staged workflow. The report stage generates the HTML/PDF release change report via `py -m fbdi report`; also triggers on report-only phrases like "generate the release change report" or "regenerate the PDF".
- **`tests/`** — 320 unit tests + ~1,720 environment-gated baseline checks (`test_blanks_integrity.py`, `TestRapidImplBlank`) that run only when `baselines/` is present (`py -m pytest tests/`)
- **Outputs:**
  - `Comparison_Report_<OLD>_<NEW>.xlsx` — 7-column diff for VBA validation (unchanged)
  - `FBDI_Master_Catalog.xlsx` — per-release snapshots + Issues + Drift tabs (gitignored; regenerable via `catalog`)
  - `FBDI_Change_Report_<OLD>_<NEW>.html` — module-grouped release change report; default output (gitignored)
  - `FBDI_Change_Report_<OLD>_<NEW>.pdf` — same content, print-rendered; opt-in via `--pdf` (gitignored)
- **Baseline layout** — `baselines/26A/originals/` (as-downloaded), `baselines/26A/blanks/` (smart-cleared copies for client use), and `baselines/26A/file_modules.json` (per-release `{filename: module}` map written by the downloader, consumed by the report for module grouping)

---

## Current Frontier

- **Phase 2 — Report redesign — landed (PR #2).** `fbdi/templates/report.html.j2` now implements the Definian design system at Tier B: navy cover with `ScopeTotals`, module→file/tab structure, color-coded change tables, self-contained inlined Aptos via `fbdi/templates/assets/definian.css`, HTML-first with an opt-in PDF (`--pdf`). See the `report.py` bullet above for the architecture. Design: `docs/superpowers/specs/2026-09-14-repo-cleanse-and-report-redesign-design.md` (§6).
- **26B → 26C — done (2026-09-14).** Full pipeline run via the skill: 1,582 field changes across 22 files (concentrated in Lease Accounting + Manufacturing Work Orders). Housekeeping (merged to master): `verify_rerun.py`'s stale `706 ± 50` compare-count pin replaced with a plausibility band (floor 50 / ceiling 6000, flags only gross anomalies), and `baseline_files.txt` (per-release download inventory) is now git-tracked so download verification has a real prior-release reference.
- **Baseline inventory reconciliation gate — added (branch `feat/baseline-inventory-reconciliation`).** Closes the silent-baseline-drift gap that let two 26C originals go missing undetected (inventory recorded 213, folder had 211). `verify_download.py` gained `diff_against_prior` + a `--reconcile` no-download audit; the skill runs it as a hard gate (**Stage 4.5**) for OLD and NEW before compare (and on resume), and `verify_run.py` re-asserts it at Stage 8 with a new **exit 2 = inventory hard-fail** (blocking; distinct from the non-blocking soft-regression exit 1). Match is asymmetric — a file present last quarter but absent now hard-fails and is named; net additions are informational.
- **26C → 26D — done (2026-11-05).**
  - **Results.** 1,139 compare changes across 32 files. 1,659 Drift rows across 36 files, 282 of them from the new `EquipmentRevenueLeaseContractImportTemplate`. Release change report: `FBDI_Change_Report_26C_26D.html`. PR #12 blanks were verified on 26D.
  - **Quarantined files.** IT released the two SentinelOne-quarantined 26D files. `baselines/26D/originals/` holds Oracle's real copies and `SUBSTITUTIONS.md` is resolved.
  - **Catalog rebuild.** 26B–26D were rebuilt after the gutter fix. That also removed 3 false PayablesBankReturns `REMOVED` Drift rows: the 26D catalog had been built while that file was quarantined.
- **26D follow-ups — done (branch `chore/26d-followups`).**
  - static `toc.htm` discovery replaces Selenium;
  - the vanished-file check;
  - catalog row-label gutter fix (28 phantom tabs per release, plus EgpCatalog metadata below the header);
  - `quiet_openpyxl`;
  - `tools/verify_blanks.py` + Stage 4.6;
  - `*.bak.xlsx` is gitignored and deleted after Stage 8.
- **Next:** 26D → 27A. Expect SentinelOne to quarantine `PayablesBankReturns`/`SupplierAttachmentImportTemplate` again on a fresh download: IT released the copies but didn't whitelist the hashes, and a 26D re-download on 2026-11-05 was quarantined again. Ask IT to whitelist them.

---

## Key Architectural Decisions (Closed — Do Not Re-litigate)

| Decision | Choice |
|---|---|
| Baseline storage | Folder-based: `baselines/26A/originals/`, `baselines/26A/blanks/` — gitignored, not Git-tracked |
| Header detection | Dynamic content scoring per tab — no hardcoded filename-to-header map |
| Excel reading | `openpyxl` with `data_only=True` where formula evaluation is needed |
| Comparison output | 7-column `.xlsx` — columns A–G as specified |
| Column scan cap | Max 500 columns per tab (avoids phantom `max_column=16384` from corrupt xlsm metadata) |

---

## Known Hazards

- **`RapidImplementationForCashManagement.xlsm` is not auto-downloadable** — this is an Oracle Rapid Implementation (FSM) template, not a standard FBDI template. It is not hosted on Oracle docs pages so the downloader never finds it. Must be obtained manually from Oracle Fusion: Setup and Maintenance → hamburger menu (top-right) → Search → search "Create Banks, Branches, and Accounts in Spreadsheet" → click the task to download. Place in `baselines/<VER>/originals/` before running compare. The `download_and_clear.py` script will warn if it's missing after a download run. Once placed, the compare engine picks it up automatically.
- **Never write blanks by re-saving through openpyxl.** openpyxl re-serializes the package and drops what it doesn't model. When `clear.py` did that, every 26A–26C blank lost Oracle's ActiveX Generate button (458 sheets), plus x14 cross-sheet dropdowns (93), drawings (72) and printer settings. Clearing hidden sheets also wiped the LOVs behind named-range dropdowns (51 sheets). `clear.py` now patches the original zip via `fbdi.ooxml.Package`. `tests/test_blanks_integrity.py` (baseline-gated) fails if any blank loses a control or hidden-sheet data its original has. Keep both.
- **RapidImpl blank fixes are hash-guarded, by design.** `blank_patches` swaps in a committed `vbaProject.bin` only when the original's sha256 matches `RAPID_IMPL_ORIGINAL_VBA_SHA256`. The A7 anchor requires the Bank Account header to still be row 7. If Oracle changes either, that one file fails in `download_and_clear.py` (the blank isn't written). Rebuild with `py tools/build_vba_patch.py <original>` and update the hash. Don't loosen the guard to "skip if mismatched". Excel drops the `Attribute VB_Control = "CommandButton1, ..."` binding on save when ActiveX is disabled (the Microsoft 365 "Disable all controls" setting). The tool checks for that, so a run that refuses to write the asset is correct. The same ActiveX setting stops the Generate button on any consultant's machine.
- **SentinelOne silently quarantines some Oracle templates.** In 26D it deleted `PayablesBankReturns.xlsm` and `SupplierAttachmentImportTemplate.xlsm` seconds after they were written.
  - It did the same to curl re-downloads, and to a local copy of the byte-identical 26C file.
  - Even after IT released those copies, a fresh 26D download was quarantined again, because releasing a file does not whitelist its hash.
  - Older copies already in `baselines/` survive, and so do the cleared blanks. Defender logs nothing.
  - **Detection.** The downloader re-checks every file on disk after a settle delay and again after the clear. It names the vanished ones and exits 1 (skill HITL #10).
  - **Resolution.** Check the SentinelOne tray app → Quarantined Files. Only IT can release a file or whitelist it by hash.
  - **Recovery.** Oracle hosts every template at `https://www.oracle.com/webfolder/technetwork/docs/fbdi-<ver>/fbdi/xlsm/<File>.xlsm`, and the downloader prints that URL.
- **Open question — Conversion Dashboard paste column on gutter tabs.**
  - **Affected tabs.** About 28 tabs per release (the PriceListsImportBatch/DiscountList/CostLists QP_* tabs, three DosSupplyOrder tabs, and five PONNegotiationLines tabs) keep row labels in column A, and their fields start at column B.
  - **What's verified.** Oracle's Generate-CSV macro copies from `B<header>`. `verify_blanks` confirms the paste *row* is right: the last column-A value is the header row.
  - **What isn't.** Whether the Dashboard pastes starting at column A or at the first header column can't be checked from this repo. If it pastes at A, these tabs would shift one column.
- **Phantom columns (`max_column=16384`)** — some xlsm files report 16384 columns due to corrupt metadata. The engine caps column scanning at 500.
- **Sample/example data rows below the header defeat header detection** — some templates ship an example data row directly under the header (Oracle 26C added these across the Lease cluster, e.g. `RevenueLeaseContractAmendTemplate::Payments` row 4). Its enum/code values (`NEW`, `USD`, `BASE_RENT`, `REV_TEMPLATE_STD`) match the UPPER_SNAKE pattern, so it used to win `detect_header.py` Tier 1 and the catalog stored *data values as field names* → alignment reported every real field REMOVED + every sample value ADDED (a fully-disjoint add/remove set is the tell). Guarded by `TIER1_MIN_STR_RATIO` (0.9): a genuine technical-header row is 100% strings, so a Tier-1 candidate carrying numbers/dates (a data row) is now excluded. Don't loosen this without a replacement discriminator.
- **Corrupt XML in some xlsm files** — handled gracefully; engine catches `zipfile.BadZipFile` and logs the file as unreadable. Post-run `verify_run` flags a regression if the per-release FILE_ERROR count jumps vs. the prior release.
- **VBA-produced `Comparison_Report_*.xlsx` files** — may ship with a corrupt stylesheet that fails standard `openpyxl.load_workbook`. Use `read_only=True` or `data_only=True` with exception handling if you need to ingest one.
- **Diagnose is still bounded by `MAX_FILE_SIZE_BYTES` (5MB)** — it loads workbooks in full (non-read_only) mode for memory reasons. Comparison is unbounded and streams via `iter_rows`.
- **Keep the Stage 4.5 reconciliation gate even though the downloader now fails loudly.**
  - **What the downloader catches.** `tools/download_and_clear.py` exits 1 on discovery errors, failed downloads and vanished files.
  - **What it can't catch.** It only knows what Oracle's docs link this quarter. It can't see a template Oracle stopped linking, or a baseline folder that drifted after the run. That is how two 26C originals went missing: inventory 213, folder 211.
  - **The authoritative check.** `verify_download.py --reconcile` is the skill's **Stage 4.5** hard gate. It runs for OLD and NEW before compare and on resume. `verify_run.py` re-asserts it at Stage 8 (**exit 2 = inventory hard-fail**). Both fail asymmetrically: a file present last quarter but absent now hard-fails; net additions are informational.
- **Discovery depends on Oracle's per-book `toc.htm`.** If Oracle moves or renames a docs book, discovery records `TOC ...: 404 / no topic pages found` and exits 1. Update `BOOK_TOC_TEMPLATES` and re-check with `--discover-only`. As of 2026-11-05, `--discover-only` finds exactly every reconciled baseline file for 26A–26D, minus the manual RapidImpl.
- **`RapidImplementationForCashManagement.xlsm` fallback when both baselines wiped** — skill HITL default "copy from prior baseline" assumes a prior is present. If a rerun wipes both 26A and 26B at once, fall back to an external archive (e.g., `C:/Users/10193/Definian/<old release>_*_Compare/<old>_FBDI/Manual/RapidImplementationForCashManagement.xlsm`).
- **Master catalog is gitignored → usually absent at run start; the report needs BOTH releases' tabs.** `FBDI_Master_Catalog.xlsx` regenerates each quarter and isn't tracked, so it's typically missing when a run begins. `generate_catalog` merges (`_load_existing_release_rows` preserves other release tabs), but against no existing master, `catalog --release <NEW>` alone yields a single-release catalog — and `report --old <OLD> --new <NEW>` (which aligns both per-release sheets) then has no OLD baseline. **When the catalog is absent, build BOTH:** `py -m fbdi catalog --release <OLD>` then `--release <NEW>`. The skill's Stage 6 assumes a pre-existing catalog and doesn't spell this out.
- **`openpyxl read_only=True` drops cell comments** — `cell.comment` is always None in read-only mode. `catalog.extract_file` uses full mode (`load_workbook(path, data_only=True)`) specifically so `_extract_metadata_from_comment` can see them; subprocess isolation in `_subprocess_util.run_worker` bounds memory for large files. Don't switch the catalog back to read_only without a replacement comment-loading path.
- **`type_parser.parse_data_type` is intentionally permissive** — it accepts any bare alpha token as a type for forward-compat with row-based extraction (so `parse_data_type("Required")` returns `data_type="REQUIRED"`). When mining cell comments or other unknown-shape inputs, gate via `catalog._ORACLE_TYPE_ALLOWLIST` (and the `_TYPE_SPEC_PREFIX_RE` helper for "VARCHAR2(N) trailing prose" patterns).
- **PDF rendering needs MSYS2 mingw64 GTK on Windows (only for `--pdf`)** — the default HTML path has no heavy dependency; `fbdi/report.py` imports weasyprint lazily only when `--pdf` is requested, and weasyprint depends on libgobject/libpango/libcairo. The standalone GtkD installer (winget id `GtkD.GtkPlusRuntime.x64`) ships Pango 1.43 and fails on weasyprint ≥53 with `pango_context_set_round_glyph_positions not found`. MSYS2 mingw64 ships Pango 1.56+ and works. Install: `winget install --id MSYS2.MSYS2 --silent`, then `C:/msys64/usr/bin/bash.exe -lc "pacman -S --needed --noconfirm mingw-w64-x86_64-pango mingw-w64-x86_64-gtk3 mingw-w64-x86_64-pkg-config"`. `_GTK_WINDOWS_BIN_CANDIDATES` in `report.py` probes MSYS2 first — keep that ordering. Do not lower the `weasyprint>=62.0` pin to work around an older GTK; older weasyprint lacks the flexbox/grid layout the Phase-2 report redesign needs.

## Resolved Hazards (historical note)

All fixed — listed so they aren't reintroduced. Forward-looking guidance for the still-live constraints lives in Known Hazards above.

- **Selenium scraper silently missed docs pages** — the JET `<oj-tree-view>` click-through raced the tree's async population (commit 82cd568 added a wait) and still missed topic pages (26D `RequirementImportTemplate`). Replaced by static `toc.htm` discovery; don't reintroduce a browser.
- **Catalog phantom position-1 field on row-label-gutter tabs** — asterisked technical names (`*BATCH_NAME`) made `_is_tier1_header` reject the row, so the tab fell to `_extract_thin`. Column A then became a field ("Technical Name" / "Column name of the Table …" / "Name"), and technical and type were never read. Fixed by `_find_gutter`. Don't "fix" it by stripping asterisks in `_is_tier1_header`: that flips about 62 tabs per release, including FixedAsset*/PO* tabs where column A is a real field.

- **>5MB files skipped / ~75min runtime** — fixed by subprocess isolation + `iter_rows` streaming; comparison is now unbounded (~0.02s per wide tab).
- **8 tabs with non-standard headers fail detection** — fixed in Phase 3; diagnose reports `NO_HEADER: 0`.
- **`ChangeOrderImportTemplate` / `ItemImportTemplate` bogus catalog TIMEOUT** — fixed by `_subprocess_util.run_worker` drain-before-join (~1,400 rows each ingest cleanly).
- **463 TYPE_PARSE_WARNING rows** — collapsed to 9 (genuinely-broken strings only) after `type_parser.py` learned temporal format masks and the trailing-period typo.

---

## Reference Files

Two read-only archives, distinct purposes:

- **`reference/`** — pre-Python pipeline artifacts (legacy VBA macros, Dan's original Selenium downloader). Do not modify.
- **`docs/archive/`** — historical narrative docs (audit notes, scraper gap findings, and stale specs/plans from retired subsystems) preserved via `git mv` so blame history survives. Do not modify.

`reference/` contents:

| File | What It Is |
|---|---|
| `fbdi_compare.xlsm` | Legacy VBA macro that did the comparison before the Python engine |
| `Clear_FBDIs - 20210412.xlsm` | Legacy VBA macro that cleared template files before re-download |
| `Oracle_26A_Comparison_Report.docx` | Sample output from the VBA macro for release 26A — used as a reference baseline during Phase 1 |
| `test.py` | Dan's original Selenium downloader — scrapes Oracle docs and downloads xlsm files |

---

## Testing

- `py -m pytest tests/` — run full suite (320 unit tests; plus ~1,720 environment-gated baseline checks when `baselines/` is present)
- `py -m pytest tests/test_clear.py -v` — run one module
- `tests/validate_against_vba.py` and `tests/vba_fieldrow_map.json` — ad-hoc validation against the legacy VBA macro's expected header rows (not pytest, kept for spot-checks against regressions)

Tests use `openpyxl.Workbook` to build synthetic FBDI files per test (see `_make_fbdi_workbook` / `_create_fbdi_workbook` helpers). There are no fixtures — fixture-like workbooks are built inline in each test so the expected layout is visible next to the assertions.

**Test-data gotcha:** `detect_header_row` scores rows by UPPER_SNAKE_CASE content. Synthetic sample data like `"CREATE"`, `"V1"`, `"DR_ECO_1"` will false-positive as headers. Use lowercase/mixed-case values in test data rows (e.g. `"Create Order"`, `"V1-org"`).

---

## Docs & Planning

Two patterns have been used in this repo — both are still valid:

- **Old:** `handoff_*.md` files (written in Claude Chat, executed by Claude Code) — now gitignored, kept in conversation or local scratch.
- **Current:** `docs/superpowers/specs/*.md` (design) and `docs/superpowers/plans/*.md` (implementation plans) — produced via the `superpowers:brainstorming` and `superpowers:writing-plans` skills and committed to the repo so the history is auditable.

Completed-project narrative docs (audit notes, one-off findings) live in `docs/archive/`. The two user-facing guides live at `docs/operator-guide.md` and `docs/developer-guide.md`.

---

## Plugins / Tooling

Project uses the `superpowers` skill family (brainstorming, writing-plans, executing-plans, systematic-debugging, verification-before-completion). See user-level `~/.claude/` config for the full plugin list — no project-specific plugin requirements.
