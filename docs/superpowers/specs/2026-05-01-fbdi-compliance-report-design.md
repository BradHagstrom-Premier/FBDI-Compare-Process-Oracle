# FBDI Compliance Report — Design Spec

**Date:** 2026-05-01
**Author:** Brad Hagstrom (with Claude Code)
**Status:** Approved for implementation planning

## Purpose

Replace the manually-built Oracle FBDI Compliance Report (Word doc, exemplified by `reference/Oracle_26A_Comparison_Report.docx`) with a data-driven generator that emits two artifacts per release pair:

- `FBDI_Compliance_Report_<OLD>_<NEW>.html` — interactive browseable report with collapsible SHIFTED tables, in-page jump-links, brand-styled visuals.
- `FBDI_Compliance_Report_<OLD>_<NEW>.pdf` — formal deliverable, same content, print-rendered (collapsibles auto-expanded into a compact form).

Both files derive from one Jinja2 template via a `print_mode` flag, sharing one source of truth so HTML and PDF stay in lockstep. The deliverable goes to Definian consultants who maintain Applaud installations on client servers; they need to know what changed in Oracle's FBDI templates and what to update in their Applaud install (DB tables, Import Forms, Export Forms).

## Background

Three artifacts exist today, all gitignored:

- `Comparison_Report_26A_26B.xlsx` — 7-column field diff (label-only, no type info), produced by `fbdi/compare.py`. Used as a VBA-validation artifact.
- `FBDI_to_ApplaudTables_Mapping.xlsx` — Brad's hand-curated mapping (639 rows; 154 MAPPED, 485 UNMAPPED) with `FBDI Template, FBDI Tab, Applaud Table, Prefix, Status, Module, In Base System?` columns.
- `FBDI_Master_Catalog.xlsx` — per-release snapshots produced by `fbdi/catalog.py`. Sheets `26A`, `26B` carry the per-tab field metadata (position, label, technical name, data type, length, scale, required). Sheet `Drift` is supposed to be a per-position comparison.

**Critical data finding from brainstorming:** the catalog's `Drift` sheet uses a naive per-position diff with no alignment. When a real change is "1 ADDED at position 19 + 20 fields shifted from positions 19–39 to 20–40" (the actual 26A→26B story for `WorkDefinitionTemplate / Work Definition Headers`), the per-position diff misclassifies it as `{'ADDED': 6, 'MULTI': 2, 'RENAMED': 19}` — wrong on every count. Across 26A→26B, the Drift sheet emits 460 RENAMED + 236 MULTI + 78 ADDED + 2 REMOVED, the bulk of which are shift-cascade artifacts.

The reference Word doc (25D→26A) was hand-built so it correctly identified shifts and removals. Reproducing that quality programmatically requires a real alignment pass.

**In-scope footprint for 26B:** of 22 FBDI files with any change, only 5 are MAPPED, accounting for 6 (file, tab) pairs in scope total (5 SCM, 1 Financials). One of those pairs (`WorkOrderTemplate / Work Order Header`) is flagged `Needs to be created in base system` and routes to the pending-base list, leaving 5 in-scope tabs across 4 distinct FBDI files for the main per-file body (`ItemImportTemplate` contributes 2 tabs).

## Decisions (closed — do not re-litigate)

| Decision | Choice | Rationale |
|---|---|---|
| Output format | PDF (formal deliverable) + HTML (interactive working copy) | PDF is what consultants archive/email; HTML is what they browse. Brand-color polish needs both. |
| PDF render path | `weasyprint` from the same Jinja2 HTML template via `print_mode=True` | Pure Python, no headless browser, excellent CSS support; one template = one source of truth. |
| Source of truth | `FBDI_Master_Catalog.xlsx` per-release sheets (`26A`, `26B`) | The only source with type/length/required data. Comparison_Report stays untouched (VBA-validation artifact). |
| Alignment algorithm | LCS-style match by `(technical_name, label)` with tie-breaks; shared module | Required for SHIFTED detection and to stop misclassifying shift cascades as RENAMED/MULTI. |
| Catalog Drift fix | In scope — root-cause fix using the same shared `align.py` module | Drift is a data product; a wrong data product downstream is worse than a missing one. Aligns with Brad's no-temp-fixes preference. |
| Report scope | MAPPED in-scope tabs only in main body; pending-base tabs in a separate compact list section; UNMAPPED files excluded entirely | Consultants can't action what isn't on their Applaud install — unmapped data is noise. |
| `NEEDS_REVIEW` handling | Visual flag in summary + per-file header (no exclusion) | Currently zero `NEEDS_REVIEW` rows in the mapping data, but design supports them for future use. |
| `In Base System? = "Multiple mapping is possible..."` rows | Treated as MAPPED (in main body) with a small advisory note in the per-file header | 6 such rows exist; they're real mappings, just with an annotation. |
| Section structure | Cover · Module rollup · Summary table · Per-file sections · Pending base-system tables | Matches the reference doc's content while restructuring for scannability. |
| Per-file change types | ADDED, REMOVED, MODIFIED (type/length, required), RENAMED (label-only), SHIFTED, MULTI | Derived from the alignment, not from naive per-position diff. |
| Action matrix per change type | See "Action matrix" below | Encodes which of DB/IF/EF actually need consultant action; required-flag and rename are flag-only / low-priority. |
| SHIFTED presentation | HTML: collapsed `<details>` (default closed) with full per-field old→new table; PDF: summary + compact 2-col grid auto-rendered (no toggle) | Same data, presentation per medium. PDF can't toggle. |
| Applaud field name truncation | Always show `<prefix><technical>` at full length; flag with ⚠ chip when length > 30 chars; never auto-truncate | Truncation is an irreversible naming decision a human must own. |
| Applaud type translation | Programmatic: `VARCHAR2(N)` → `char N`, `NUMBER(p,s)` → `numeric p,s`, `NUMBER` → `numeric` (no defaults), `DATE`/`TIMESTAMP` → `date` | Mirrors the convention in the reference doc. |
| Color palette | Strict Definian palette only (`#0D2C71`, `#00AB63`, `#02072D`, `#3C405B`, `#D8D7EE`, `#FFFFFF`) plus minimal accent (warning amber `#B8860B`, removal red `#C0392B`) | Per `reference/colorguide.pdf`. No off-palette colors anywhere. |
| Filename convention | `FBDI_Compliance_Report_<OLD>_<NEW>.html` and `.pdf` | Matches `FBDI_Master_Catalog.xlsx` style; preserves "FBDI" + "Compliance" naming from the reference doc title. |
| CLI surface | `python -m fbdi report --old 26A --new 26B [--out-dir .] [--catalog FBDI_Master_Catalog.xlsx] [--mapping FBDI_to_ApplaudTables_Mapping.xlsx]` | Matches existing subcommand style (`compare`, `catalog`, `populate-module`). |
| Generator side effects | Read-only; emits files only | No coupling between report and catalog state; catalog is regenerated separately. |

## Action matrix

Encodes which of DB / IF / EF need a consultant action per change type. Renders as visible checkbox columns in each per-file section table.

| Change type | DB action | IF action | EF action | Visual |
|---|---|---|---|---|
| **ADDED** | Add column | Add field | Add field | Solid blue checkbox each |
| **REMOVED** | Drop column | Remove field | Remove field | Red checkbox each |
| **MODIFIED — type/length** | Alter column | Update length validation | Update length validation | Amber checkbox each |
| **MODIFIED — required flag** | Alter NULL constraint | — *(IF cannot validate required at field level)* | — | Amber DB checkbox; "flag only" badge in change cell; em-dash for IF/EF |
| **RENAMED — label only** | Optional: update DB data element description | — | — | Dashed-border DB checkbox; em-dash for IF/EF; advisory note "low priority" |
| **SHIFTED** | — | Reorder field | Reorder field | Em-dash DB; blue checkbox IF/EF |
| **MULTI** | Union of applicable actions per the underlying change components | | | One row per affected position; checkboxes per applicable action |

## Architecture

Five new/modified pieces, each independently testable.

### 1. Alignment algorithm — `fbdi/align.py` (new)

Pure function `align_tabs(old_rows, new_rows) -> AlignmentResult`. Inputs are the per-tab row sets read from the catalog's per-release sheets. Algorithm:

1. **Match pass**: longest common subsequence over `technical_name` (preferred) with `label` as fallback when `technical_name` is None. Produces matched pairs (old_pos, new_pos) and unmatched-old / unmatched-new sets.
2. **Classify each matched pair across three independent axes**:
   - `label_changed`: labels differ
   - `metadata_changed`: any of (data_type, length, required) differs
   - `position_changed`: old_pos ≠ new_pos
3. **Map axis combinations to change types**:
   - 0 axes changed → unchanged (not emitted)
   - 1 axis changed → `RENAMED` (label only) / `MODIFIED` (metadata only) / `SHIFTED` (position only)
   - 2+ axes changed → `MULTI` with `sub_kinds` listing the axes that changed (e.g., `"position,metadata"` for a shift+type change). For MULTI, additionally store which metadata sub-kinds applied (`type`, `length`, `required`).
4. **Classify unmatched**:
   - `ADDED` for unmatched-new
   - `REMOVED` for unmatched-old

Returns a typed dataclass list of `Change(file, tab, change_type, old_pos, new_pos, old_field, new_field, axes, sub_kinds)`.

Pure, deterministic, no I/O. Tested in isolation against synthetic alignment scenarios per `superpowers:test-driven-development`.

### 2. Catalog Drift fix — `fbdi/catalog.py` (modified)

Existing `Drift` writer is replaced by one that calls `align.align_tabs()` and emits one row per `Change`. New schema:

| Column | Type |
|---|---|
| `file` | str |
| `tab` | str |
| `change_type` | enum (`ADDED`, `REMOVED`, `MODIFIED`, `RENAMED`, `SHIFTED`, `MULTI`) |
| `old_position` | int or `None` (for ADDED) |
| `new_position` | int or `None` (for REMOVED) |
| `old_label`, `new_label` | str or `None` |
| `old_technical`, `new_technical` | str or `None` |
| `old_data_type`, `new_data_type` | str or `None` |
| `old_length`, `new_length` | int or `None` |
| `old_required`, `new_required` | bool or `None` |
| `sub_kinds` | str or `None` (e.g., `"type,required"` for MULTI) |

Existing tests for catalog Drift are updated to match the new schema; new tests cover the alignment-based classifications.

### 3. Oracle → Applaud type translator — `fbdi/applaud_type.py` (new)

Pure function `applaud_type_for(oracle_type: ParsedType) -> str` consuming the parsed-type dataclass from `fbdi/type_parser.py`. Mapping:

| Oracle | Applaud |
|---|---|
| `VARCHAR2(N)` | `char N` |
| `VARCHAR2(N CHAR)` | `char N` |
| `NUMBER(p, s)` | `numeric p,s` |
| `NUMBER` (no precision) | `numeric` |
| `DATE`, `DATE(format)` | `date` |
| `TIMESTAMP`, `TimeStamp(format)` | `date` |
| `CLOB`, `BLOB`, `RAW` | passthrough as `<type>` |
| Unparseable / unknown | passthrough as `<raw>` |

Tested in isolation; covers the type strings actually present in the catalog plus a few synthetic edge cases.

### 4. Report generator — `fbdi/report.py` (new)

Public function `generate_report(catalog_path, mapping_path, old_release, new_release, out_dir)` returns `(html_path, pdf_path)`. Pipeline:

1. Load mapping → in-memory dict `{(template, tab): {applaud_table, prefix, module, status, in_base}}`
2. Load catalog `26A` and `26B` sheets → group rows by `(file, tab)`
3. For each `(file, tab)` present in either release:
   - Lookup mapping. If `Status = UNMAPPED` or `(file, tab) not in mapping`, **skip entirely**.
   - If `In Base System? contains "Needs to be created"`, **route to pending-base list** (not main body).
   - Otherwise call `align.align_tabs()` → list of `Change`
   - Build a `FileSection` view-model with: file, tab, applaud_table, prefix, module, status, change-type buckets (ADDED/REMOVED/MODIFIED/RENAMED/SHIFTED/MULTI), per-row Applaud field names (`prefix + technical_name` when present, else `prefix + catalog_normalize.normalize_label(label)`), per-row Applaud types via `applaud_type_for`, per-row 30-char-limit warnings.
4. Build top-level `ReportContext`: cover meta, per-module rollup, summary table, file sections sorted by `(module, file, tab)`, pending-base list.
5. Render HTML: `templates/report.html.j2` with `print_mode=False` → write to `out_dir/FBDI_Compliance_Report_<OLD>_<NEW>.html`
6. Render PDF: same template with `print_mode=True` → pass through `weasyprint.HTML(string=...).write_pdf(out_dir/...pdf)`

Returns the two paths. Logs counts: in-scope tabs, pending-base tabs, excluded UNMAPPED files (for sanity).

### 5. Templates — `fbdi/templates/`

- `report.html.j2` — single Jinja2 template; uses `{% if not print_mode %}<details>{% endif %}`-style conditionals for collapsibles. Self-contained CSS embedded in `<style>` tag at the top.
- `_partials/*.j2` — small shared partials (per-file section, per-change-block table, summary row).

The CSS is hand-written using the Definian palette — no external framework, no Tailwind, no off-palette grays.

Generated narrative prose (lede sentences, per-file shift summaries) is authored with the `humanizer-skill` patterns in mind: no AI-tells, no rule-of-three, no "in today's evolving landscape" phrasings, no inflated significance. Plain factual sentences.

### 6. CLI integration — `fbdi/cli.py` (modified)

Add `report` subcommand:
```
python -m fbdi report --old 26A --new 26B [--out-dir .] [--catalog FBDI_Master_Catalog.xlsx] [--mapping FBDI_to_ApplaudTables_Mapping.xlsx]
```
Default `--out-dir` is repo root. Default `--catalog` and `--mapping` resolve to the working files at repo root (matching existing patterns in `populate-module`).

## Data flow

```
FBDI_Master_Catalog.xlsx (26A sheet, 26B sheet)
                           │
                           ▼
                    align_tabs()  ◄────── shared module
                    │           │
                    ▼           ▼
       fbdi/catalog.py    fbdi/report.py
       (Drift sheet        │
        regenerated)       │
                           ▼
                    + FBDI_to_ApplaudTables_Mapping.xlsx
                    + applaud_type_for()
                           │
                           ▼
                    ReportContext (view-model)
                           │
            ┌──────────────┴──────────────┐
            ▼                             ▼
      report.html.j2                report.html.j2
      (print_mode=False)            (print_mode=True)
            │                             │
            ▼                             ▼
  Compliance_Report.html       weasyprint → Compliance_Report.pdf
```

## Output structure

| Section | Content |
|---|---|
| **Cover** | Title, release pair, generated date, Definian brand mark |
| **1. At a glance** | Module rollup cards (one per module with in-scope changes) — counts of tabs, added, shifted, removed |
| **2. Summary by FBDI tab** | Table of all in-scope tabs with file·tab, applaud table, prefix, module pill, per-change-type counts |
| **3. Required changes by FBDI file** | Per-file sections, each with: header (file, tab, applaud table, prefix, module pill), then per-change-type tables (ADDED, REMOVED, MODIFIED, RENAMED, SHIFTED) |
| **4. Pending base-system tables** | Compact list of `(file, tab → applaud_table · prefix · module · change count)` for tables flagged `Needs to be created in base system`. Footer note pointing to the catalog for full detail. |

Sections 1, 2, 3 only show MAPPED in-scope content. UNMAPPED files are silently excluded.

## Skills used during implementation

| Skill | Where it applies |
|---|---|
| **`superpowers:test-driven-development`** | All new code: `align.py`, `applaud_type.py`, `report.py`. Tests authored before implementation; one assertion per scenario. |
| **`superpowers:writing-plans`** | Immediate next step after this spec is approved — produces a paired plan in `docs/superpowers/plans/`. |
| **`frontend-design`** | Authoring `report.html.j2` and embedded CSS. Visual hierarchy, typography, spacing rhythm, deliberate color choices, no generic AI defaults. |
| **`humanizer-skill:humanizer`** | Authoring all generated prose in templates and `report.py` (lede sentences, narrative summaries, advisory notes). Filters out AI-tell patterns. |
| **`chrome-devtools-mcp:chrome-devtools`** | Verifying rendered HTML and PDF output post-implementation. Screenshot per section to validate layout, palette adherence, table widths, the SHIFTED collapsible behavior. |
| **`superpowers:verification-before-completion`** | Before marking the implementation complete: run the generator end-to-end, visually inspect both outputs, confirm the report matches this spec's content. |

## Testing strategy

- `align.py` — unit tests for: pure-add, pure-remove, pure-shift, type-modified, length-modified, required-flipped, rename-only, multi-change, swap (a moves down + b moves up), insertion at start/middle/end, empty-old, empty-new.
- `applaud_type.py` — unit tests for each Oracle type variant present in the catalog plus synthetic edge cases (unparseable, unknown, trailing-period typos).
- `report.py` — unit tests for: scope filtering (UNMAPPED excluded, pending-base routed), 30-char warning emission, MULTI-row composition, action-matrix correctness per change type. Plus one integration test: load the actual 26A/26B catalog + mapping, run end-to-end, assert that the in-scope file count and per-file section content match expectations.
- HTML/PDF rendering — golden-file test for one synthetic small input; visual verification via `chrome-devtools` on the real 26A→26B output.

## Out of scope (explicit)

- **Applaud-mcp integration** — comparing FBDIs directly to a client's Applaud MDB file. Brad will tackle this in a later session.
- **Auto-truncation of Applaud field names exceeding 30 chars** — flagged with chip; truncation decision stays human.
- **Inline Oracle docs URLs per FBDI tab** — would require a separate scrape of Oracle's table-detail pages (different URL space than where the FBDI xlsm lives). Possible v2 enhancement.
- **`NEEDS_REVIEW` workflow beyond visual flag** — none currently in the data; design supports them when they appear.
- **DB/IF/EF action checklist as a tracking artifact** — checkboxes are visual cues in the report, not a stateful workbench. Stateful tracking lives in consultant-side tooling.
- **Section 3 "Oracle Documentation References"** from the reference doc — dropped (we don't have the per-table Oracle docs URLs).
- **Section 2 "Unresolved FBDI Files"** from the reference doc — dropped (UNMAPPED files are noise per the consultant audience).

## Open questions / future work

- Once Applaud-mcp lands, the report could be filtered to "changes that affect *this* client's Applaud install" rather than the base-system superset. The current design's clean separation between mapping lookup and rendering makes this drop-in.
- A "diff vs. previous report" mode (e.g., "what's new in 26B that wasn't in 26A→26B compared to 25D→26A") — useful if Definian tracks compliance burden across releases.
- HTML accessibility audit via `chrome-devtools-mcp:a11y-debugging` — not blocking initial release but worth adding once the template is stable.
