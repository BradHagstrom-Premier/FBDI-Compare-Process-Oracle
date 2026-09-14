# Repo Cleanse + Consultant Report Redesign — Design

- **Date:** 2026-09-14
- **Status:** Approved (design) — pending spec review
- **Author:** Brad Hagstrom + Claude (Opus 4.8)
- **Supersedes:** `2026-05-01-fbdi-compliance-report-design.md`,
  `2026-05-04-compliance-report-skill-stage9-design.md`,
  `2026-05-04-fbdi-run-headless-pipeline-design.md` (all to be archived)

---

## 1. Context & Motivation

This repo was built before three sibling assets existed at Definian:

- **ApplaudMCP** (`C:\Users\10193\Definian\ApplaudMCP`) — the MCP server that is
  now the source of truth for everything Applaud (tables, fields, imports, code).
- **CodingAssistant** (`C:\Users\10193\Definian\CodingAssistant`) — home of the
  Definian design system and the `impeccable` tooling.
- **The Definian design system** — brand tokens, Aptos fonts, and the
  `definian-design` skill under `CodingAssistant/local/impeccable`.

Over ~4 months (PRs #2–#4) this repo grew a large **FBDI→Applaud** audit /
mapping / field-correspondence subsystem. That subsystem is now **superseded by
ApplaudMCP**. The durable value that remains is exactly what it was at the
start: **pull Oracle FBDI templates down, blank them out, and compare releases**
— plus the per-release **catalog** and a **consultant-facing change report**.

Before running the 26B→26C quarterly refresh, we cleanse the repo down to that
durable core and redesign the report as a Definian-branded consultant
deliverable ("here's what changed from 26B to 26C").

CLAUDE.md is itself stale — it does not mention `audit_applaud.py`,
`correspondence.py`, `applaud_snapshot.py`, or `applaud_appmap.py` — so it is
rewritten as part of this work.

## 2. Goals

- Remove the entire Applaud layer (modules, tools, tests, CLI subcommands, data,
  tracked workbooks) — it is superseded by ApplaudMCP.
- Keep and clean the durable core: **download → clear → compare → diagnose**,
  plus the **catalog** (per-release schema snapshot + Drift).
- **Repurpose the report** from an "Applaud compliance report" into a
  **consultant-facing FBDI release-change report**, decoupled from Applaud and
  driven by the catalog + `file_modules.json`.
- **Redesign the report** with the Definian design system: self-contained HTML
  as the primary deliverable, with an **opt-in PDF** (`--pdf`).
- Rewrite CLAUDE.md, README, DESIGN.md, PRODUCT.md, and the operator/developer
  guides to match the cleansed reality; trim the `fbdi-compare-release` skill.
- Leave `master` green (full test suite passing) after **each** PR.

## 3. Non-Goals

- No new comparison features. The compare/catalog engines are unchanged except
  where the Applaud cut forces an edit.
- No `python -m fbdi run` headless pipeline (its old design is archived; it can
  be re-designed later against the cleansed shape).
- No integration with ApplaudMCP from this repo. Consultants use ApplaudMCP
  separately after reading the report.
- No change to Oracle scraping behavior in `download_and_clear.py` beyond
  incidental doc/comment updates.

## 4. Current-State Inventory (four layers)

| Layer | Modules | CLI cmds | Imports core? | Disposition |
|---|---|---|---|---|
| **1. Core** | `clear`, `compare`, `diagnose`, `detect_header`, `config`, `utils`, `_subprocess_util`, `cli`, `__main__`, `tools/download_and_clear.py` | `compare`, `diagnose` | — | **Keep** (clean) |
| **2. Catalog** | `catalog`, `type_parser`, `align`, `catalog_normalize`* | `catalog` | no | **Keep** (evaluate `catalog_normalize`) |
| **3. Report** | `report`, `templates/`, (`applaud_type`) | `report` | no | **Repurpose + redesign** |
| **4. Applaud** | `audit`, `audit_applaud`, `applaud_appmap`, `applaud_snapshot`, `correspondence`, `build_mapping`, `populate_module`, `applaud_type` | `populate-module`, `audit-applaud`, `corr-derive`, `corr-confirm` | no | **Remove** |

\* `catalog_normalize` (Applaud-MDB label normalization) is shared by catalog +
report. See §5.1 for its disposition.

**Key fact:** the core imports nothing from layers 2–4, and layer 4 imports
nothing into the core, so the Applaud removal is surgical.

---

## 5. Phase 1 — The Cleanse (PR #1)

Mechanical, low-risk. Goal: focused core+catalog+working-report repo, green
tests, ready to run 26B→26C.

### 5.1 Remove the Applaud layer

Delete (via `git rm`):

- **Modules:** `fbdi/audit.py`, `fbdi/audit_applaud.py`, `fbdi/applaud_appmap.py`,
  `fbdi/applaud_snapshot.py`, `fbdi/correspondence.py`, `fbdi/build_mapping.py`,
  `fbdi/populate_module.py`, `fbdi/applaud_type.py`.
- **Tools:** `tools/_gen_appmap_full.py`, `tools/assemble_applaud_snapshot.py`,
  `tools/extract_applaud_snapshot.mjs`.
- **Tests:** `tests/test_audit.py`, `tests/test_audit_applaud.py`,
  `tests/test_correspondence.py`, `tests/test_applaud_appmap.py`,
  `tests/test_applaud_snapshot.py`, `tests/test_populate_module.py`,
  `tests/test_module_classifier.py`, `tests/test_applaud_type.py`.
  Prune Applaud cases from `tests/test_cli.py`, `tests/test_report.py`,
  `tests/test_skill_scripts.py`.
- **Config:** remove `config.applaud_snapshot_path()` and the Applaud comment
  block from `fbdi/config.py`.
- **CLI:** remove the `populate-module`, `audit-applaud`, `corr-derive`,
  `corr-confirm` subparsers + their handler functions + lazy imports from
  `fbdi/cli.py`.
- **Data:** delete `baselines/applaud/` (gitignored; on-disk only).

**`catalog_normalize` decision:** it exists only for Applaud MDB compatibility.
During implementation, check whether the catalog still writes a normalized-label
column and whether anything but the (removed) report consumes it. If nothing
non-Applaud needs it, remove `catalog_normalize.py`, its test
(`test_catalog_normalize.py`), the catalog column, and the catalog import. If
removing it perturbs the catalog schema in a way the report reader depends on,
keep the module but drop the Applaud framing from its docstring. Default lean:
**remove it** (root-cause cleanse over leaving dead weight).

### 5.2 Decouple the report to a working state

Rewrite `fbdi/report.py` so it no longer touches Applaud:

- **Remove** `load_mapping()`, the `applaud_type`/`catalog_normalize` imports,
  and the Applaud-only view-model fields (`applaud_field_name`, `name_length`,
  `name_exceeds_30`, `applaud_type_str`, `applaud_table`, `prefix`).
- **Scope** becomes *every `(file, tab)` in the catalog that has changes* — no
  mapping filter, no `pending_base` routing. (Consultants want all changes, not
  just previously-mapped ones.)
- **Module grouping** comes from `baselines/<new>/file_modules.json`, falling
  back to `baselines/<old>/file_modules.json` (NEW wins, mirroring the retired
  `populate_module`), keyed by `file_name`. Files with no module entry group
  under an "Unclassified" bucket.
- **Keep** `align_tabs` classification (ADDED/REMOVED/MODIFIED/RENAMED/SHIFTED/
  MULTI), `load_catalog_release`, `_oracle_type_str`, shift summaries,
  `ScopeTotals`. These are Applaud-free.
- **PDF opt-in:** `generate_report(..., pdf: bool = False)`. HTML always
  written; weasyprint imported lazily **only** when `pdf=True`, with a clear
  actionable error if GTK is missing (keep `_register_windows_gtk_dlls`).
- **Rename outputs** to `FBDI_Change_Report_<OLD>_<NEW>.html` / `.pdf`; retitle
  in-document copy from "Compliance Report" to "FBDI Release Change Report".
- **CLI:** `python -m fbdi report --old 26B --new 26C [--pdf]`. Drop the
  `--mapping` argument.

The Phase-1 template stays **functional-but-plain** (existing markup minus the
Applaud columns). Visual redesign is Phase 2. `tests/test_report.py` is updated
to the Applaud-free view-model and the new scope/module logic.

### 5.3 De-clutter the working tree

- `git rm --cached FBDI_Master_Catalog.xlsx` (regenerable) and add it to
  `.gitignore`.
- `git rm FBDI_to_ApplaudTables_Mapping.xlsx FBDI_to_Applaud_AppMap.xlsx`.
- Delete on-disk (gitignored) clutter: `Archive/`,
  `FBDI_Compliance_Report_26A_26B.html/.pdf`, `Diagnostic_Report_26B.xlsx`,
  `Comparison_Report_26A_26B.xlsx`, `Applaud_Compliance_Report_*.xlsx`,
  `Applaud_FieldMap_Review_*.xlsx`.
- `.gitignore`: add `FBDI_Master_Catalog.xlsx`; rename the report ignore pattern
  to `FBDI_Change_Report_*.html` / `.pdf`; remove the Applaud-specific lines
  (`Applaud_Compliance_Report_*`, `Applaud_FieldMap_Review_*`,
  `applaud_table_coverage.xlsx`, the FBDI_to_Applaud tracked-file notes).

### 5.4 Docs

- `git mv` the stale Applaud + old-report docs into `docs/archive/` (preserves
  blame, per repo convention). See Appendix B for the full list.
- Update `docs/operator-guide.md` and `docs/developer-guide.md`: remove the
  populate-module and Applaud steps; document the new `report --pdf` flow.

### 5.5 Trim the `fbdi-compare-release` skill

- `.claude/skills/fbdi-compare-release/SKILL.md`: remove the populate-module
  stage; refit the report stage (Stage 9) to `report --old --new [--pdf]` with
  the new name; strip Applaud references from stage text and `references/`.
- Update bundled scripts (`summarize_report.py`, `verify_run.py`,
  `verify_rerun.py`) where they reference the old report name or Applaud.

### 5.6 Rewrite CLAUDE.md + top-level docs

- Rewrite `CLAUDE.md` (using the `claude-md-management:claude-md-improver`
  skill) to the cleansed reality: drop the Applaud "Current Frontier" and
  Applaud modules from the pipeline list; drop Applaud hazards; **remove the
  Inter-font hazard note** (Phase 2 replaces it with Definian fonts — but if
  Phase 2 lands in a later PR, keep the note until then; see §7); update Quick
  Start (remove `populate-module`, `audit-applaud`, `corr-*`); document the new
  report command and `--pdf`.
- Update `README.md`, `DESIGN.md`, `PRODUCT.md` to drop Applaud framing.

### 5.7 Phase 1 acceptance criteria

- `python -m pytest tests/` passes (expected ~ down from 354 to the core+catalog
  +report subset; the exact count is recorded in the plan).
- `python -m fbdi compare --old 26A --new 26B` and `--old 26B --new 26C`
  (once 26C exists) run unchanged.
- `python -m fbdi catalog --release 26B` runs unchanged.
- `python -m fbdi report --old 26A --new 26B` writes
  `FBDI_Change_Report_26A_26B.html`; `--pdf` additionally writes the PDF (or
  errors clearly if GTK is absent).
- `grep -ri applaud fbdi/ tools/ tests/` returns nothing (except intentional
  archived docs).
- No import of a deleted module remains (`python -c "import fbdi.cli"` clean).

---

## 6. Phase 2 — Report Redesign (PR #2)

Creative. The report's **data layer is already Applaud-free from Phase 1**, so
this phase is primarily the template + presentation.

### 6.1 Design system application

- Apply the **Definian design system** from
  `CodingAssistant/local/impeccable/definian-design-system/project` via the
  `definian-design` / `impeccable` skill at **Tier B** (internal/operational —
  brand-aligned, holds the brand spirit; elevate to Tier A if the report becomes
  client-facing).
- Use **`definian.css`** (self-contained, base64-inlined Aptos fonts) **inlined
  into the report HTML** so the file is a single portable artifact (renders in
  browsers, email, and weasyprint without external font fetches). This
  **retires the Inter/@font-face hazard** currently documented in CLAUDE.md.
- Brand tokens: Definian Blue `#0D2C71`, Green `#00AB63`, Midnight `#02072D`,
  Cool Gray `#D8D7EE`, Dark Gray `#3C405B`; Aptos SemiBold (display) / Aptos
  Regular (body). Charts/severity accents use the brand palette in order.

### 6.2 Content structure

- **Cover:** Definian wordmark, "FBDI Release Change Report — 26B → 26C",
  generated date, and `ScopeTotals` (tabs changed; added / removed / modified /
  shifted counts) so a lead reads the cover and understands scope.
- **Body grouped by module → file/tab.** Each tab section shows per-type change
  tables (ADDED / REMOVED / MODIFIED / RENAMED / SHIFTED) with Oracle type,
  position(s), and required flag; RENAMED/MODIFIED/MULTI render old-vs-new
  side by side.
- **SHIFTED** collapses to the one-line summary when the shift is uniform
  (existing `shift_is_uniform` logic); otherwise a collapsible detail table.

### 6.3 Interactivity (HTML) & PDF

- HTML: collapsible sections, a module/file filter, and a text search across
  field names — all inline vanilla JS, no external deps (email/offline safe).
- PDF (`--pdf`): print CSS renders the same content statically (interactivity
  degrades to expanded). definian.css base64 fonts make PDF glyphs
  deterministic.

### 6.4 Phase 2 acceptance criteria

- `impeccable` audit pass on the rendered HTML (brand tokens, WCAG AA contrast,
  no off-palette values, no lorem/emoji-as-icon).
- HTML is a single self-contained file (no external CSS/font/JS references).
- `--pdf` produces a visually consistent PDF via weasyprint.
- CLAUDE.md Inter-font hazard note removed (fonts now via definian.css).
- `tests/test_report.py` still green (view-model contract unchanged from Phase 1;
  template-render smoke test asserts key sections + no Applaud strings).

---

## 7. Rollout / Packaging

- **Two PRs**, both off `master`, each branch → commit → push → PR → **wait for
  review** (never push to master, never self-merge).
  - **PR #1 (`chore/repo-cleanse`)** — this spec + all of Phase 1.
  - **PR #2** — Phase 2 redesign, branched after PR #1 merges.
- **Sequencing vs. 26C (default, correctable):** aim to land **both** phases
  before running 26B→26C so consultants get the redesigned report. Phase 1 alone
  already unblocks the 26C run (catalog + plain report work), so Phase 2 may slip
  to after the run if the timeline demands.
- The CLAUDE.md **Inter-font hazard** note is removed in the PR that lands
  Phase 2 (§6.4). If Phase 2 slips, the note stays accurate until then.

## 8. Testing Strategy

- Full suite green after each PR (`python -m pytest tests/`).
- Phase 1: updated `test_report.py`, `test_cli.py`, `test_skill_scripts.py`;
  deleted Applaud test modules; core/catalog tests untouched.
- Phase 2: template-render smoke test (sections present, no Applaud strings,
  self-contained), plus the impeccable audit as a manual gate.
- Spot-check against a real run: `report --old 26A --new 26B` on the existing
  catalog and eyeball the HTML.

## 9. Risks & Open Questions

- **`catalog_normalize` coupling** (§5.1) — resolved during implementation;
  default is removal.
- **Module coverage** — a new 26C file absent from `file_modules.json` groups
  under "Unclassified" rather than failing. Acceptable; the downloader writes
  `file_modules.json`, so new files are normally captured.
- **`RapidImplementationForCashManagement.xlsm`** manual-download hazard is
  unaffected and stays documented in CLAUDE.md.
- **Test count** drops materially (~8 test files removed). Expected and healthy;
  the plan records the before/after count.

---

## Appendix A — File disposition (modules/tools/tests)

**Remove:** `fbdi/{audit,audit_applaud,applaud_appmap,applaud_snapshot,`
`correspondence,build_mapping,populate_module,applaud_type}.py`;
`tools/{_gen_appmap_full.py,assemble_applaud_snapshot.py,`
`extract_applaud_snapshot.mjs}`; `tests/{test_audit,test_audit_applaud,`
`test_correspondence,test_applaud_appmap,test_applaud_snapshot,`
`test_populate_module,test_module_classifier,test_applaud_type}.py`;
`FBDI_to_ApplaudTables_Mapping.xlsx`; `FBDI_to_Applaud_AppMap.xlsx`;
`baselines/applaud/`. Conditional: `fbdi/catalog_normalize.py` +
`tests/test_catalog_normalize.py`.

**Untrack + ignore:** `FBDI_Master_Catalog.xlsx`.

**Rewrite:** `fbdi/report.py`, `fbdi/cli.py`, `fbdi/config.py`,
`fbdi/templates/report.html.j2` (Phase 2), `CLAUDE.md`, `README.md`,
`DESIGN.md`, `PRODUCT.md`, `docs/operator-guide.md`, `docs/developer-guide.md`,
`.claude/skills/fbdi-compare-release/**`, `.gitignore`,
`tests/{test_report,test_cli,test_skill_scripts}.py`.

**Keep untouched:** `fbdi/{clear,compare,diagnose,detect_header,utils,`
`_subprocess_util,type_parser,align,__init__,__main__}.py`;
`tools/download_and_clear.py`; core/catalog tests.

## Appendix B — Docs to archive (`git mv` → `docs/archive/`)

Plans: `2026-04-21-applaud-mapping-audit.md`,
`2026-05-01-fbdi-compliance-report.md`,
`2026-05-04-compliance-report-skill-stage9.md`,
`2026-05-04-fbdi-run-headless-pipeline.md` (+ `.png`),
`2026-06-02-applaud-compliance-audit.md`,
`2026-06-09-applaud-audit-first-run.md`,
`2026-06-11-applaud-field-correspondence.md`.

Specs: the design twins of each plan above
(`applaud-mapping-audit-design-revised`, `fbdi-compliance-report-design`,
`compliance-report-skill-stage9-design`, `fbdi-run-headless-pipeline-design`,
`applaud-compliance-audit-design`, `applaud-audit-first-run-design`,
`applaud-field-correspondence-design`).

Loose: `applaud-audit-first-run-notes.md`,
`AUDIT_RESULTS_applaud-compliance-audit.md`,
`AUDIT_RESULTS_field-correspondence*.md`, `AUDIT_RESULTS_plan_pass2.md`,
`AUDIT_RESULTS_plan_pass3.md`, `references/applaud-snapshot-extraction.md`.
(`handoff_phase2_rerun.md` is gitignored — delete on disk.)

**Keep in place:** `2026-04-14-smart-clear-*`, `2026-04-15-fbdi-master-catalog-*`,
`2026-04-20-catalog-subprocess-deadlock-*`, `2026-04-23-fbdi-compare-release-skill-*`,
`2026-04-24-handoff-docs-and-cleanup-*`, `2026-04-30-quarterly-rerun-and-module-capture-*`,
and this spec.
