# Oracle FBDI Pulldown

Automates comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) across Oracle Cloud quarterly releases. Produces three deliverables: a field-level diff report (`Comparison_Report_<OLD>_<NEW>.xlsx`), a per-release snapshot catalog (`FBDI_Master_Catalog.xlsx`), and an HTML/PDF release change report (`FBDI_Change_Report_<OLD>_<NEW>.html`/`.pdf`).

> **Running it:** see [`docs/operator-guide.md`](docs/operator-guide.md).
> **Developing on it:** see [`docs/developer-guide.md`](docs/developer-guide.md).

---

## Setup

You need Python 3.14+, Google Chrome (the downloader drives Chrome via Selenium), and Windows. Mac and Linux are untested. PDF rendering also needs MSYS2 mingw64 GTK; see CLAUDE.md for the install steps.

```bash
pip install -r requirements.txt
```

---

## Running a quarterly refresh

### Option A: through Claude Code (recommended)

The repo ships the `fbdi-compare-release` skill at `.claude/skills/fbdi-compare-release/`. In a Claude Code session, say something like:

> Compare 26A to 26B

Claude invokes the skill and walks you through the orchestrated pipeline: preflight, version resolve, download, smart-clear, compare, catalog, summary, post-run verification, and a final HTML/PDF release change report. There are human-in-the-loop checkpoints along the way for the edge cases. Expect 35–50 minutes end to end; downloads dominate. The full workflow lives in `.claude/skills/fbdi-compare-release/SKILL.md`.

### Option B: CLI directly

For Python-first workflows:

```bash
# Download + smart-clear templates for a new release (~15–20 min)
python tools/download_and_clear.py 26B

# Compare two releases → Comparison_Report_26A_26B.xlsx
python -m fbdi compare --old 26A --new 26B

# Update the per-release snapshot catalog
python -m fbdi catalog --release 26B

# Generate the release change report (HTML; add --pdf for the PDF)
python -m fbdi report --old 26A --new 26B

# Diagnose header-detection outcomes per tab
python -m fbdi diagnose --old baselines/26A/originals --new baselines/26B/originals
```

Run `python -m fbdi --help` or `python -m fbdi <cmd> --help` for flag details.

---

## Known hazards

- **`RapidImplementationForCashManagement.xlsm` is not auto-downloadable.** It's an Oracle Rapid Implementation (FSM) template and isn't hosted on the Oracle docs pages the scraper walks. Fetch it manually from Oracle Fusion (Setup and Maintenance, hamburger menu, Search, "Create Banks, Branches, and Accounts in Spreadsheet") and drop it into `baselines/<VER>/originals/` before comparing. HITL #2 in the skill walks you through this.

See `CLAUDE.md` for the full list of hazards and the resolved-issues log.

---

## Repo structure

```
FBDI-Compare-Process-Oracle/
├── fbdi/                      # Python comparison/catalog/clear engine
├── tools/                     # Selenium downloader (download_and_clear.py)
├── tests/                     # 233 tests (pytest)
├── .claude/skills/            # Project-level Claude Code skills
│   └── fbdi-compare-release/  # Orchestrator for quarterly refreshes
├── docs/
│   ├── operator-guide.md      # End-to-end pipeline walkthrough
│   ├── developer-guide.md     # Codebase tour and extension guide
│   ├── archive/               # Historical narrative docs (audits, gap findings)
│   └── superpowers/           # Design specs and implementation plans
├── baselines/                 # gitignored: downloaded xlsm + file_modules.json per release
├── reference/                 # Read-only archive of legacy VBA + scripts
├── baseline_files.txt         # Inventory of expected downloads per release
├── FBDI_Master_Catalog.xlsx   # Per-release snapshot catalog (gitignored; regenerable)
├── requirements.txt
├── CLAUDE.md                  # Persistent Claude Code context
└── README.md
```

---

## Testing

```bash
python -m pytest tests/              # full suite (233 tests)
python -m pytest tests/test_clear.py -v
```

---

## Reference files

`reference/` is a read-only archive of the pre-Python pipeline.

| File | Description |
|---|---|
| `fbdi_compare.xlsm` | Legacy VBA macro that compared FBDI templates |
| `Clear_FBDIs - 20210412.xlsm` | Legacy VBA macro that cleared template files |
| `Oracle_26A_Comparison_Report.docx` | Sample VBA comparison output for 26A |
| `test.py` | Dan's original Selenium downloader |

---

## Status

Shipped: the comparison engine, the CLI (`fbdi compare`, `catalog`, `diagnose`, `report`), smart clearing, the `download_and_clear` Selenium driver, the FBDI master catalog, the `fbdi-compare-release` Claude Code skill, and the HTML/PDF release change report. 233 tests passing.

Planned: Phase 2 restyles the release change report with the Definian design system (HTML-first, opt-in PDF). Design at `docs/superpowers/specs/2026-09-14-repo-cleanse-and-report-redesign-design.md`.
