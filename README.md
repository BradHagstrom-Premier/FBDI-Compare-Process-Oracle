# Oracle FBDI Pulldown

Automates comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) across Oracle Cloud quarterly releases. Produces two outputs: a field-level diff report (`Comparison_Report_<OLD>_<NEW>.xlsx`) and a per-release snapshot catalog (`FBDI_Master_Catalog.xlsx`).

> **Running it:** see [`docs/operator-guide.md`](docs/operator-guide.md).
> **Developing on it:** see [`docs/developer-guide.md`](docs/developer-guide.md).

---

## Setup

**Required:**
- Python 3.14+
- Google Chrome (Selenium dependency for the downloader)
- Windows (the supported platform; Mac/Linux may work but are untested)

```bash
pip install -r requirements.txt
```

---

## Running a quarterly refresh

### Option A — through Claude Code (recommended)

The repo ships with the `fbdi-compare-release` skill at `.claude/skills/fbdi-compare-release/`. In a Claude Code session, say something like:

> Compare 26A to 26B

Claude invokes the skill and walks you through an 8-stage orchestrated pipeline (environment preflight → version resolve → download → smart-clear → compare → catalog → summary → post-run verification), with six human-in-the-loop checkpoints for edge cases. Expected wall time ≈ 35–50 minutes (downloads dominate). See `.claude/skills/fbdi-compare-release/SKILL.md` for the full workflow.

### Option B — CLI directly

For Python-first workflows:

```bash
# Download + smart-clear templates for a new release (~15–20 min)
python tools/download_and_clear.py 26B

# Compare two releases → Comparison_Report_26A_26B.xlsx
python -m fbdi compare --old 26A --new 26B

# Update the per-release snapshot catalog
python -m fbdi catalog --release 26B

# Diagnose header-detection outcomes per tab
python -m fbdi diagnose --old baselines/26A/originals --new baselines/26B/originals
```

Run `python -m fbdi --help` or `python -m fbdi <cmd> --help` for flag details.

---

## Known hazards

- **`RapidImplementationForCashManagement.xlsm` is not auto-downloadable.** It's an Oracle Rapid Implementation (FSM) template, not hosted on Oracle docs pages. Fetch it manually from Oracle Fusion (Setup and Maintenance → hamburger menu → Search → "Create Banks, Branches, and Accounts in Spreadsheet") and drop it into `baselines/<VER>/originals/` before comparing. The skill's HITL #2 walks you through this.

See `CLAUDE.md` for the full list of hazards and the resolved-issues log.

---

## Repo structure

```
FBDI-Compare-Process-Oracle/
├── fbdi/                      # Python comparison/catalog/clear engine
├── tools/                     # Selenium downloader (download_and_clear.py)
├── tests/                     # 241 unit tests (pytest)
├── .claude/skills/            # Project-level Claude Code skills
│   └── fbdi-compare-release/  # Orchestrator for quarterly refreshes
├── docs/
│   ├── operator-guide.md      # End-to-end pipeline walkthrough
│   ├── developer-guide.md     # Codebase tour and extension guide
│   ├── archive/               # Historical narrative docs (audits, gap findings)
│   └── superpowers/           # Design specs and implementation plans
├── baselines/                 # GITIGNORED — downloaded xlsm per release + applaud_snapshot.json
├── reference/                 # Read-only archive of legacy VBA + scripts
├── baseline_files.txt         # Inventory of expected downloads per release
├── FBDI_Master_Catalog.xlsx   # Per-release snapshot catalog (git-tracked)
├── requirements.txt
├── CLAUDE.md                  # Persistent Claude Code context
└── README.md
```

---

## Testing

```bash
python -m pytest tests/              # full suite (241 tests)
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

- **Shipped:** comparison engine, CLI (`fbdi compare` / `catalog` / `diagnose` / `report`), smart clearing, `download_and_clear` Selenium driver, FBDI master catalog, Applaud mapping audit, `fbdi-compare-release` Claude Code skill, `FBDI_to_ApplaudTables_Mapping.xlsx` mapping complete (no TBD rows), 320-test suite.
- **Planned:** `python -m fbdi run` (chained pipeline).
