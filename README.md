# Oracle FBDI Pulldown

## What This Repo Does

This repo automates the comparison of Oracle FBDI (File-Based Data Import) template files (`.xlsm`) between release versions. The core asset is a Python comparison engine (`fbdi/` package) that detects template headers dynamically, diffs field-level changes between two releases, and outputs a structured Excel report. The legacy Selenium-based downloader and VBA macros that preceded the Python engine are archived in `reference/`.

---

## Repo Structure

```
oracle-fbdi-pulldown/
├── fbdi/           # Python package — comparison engine and CLI
├── tests/          # Unit tests (33 passing)
├── baselines/      # GITIGNORED — release template folders populated by downloader
├── reference/      # Version-controlled archive of legacy files — read-only
├── .gitignore
├── CLAUDE.md       # Persistent context for Claude Code sessions
├── NEXT_STEPS.md   # Ranked recommendations for next development work
└── README.md
```

---

## Setup

Requires Python 3.10+. No package installation needed — all dependencies are standard library plus `openpyxl`.

```bash
pip install openpyxl
```

For running tests:

```bash
pip install pytest
```

The legacy downloader (`reference/test.py`) also requires Chrome and Selenium:

```bash
pip install selenium webdriver-manager
```

---

## Usage

```bash
python -m fbdi compare --old 25d --new 26a
```

This compares all FBDI templates found in `baselines/25d/` against `baselines/26a/`, detecting headers dynamically and producing a `Comparison_Report_25D_26A.xlsx` at repo root.

```bash
python -m fbdi --help
python -m fbdi compare --help
```

---

## Baseline Management

`baselines/` holds release-named subdirectories (`25d/`, `26a/`, etc.) containing the downloaded `.xlsm` template files. This folder is gitignored — files are large binaries that should not be committed. Populate it by running the downloader (`reference/test.py`) or manually placing template files in the appropriate release folder.

---

## Reference Files

`reference/` is a read-only archive of legacy artifacts that preceded the Python engine. These files are version-controlled for historical context but are not part of the active pipeline:

| File | Description |
|---|---|
| `fbdi_compare.xlsm` | Legacy VBA macro workbook that compared FBDI templates |
| `Clear_FBDIs - 20210412.xlsm` | Legacy VBA macro that cleared/reset template files |
| `Oracle_26A_Comparison_Report.docx` | Sample VBA comparison output for release 26A |
| `test.py` | Dan's original Selenium downloader |

---

## Status

- **Phase 1 — Complete:** Python comparison engine (`fbdi/` package), CLI (`python -m fbdi compare`), 33 unit tests, dynamic header detection, 7-column Excel output
- **Phase 2 — In Progress:** Applaud compliance report automation (`report.py`) — reformatting comparison output for the Applaud change-tracking format
- **Phase 3 — Planned:** Full integrated pipeline (`python -m fbdi run`) — download, compare, and report in a single command
