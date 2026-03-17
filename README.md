# Oracle FBDI Pulldown

Automates downloading Oracle FBDI (File-Based Data Import) template files (`.xlsm`) from Oracle's documentation site and running comparison macros against a baseline.

## What It Does

1. **Creates output folders** — `Blank Copies/` and `Originals/` are created (or cleared if they already exist)
2. **Scrapes & downloads FBDI templates** — Uses Selenium to navigate Oracle's documentation for the following modules and download all `.xlsm` template files:
   - Project Management
   - Financials
   - Procurement
   - Supply Chain & Manufacturing
3. **Runs Excel macros** — Uses `xlwings` to run macros in `Clear_FBDIs` and `fbdi_compare.xlsm` to process and compare the downloaded templates against a baseline

## Files

| File | Description |
|---|---|
| `test.py` | Main script — downloads templates and runs macros |
| `Clear_FBDIs - 20210412.xlsm` | Excel macro workbook that clears/resets FBDI templates |
| `fbdi_compare.xlsm` | Excel macro workbook that compares new templates against a baseline |

## Requirements

- Python 3.x
- Google Chrome
- Microsoft Excel (for `xlwings` macro execution)

Install dependencies:
```bash
pip install selenium webdriver-manager xlwings requests
```

## Usage

Update the Oracle release version in the `base_urls` inside `test.py` if needed (currently set to `26a`), then run:

```bash
python test.py
```

Downloaded templates will be saved to the `Originals/` folder. The `Blank Copies/` folder is used for processed output.
