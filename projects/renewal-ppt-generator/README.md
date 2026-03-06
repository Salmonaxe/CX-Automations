# Renewal and Opportunities Automation

This project contains automation scripts for Cisco renewals and new opportunities data exported from CS Console.

## What This Project Includes

- PPT generation for renewals (baseline and enhanced)
- PPT generation for new opportunities
- Interactive web viewer that combines renewals and new opportunities

## Recommended Script for Most Users

Use `src/create_renew_ops_ppt.py` for renewals. It includes threshold filtering and clearer summaries for multi-customer files.

## Prerequisites

- Python 3.11+
- PowerShell terminal
- CS Console Excel exports (`.xlsx`)

## Setup

```powershell
cd .\projects\renewal-ppt-generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest
```

## Where To Put Input Files

Store local exports under `data/`:

- `data/renewals/`
- `data/new-ops/`

Optional template files:

- `templates/<your-company-template>.pptx`
- Repo seed file: `templates/company-template.potx`
- Runtime file for `--template-pptx`: `templates/company-template.pptx`

## Script Reference

### 1) Baseline Renewals

File: `src/create_renewal_ppt.py`

Use when you need the original renewals flow.

```powershell
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx --template-pptx .\templates\company-template.pptx
```

Output:

- `<input>_product_<FY-range>.pptx`
- `<input>_service_<FY-range>.pptx`

### 2) Enhanced Renewals

File: `src/create_renew_ops_ppt.py`

Use for production renewals reporting.

```powershell
python .\src\create_renew_ops_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx --min-atr 100
python .\src\create_renew_ops_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx --min-atr 100 --template-pptx .\templates\company-template.pptx
```

Key behavior:

- Aggregates by `Deal Id`
- Supports all-customer files
- Adds summary slides for:
  - `All Customers`
  - each customer individually
- Adds monthly timeline slides

Output:

- `<input>_product_<FY-range>[_MIN_ATR_###K].pptx`
- `<input>_service_<FY-range>[_MIN_ATR_###K].pptx`

### 3) New Opportunities

File: `src/create_new_ops_ppt.py`

```powershell
python .\src\create_new_ops_ppt.py Q1FY26 Q3FY26 .\data\new-ops\new_ops.xlsx --min-tcv 100
python .\src\create_new_ops_ppt.py Q1FY26 Q3FY26 .\data\new-ops\new_ops.xlsx --min-tcv 100 --template-pptx .\templates\company-template.pptx
```

Key behavior:

- Aggregates by `Deal Id`
- Supports all-customer files
- Adds summary slides for:
  - `All Customers`
  - each customer individually
- Uses stage-based timeline colors

Output:

- `<input>_<FY-range>_TCV_MIN_<value>.pptx`

### 4) Interactive Viewer

File: `src/opps_viewer.py`

```powershell
streamlit run .\src\opps_viewer.py
```

Key behavior:

- Upload renewals and/or new opportunities exports
- Filter by fiscal range, account, stage, pulses, and thresholds
- Download filtered details as CSV

## Multi-Customer Behavior

For `create_renew_ops_ppt.py` and `create_new_ops_ppt.py`:

- Title slide shows customer scope
- Summary section includes one overall summary plus one per customer
- Account-level table and timeline slides still break out by `Account Name`

## Corporate Template Support

All three PPT generators support `--template-pptx` so output inherits your company theme.

- Supported template extension: `.pptx`
- The repo currently includes `templates/company-template.potx` as a starter asset
- Save or convert that `.potx` file to `templates/company-template.pptx` before using `--template-pptx`
- Example:
  `--template-pptx .\templates\company-template.pptx`
- Team convention: keep one canonical template in `templates/` and reuse it across all scripts

## Verification Coverage

The committed test suite is intentionally lightweight because customer exports
are not stored in git. Current smoke coverage checks:

- CLI `--help` execution for all three PPT generators
- `.pptx` template validation behavior without needing sample Excel files
- Python import/compile health for the Streamlit viewer and PPT scripts

## Git Hygiene

`projects/renewal-ppt-generator/.gitignore` excludes local data and generated presentations (`.xlsx`, `.xls`, `.pptx`). Keep these local and do not commit them.

## Data Source Notes

From CS Console:

1. Renewals: `Manage Pipeline -> Renewals Opportunities` (Line Details)
2. New opportunities: `Manage Pipeline -> New Opportunities` (Line Details)
3. Export and save `.xlsx` files locally under `data/`
