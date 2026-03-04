# Renewal and Opportunities Automation

This project contains Cisco opportunity automation scripts for:

- Renewals PPT generation (baseline and enhanced)
- New Opportunities PPT generation
- Interactive timeline analysis in a web UI

## Scripts

- `src/create_renewal_ppt.py`
Classic renewals PPT generator (product + service outputs).

- `src/create_renew_ops_ppt.py`
Enhanced renewals generator with minimum ATR filtering and additional monthly timeline slides.

- `src/create_new_ops_ppt.py`
New opportunities PPT generator with stage-based timeline coloring and minimum TCV filtering.

- `src/opps_viewer.py`
Streamlit web app to explore renewals and new opportunities together with interactive filters.

## Setup

```powershell
cd .\projects\renewal-ppt-generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## Run

### 1) Baseline Renewals PPT

```powershell
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\renewals.xlsx
```

### 2) Enhanced Renewals PPT

```powershell
python .\src\create_renew_ops_ppt.py Q1FY26 Q3FY26 .\data\renewals.xlsx --min-atr 100
```

### 3) New Opportunities PPT

```powershell
python .\src\create_new_ops_ppt.py Q1FY26 Q3FY26 .\data\new_ops.xlsx --min-tcv 100
```

### 4) Interactive Viewer

```powershell
streamlit run .\src\opps_viewer.py
```

## Typical Outputs

- Renewals scripts produce `.pptx` files in the current working directory.
- Viewer provides an integrated timeline and CSV export from the browser.

## Multi-Customer Runs

- `create_renew_ops_ppt.py` and `create_new_ops_ppt.py` support all-customer input files.
- Title slides list multiple customers.
- Summary section now includes:
  - one overall `All Customers` summary slide
  - one summary slide per customer

## Git Hygiene

- `projects/renewal-ppt-generator/.gitignore` excludes local `.xlsx` and generated `.pptx` files.
- Keep customer data exports and generated decks local to your machine.

## Data Source Notes

From CS Console:

1. Renewals: `Manage Pipeline -> Renewals Opportunities` (Line Details export)
2. New Opportunities: `Manage Pipeline -> New Opportunities` (Line Details export)
3. Save exported `.xlsx` files and pass them to the appropriate script
