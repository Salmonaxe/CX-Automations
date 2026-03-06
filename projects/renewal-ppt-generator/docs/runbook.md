# Runbook

## Purpose

Operational steps for running renewals and new opportunities reports from CS Console exports.

## 1) One-Time Local Setup

```powershell
cd .\projects\renewal-ppt-generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## 2) Prepare Inputs

Place files here:

- Renewals export: `data/renewals/<file>.xlsx`
- New opportunities export: `data/new-ops/<file>.xlsx`
- Shared corporate template: `templates/company-template.pptx`

## 3) Run Commands

### Enhanced Renewals (recommended)

```powershell
python .\src\create_renew_ops_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx --min-atr 100
python .\src\create_renew_ops_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx --min-atr 100 --template-pptx .\templates\company-template.pptx
```

### New Opportunities

```powershell
python .\src\create_new_ops_ppt.py Q1FY26 Q3FY26 .\data\new-ops\new_ops.xlsx --min-tcv 100
python .\src\create_new_ops_ppt.py Q1FY26 Q3FY26 .\data\new-ops\new_ops.xlsx --min-tcv 100 --template-pptx .\templates\company-template.pptx
```

### Baseline Renewals (legacy)

```powershell
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\renewals\renewals.xlsx --template-pptx .\templates\company-template.pptx
```

### Interactive Viewer

```powershell
streamlit run .\src\opps_viewer.py
```

## 4) Output Locations

Scripts write output PPT files to the current working directory (normally `projects/renewal-ppt-generator`).

## 5) Multi-Customer Inputs

Enhanced renewals and new opportunities scripts support all-customer exports and produce:

- one overall summary (`All Customers`)
- one summary slide per customer

## Troubleshooting

- If VS Code shows missing imports, ensure interpreter is `projects/renewal-ppt-generator/.venv/Scripts/python.exe`.
- If `streamlit` is not found, run `python -m streamlit run .\src\opps_viewer.py`.
- If no output is generated, verify date range (`Q?FY??`) and required Excel columns.
- Keep input files as `.xlsx`.
