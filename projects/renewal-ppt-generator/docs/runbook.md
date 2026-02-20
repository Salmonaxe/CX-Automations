# Runbook

## Local Setup

```powershell
cd .\projects\renewal-ppt-generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## Execute

```powershell
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\input.xlsx
```

## Troubleshooting

- If script execution is blocked by policy, run Python directly as above (not `.ps1`).
- Confirm the input file includes all required columns from CS Console export.
- Confirm file extension is `.xlsx`.
