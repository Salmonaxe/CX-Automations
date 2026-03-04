# Runbook

## Local Setup

```powershell
cd .\projects\renewal-ppt-generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## Script Selection

- Use `src/create_renewal_ppt.py` for baseline renewals PPT output.
- Use `src/create_renew_ops_ppt.py` for enhanced renewals output (`--min-atr`, monthly timeline slides).
- Use `src/create_new_ops_ppt.py` for new opportunities PPT output (`--min-tcv`).
- Use `src/opps_viewer.py` for interactive filtering and integrated timeline review.

## Execute

```powershell
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\renewals.xlsx
python .\src\create_renew_ops_ppt.py Q1FY26 Q3FY26 .\data\renewals.xlsx --min-atr 100
python .\src\create_new_ops_ppt.py Q1FY26 Q3FY26 .\data\new_ops.xlsx --min-tcv 100
streamlit run .\src\opps_viewer.py
```

## Troubleshooting

- If PowerShell execution policy blocks `.ps1` scripts, run the Python commands directly as shown above.
- Confirm required Excel columns are present for the selected script type.
- Confirm input files are `.xlsx`.
- If `streamlit` command is not found, run `python -m streamlit run .\src\opps_viewer.py`.
