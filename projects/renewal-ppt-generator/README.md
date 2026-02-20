# Renewal PPT Generator

Generates Cisco renewal opportunity PowerPoints for product and service opportunities using a CS Console Excel export.

## Inputs

- `initial_fy` format `Q?FY??` (example: `Q1FY26`)
- `final_fy` format `Q?FY??` (example: `Q3FY26`)
- `excel_filename` `.xlsx` exported from CS Console renewals opportunities

## Setup

```powershell
cd .\projects\renewal-ppt-generator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## Run

```powershell
python .\src\create_renewal_ppt.py Q1FY26 Q3FY26 .\data\_Renewal_Opportunities_durgell_1770026732.xlsx
```

## Output

For input `your_file.xlsx`, script generates:

- `your_file_product_Q1FY26-Q3FY26.pptx`
- `your_file_service_Q1FY26-Q3FY26.pptx`

## Data Source Notes

From CS Console:

1. Select customer and go to `Manage Pipeline -> Renewals Opportunities`
2. Choose `All Risk ATR` or `High Risk ATR`, and `Line Details`
3. Export and save the `.xlsx`
4. Use that file as script input
