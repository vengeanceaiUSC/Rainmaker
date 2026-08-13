# LBO — FICO leveraged buyout model

Editable Excel LBO built from SEC-sourced FICO figures. Open the completed workbook in Excel / Google Sheets so formulas recalculate.

## Click to download (raw)

| File | Raw download |
|---|---|
| **FICO LBO (completed, editable)** | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/FICO_LBO_Completed_Model.xlsx |
| Same file under `LBO/output/` | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/output/FICO_LBO_Completed_Model.xlsx |
| LBO Model ABC template | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/LBO_Model_ABC.xlsx |
| Wall Street Prep LBO + DCF sample | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/WSP_LBO_with_DCF_Sample.xlsx |
| CFI Leveraged Finance | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/CFI_Leveraged_Finance_Template.xlsx |
| CFI Debt Capacity | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/CFI_Debt_Capacity_Model_Template.xlsx |
| Exinfm LBO + DCF (xlsx) | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/exinfm_LBO_DCF_Model.xlsx |

GitHub “View” (may not preview binary Excel — use raw links above):

- https://github.com/vengeanceaiUSC/Rainmaker/blob/cursor/lbo-44dc/FICO_LBO_Completed_Model.xlsx

## Sheets in `FICO_LBO_Completed_Model.xlsx`

| Sheet | Purpose |
|---|---|
| `Cover` | Target, units, SEC source hyperlinks |
| `01_Assumptions` | Yellow/blue inputs — premium, leverage turns, growth, exit multiple |
| `02_Sources_Uses` | Purchase price, debt, sponsor equity plug |
| `03_Operating` | FY2023–25 hist + 5-year forecast |
| `04_Debt_Schedule` | Interest, TLA amort, cash sweep waterfall |
| `05_Returns` | Exit equity, MoIC, IRR |
| `06_Sensitivity` | Exit multiple × premium MoIC / IRR grid |

## Rebuild

```bash
python3 LBO/build_fico_lbo_model.py
```

Writes:

- `LBO/output/FICO_LBO_Completed_Model.xlsx`
- `FICO_LBO_Completed_Model.xlsx` (repo root, easy raw download)

Educational / research use only — not investment advice.
