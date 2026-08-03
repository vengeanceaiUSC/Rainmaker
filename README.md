# Rainmaker — FICO DCF / SEC Valuation Pack

SEC EDGAR fundamentals for **Fair Isaac Corp (FICO)** plus a working investor DCF Excel model.

## Quick start

1. Open [`fico-valuation/FICO_DCF_Valuation_Model.xlsx`](fico-valuation/FICO_DCF_Valuation_Model.xlsx)
2. Read tab `00_ReadMe`, then `01_Template_Links` for external downloadable templates
3. Confirm history on `03_IS_Historical` / `04_BS_Historical` / `05_CF_Historical`
4. Edit yellow/blue inputs on `06_Assumptions`
5. Read value on `07_DCF` and `10_Output`

Rebuild the workbook anytime:

```bash
python3 fico-valuation/build_fico_model.py
```

## Key files

| Path | Description |
|---|---|
| `fico-valuation/FICO_DCF_Valuation_Model.xlsx` | Full DCF + historical statements + checklist |
| `fico-valuation/TEMPLATE_LINKS.md` | Downloadable valuation template URLs |
| `fico-valuation/FICO_NUMBERS_CHEATSHEET.md` | Audited line items you need for inputs |
| `fico-valuation/data/fico_companyfacts.json` | Raw SEC companyfacts XBRL pull |
| `fico-valuation/data/fico_xbrl_by_period_end.csv` | Cleaned annual XBRL extract |
| `fico-valuation/build_fico_model.py` | Regenerates the Excel model |

## Primary SEC sources

- CIK: `0000814547`
- [FY2025 10-K](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)
- [Q3 FY2026 10-Q](https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm)
- [Companyfacts API](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## Recommended external templates (direct downloads)

- Damodaran FCFF simple ginzu: https://pages.stern.nyu.edu/~adamodar/pc/fcffsimpleginzu.xlsx
- Damodaran FCFF full ginzu: https://pages.stern.nyu.edu/~adamodar/pc/fcffginzu.xls
- CFI compact DCF: https://corporatefinanceinstitute.com/assets/DCF-Valuation-Compact-Complete.xlsx

See [`fico-valuation/TEMPLATE_LINKS.md`](fico-valuation/TEMPLATE_LINKS.md) for the full list.

Educational / research use only — not investment advice.
