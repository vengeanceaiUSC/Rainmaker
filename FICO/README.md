# FICO — Deterministic 3-Statement + DCF Pipeline

Architecture (no LLM numbers):

1. **Extraction** — SEC EDGAR companyfacts XBRL JSON (`us-gaap_*` tags)
2. **Standardization** — strict Pydantic models (IS / BS / CF)
3. **Forecast** — rules-based assumptions applied with Python/pandas math
4. **Valuation** — hardcoded DCF math on FCFF from the 3-statement model
5. **Excel inject (optional)** — CSV inputs → CFI template **Named Ranges** only

## CSV outputs (source of truth for agents)

- `output/3S_FICO_income_statement.csv` / `_balance_sheet.csv` / `_cash_flow.csv`
- `output/DCF_FICO_summary.csv` / `_annual_fcff.csv` / `_market_inputs.csv`

## Excel Completed Model (humans)

```bash
# 1) Ensure CSVs exist
python3 -m FICO.pipeline.run --ticker FICO

# 2) Stamp Named Ranges on CFI template + inject INPUTS only
python3 -m FICO.pipeline.prepare_template
python3 -m FICO.pipeline.inject_template

# Or one-shot:
python3 -m FICO.pipeline.run --ticker FICO --inject-excel
```

Produces `output/FICO_Completed_Model.xlsx`.

### Formula protection

- Named Ranges are attached **only** to input/driver cells (Revenue, COGS, SG&A, CapEx, WACC, shares, etc.).
- Formula cells (Gross Profit, Net Income, UFCF, Enterprise Value) have **no** Named Ranges and are never written.
- Loader always uses `openpyxl.load_workbook(..., data_only=False)`.
- **Open the xlsx in Excel** to recalculate formulas from the injected SEC inputs.

Horizontal fill uses `*_Start` Named Ranges (e.g. `IS_Revenue_Start`) so inserting rows in the template does not break the injector — the name moves with the cell.

## Package layout

```
FICO/pipeline/
  edgar.py / map_xbrl.py / models.py / three_statement.py / dcf.py
  export_csv.py
  named_range_map.py   # INPUT-only Named Range registry + CSV maps
  prepare_template.py  # build FICO/templates/CFI_Template.xlsx
  inject_template.py   # CSV → Named Ranges → FICO_Completed_Model.xlsx
  run.py
```
