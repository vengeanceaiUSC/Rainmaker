# FICO — Deterministic 3-Statement + DCF Pipeline

Architecture (no LLM numbers):

1. **Extraction** — SEC EDGAR companyfacts XBRL JSON (`us-gaap_*` tags)
2. **Standardization** — strict Pydantic models (IS / BS / CF)
3. **Forecast** — rules-based assumptions applied with Python/pandas math
4. **Valuation** — hardcoded DCF math on FCFF from the 3-statement model

Outputs are **separate CSV files** (readable text; no binary `.xlsx`):

- `output/3S_FICO_income_statement.csv`
- `output/3S_FICO_balance_sheet.csv`
- `output/3S_FICO_cash_flow.csv`
- `output/3S_FICO_assumptions.csv`
- `output/3S_FICO_validation.csv`
- `output/DCF_FICO_summary.csv`
- `output/DCF_FICO_annual_fcff.csv`
- `output/DCF_FICO_market_inputs.csv`
- `output/DCF_FICO_notes.csv`

## Run

```bash
pip install -r requirements.txt
python3 -m FICO.pipeline.run --ticker FICO
```

Useful flags: `--force-refresh`, `--wacc 0.096`, `--g 0.03`, `--share-price 1046.23`, `--shares 21597.635`, `--use-fy-net-debt`, `--tv-method exit|gordon`.

## Package layout

```
FICO/pipeline/
  edgar.py           # SEC companyfacts HTTP client + cache
  map_xbrl.py        # us-gaap tag map → Pydantic
  models.py          # IncomeStatement / BalanceSheet / CashFlowStatement
  three_statement.py # pandas frames, BS balance asserts, forecast
  dcf.py             # FCFF + DCF = Σ CF/(1+WACC)^t + TV/(1+WACC)^n
  export_csv.py      # separate 3S and DCF CSV exporters (pandas.to_csv)
  run.py             # CLI orchestration
```
