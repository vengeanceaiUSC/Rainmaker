# FICO — Deterministic 3-Statement + DCF Pipeline

Architecture (no LLM numbers):

1. **Extraction** — SEC EDGAR companyfacts XBRL JSON (`us-gaap_*` tags)
2. **Standardization** — strict Pydantic models (IS / BS / CF)
3. **Forecast** — rules-based assumptions applied with Python/pandas math
4. **Valuation** — hardcoded DCF math on FCFF from the 3-statement model

Outputs are **separate** workbooks (never joined):

- `output/3S_FICO.xlsx` — historical + forecast 3 statements
- `output/DCF_FICO.xlsx` — DCF fed by 3-statement FCFF

## Run

```bash
pip install -r requirements.txt
python3 -m FICO.pipeline.run --ticker FICO
```

Useful flags: `--force-refresh`, `--wacc 0.096`, `--g 0.03`, `--share-price 1046.23`, `--shares 21597.635`, `--use-fy-net-debt`.

## Package layout

```
FICO/pipeline/
  edgar.py           # SEC companyfacts HTTP client + cache
  map_xbrl.py        # us-gaap tag map → Pydantic
  models.py          # IncomeStatement / BalanceSheet / CashFlowStatement
  three_statement.py # pandas frames, BS balance asserts, forecast
  dcf.py             # FCFF + DCF = Σ CF/(1+WACC)^t + TV/(1+WACC)^n
  export_excel.py    # separate 3S and DCF Excel exporters
  run.py             # CLI orchestration
```
