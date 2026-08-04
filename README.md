# Rainmaker — Deterministic SEC → 3-Statement → DCF

Python pipeline that builds a **3-statement financial model** and **DCF valuation** from SEC 10-K XBRL — without LLM number extraction.

## Architecture

| Stage | Method | Role |
|---|---|---|
| Extraction | SEC EDGAR companyfacts API | Pull `us-gaap` XBRL tags (not HTML scrape) |
| Standardization | Pydantic | Force tags into IS / BS / CF models |
| Forecasting | Rules + Python math | Growth/margins → projected statements |
| Valuation | Hardcoded DCF math | FCFF discounted at WACC |

Flow: **XBRL → 3-statement → DCF** (separate Excel outputs, never one joined workbook).

## Quick start

```bash
pip install -r requirements.txt
python3 -m FICO.pipeline.run --ticker FICO
```

Outputs:

- [`FICO/output/3S_FICO.xlsx`](FICO/output/3S_FICO.xlsx)
- [`FICO/output/DCF_FICO.xlsx`](FICO/output/DCF_FICO.xlsx)
- [`FICO/output/FICO_summary.json`](FICO/output/FICO_summary.json)

Details: [`FICO/README.md`](FICO/README.md)

## Primary SEC sources (FICO)

- CIK: `0000814547`
- [FY2025 10-K](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)
- [Companyfacts API](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## Other folders

- [`DCF resources/`](DCF%20resources/) — pristine CFI templates (reference only)
- [`fico-valuation/`](fico-valuation/) — earlier research notes / cheat sheets

Educational / research use only — not investment advice.
