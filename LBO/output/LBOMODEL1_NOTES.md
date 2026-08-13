# vengeanceaiUSC-LBOMODEL1 — ZoomInfo inputs in LBO

Workbook: [`vengeanceaiUSC_LBOMODEL1.xlsx`](../../vengeanceaiUSC_LBOMODEL1.xlsx)

## What changed
Core **`LBO`** sheet inputs overwritten from ZoomInfo public filings (BMC sample replaced). Original IRR / cash-on-cash returns block retained and refreshed for the ZoomInfo case. `ZI_*` tabs kept for S&M → AI detail.

## Key inputs ($ millions except per share)

| Input | Value | Source |
|---|---|---|
| Ticker | GTM (legacy ZI) | SEC |
| Share price | $4.31 | Yahoo/GTM ~2026-08-12 |
| Diluted / shares out | 292.3m | 10-Q 2026-06-30 |
| Offer premium | 25% → **$5.39/sh** | Model assumption |
| LTM EBITDA | **$314.5** | FY2025 OpInc $225.7 + D&A $88.8 |
| Gross debt | **$1,269.9** | LT + current, 10-Q 2026-06-30 |
| Cash + ST investments | **$150.1** | 10-Q 2026-06-30 |
| Exit EV/EBITDA | **8.0x** | Model assumption |
| New debt | **5.0x** EBITDA (TLA/TLB/Notes) | Model assumption |

## Illustrative returns (open in Excel)
Base case in-file: sponsor **MoIC ~3.0x**, **IRR ~24%** at 8x exit (depends on forecast / paydown). SG&A margin path embeds ~500 bps S&M AI efficiency by Y3.

## Filings
- FY2025 10-K: https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm
- Companyfacts: https://data.sec.gov/api/xbrl/companyfacts/CIK0001794515.json

Educational / research use only — not investment advice.
