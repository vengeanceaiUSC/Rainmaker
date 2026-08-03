# Model errors found and fixed

## 3-Statement
| Error | Fix |
|---|---|
| Forecast COGS% was **37%** (CFI sample) | **17.77%** from FY25 |
| Forecast SG&A/R&D was **$25k/$10k** | FY25 **513,028/188,347** grown with revenue |
| Forecast tax was **28%** | **18.77%** FY25 ETR |
| AR/Inv days 18/73 | ~97 AR days, **0** inventory |
| Capex forecast ~15k sample | From FY25 **39,407** growing 8%/yr |
| Year headers blank (formulas) | Hardcoded **2021–2030** |

## DCF
| Error | Fix |
|---|---|
| Year-1 EBIT ~1,048k > FY25 OpInc | Re-anchored to FY25 OpInc **924,850** then grown |
| D&A ~18k vs FY25 14,952 | Anchored to **14,952** × rev scale |
| Missing FY25 base year display | Row 16 shows FY25 EBIT/D&A/Capex/Tax |
| Capex flat | Grows 8%/yr |
| EV/EBITDA 25x extreme vs PG TV | **22x**; still show PG / multiple / average |
| Blank taxes/UFCF/TV in viewer | Computed values written + chart |

Intrinsic equity/share (model): **$687.55** vs price $1,046.23

10-K: https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm

## Forecast path (USD thousands)

| Year | Growth | EBIT | Capex | UFCF |
|---|---:|---:|---:|---:|
| 2026 | 12% | 1,031,318 | 42,560 | 752,304 |
| 2027 | 10% | 1,149,880 | 45,964 | 850,856 |
| 2028 | 9% | 1,270,112 | 49,641 | 947,058 |
| 2029 | 8% | 1,398,473 | 53,613 | 1,050,674 |
| 2030 | 7% | 1,533,660 | 57,902 | 1,160,650 |
