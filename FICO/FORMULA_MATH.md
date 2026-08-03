# Updated DCF / 3-Statement math (Excel formulas)

Only **yellow input cells** are typed numbers. Calculated lines are Excel formulas.

## DCF identities

| Line | Formula |
|---|---|
| Cash Taxes | `=EBIT × TaxRate` → `E22 = E21*$D$5` |
| Capex (standalone) | `=Capex0 × (1+g_capex)^period` → `E24 = $D$15*(1+$D$16)^E19` |
| Capex (joined) | `='03_Three_Statement'!J68` |
| Unlevered FCF | `=EBIT − CashTaxes + D&A − Capex − ΔNWC` |
| TV (perpetuity) | `=UFCF_t × (1+g) / (WACC − g)` → `N18` |
| TV (exit multiple) | `=EV/EBITDA × (EBIT_t + D&A_t)` → `N19` |
| TV used | `=AVERAGE(N18:N19)` → `N20` |
| Enterprise Value | `=XNPV(WACC, TransactionCFs, Dates)` → `D32` |
| Equity Value | `=EV + Cash − Debt` → `D35 = D32+D33-D34` |
| Equity / share | `=Equity / Shares` → `D37 = D35/D12` |

## EBIT path

- **Standalone:** `E21 = $D$3*(1+E4)` then `F21 = E21*(1+F4)` … (D3 = FY25 OpInc 924,850)
- **Joined:** `E21 = '03_Three_Statement'!J26 - J28 - J29 - J30` (GP − SG&A − R&D − D&A)

## 3-Statement forecast dependencies

| Line | Formula |
|---|---|
| Revenue | `J24 = I24*(1+J7)` |
| COGS | `J25 = J24*J8` |
| Gross Profit | `J26 = J24-J25` |
| SG&A | `J28 = J9` |
| R&D | `J29 = J10` |
| Capex | `J68 = J18` |
| ΔNWC | `J89 = J88-I88` |

## Assumptions used

- Tax rate `18.77%` · WACC `9.6%` · g `3%` · EV/EBITDA `22x`
- Capex0 `39,407` · Capex growth `8%`
- FY25 OpInc `924,850` · D&A `14,952`

Source: https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm
