# vengeanceaiUSCMODEL8 — Assumptions List

**PDF download:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusmodel8-d3ac/FICO/output/MODEL8_FICO_ASSUMPTIONS_LIST.pdf

Plain list only (no charts). Paste into Google Docs via File → Open if needed.

## 7: Revenue Growth
- **What it is:** Year-over-year % increase in revenue for each forecast year.
- **How set in model:** Yellow POLICY input: 27% / 16% / 13% / 10% / 7%.
- **Why this choice:** Y1 ≈ company FY2026 revenue guidance (~$2.53B / FY25 $1.991B − 1 ≈ 27%). Then fade toward terminal g=3% (not straight-lined). This is the main forward-looking judgment; it drives Rev → GP → EBT → Net Earnings → CF.
- **Source:** [SEC EX-99.1 Q3 FY2026 — updated FY2026 revenue guidance $2.53B](https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm)

## 8: COGS % of Revenue
- **What it is:** Cost of revenues as a % of sales (gross margin = 1 − this).
- **How set in model:** Excel equation: I25/I24 (FY25 COGS ÷ FY25 Revenue) held flat.
- **Why this choice:** Locks in latest reported cost structure (~17.8%). Scores mix is high-margin; holding FY25 is conservative vs further mix shift.
- **Source:** [SEC 10-K FY2025 — Cost of revenues / Total revenues](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 9: SG&A % of Revenue
- **What it is:** Operating opex (SG&A) as % of sales.
- **How set in model:** Equation: MAX(15%, (I28 − 10,922)/I24 − 75bps × year). Strips FY25 restructuring; then −0.75% of sales each year.
- **Why this choice:** MODEL8: restore software operating leverage. Floor 15%; grind −75 bps/yr so SG&A can scale toward mid-teens as revenue expands (not stuck ~24–25%).
- **Source:** [SEC 10-K FY2025 — SG&A + restructuring note](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 10: R&D % of Revenue
- **What it is:** Research & development expense as % of sales (IS label: R&D, not Rent).
- **How set in model:** Excel equation: I29/I24 held flat (~9.46%).
- **Why this choice:** FICO reinvests steadily in analytics/software. Flat % grows dollars with revenue.
- **Source:** [SEC 10-K FY2025 — Research and development](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 11: D&A % of Avg PP&E
- **What it is:** Depreciation rate from FY25; dollars applied to average PP&E so new CapEx is depreciated.
- **How set in model:** Rate = I30/I44 (FY25 DA ÷ FY25 PPE). Forecast DA$ = ((Open + Open+CapEx)/2) × rate.
- **Why this choice:** MODEL6 fix: prior model depreciated only opening PP&E, so ~$96M of forecast CapEx never hit D&A/FCFF add-back. Average PP&E (Open ↔ pre-DA close) depreciates additions in the year they are placed in service.
- **Source:** [SEC 10-K FY2025 — D&A and PP&E](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 12: Interest % of Debt Open
- **What it is:** Interest expense as % of opening debt balance.
- **How set in model:** Excel equation: I101/I98 (FY25 interest ÷ FY25 opening debt in schedule).
- **Why this choice:** Approximates average coupon / cost of debt on the book. Debt held flat (issuance = 0), so interest stays linked to that stock.
- **Source:** [SEC 10-K FY2025 — Interest expense; see also 8-K 6.250% notes](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 13: Tax Rate (% of EBT)
- **What it is:** Effective book tax rate applied to forecast EBT.
- **How set in model:** Excel equation: I35/I33 (FY25 tax ÷ simplified FY25 EBT in this template).
- **Why this choice:** Uses reported effective rate (~19% on template EBT; ~18.8% on 10-K EBT). UFCF in DCF also uses unlevered EBIT × t.
- **Source:** [SEC 10-K FY2025 — Provision for income taxes / Income before taxes](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 15: Operating NWC % of Revenue
- **What it is:** Single consolidated operating NWC ratio: NWC_t = Revenue_t × NWC%.
- **How set in model:** Yellow POLICY input: flat 2.5% of revenue every forecast year.
- **Why this choice:** MODEL8: unify AR/Deferred into one NWC% (ends ~$297M double-count drain). Flat 2.5% = asset-light software; AR plugs to NWC. Y1 ΔNWC uses hist NWC as prior so the AR step-down releases cash and the BS stays balanced.
- **Source:** [SEC companyfacts XBRL — AR / AP / DeferredRevenueCurrent](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 16: Inventory (Days)
- **What it is:** Inventory days (Inv = COGS × days/365).
- **How set in model:** Hard zero — FICO is software / scores; inventory is immaterial.
- **Why this choice:** No inventory cycle to fund.
- **Source:** [SEC 10-K FY2025 — Inventory ≈ $0](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 17: Accounts Payable (Days)
- **What it is:** AP days used to project AP = COGS × days/365.
- **How set in model:** Excel equation: ROUND(I48/I25×365, 0) from FY25.
- **Why this choice:** Payable timing offsets AR in operating NWC.
- **Source:** [SEC 10-K FY2025 — Accounts payable](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 18: CapEx % of Revenue
- **What it is:** Capital investment (PPE + capitalized software) as % of sales.
- **How set in model:** Equation: fade from I68/I24 toward 1.0% steady. MODEL8 weights 40%→20%→10%→0%→0% on the peak (fast fade).
- **Why this choice:** MODEL8 checklist: prior path left CapEx ≫ D&A (~$80M cumulative drag). Fade to maintenance ~1% so CapEx ≈ D&A by Y5 (terminal cash conversion).
- **Source:** [SEC 10-K FY2025 — PP&E purchases + capitalized software](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 19: Debt Issuance (Repayment)
- **What it is:** New borrowing (+) or repayment (−) in the forecast.
- **How set in model:** POLICY input = 0 every year.
- **Why this choice:** Hold debt stock flat at FY25 level for the explicit period; interest follows. Net debt for DCF equity bridge uses the later 10-Q amount separately.
- **Source:** [SEC 10-Q Q3 FY2026 — debt stock (financing excluded from FCFF)](https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm)

## 20: Equity Issued (Repaid)
- **What it is:** New equity issued (+) or buybacks (−) in the forecast BS/CF.
- **How set in model:** POLICY input = 0 every year.
- **Why this choice:** Buybacks are real historically but are financing — excluded from FCFF. Share count for $/share is the spot shares outstanding assumption.
- **Source:** [Damodaran FCFF framework — financing (buybacks) not in FCFF](https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histfcff.html)

## DCF-D5: Tax Rate (unlevered)
- **What it is:** Effective tax rate used for FCFF NOPAT and after-tax cost of debt.
- **How set in model:** Linked live to 3-Statement!J13 (= FY25 tax / simplified EBT).
- **Why this choice:** DCF must use the same t as the 3-statement forecast.
- **Source:** [SEC 10-K FY2025 — via 3S J13](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## DCF-D6: WACC (discount rate)
- **What it is:** Weighted average cost of capital for XNPV of FCFF + TV.
- **How set in model:** CAPM live formula ≈ 9.24% (We×Ke + Wd×Rd(1−t)).
- **Why this choice:** Market-consistent discount rate from Rf, β, ERP, and note coupon.
- **Source:** [FRED DGS10 / Damodaran ERP / Yahoo β / 8-K 6.250% notes](https://fred.stlouisfed.org/series/DGS10)

## DCF-D7: Perpetual Growth (g)
- **What it is:** Long-run growth used only in the Gordon TV cross-check.
- **How set in model:** POLICY input = 3.0%.
- **Why this choice:** Long-run nominal GDP-like floor; primary TV uses exit multiple.
- **Source:** [Damodaran long-run growth framing](https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/histimpl.html)

## DCF-D8: Exit EV/EBITDA (BASE)
- **What it is:** Terminal enterprise value multiple on Year-5 EBITDA.
- **How set in model:** POLICY BASE 17.5x; Bull 25.0x; Bear 12.8x. Peer set: EFX 13.88x, VRSK 17.42x, SPGI 18.11x, MCO 22.20x, MSCI 23.74x. Median ≈ 18.11x.
- **Why this choice:** Base 17.5x is near the public peer median (18.1x), between Gordon-implied bear and spot/bull.
- **Source:** [VCP Scanner FICO peer EV/EBITDA comps](https://vcpscanner.com/valuation/fico/relative)

## DCF-D13/D14: Debt & Cash (equity bridge)
- **What it is:** Gross debt and cash+marketable securities for EV → equity.
- **How set in model:** 10-Q Q3 FY2026 totals (more current than FY25 3S balances).
- **Why this choice:** Bridge should reflect the latest capital structure; 3S FY25 shown as reference.
- **Source:** [SEC 10-Q Q3 FY2026](https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm)
