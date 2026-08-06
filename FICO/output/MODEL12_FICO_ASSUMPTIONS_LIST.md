# VengeanceUSCModel12.0 — Assumptions List

**PDF download:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceuscmodel12-d3ac/FICO/output/MODEL12_FICO_ASSUMPTIONS_LIST.pdf

Plain list only (no charts). Paste into Google Docs via File → Open if needed.

> Working Capital uses segment cash-conversion DSOs blended by FY2025 10-K revenue mix (SaaS, on-prem software, B2C myFICO, B2B Scores, Professional services). NWC = AR + Inventory − AP − Deferred Revenue (output); ΔNWC = NWCt − NWCt−1. Mix percentages offset within Total Revenue — they do not add a second revenue stack.

> FY2025 10-K disaggregation ($000s): B2B Scores 948,595; B2C myFICO 219,980; SaaS software 419,720; on-premises software 320,425; Professional services 82,149 (sum = Total revenues 1,990,869). FICO Platform ARR was $263.6M (35% of software ARR) at 9/30/2025 — ARR is a KPI, not an incremental IS line. SaaS/on-prem billings create Deferred Revenue (contract liability); B2C is near-zero DSO (card-settled); B2B Scores use ~30-day trade terms; Professional services convert cash quickly but at much lower gross margin (~28% vs software/scores).

> Deferred Revenue is an explicit Balance Sheet liability included in Total Liabilities and in NWC (= AR+Inv−AP−Deferred). Forecast Deferred = Total Rev × (SaaS% + OnPrem%) × (FY25 Deferred ÷ FY25 software revenue). Only software subscription / maintenance billings drive the liability — Scores and myFICO are not double-counted into Deferred.

> D&A is forecast as a % of Revenue (3yr FY23–25 average of total DepreciationDepletionAndAmortization / Revenue ≈ 0.84%), not a PP&E-only rate — capturing software amortization and intangibles.

> Stock-Based Compensation is added back in CFO and FCFF at the 3yr average SBC/Revenue (≈ 8.25%) from ShareBasedCompensation in the 10-K cash flow. DCF shares outstanding grow each year by SBC$ / share price so Equity Value/Share reflects dilution from the SBC add-back.

> Starting diluted shares (D12) roll forward each forecast year: Shares_t = Shares_t−1 + SBC_t($000s) / Current Price. SBC$ comes from 3-statement row 64 (Rev × 8.25% SBC). Terminal Equity Value/Share uses Year-5 diluted shares, not the static starting share count.

> CapEx uses a flat 3yr average of (PPE purchases + capitalized software) / Revenue (≈ 1.25%). The prior hardcoded fade to 1% was removed so reinvestment reflects actual history rather than forcing CapEx ≈ D&A by Year 5.

> DCF unlevered cash taxes use the 3yr average cash tax rate (IncomeTaxesPaidNet ÷ EBT ≈ 22.87%), not the book effective tax rate. The 3-statement income statement still applies book tax for NI.

> Financing CF restores realism outside FCFF: debt uses the 3yr average net debt cash flow (senior-note proceeds − line-of-credit repayments ≈ −$291M/yr). Equity buybacks distribute residual levered FCF after that debt CF so cash does not artificially stockpile (historical buybacks averaged ≈ $881M/yr in FY23–25).

> Balance sheet identity is enforced every year: Equity Capital = Total Assets − Total Liabilities − Retained Earnings (hard plug). Total Liabilities = AP + Debt + Deferred Revenue. Row-3 Balance Sheet Check flags any residual ≠ 0.

SEC filings hub: https://investors.fico.com/financial-information/sec-filings

FY2025 10-K: https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm

## 7: Revenue Growth
- **What it is:** Year-over-year % increase in revenue for each forecast year.
- **How set in model:** Yellow POLICY input: 27% / 16% / 13% / 10% / 7%.
- **Why this choice:** Y1 ≈ company FY2026 revenue guidance (~$2.53B / FY25 $1.991B − 1 ≈ 27%). Then fade toward terminal g=3% (not straight-lined). This is the main forward-looking judgment; it drives Rev → GP → EBT → Net Earnings → CF.
- **Source:** [SEC EX-99.1 Q3 FY2026 — updated FY2026 revenue guidance $2.53B](https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm)

## 8: COGS % of Revenue
- **What it is:** Cost of revenues as a % of sales (gross margin = 1 − this).
- **How set in model:** Excel equation: =$J$122 blended COGS from segment GMs (≈ 17.60%; calibrated near FY25 ~17.8%).
- **Why this choice:** MODEL12: COGS reflects segment margin mix — B2B Scores high GM, SaaS/on-prem high GM, Professional Services much lower GM (~28%). Mix % offset within Total Revenue (not a second revenue stack).
- **Source:** [SEC 10-K FY2025 — Note 9 / MD&A disaggregated revenue + Cost of revenues](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 9: SG&A % of Revenue
- **What it is:** Operating opex (SG&A) as % of sales.
- **How set in model:** Equation: MAX(15%, (I28 − 10,922)/I24 − 75bps × year). Strips FY25 restructuring; then −0.75% of sales each year.
- **Why this choice:** MODEL12: restore software operating leverage. Floor 15%; grind −75 bps/yr so SG&A can scale toward mid-teens as revenue expands (not stuck ~24–25%).
- **Source:** [SEC 10-K FY2025 — SG&A + restructuring note](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 10: R&D % of Revenue
- **What it is:** Research & development expense as % of sales (IS label: R&D, not Rent).
- **How set in model:** Excel equation: I29/I24 held flat (~9.46%).
- **Why this choice:** FICO reinvests steadily in analytics/software. Flat % grows dollars with revenue.
- **Source:** [SEC 10-K FY2025 — Research and development](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 11: D&A % of Revenue
- **What it is:** Total depreciation & amortization as % of sales (software-industry driver).
- **How set in model:** Yellow POLICY input: flat 0.84% (3yr FY23–25 avg). Forecast DA$ = Revenue × DA%.
- **Why this choice:** MODEL12: D&A is forecast as a % of Revenue (3yr FY23–25 average of total DepreciationDepletionAndAmortization / Revenue ≈ 0.84%), not a PP&E-only rate — capturing software amortization and intangibles.
- **Source:** [SEC companyfacts — DepreciationDepletionAndAmortization / Revenue](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 12: Interest % of Debt Open
- **What it is:** Interest expense as % of opening debt balance.
- **How set in model:** Excel equation: I101/I98 (FY25 interest ÷ FY25 opening debt in schedule).
- **Why this choice:** Approximates average coupon / cost of debt on the book. Debt stock follows the net debt issuance assumption (row 19).
- **Source:** [SEC 10-K FY2025 — Interest expense; see also 8-K 6.250% notes](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 13: Book Tax Rate (% of EBT)
- **What it is:** Effective book tax rate applied to forecast EBT → Net Income.
- **How set in model:** Excel equation: I35/I33 (FY25 tax ÷ simplified FY25 EBT in this template).
- **Why this choice:** Book rate (~19%) remains on the income statement. DCF unlevered cash taxes use a separate cash tax rate (22.87% = 3yr IncomeTaxesPaidNet/EBT). DCF unlevered cash taxes use the 3yr average cash tax rate (IncomeTaxesPaidNet ÷ EBT ≈ 22.87%), not the book effective tax rate. The 3-statement income statement still applies book tax for NI.
- **Source:** [SEC 10-K FY2025 — tax expense; cash taxes from IncomeTaxesPaidNet](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 15: DSO — Accounts Receivable (Days)
- **What it is:** Days Sales Outstanding. AR = Revenue × DSO / 365.
- **How set in model:** Excel equation: ROUND($J$115,1) = blended segment DSO ≈ 31.9 days (SaaS 45d / B2C 2d / B2B 30d / PS 15d / On-prem 45d × FY25 mix).
- **Why this choice:** MODEL12: Working Capital uses segment cash-conversion DSOs blended by FY2025 10-K revenue mix (SaaS, on-prem software, B2C myFICO, B2B Scores, Professional services). NWC = AR + Inventory − AP − Deferred Revenue (output); ΔNWC = NWCt − NWCt−1. Mix percentages offset within Total Revenue — they do not add a second revenue stack. FY2025 10-K disaggregation ($000s): B2B Scores 948,595; B2C myFICO 219,980; SaaS software 419,720; on-premises software 320,425; Professional services 82,149 (sum = Total revenues 1,990,869). FICO Platform ARR was $263.6M (35% of software ARR) at 9/30/2025 — ARR is a KPI, not an incremental IS line. SaaS/on-prem billings create Deferred Revenue (contract liability); B2C is near-zero DSO (card-settled); B2B Scores use ~30-day trade terms; Professional services convert cash quickly but at much lower gross margin (~28% vs software/scores). AR is organic from blended collections days — no top-down NWC% AR plug. Deferred Revenue is an explicit Balance Sheet liability included in Total Liabilities and in NWC (= AR+Inv−AP−Deferred). Forecast Deferred = Total Rev × (SaaS% + OnPrem%) × (FY25 Deferred ÷ FY25 software revenue). Only software subscription / maintenance billings drive the liability — Scores and myFICO are not double-counted into Deferred.
- **Source:** [SEC 10-K FY2025 — Note 9 disaggregated revenue; payment terms 30–60 days](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 16: Inventory (Days)
- **What it is:** Inventory days (Inv = COGS × days/365).
- **How set in model:** Hard zero — FICO is software / scores; inventory is immaterial.
- **Why this choice:** No inventory cycle to fund.
- **Source:** [SEC 10-K FY2025 — Inventory ≈ $0](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 17: DPO — Accounts Payable (Days)
- **What it is:** Days Payable Outstanding. AP = COGS × DPO / 365.
- **How set in model:** Excel equation: ROUND(I48/I25×365, 0) from FY25; held flat.
- **Why this choice:** MODEL12: DPO is an explicit WC driver paired with DSO. NWC = AR + Inventory − AP − Deferred (output); ΔNWC = NWCt − NWCt−1.
- **Source:** [SEC 10-K FY2025 — Accounts payable](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 18: CapEx % of Revenue
- **What it is:** Capital investment (PPE + capitalized software) as % of sales.
- **How set in model:** Yellow POLICY input: flat 1.25% (3yr avg of PPE purchases + PaymentsToDevelopSoftware / Revenue).
- **Why this choice:** MODEL12: CapEx uses a flat 3yr average of (PPE purchases + capitalized software) / Revenue (≈ 1.25%). The prior hardcoded fade to 1% was removed so reinvestment reflects actual history rather than forcing CapEx ≈ D&A by Year 5.
- **Source:** [SEC companyfacts — PPE purchases + PaymentsToDevelopSoftware](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 19: Debt Issuance (Repayment)
- **What it is:** New borrowing (+) or repayment (−) in the forecast ($000s).
- **How set in model:** Yellow POLICY input: -290,917 each year (3yr avg senior-note proceeds − line-of-credit repayments).
- **Why this choice:** MODEL12: Financing CF restores realism outside FCFF: debt uses the 3yr average net debt cash flow (senior-note proceeds − line-of-credit repayments ≈ −$291M/yr). Equity buybacks distribute residual levered FCF after that debt CF so cash does not artificially stockpile (historical buybacks averaged ≈ $881M/yr in FY23–25). Financing is excluded from FCFF.
- **Source:** [SEC companyfacts — ProceedsFromIssuanceOfSeniorLongTermDebt / RepaymentsOfLinesOfCredit](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 20: Equity Issued (Repaid)
- **What it is:** New equity issued (+) or buybacks (−) in the forecast BS/CF.
- **How set in model:** Equation: −(CFO − CapEx) − DebtIssuance so residual levered FCF funds buybacks (hist 3yr avg buybacks ≈ $881M/yr).
- **Why this choice:** Restores financing realism so cash does not artificially stockpile. Buybacks are financing — excluded from FCFF. DCF $/share uses SBC-diluted shares (row 16), not a static count. Starting diluted shares (D12) roll forward each forecast year: Shares_t = Shares_t−1 + SBC_t($000s) / Current Price. SBC$ comes from 3-statement row 64 (Rev × 8.25% SBC). Terminal Equity Value/Share uses Year-5 diluted shares, not the static starting share count.
- **Source:** [SEC companyfacts — PaymentsForRepurchaseOfCommonStock](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 21: SBC % of Revenue
- **What it is:** Stock-based compensation as % of sales; non-cash add-back in CFO and FCFF.
- **How set in model:** Yellow POLICY input: flat 8.25% (3yr FY23–25 avg ShareBasedCompensation / Revenue).
- **Why this choice:** MODEL12: Stock-Based Compensation is added back in CFO and FCFF at the 3yr average SBC/Revenue (≈ 8.25%) from ShareBasedCompensation in the 10-K cash flow. DCF shares outstanding grow each year by SBC$ / share price so Equity Value/Share reflects dilution from the SBC add-back. SBC$ (3S row 64) also increases DCF diluted shares each year (ΔShares = SBC$000s / Price).
- **Source:** [SEC 10-K cash flow — ShareBasedCompensation](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 104: Mix % — SaaS / Platform software
- **What it is:** Share of Total Revenue from SaaS / cloud software (incl. FICO Platform cloud).
- **How set in model:** Yellow POLICY = 21.08% (FY25 SaaS $419,720k / Total $1,990,869k). Platform ARR KPI was $263.6M (35% of software ARR) — ARR is not added on top of IS revenue.
- **Why this choice:** SaaS is billed largely annually in advance → primary Deferred Revenue driver with on-prem maintenance. 30–60 day payment terms (model uses 45-day DSO). Mix offsets within Total Rev — does not create a second revenue line.
- **Source:** [SEC 10-K FY2025 — Software SaaS disaggregation + Platform ARR MD&A](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 105: Mix % — B2C Subscriptions (myFICO)
- **What it is:** Share of Total Revenue from B2C scoring / myFICO.com subscriptions.
- **How set in model:** Yellow POLICY = 11.05% (FY25 B2C Scores $219,980k / Total $1,990,869k).
- **Why this choice:** Near-zero DSO (card-settled consumer subscriptions). Minimal Deferred Revenue vs annual SaaS invoices. High incremental margin; WC-light cash conversion.
- **Source:** [SEC 10-K FY2025 — Scores segment B2C / myFICO disaggregation](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 106: Mix % — B2B Scores
- **What it is:** Share of Total Revenue from B2B scoring (mortgage, auto, card via CRAs).
- **How set in model:** Yellow POLICY = 47.65% (FY25 B2B Scores $948,595k / Total $1,990,869k).
- **Why this choice:** Transactional volume through Experian/TransUnion/Equifax; ~30-day DSO. Does not create SaaS-style Deferred Revenue (usage-based recognition).
- **Source:** [SEC 10-K FY2025 — Scores segment B2B disaggregation](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 107: Mix % — Professional Services
- **What it is:** Share of Total Revenue from implementation / consulting / training fees.
- **How set in model:** Yellow POLICY = 4.13% (FY25 PS $82,149k / Total $1,990,869k).
- **Why this choice:** Upfront integration fees convert cash quickly (low DSO) but carry significantly lower gross margin than software/scores — pulls blended COGS up vs pure SaaS.
- **Source:** [SEC 10-K FY2025 — Software segment Professional services line](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 108: Mix % — On-Premises Software
- **What it is:** Share of Total Revenue from on-premises license + maintenance software.
- **How set in model:** Yellow POLICY = 16.09% (FY25 on-prem $320,425k / Total $1,990,869k).
- **Why this choice:** With SaaS mix, drives Deferred Revenue (maintenance billed in advance). Same 45-day DSO policy as SaaS. Completes mix identity to 100% of Total Rev.
- **Source:** [SEC 10-K FY2025 — Software on-premises vs SaaS deployment table](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## DCF-D5: Cash Tax Rate (unlevered)
- **What it is:** Cash tax rate used for FCFF unlevered taxes and after-tax cost of debt.
- **How set in model:** Yellow POLICY = 22.87% (3yr avg IncomeTaxesPaidNet ÷ EBT). NOT linked to book tax J13.
- **Why this choice:** DCF unlevered cash taxes use the 3yr average cash tax rate (IncomeTaxesPaidNet ÷ EBT ≈ 22.87%), not the book effective tax rate. The 3-statement income statement still applies book tax for NI.
- **Source:** [SEC companyfacts — IncomeTaxesPaidNet / EBT](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

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
- **How set in model:** Latest 10-Q bridge totals (more current than FY25 3S balances).
- **Why this choice:** Bridge should reflect the latest capital structure; 3S FY25 shown as reference.
- **Source:** [SEC 10-Q — Fair Isaac Corp](https://www.sec.gov/Archives/edgar/data/814547/000081454725000016/fico-20250331.htm)

## DCF-D12/I16: Shares Outstanding (SBC dilution)
- **What it is:** Starting diluted shares (D12) roll forward each year by SBC$ ÷ share price; Equity Value/Share uses Year-5 diluted shares (I16).
- **How set in model:** D16=$D$12; E16:I16 = prior + 3S!{J–N}64 / $D$11; D37=$D$35/$I$16.
- **Why this choice:** Starting diluted shares (D12) roll forward each forecast year: Shares_t = Shares_t−1 + SBC_t($000s) / Current Price. SBC$ comes from 3-statement row 64 (Rev × 8.25% SBC). Terminal Equity Value/Share uses Year-5 diluted shares, not the static starting share count.
- **Source:** [3S SBC row 64 (Rev × 8.25%); SEC ShareBasedCompensation](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## DCF-E25: ΔNWC (includes Deferred Revenue)
- **What it is:** Change in NWC = AR + Inventory − AP − Deferred Revenue. Linked live from 3S row 90 into FCFF.
- **How set in model:** DCF E25:I25 = 3S!J90:N90 (organic WC schedule).
- **Why this choice:** Deferred Revenue is an explicit Balance Sheet liability included in Total Liabilities and in NWC (= AR+Inv−AP−Deferred). Forecast Deferred = Total Rev × (SaaS% + OnPrem%) × (FY25 Deferred ÷ FY25 software revenue). Only software subscription / maintenance billings drive the liability — Scores and myFICO are not double-counted into Deferred.
- **Source:** [SEC 10-K FY2025 — Deferred revenue / contract liabilities](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)
