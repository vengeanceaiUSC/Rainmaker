# vengeanceaiUSCMODEL17.0 — Assumptions List

**PDF download:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusmodel17-d3ac/FICO/output/MODEL17_FICO_ASSUMPTIONS_LIST.pdf

Plain list only (no charts). Paste into Google Docs via File → Open if needed.

> Yacktman lens (MODEL16 rated 6.5/10): assign AAA credit to Scores coupons → lower β (1.32→0.854), raise FCFF via Scores-aware margins/CapEx/SBC, and shrink shares with levered buybacks. Every WACC term is algebraic — no naked hardcoded discount rate.

> DSO path unchanged vs Model16 (97→28d). Op NWC = AR+Inv−AP ONLY. +ΔDeferred is a SEPARATE CFO/FCFF source. No AR/unearned double-count. Source: https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm.

> Scores OM ≈ 91% (IR https://investors.fico.com/static-files/66ef723c-1f2e-4501-a452-ab2f3d02ae85). Updated GM_B2B from Model16 (Previous: 90%) to Model17 (New: 97%). Q3 Scores +41% / B2B +49% (https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm).

> Deferred = Rev × (SaaS%+OnPrem%) × (FY25 Def/FY25 soft Rev); +ΔDeferred in CFO row 81. FY25 deferred ≈ $187.4M. Detail: https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm.

> TOTAL D&A ≈ 0.84% of Rev. SBC is a SEPARATE add-back — no double-count.

> Updated SBC from Model16 (Previous: flat 8.25%) to Model17 (New: fade 8.25% / 7.80% / 7.40% / 7.00% / 6.60%) because royalty price dollars do not require 1:1 SBC. Source: https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json.

> Updated shares from Model16 (Previous: 23,764k @ $1,046) to Model17 (New: 21,597.635k @ $1,149 post–Q3). SBC add-back does NOT dilute; shares fall via buybacks only.

> Updated CapEx from Model16 (Previous: 1.25%→0.60%) to Model17 (New: 1.00% → 0.85% → 0.70% → 0.55% → 0.40%) — Scores royalties need ~0 CapEx; do not tax the coupon with Software capitalization.

> Updated margins from Model16 (Previous: COGS −75 / SG&A −125 / R&D flat, COGS floor 10%) to Model17 (New: COGS −120 / SG&A −175 / R&D −50; floor 8%; GM_B2B 97%). Mortgage royalty $4.95→$10.00.

> Keep cash tax 22.87% (3yr IncomeTaxesPaidNet/EBT). Book ETR ~18.8% used in CAPM Rd(1−T).

> Updated buybacks from Model16 (Previous: residual ≈1.0× FCF) to Model17 (New: 1.40×(CFO−CapEx); debt funds MULT−1). Q3 FY2026: buybacks $1960M vs FCF $370.3M (https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm).

> Equity = Assets − Liab − RE every year. Liab = AP + Debt + Deferred. Row-3 OK; cash check ≈ 0.

SEC filings hub: https://investors.fico.com/financial-information/sec-filings

FY2025 10-K: https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm

## 7: Revenue Growth
- **What it is:** Year-over-year % increase in revenue for each forecast year.
- **How set in model:** Yellow POLICY input: 30.0% / 18.0% / 14.5% / 11.0% / 8.5%.
- **Why this choice:** Updated revenue growth from Model16 (Previous: 27%/16%/13%/10%/7%) to Model17 (New: 30%/18%/14.5%/11%/8.5%) — Q3 Scores +41% / B2B +49% mortgage unit price (https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm); royalty $4.95→$10.00. Terminal g 3.0%→3.5%.
- **Source:** [SEC EX-99.1 Q3 FY2026 — Scores +41% / B2B +49%; FY2026 guidance](https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm)

## 8: COGS % of Revenue
- **What it is:** Cost of revenues as a % of sales (gross margin = 1 − this).
- **How set in model:** Excel: MAX(8%, segment blended COGS base≈12.98% − 120bps × year).
- **Why this choice:** MODEL17 Yacktman: Updated margins from Model16 (Previous: COGS −75 / SG&A −125 / R&D flat, COGS floor 10%) to Model17 (New: COGS −120 / SG&A −175 / R&D −50; floor 8%; GM_B2B 97%). Mortgage royalty $4.95→$10.00. Mix % offset within Total Revenue (not a second revenue stack).
- **Source:** [SEC 10-K FY2025 — Scores +$249M YoY 'primarily attributable to a higher unit price'](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 9: SG&A % of Revenue
- **What it is:** Operating opex (SG&A) as % of sales.
- **How set in model:** Equation: MAX(14%, (I28 − 10,922)/I24 − 175bps × year). Strips FY25 restructuring; then −1.75% of sales each year.
- **Why this choice:** Updated SG&A from Model16 (Previous: −125 bps, floor 15%) to Model17 (New: −175 bps/yr, floor 14%) — Scores near-100% incremental margins drop dollars below the line.
- **Source:** [SEC 10-K FY2025 — SG&A + Scores pricing commentary](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 10: R&D % of Revenue
- **What it is:** Research & development expense as % of sales (IS label: R&D, not Rent).
- **How set in model:** Excel: MAX(7%, I29/I24 − 50bps × year).
- **Why this choice:** Updated R&D from Model16 (Previous: flat ~9.5%) to Model17 (New: −50 bps/yr, floor 7%) — royalty price dollars do not require 1:1 R&D reinvestment.
- **Source:** [SEC 10-K FY2025 — Research and development](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 11: D&A % of Revenue
- **What it is:** Total depreciation & amortization as % of sales (software-industry driver).
- **How set in model:** Yellow POLICY input: flat 0.84% (3yr FY23–25 avg TOTAL D&A). Forecast DA$ = Revenue × DA%; CF row 63 adds it back once.
- **Why this choice:** MODEL17: TOTAL D&A ≈ 0.84% of Rev. SBC is a SEPARATE add-back — no double-count.
- **Source:** [SEC companyfacts — DepreciationDepletionAndAmortization (+ AmortizationOfIntangibleAssets)](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 12: Interest % of Debt Open
- **What it is:** Interest expense as % of opening debt balance.
- **How set in model:** Excel equation: I101/I98 (FY25 interest ÷ FY25 opening debt in schedule).
- **Why this choice:** Approximates average coupon / cost of debt on the book. Debt stock follows the net debt issuance assumption (row 19).
- **Source:** [SEC 10-K FY2025 — Interest expense; see also 8-K 6.250% notes](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 13: Book Tax Rate (% of EBT)
- **What it is:** Effective book tax rate applied to forecast EBT → Net Income.
- **How set in model:** Excel equation: I35/I33 (FY25 tax ÷ simplified FY25 EBT in this template).
- **Why this choice:** Book rate (~19%) remains on the income statement. DCF unlevered cash taxes use a separate cash tax rate (22.87% = 3yr IncomeTaxesPaidNet/EBT). Keep cash tax 22.87% (3yr IncomeTaxesPaidNet/EBT). Book ETR ~18.8% used in CAPM Rd(1−T).
- **Source:** [SEC 10-K FY2025 — tax expense; cash taxes from IncomeTaxesPaidNet](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 15: DSO — Accounts Receivable (Days)
- **What it is:** Days Sales Outstanding. AR = Revenue × DSO / 365.
- **How set in model:** Yellow POLICY path: phases 97d (FY25 hist) → 28d over 5 years (Y1≈83.2d … Y5=28d). Segment blend ref ≈ 27.5d on schedule row 115 (SaaS 40/B2C 2/B2B 25/PS 12/On-prem 40).
- **Why this choice:** DSO path unchanged vs Model16 (97→28d). DSO path unchanged vs Model16 (97→28d). Op NWC = AR+Inv−AP ONLY. +ΔDeferred is a SEPARATE CFO/FCFF source. No AR/unearned double-count. Source: https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm. Deferred = Rev × (SaaS%+OnPrem%) × (FY25 Def/FY25 soft Rev); +ΔDeferred in CFO row 81. FY25 deferred ≈ $187.4M. Detail: https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm.
- **Source:** [SEC 10-K FY2025 — AR; Q1 FY2026 deferred detail (R45)](https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm)

## 16: Inventory (Days)
- **What it is:** Inventory days (Inv = COGS × days/365).
- **How set in model:** Hard zero — FICO is software / scores; inventory is immaterial.
- **Why this choice:** No inventory cycle to fund.
- **Source:** [SEC 10-K FY2025 — Inventory ≈ $0](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 17: DPO — Accounts Payable (Days)
- **What it is:** Days Payable Outstanding. AP = COGS × DPO / 365.
- **How set in model:** Excel equation: ROUND(I48/I25×365, 0) from FY25; held flat.
- **Why this choice:** MODEL17: DPO pairs with phased DSO for Op NWC = AR+Inv−AP. Deferred is a SEPARATE CFO cash source (row 81) — never mixed into AR/DPO.
- **Source:** [SEC 10-K FY2025 — Accounts payable](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 18: CapEx % of Revenue
- **What it is:** Capital investment (PPE + capitalized software) as % of sales.
- **How set in model:** Yellow POLICY path: 1.00% / 0.85% / 0.70% / 0.55% / 0.40% (fade with operating leverage).
- **Why this choice:** MODEL17 Yacktman: Updated CapEx from Model16 (Previous: 1.25%→0.60%) to Model17 (New: 1.00% → 0.85% → 0.70% → 0.55% → 0.40%) — Scores royalties need ~0 CapEx; do not tax the coupon with Software capitalization.
- **Source:** [SEC 10-K / companyfacts — PPE purchases + capitalized software; fade to maintenance](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 19: Debt Issuance (Repayment)
- **What it is:** New borrowing (+) or repayment (−) in the forecast ($000s).
- **How set in model:** Equation: −(CFO−CapEx) − EquityCF = (1.40−1)×FCF (debt funds buybacks above organic FCF).
- **Why this choice:** MODEL17: Updated buybacks from Model16 (Previous: residual ≈1.0× FCF) to Model17 (New: 1.40×(CFO−CapEx); debt funds MULT−1). Q3 FY2026: buybacks $1960M vs FCF $370.3M (https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm). Financing is excluded from FCFF.
- **Source:** [SEC EX-99.1 Q3 FY2026 — levered buybacks; companyfacts senior notes / LOC](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 20: Equity Issued (Repaid)
- **What it is:** New equity issued (+) or buybacks (−) in the forecast BS/CF.
- **How set in model:** Equation: −1.40×(CFO−CapEx) levered buybacks (hist 3yr avg ≈ $881M/yr; Q3 FY26 >> FCF).
- **Why this choice:** Updated buybacks from Model16 (Previous: residual ≈1.0× FCF) to Model17 (New: 1.40× FCF). Buybacks are financing — excluded from FCFF. DCF share count FALLS with this outflow. Updated shares from Model16 (Previous: 23,764k @ $1,046) to Model17 (New: 21,597.635k @ $1,149 post–Q3). SBC add-back does NOT dilute; shares fall via buybacks only.
- **Source:** [SEC EX-99.1 Q3 FY2026 — buybacks vs FCF; 10-K FY2025 repurchase footnote](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 21: SBC % of Revenue
- **What it is:** Stock-based compensation as % of sales; non-cash add-back in CFO and FCFF.
- **How set in model:** Yellow POLICY fade path: 8.25% / 7.80% / 7.40% / 7.00% / 6.60% (starts at 3yr avg 8.25%).
- **Why this choice:** MODEL17 Yacktman: Updated SBC from Model16 (Previous: flat 8.25%) to Model17 (New: fade 8.25% / 7.80% / 7.40% / 7.00% / 6.60%) because royalty price dollars do not require 1:1 SBC. Source: https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json.
- **Source:** [SEC 10-K cash flow — ShareBasedCompensation (add-back only; no share dilution)](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 104: Mix % — SaaS / Platform software
- **What it is:** Share of Total Revenue from SaaS / cloud software (incl. FICO Platform cloud).
- **How set in model:** FY25 base 21.08% (SaaS $419,720k / Total $1,990,869k), then +250 bps/yr taken from on-prem. Platform ARR KPI $263.6M is not additive IS revenue.
- **Why this choice:** SaaS mix shift unchanged vs Model16 (+250 bps/yr). SaaS billed annually in advance → Deferred Revenue growth → explicit CFO cash source (row 81). Source: https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/fico-20251231.htm.
- **Source:** [SEC 10-Q Q1 FY2026 — Deferred revenue / contract liabilities detail](https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm)

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
- **Why this choice:** Keep cash tax 22.87% (3yr IncomeTaxesPaidNet/EBT). Book ETR ~18.8% used in CAPM Rd(1−T).
- **Source:** [SEC companyfacts — IncomeTaxesPaidNet / EBT](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## DCF-D6: WACC (Yacktman-adj CAPM — D6 = R15)
- **What it is:** Weighted average cost of capital for XNPV of FCFF + TV.
- **How set in model:** D6 = R15 (Yacktman-adj CAPM) ≈ 7.68%. Updated from Model16 (Previous: raw CAPM 9.27%) to Model17 (New: Yacktman-adj CAPM 7.68%). Rf/β/ERP/Rd inputs in Q6:R15 with clickable sources.
- **Why this choice:** Updated WACC from Model16 (Previous: raw CAPM 9.27% with Yahoo β=1.32) to Model17 (New: Yacktman-adjusted CAPM 7.68%). Algebra: β_Blume=(2/3)×1.32+(1/3)×1.0=1.213; β_AAA=0.70 (GSE Scores coupon credit); β_Yacktman=0.30×β_Blume+0.70×β_AAA=0.854; Ke=Rf(4.67%)+β×ERP(4.28%)=8.33%; Rd_at=6.250%×(1−18.77%)=5.08%; WACC=80%×Ke+20%×Rd_at=7.68%. Raw CAPM kept at R15 as reference. Sources: https://finance.yahoo.com/quote/FICO/key-statistics/; https://www.jstor.org/stable/2329867; https://investors.fico.com/static-files/66ef723c-1f2e-4501-a452-ab2f3d02ae85; https://fred.stlouisfed.org/series/DGS10; https://www.sec.gov/Archives/edgar/data/814547/000119312526117936/d56220d8k.htm.
- **Source:** [Blume 1971; FRED DGS10; Yahoo β; Damodaran ERP; SEC 8-K 6.250% Rd](https://fred.stlouisfed.org/series/DGS10)

## DCF-D7: Perpetual Growth (g)
- **What it is:** Long-run growth used only in the Gordon TV cross-check.
- **How set in model:** POLICY input = 3.5%. Updated from Model16 (Previous: 3.0%) to Model17 (New: 3.5%).
- **Why this choice:** Long-run nominal GDP-like floor; primary TV uses exit multiple.
- **Source:** [Damodaran long-run growth framing](https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/histimpl.html)

## DCF-D8: Exit EV/EBITDA (BASE)
- **What it is:** Terminal enterprise value multiple on Year-5 EBITDA.
- **How set in model:** POLICY BASE 23.0x; Bull 27.0x; Bear 17.0x. Peer set: EFX 13.88x, VRSK 17.42x, SPGI 18.11x, MCO 22.20x, MSCI 23.74x. Median ≈ 18.11x.
- **Why this choice:** Updated exit from Model16 (Previous: 21.0x) to Model17 (New: BASE 23.0x; Bull 27.0x / Bear 17.0x). Still below spot.
- **Source:** [VCP Scanner FICO peer EV/EBITDA comps + Yacktman premium](https://vcpscanner.com/valuation/fico/relative)

## DCF-D13/D14: Debt & Cash (equity bridge)
- **What it is:** Gross debt and cash+marketable securities for EV → equity.
- **How set in model:** Latest 10-Q bridge totals (more current than FY25 3S balances).
- **Why this choice:** Bridge should reflect the latest capital structure; 3S FY25 shown as reference.
- **Source:** [SEC 10-Q — Fair Isaac Corp](https://www.sec.gov/Archives/edgar/data/814547/000081454725000016/fico-20250331.htm)

## DCF-D12/I16: Shares Outstanding (buyback-adjusted)
- **What it is:** Starting shares (D12 ≈ 21,597.6k post–Q3 @ $1,149) fall each year with 1.4×FCF buybacks (3S equity CF / price). SBC is added back in FCFF but does NOT increase shares — avoids double penalty. $/share uses Y5 shares (I16).
- **How set in model:** D16=$D$12; E16:I16 = MAX(1000, prior + 3S!{J–N}20 / $D$11); D37=$D$35/$I$16.
- **Why this choice:** Updated shares from Model16 (Previous: 23,764k @ $1,046) to Model17 (New: 21,597.635k @ $1,149 post–Q3). SBC add-back does NOT dilute; shares fall via buybacks only.
- **Source:** [EX-99.1 Q3 FY2026 — shares ≈ 21,597.6k @ $1,149; FY25 buybacks $1.415B](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## DCF-E25: Net WC investment (ΔOpNWC − ΔDeferred)
- **What it is:** FCFF subtracts Δ Operating NWC (AR+Inv−AP) and adds Increase in Deferred Revenue. AR and Deferred are strictly separated.
- **How set in model:** DCF E25:I25 = 3S!{J–N}90 − (3S Def_t − Def_t−1). Equivalent to −ΔOpNWC + ΔDeferred inside UFCFF.
- **Why this choice:** DSO path unchanged vs Model16 (97→28d). Op NWC = AR+Inv−AP ONLY. +ΔDeferred is a SEPARATE CFO/FCFF source. No AR/unearned double-count. Source: https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm. Deferred = Rev × (SaaS%+OnPrem%) × (FY25 Def/FY25 soft Rev); +ΔDeferred in CFO row 81. FY25 deferred ≈ $187.4M. Detail: https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm.
- **Source:** [SEC 10-Q Q1 FY2026 — Deferred revenue roll-forward (R45)](https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm)
