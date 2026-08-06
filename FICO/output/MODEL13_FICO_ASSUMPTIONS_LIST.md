# vengeanceaiUSCMODEL13.0 — Assumptions List

**PDF download:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusmodel13-d3ac/FICO/output/MODEL13_FICO_ASSUMPTIONS_LIST.pdf

Plain list only (no charts). Paste into Google Docs via File → Open if needed.

> Yacktman lens (MODEL12 rated 6/10): think of FICO equity as a long-duration AAA-equivalent bond — monopolistic B2B Scores pricing, rising SaaS deferred billings, and persistent buybacks. MODEL13 lowers WACC to 7.80%, lifts exit to 21.0x, separates Deferred cash inflows from Op NWC/AR, phases DSO, and sharpens incremental margins — optimistic cash conversion without fake plugs; BS identity always holds.

> MODEL13 splits working capital for cash-flow clarity (no double count): Operating NWC = AR + Inventory − AP (excludes Deferred). Δ Op NWC uses cash when AR/AP/Inv rise. Separately, Increase in Deferred Revenue is an explicit positive CFO / FCFF line (SaaS & maintenance billed upfront). DSO phases from FY25 hist ≈ 97 days → policy target 45 days over 5 years (removes MODEL12's punitive AR cliff). AR changes never absorb Deferred billings.

> FY2025 10-K disaggregation ($000s): B2B Scores 948,595; B2C myFICO 219,980; SaaS software 419,720; on-premises software 320,425; Professional services 82,149 (sum = Total revenues 1,990,869). FICO Platform ARR was $263.6M (35% of software ARR) at 9/30/2025 — ARR is a KPI, not an incremental IS line. MODEL13 shifts +200 bps/yr from on-prem mix into SaaS to reflect the cloud transition (drives Deferred Revenue growth).

> Deferred Revenue is a BS liability in Total Liabilities (AP + Debt + Deferred). It is NOT netted inside Op NWC for CFO. Forecast Deferred = Total Rev × (SaaS%+OnPrem%) × (FY25 Deferred ÷ FY25 software Rev). FY25 current deferred ≈ $187.4M (10-K). Q1 FY2026 10-Q shows deferred roll-forward and notes maintenance/SaaS billed annually in advance (https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm). Increase in Deferred = Def_t − Def_t−1 is added in CFO and FCFF. Scores/myFICO usage fees are NOT forced into Deferred (no double count with AR).

> CFO/FCFF add back TOTAL D&A at ≈ 0.84% of Revenue (3yr avg DepreciationDepletionAndAmortization / Rev). That XBRL total already includes AmortizationOfIntangibleAssets — intangibles amortization is fully added back once (no separate second add-back).

> SBC is added back once in CFO and FCFF at ≈ 8.25% of Revenue (ShareBasedCompensation). MODEL13 does NOT also dilute shares by SBC$/Price (that would double-penalize). Share count instead falls with actual buyback cash outflows from residual levered FCF (3S equity CF row 20).

> Starting shares = 23,764k (FYE Sep 30, 2025 outstanding). Each forecast year: Shares_t = Shares_t−1 + EquityIssuanceCF_t / Price. EquityIssuanceCF is the buyback outflow (negative) from 3S row 20 — so repurchases REDUCE the share count. SBC is added back in FCFF but does NOT increase shares (avoids double penalty). $/share uses Year-5 buyback-adjusted shares.

> CapEx % of Revenue fades with operating leverage: 1.25% → 1.05% → 0.90% → 0.75% → 0.60%. FY1 starts near the 3yr hist avg of (PPE + capitalized software)/Rev; Y5 approaches asset-light maintenance intensity as Scores/SaaS mix dominates.

> Pricing power: blended COGS% grinds −75 bps/yr from segment GM base; SG&A grinds −125 bps/yr toward a 15% floor. B2B Scores unit-price increases have near-100% incremental gross margin (little variable COGS) — MODEL12's linear −40/−100 bps understated that cash conversion. FY25 10-K: Scores revenue +$249M YoY 'primarily attributable to a higher unit price'.

> DCF unlevered cash taxes use the 3yr average cash tax rate (IncomeTaxesPaidNet ÷ EBT ≈ 22.87%), not the book effective tax rate. The 3-statement income statement still applies book tax for NI.

> Debt CF uses the 3yr avg net debt cash flow (≈ −$291M/yr). Equity buybacks distribute residual levered FCF so cash does not stockpile. 10-K CF buybacks: FY23 $406M / FY24 $822M / FY25 $1415M (avg ≈ $881M). June 2025 authorization: $1.0B program. Accretive buyback engine: robust FCFF → residual equity CF → lower forward share count in $/share.

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
- **How set in model:** Excel: MAX(10%, segment blended COGS base≈17.60% − 75bps × year).
- **Why this choice:** MODEL13 Yacktman: Pricing power: blended COGS% grinds −75 bps/yr from segment GM base; SG&A grinds −125 bps/yr toward a 15% floor. B2B Scores unit-price increases have near-100% incremental gross margin (little variable COGS) — MODEL12's linear −40/−100 bps understated that cash conversion. FY25 10-K: Scores revenue +$249M YoY 'primarily attributable to a higher unit price'. Mix % offset within Total Revenue (not a second revenue stack).
- **Source:** [SEC 10-K FY2025 — Scores +$249M YoY 'primarily attributable to a higher unit price'](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 9: SG&A % of Revenue
- **What it is:** Operating opex (SG&A) as % of sales.
- **How set in model:** Equation: MAX(15%, (I28 − 10,922)/I24 − 125bps × year). Strips FY25 restructuring; then −1.25% of sales each year.
- **Why this choice:** MODEL13 Yacktman: faster opex leverage (−125 bps/yr, floor 15%) as Scores near-100% incremental margins and SaaS scale drop incremental dollars below the line (MODEL12 used −100 bps).
- **Source:** [SEC 10-K FY2025 — SG&A + Scores pricing commentary](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 10: R&D % of Revenue
- **What it is:** Research & development expense as % of sales (IS label: R&D, not Rent).
- **How set in model:** Excel equation: I29/I24 held flat (~9.46%).
- **Why this choice:** FICO reinvests steadily in analytics/software. Flat % grows dollars with revenue.
- **Source:** [SEC 10-K FY2025 — Research and development](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 11: D&A % of Revenue
- **What it is:** Total depreciation & amortization as % of sales (software-industry driver).
- **How set in model:** Yellow POLICY input: flat 0.84% (3yr FY23–25 avg TOTAL D&A). Forecast DA$ = Revenue × DA%; CF row 63 adds it back once.
- **Why this choice:** MODEL13: CFO/FCFF add back TOTAL D&A at ≈ 0.84% of Revenue (3yr avg DepreciationDepletionAndAmortization / Rev). That XBRL total already includes AmortizationOfIntangibleAssets — intangibles amortization is fully added back once (no separate second add-back).
- **Source:** [SEC companyfacts — DepreciationDepletionAndAmortization (+ AmortizationOfIntangibleAssets)](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

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
- **How set in model:** Yellow POLICY path: phases 97d (FY25 hist) → 45d over 5 years (Y1≈86.6d … Y5=45d). Segment blend ref ≈ 31.9d on schedule row 115 (SaaS 45/B2C 2/B2B 30/PS 15/On-prem 45).
- **Why this choice:** MODEL13 fix vs MODEL12 (6/10): removes the punitive one-year AR cliff (hist ~97d → ~32d). MODEL13 splits working capital for cash-flow clarity (no double count): Operating NWC = AR + Inventory − AP (excludes Deferred). Δ Op NWC uses cash when AR/AP/Inv rise. Separately, Increase in Deferred Revenue is an explicit positive CFO / FCFF line (SaaS & maintenance billed upfront). DSO phases from FY25 hist ≈ 97 days → policy target 45 days over 5 years (removes MODEL12's punitive AR cliff). AR changes never absorb Deferred billings. Deferred Revenue is a BS liability in Total Liabilities (AP + Debt + Deferred). It is NOT netted inside Op NWC for CFO. Forecast Deferred = Total Rev × (SaaS%+OnPrem%) × (FY25 Deferred ÷ FY25 software Rev). FY25 current deferred ≈ $187.4M (10-K). Q1 FY2026 10-Q shows deferred roll-forward and notes maintenance/SaaS billed annually in advance (https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm). Increase in Deferred = Def_t − Def_t−1 is added in CFO and FCFF. Scores/myFICO usage fees are NOT forced into Deferred (no double count with AR).
- **Source:** [SEC 10-K FY2025 — AR; payment terms 30–60 days; Q1 FY2026 deferred detail](https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm)

## 16: Inventory (Days)
- **What it is:** Inventory days (Inv = COGS × days/365).
- **How set in model:** Hard zero — FICO is software / scores; inventory is immaterial.
- **Why this choice:** No inventory cycle to fund.
- **Source:** [SEC 10-K FY2025 — Inventory ≈ $0](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 17: DPO — Accounts Payable (Days)
- **What it is:** Days Payable Outstanding. AP = COGS × DPO / 365.
- **How set in model:** Excel equation: ROUND(I48/I25×365, 0) from FY25; held flat.
- **Why this choice:** MODEL13: DPO pairs with phased DSO for Op NWC = AR+Inv−AP. Deferred is a SEPARATE CFO cash source (row 81) — never mixed into AR/DPO.
- **Source:** [SEC 10-K FY2025 — Accounts payable](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 18: CapEx % of Revenue
- **What it is:** Capital investment (PPE + capitalized software) as % of sales.
- **How set in model:** Yellow POLICY path: 1.25% / 1.05% / 0.90% / 0.75% / 0.60% (fade with operating leverage).
- **Why this choice:** MODEL13 Yacktman: CapEx % of Revenue fades with operating leverage: 1.25% → 1.05% → 0.90% → 0.75% → 0.60%. FY1 starts near the 3yr hist avg of (PPE + capitalized software)/Rev; Y5 approaches asset-light maintenance intensity as Scores/SaaS mix dominates.
- **Source:** [SEC 10-K / companyfacts — PPE purchases + capitalized software; fade to maintenance](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 19: Debt Issuance (Repayment)
- **What it is:** New borrowing (+) or repayment (−) in the forecast ($000s).
- **How set in model:** Yellow POLICY input: -290,917 each year (3yr avg senior-note proceeds − line-of-credit repayments).
- **Why this choice:** MODEL13: Debt CF uses the 3yr avg net debt cash flow (≈ −$291M/yr). Equity buybacks distribute residual levered FCF so cash does not stockpile. 10-K CF buybacks: FY23 $406M / FY24 $822M / FY25 $1415M (avg ≈ $881M). June 2025 authorization: $1.0B program. Accretive buyback engine: robust FCFF → residual equity CF → lower forward share count in $/share. Financing is excluded from FCFF.
- **Source:** [SEC companyfacts — ProceedsFromIssuanceOfSeniorLongTermDebt / RepaymentsOfLinesOfCredit](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 20: Equity Issued (Repaid)
- **What it is:** New equity issued (+) or buybacks (−) in the forecast BS/CF.
- **How set in model:** Equation: −(CFO − CapEx) − DebtIssuance so residual levered FCF funds buybacks (hist 3yr avg buybacks ≈ $881M/yr).
- **Why this choice:** Restores financing realism so cash does not artificially stockpile. Buybacks are financing — excluded from FCFF. DCF share count FALLS with this buyback outflow. Starting shares = 23,764k (FYE Sep 30, 2025 outstanding). Each forecast year: Shares_t = Shares_t−1 + EquityIssuanceCF_t / Price. EquityIssuanceCF is the buyback outflow (negative) from 3S row 20 — so repurchases REDUCE the share count. SBC is added back in FCFF but does NOT increase shares (avoids double penalty). $/share uses Year-5 buyback-adjusted shares.
- **Source:** [SEC 10-K FY2025 — Repurchases of common stock ($1.415B); June 2025 $1B authorization](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## 21: SBC % of Revenue
- **What it is:** Stock-based compensation as % of sales; non-cash add-back in CFO and FCFF.
- **How set in model:** Yellow POLICY input: flat 8.25% (3yr FY23–25 avg ShareBasedCompensation / Revenue).
- **Why this choice:** MODEL13 Yacktman: SBC is added back once in CFO and FCFF at ≈ 8.25% of Revenue (ShareBasedCompensation). MODEL13 does NOT also dilute shares by SBC$/Price (that would double-penalize). Share count instead falls with actual buyback cash outflows from residual levered FCF (3S equity CF row 20).
- **Source:** [SEC 10-K cash flow — ShareBasedCompensation (add-back only; no share dilution)](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## 104: Mix % — SaaS / Platform software
- **What it is:** Share of Total Revenue from SaaS / cloud software (incl. FICO Platform cloud).
- **How set in model:** FY25 base 21.08% (SaaS $419,720k / Total $1,990,869k), then +200 bps/yr taken from on-prem. Platform ARR KPI $263.6M is not additive IS revenue.
- **Why this choice:** SaaS billed annually in advance → Deferred Revenue growth → explicit CFO cash source (3S row 81). Source: Q1 FY2026 10-Q deferred roll-forward (https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/fico-20251231.htm). Mix offsets within Total Rev — not a second revenue line.
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
- **Why this choice:** DCF unlevered cash taxes use the 3yr average cash tax rate (IncomeTaxesPaidNet ÷ EBT ≈ 22.87%), not the book effective tax rate. The 3-statement income statement still applies book tax for NI.
- **Source:** [SEC companyfacts — IncomeTaxesPaidNet / EBT](https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json)

## DCF-D6: WACC (Yacktman AAA-equity discount rate)
- **What it is:** Weighted average cost of capital for XNPV of FCFF + TV.
- **How set in model:** Yellow POLICY D6 = 7.80% (band 7.5–8.2%). CAPM stack kept at R15 ≈ 9.24% as reference only (not used for XNPV).
- **Why this choice:** Yacktman AAA-equity WACC = 7.80% (policy band 7.5%–8.2%). FICO's Scores toll-bridge has ultra-low default risk and bond-like cash predictability; CAPM (~9.24% with β≈1.32) overstates the long-horizon discount rate for this franchise. Live CAPM stack remains in DCF!R15 as a reference; D6 is the Yacktman policy rate used for XNPV.
- **Source:** [Yacktman AAA-equity framework; CAPM ref FRED/Damodaran/Yahoo/8-K](https://fred.stlouisfed.org/series/DGS10)

## DCF-D7: Perpetual Growth (g)
- **What it is:** Long-run growth used only in the Gordon TV cross-check.
- **How set in model:** POLICY input = 3.0%.
- **Why this choice:** Long-run nominal GDP-like floor; primary TV uses exit multiple.
- **Source:** [Damodaran long-run growth framing](https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/histimpl.html)

## DCF-D8: Exit EV/EBITDA (BASE)
- **What it is:** Terminal enterprise value multiple on Year-5 EBITDA.
- **How set in model:** POLICY BASE 21.0x; Bull 26.0x; Bear 15.5x. Peer set: EFX 13.88x, VRSK 17.42x, SPGI 18.11x, MCO 22.20x, MSCI 23.74x. Median ≈ 18.11x.
- **Why this choice:** Base exit EV/EBITDA = 21.0x (raised from MODEL12 17.5x). A perpetual pricing-power compounder deserves a premium to the peer median (~18.1x); band Bull 26.0x / Bear 15.5x. Still below FICO spot (~25–27x).
- **Source:** [VCP Scanner FICO peer EV/EBITDA comps + Yacktman premium](https://vcpscanner.com/valuation/fico/relative)

## DCF-D13/D14: Debt & Cash (equity bridge)
- **What it is:** Gross debt and cash+marketable securities for EV → equity.
- **How set in model:** Latest 10-Q bridge totals (more current than FY25 3S balances).
- **Why this choice:** Bridge should reflect the latest capital structure; 3S FY25 shown as reference.
- **Source:** [SEC 10-Q — Fair Isaac Corp](https://www.sec.gov/Archives/edgar/data/814547/000081454725000016/fico-20250331.htm)

## DCF-D12/I16: Shares Outstanding (buyback-adjusted)
- **What it is:** Starting shares (D12 = 23,764k FYE25) fall each year with buybacks (3S equity CF / price). SBC is added back in FCFF but does NOT increase shares — avoids double penalty. $/share uses Y5 shares (I16).
- **How set in model:** D16=$D$12; E16:I16 = MAX(1000, prior + 3S!{J–N}20 / $D$11); D37=$D$35/$I$16.
- **Why this choice:** Starting shares = 23,764k (FYE Sep 30, 2025 outstanding). Each forecast year: Shares_t = Shares_t−1 + EquityIssuanceCF_t / Price. EquityIssuanceCF is the buyback outflow (negative) from 3S row 20 — so repurchases REDUCE the share count. SBC is added back in FCFF but does NOT increase shares (avoids double penalty). $/share uses Year-5 buyback-adjusted shares.
- **Source:** [SEC 10-K FY2025 — shares outstanding 23,764k; buybacks $1.415B](https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm)

## DCF-E25: Net WC investment (ΔOpNWC − ΔDeferred)
- **What it is:** FCFF subtracts Δ Operating NWC (AR+Inv−AP) and adds Increase in Deferred Revenue. AR and Deferred are strictly separated.
- **How set in model:** DCF E25:I25 = 3S!{J–N}90 − (3S Def_t − Def_t−1). Equivalent to −ΔOpNWC + ΔDeferred inside UFCFF.
- **Why this choice:** MODEL13 splits working capital for cash-flow clarity (no double count): Operating NWC = AR + Inventory − AP (excludes Deferred). Δ Op NWC uses cash when AR/AP/Inv rise. Separately, Increase in Deferred Revenue is an explicit positive CFO / FCFF line (SaaS & maintenance billed upfront). DSO phases from FY25 hist ≈ 97 days → policy target 45 days over 5 years (removes MODEL12's punitive AR cliff). AR changes never absorb Deferred billings. Deferred Revenue is a BS liability in Total Liabilities (AP + Debt + Deferred). It is NOT netted inside Op NWC for CFO. Forecast Deferred = Total Rev × (SaaS%+OnPrem%) × (FY25 Deferred ÷ FY25 software Rev). FY25 current deferred ≈ $187.4M (10-K). Q1 FY2026 10-Q shows deferred roll-forward and notes maintenance/SaaS billed annually in advance (https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm). Increase in Deferred = Def_t − Def_t−1 is added in CFO and FCFF. Scores/myFICO usage fees are NOT forced into Deferred (no double count with AR).
- **Source:** [SEC 10-Q Q1 FY2026 — Deferred revenue roll-forward (R45)](https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm)
