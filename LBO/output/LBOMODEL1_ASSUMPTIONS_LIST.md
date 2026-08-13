# vengeanceaiUSC-LBOMODEL1 — Assumptions List (every tab)

**Ticker / issuer:** GTM — ZoomInfo Technologies Inc.

**Format:** Plain list. Every item uses **What** / **How** / **Why** / **Source**.

**Workbook:** [https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx](https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx)

**PDF:** [https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.pdf](https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.pdf)

**Generated:** 2026-08-13 20:31 UTC

Educational / research use only — not investment advice.


# Tab: `Coversheet`

## 1. Model identity (B2)
- **Value:** vengeanceaiUSC-LBOMODEL1 — ZoomInfo (GTM)
- **What:** Name of this LBO case.
- **How:** Set on Coversheet and LBO!B2.
- **Why:** Separates ZoomInfo LBOMODEL1 from the blank WSP BMC sample.
- **Source:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx

## 2. Units (—)
- **Value:** $ millions except per share / rates / multiples
- **What:** Unit convention for the workbook.
- **How:** Stated on LBO!B3 / DCF!B3.
- **Why:** Matches WSP template convention and SEC $ presentation scaled to millions.
- **Source:** https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm


# Tab: `LBO`

## 3. Company / ticker (D6/D7)
- **Value:** ZoomInfo Technologies Inc. / GTM
- **What:** Target company identity.
- **How:** Hardcoded inputs on LBO general inputs.
- **Why:** Public GTM (legacy ZI) is the LBOMODEL1 case company.
- **Source:** https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm

## 4. Current share price (D8)
- **Value:** $4.31
- **What:** Spot equity price used for premium math.
- **How:** LBO!D8 market input.
- **Why:** Latest available GTM print used when MODEL1 was built (~2026-08-12).
- **Source:** https://finance.yahoo.com/quote/GTM

## 5. Price date (D9)
- **Value:** 2026-08-12
- **What:** As-of date for the spot price.
- **How:** LBO!D9.
- **Why:** Timestamps the market input for auditability.
- **Source:** https://finance.yahoo.com/quote/GTM

## 6. Acquisition premium (H21)
- **Value:** 25%
- **What:** Premium of offer price over spot.
- **How:** Offer/sh = D8×(1+premium).
- **Why:** Standard illustrative take-private premium; editable deal assumption.
- **Source:** Model assumption (LBOMODEL1)

## 7. Offer price / share (H20)
- **Value:** $5.3875
- **What:** Implied takeout price per share.
- **How:** D8 × 1.25.
- **Why:** Drives equity purchase price in Uses.
- **Source:** Model assumption + market price

## 8. Diluted shares (millions) (H18)
- **Value:** 292.312
- **What:** Share count for equity value.
- **How:** From latest shares outstanding.
- **Why:** 10-Q 2026-06-30 common shares outstanding used as diluted proxy (options ladder not fully modeled).
- **Source:** https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001794515&type=10-Q

## 9. Equity purchase price (D20)
- **Value:** $1,574.8m
- **What:** Cash to buy equity.
- **How:** Offer/sh × shares.
- **Why:** Primary Use of funds line.
- **Source:** Derived

## 10. LTM EBITDA (D13)
- **Value:** $314.5m
- **What:** Entry EBITDA for leverage / multiple math.
- **How:** FY2025 OpInc $225.7m + D&A $88.8m.
- **Why:** GAAP EBITDA proxy from latest annual 10-K (not adjusted EBITDA).
- **Source:** https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm ; https://data.sec.gov/api/xbrl/companyfacts/CIK0001794515.json

## 11. Gross debt (D14)
- **Value:** −$1,269.9m
- **What:** Debt refinanced at close (entered as negative in template).
- **How:** LT debt + current maturities.
- **Why:** 10-Q 2026-06-30: LongTermDebt + LongTermDebtCurrent.
- **Source:** https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001794515&type=10-Q

## 12. Cash (D15)
- **Value:** $150.1m
- **What:** Cash available at entry.
- **How:** Cash & equivalents + short-term investments.
- **Why:** 10-Q 2026-06-30.
- **Source:** https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001794515&type=10-Q

## 13. Minimum cash (D16)
- **Value:** $50.0m
- **What:** Cash retained for ops; excess used as a Source.
- **How:** Hardcoded policy input.
- **Why:** Leaves liquidity post-close; excess cash = 150.1 − 50 = 100.1.
- **Source:** Model assumption (LBOMODEL1)

## 14. Entry EV / LTM EBITDA (I11)
- **Value:** 8.57x
- **What:** Implied entry multiple.
- **How:** (Equity buy + debt − cash) / EBITDA.
- **Why:** Output of offer + net debt vs LTM EBITDA.
- **Source:** Derived from SEC + market inputs

## 15. Exit EV / EBITDA (D17)
- **Value:** 8.0x
- **What:** Exit multiple in hold-year 5.
- **How:** LBO!D17 policy input.
- **Why:** Near entry; assumes modest multiple compression vs entry 8.57x.
- **Source:** Model assumption (LBOMODEL1)

## 16. Term Loan A (C29/D29)
- **Value:** 2.5x / $786.2m
- **What:** New TLA quantum.
- **How:** EBITDA × turns.
- **Why:** Part of 5.0x new-money debt stack.
- **Source:** Model assumption (LBOMODEL1)

## 17. Term Loan B (C30/D30)
- **Value:** 1.5x / $471.8m
- **What:** New TLB quantum.
- **How:** EBITDA × turns.
- **Why:** Part of 5.0x new-money debt stack.
- **Source:** Model assumption (LBOMODEL1)

## 18. Senior Notes (C31/D31)
- **Value:** 1.0x / $314.5m
- **What:** New senior notes quantum.
- **How:** EBITDA × turns.
- **Why:** Completes 5.0x new debt with TLA+TLB.
- **Source:** Model assumption (LBOMODEL1)

## 19. Sponsor equity (D35)
- **Value:** $1,203.6m
- **What:** Equity check (Sources plug).
- **How:** Total Uses − excess cash − new debt.
- **Why:** Balancing figure so Sources = Uses ($2,876.2m).
- **Source:** Derived

## 20. Transaction fee % (G36)
- **Value:** 2% of equity value
- **What:** M&A / financing fee assumption.
- **How:** Fees $ = 2% × equity purchase.
- **Why:** Simple fee load for Uses; financing fee dollars also shown in fee block.
- **Source:** Model assumption (LBOMODEL1)

## 21. Revenue growth (FY26–30) (F65:J65)
- **Value:** 4% / 5% / 5% / 4% / 3%
- **What:** YoY revenue growth in forecast.
- **How:** Hardcoded on growth row.
- **Why:** Modest recovery vs FY24 dip; not a hyper-growth case.
- **Source:** Model assumption (LBOMODEL1)

## 22. Gross margin (E66/F66)
- **Value:** FY25 82.35% → Fcst 82.85%
- **What:** Gross profit % of sales.
- **How:** From FY25 COGS implied by 10-K; slight +50 bps in forecast.
- **Why:** Software gross margin structure from SEC P&L.
- **Source:** https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm

## 23. R&D margin (E67/F67)
- **Value:** 14.57% flat in forecast
- **What:** R&D % of sales.
- **How:** FY2025 R&D / Revenue held flat.
- **Why:** Keeps product investment while S&M is the efficiency lever.
- **Source:** https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm

## 24. SG&A margin path (F68:J68)
- **Value:** 47.72% → 46.22% → 44.72% → 44.22% → 44.22%
- **What:** SG&A (S&M+G&A) % of sales in forecast.
- **How:** Compresses ~500 bps of revenue by Y3 vs FY25 49.72%.
- **Why:** Operationalizes ZI AI-SDR thesis into the LBO P&L.
- **Source:** LBOMODEL1 thesis + ZI_* tabs; FY25 base from https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm

## 25. Tax rate (C69:J69)
- **Value:** 21%
- **What:** Book tax rate on EBT.
- **How:** Flat statutory-like rate in MODEL1.
- **Why:** Simplifies vs actual cash tax / NOL complexity.
- **Source:** Model assumption (LBOMODEL1)

## 26. Sponsor MoIC / IRR (I275/J275)
- **Value:** ~2.97x / ~24.4%
- **What:** Base returns at 8.0x exit.
- **How:** Exit equity / sponsor equity; IRR = MoIC^(1/5)−1.
- **Why:** Illustrative; sensitive to exit multiple, leverage, and SG&A path.
- **Source:** Derived on LBO returns block


# Tab: `DCF`

## 27. DCF reference share price (C6)
- **Value:** $5.3875 (LBO offer)
- **What:** Price shown in DCF header for football-field compare.
- **How:** Synced to LBO offer/sh.
- **Why:** Makes DCF tab speak the same deal price as the LBO.
- **Source:** LBO!H20

## 28. Basic / diluted share count (C8)
- **Value:** 292.312m
- **What:** Share count for per-share value.
- **How:** Synced to LBO shares.
- **Why:** Same 10-Q share count as LBO.
- **Source:** https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001794515&type=10-Q

## 29. CAPM inputs (C55/C56/C57)
- **Value:** Rf 4.3% / β 1.20 / ERP 5.0%
- **What:** Building blocks of cost of equity.
- **How:** Ke = Rf + β×ERP.
- **Why:** Illustrative SaaS CAPM stack for DCF tab (not Yacktman-adjusted).
- **Source:** Model assumption (LBOMODEL1)

## 30. Cost of debt / tax (C50/C51)
- **Value:** 7.0% / 21%
- **What:** After-tax debt cost inputs.
- **How:** Rd×(1−t) in WACC.
- **Why:** Rounded financing cost vs new LBO stack.
- **Source:** Model assumption (LBOMODEL1)

## 31. Capital structure weights (C53/C58)
- **Value:** Wd 45% / We 55%
- **What:** Target weights in WACC.
- **How:** Hardcoded policy on DCF tab.
- **Why:** Closer to PF leveraged structure than pre-deal equity-heavy weights.
- **Source:** Model assumption (LBOMODEL1)

## 32. WACC (C9/C59)
- **Value:** ~8.53%
- **What:** Discount rate for unlevered FCF.
- **How:** We×Ke + Wd×Rd×(1−t).
- **Why:** Used for DCF PV math on this tab.
- **Source:** Derived from CAPM + Rd assumptions

## 33. Exit EBITDA multiple (DCF) (J37)
- **Value:** 8.0x
- **What:** Terminal multiple in exit-EBITDA approach.
- **How:** Synced to LBO!D17.
- **Why:** Keeps LBO and DCF exit frameworks comparable.
- **Source:** LBO!D17

## 34. Perpetuity g (C37)
- **Value:** 0% in template cell (sensitivity uses 2–4%)
- **What:** Long-term growth in Gordon approach.
- **How:** Template default; grid varies g.
- **Why:** Football-field sensitivity still shows 2–4% g range.
- **Source:** WSP template convention + MODEL1 note

## 35. Cash-flow timing (C31)
- **Value:** Middle of period
- **What:** Mid-year discounting convention.
- **How:** Midperiod adjustment factor on DCF tab.
- **Why:** Standard IB mid-year convention for annual FCFs.
- **Source:** WSP template


# Tab: `Shares`

## 36. Offer price for dilution (E5/E6)
- **Value:** $5.3875
- **What:** Price used in treasury-stock method.
- **How:** Synced to LBO offer.
- **Why:** Aligns diluted-share math with deal price.
- **Source:** LBO!H20

## 37. Basic & net diluted shares (E7/E14)
- **Value:** 292.312m / 292.312m
- **What:** Share count used for equity value.
- **How:** Outstanding shares; options set to 0 in MODEL1.
- **Why:** 10-Q share count; detailed options ladder not loaded (net dilutive options = 0).
- **Source:** https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001794515&type=10-Q

## 38. Options / other dilutives (E8:E11)
- **Value:** 0
- **What:** Incremental dilutive securities.
- **How:** Cleared BMC ladder; set to zero pending full award rollforward.
- **Why:** Conservative vs overstating dilution without a current award table.
- **Source:** Model assumption (LBOMODEL1) — update from proxy/10-K award note when refined


# Tab: `52wkHL`

## 39. 52-week high / low (E4/E5)
- **Value:** Legacy sample tape (not GTM)
- **What:** Market range for football-field.
- **How:** Still WSP BMC sample series in MODEL1.
- **Why:** Not yet replaced with GTM price history; LBO uses spot $4.31 instead.
- **Source:** WSP sample residual — replace with GTM tape in MODEL2+


# Tab: `ZI_01_10K_Source`

## 40. FY2025 Revenue (E5)
- **Value:** $1,249.5m
- **What:** Annual revenue anchor.
- **How:** XBRL RevenueFromContractWithCustomerExcludingAssessedTax.
- **Why:** Latest annual 10-K fact.
- **Source:** https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm ; https://data.sec.gov/api/xbrl/companyfacts/CIK0001794515.json

## 41. FY2025 Selling & Marketing (E6)
- **Value:** $414.6m
- **What:** S&M expense (the bloat line).
- **How:** XBRL SellingAndMarketingExpense.
- **Why:** 33.2% of FY25 revenue — thesis starting point.
- **Source:** https://data.sec.gov/api/xbrl/companyfacts/CIK0001794515.json

## 42. OpInc / D&A / EBITDA proxy (E10/E13/E20)
- **Value:** $225.7 / $88.8 / $314.5m
- **What:** Operating income, D&A, EBITDA proxy.
- **How:** OpInc + OtherDepreciationAndAmortization.
- **Why:** Bridges to LBO!D13 LTM EBITDA.
- **Source:** https://data.sec.gov/api/xbrl/companyfacts/CIK0001794515.json

## 43. S&M employees (E28)
- **Value:** 1,370
- **What:** Headcount in sales & marketing.
- **How:** 10-K Human Capital disclosure (3,180 total).
- **Why:** Denominator for SDR % assumption.
- **Source:** https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm


# Tab: `ZI_02_SM_Isolation`

## 44. Target mature SaaS S&M % rev (C16)
- **Value:** 22%
- **What:** Benchmark efficient S&M intensity.
- **How:** Yellow input vs FY25 33.2%.
- **Why:** Defines “bloat” dollars vs a mature SaaS target.
- **Source:** Model assumption (LBOMODEL1)


# Tab: `ZI_03_Headcount_AI`

## 45. % of S&M that are outbound SDR/BDR (C6)
- **Value:** 40%
- **What:** Share of S&M headcount treated as AI-replaceable SDRs.
- **How:** Yellow input × 1,370.
- **Why:** 10-K does not disclose SDR count — explicit model assumption.
- **Source:** Model assumption (LBOMODEL1)

## 46. Fully loaded cost / SDR / yr (C8)
- **Value:** $155k
- **What:** Cash cost per outbound SDR.
- **How:** Yellow input.
- **Why:** Below blended residual S&M $/employee (commissions/ads stay).
- **Source:** Model assumption (LBOMODEL1)

## 47. AI seat / oversight cost (C13/C15)
- **Value:** $15k per seat; $150k per oversight FTE
- **What:** Replacement opex for AI agents + humans.
- **How:** Yellow inputs; 1 oversight FTE / 20 seats.
- **Why:** Net savings must clear AI + oversight to hit margin target.
- **Source:** Model assumption (LBOMODEL1)

## 48. Cumulative % SDRs replaced (C21:E21)
- **Value:** 45% / 80% / 95% (2026–28)
- **What:** Headcount reduction schedule.
- **How:** Yellow cumulative exit fractions.
- **Why:** Phased cut to avoid one-year cliff risk.
- **Source:** Model assumption (LBOMODEL1)

## 49. Severance / AI build (C16/C17)
- **Value:** $25k per SDR; $8m Y1 build
- **What:** One-time costs.
- **How:** Yellow inputs in P&L impact.
- **Why:** Hits in-year EBITDA but excluded from run-rate bps.
- **Source:** Model assumption (LBOMODEL1)


# Tab: `ZI_04_EBITDA_Bridge`

## 50. Target margin expansion (C7)
- **Value:** 500 bps
- **What:** Minimum success hurdle for the AI program.
- **How:** Yellow input; YES/NO vs run-rate bps.
- **Why:**  equates to ~$62.5m EBITDA on $1,249.5m revenue.
- **Source:** User thesis requirement


# Tab: `ZI_05_Sensitivity`

## 51. Sensitivity grid (B12:H17)
- **Value:** Replace-% × AI $/seat → bps
- **What:** Y3 run-rate bps under alternate AI costs / cut depths.
- **How:** Algebra on SDR0 & unit costs.
- **Why:** Shows what it takes to stay above 500 bps if AI is more expensive.
- **Source:** Derived from ZI_03 inputs
