# LBOMODEL2 — Assumptions_Drivers Source Map

All commentary + links live on **Assumptions_Drivers** (same row). **No** `AI_Cost_Sources` tab.

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx

**PDF:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_DRIVER_SOURCES.pdf

## Revenue growth
- **Why reasonable:** Assumes steady top-line expansion. Reasonable for a mature enterprise software company, avoiding aggressive hypergrowth assumptions while maintaining a healthy, realistic sales trajectory.
- **Source:** [ZoomInfo FY2024 results (BusinessWire)](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **Credibility:** ZI disclosed FY24 revenue of $1,214.3m (−2% YoY). Using 10% forward growth is a mature-SaaS recovery underwrite held equal in AI vs Base—not hypergrowth.
- **Feeds:** `AI_Operating!D5` (Go to AI_Operating!D5 (Revenue))
- **Secondary:** https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean

## Gross margin
- **Why reasonable:** High profitability on direct services. Highly reasonable for a SaaS company, which historically scales digital infrastructure with minimal incremental cost per user.
- **Source:** [Aleph/Benchmarkit SaaS GM median ~80%](https://www.getaleph.com/answers/saas-gross-margin-2026)
- **Credibility:** Independent SaaS benchmarks put software median GM near 80%; ZI already prints ~84% GAAP GM, so 80% is conservative and industry-standard.
- **Feeds:** `AI_Operating!D6` (Go to AI_Operating!D6 (COGS))
- **Secondary:** https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc

## Sales payroll cut (Y1)
- **Why reasonable:** Immediate headcount reduction for human SDRs. Reasonable as a private equity cost-cutting measure, easily achieved by eliminating the lower-tier outbound sales team.
- **Source:** [AI SDR vs human SDR cost/meeting analysis](https://firstsales.io/blog/cost-per-meeting-outbound/)
- **Credibility:** Published outbound economics show fully loaded human SDRs far costlier than AI stacks. −15% Y1 payroll is a modest PE-style cut of outbound headcount, not a full S&M wipeout.
- **Feeds:** `AI_Operating!D7` (Go to AI_Operating!D7 (Sales payroll))
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## Variable commission cut
- **Why reasonable:** Reduces payouts for automated sales. Reasonable because AI agents do not collect commission checks, retaining more profit per enterprise deal directly for the company.
- **Source:** [SDR pay mix ~70/30 base/variable (SyncGTM)](https://syncgtm.com/blog/how-to-pay-sales-development-rep)
- **Credibility:** SaaS SDR OTE is typically ~70% base / ~30% variable. Replacing humans with AI shrinks the commission pool; −20% is a partial cut of that variable layer.
- **Feeds:** `AI_Operating!D8` (Go to AI_Operating!D8 (Commissions))
- **Secondary:** https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies

## AI infrastructure ($m)
- **Why reasonable:** Fixed cost for AI software and compute. Reasonable estimate for enterprise-scale AI deployment, covering API calls and infrastructure needed to replace human reps.
- **Source:** [Salesforce Agentforce $2/conv + Flex Credits press](https://www.salesforce.com/news/press-releases/2025/05/15/agentforce-flexible-pricing-news/)
- **Credibility:** SF publishes $2/conversation and ~$0.10/action. At ZI scale (~378 outbound roles = 25% of 1,513 S&M), $5–10k/mo agents imply ~$23–45m; $15m is conservative.
- **Feeds:** `AI_Operating!D10` (Go to AI_Operating!D10 (AI infrastructure))
- **Secondary:** https://highticketaisystems.com/blog/ai-sdr-pricing-explained

## S&M expense growth Y2+
- **Why reasonable:** Caps future marketing cost increases. Reasonable because AI scales infinitely without needing proportional headcount additions, severing the link between revenue growth and human hiring.
- **Source:** [Gartner: agentic AI / inference scales as opex](https://www.gartner.com/en/newsroom/press-releases/2026-08-10-gartner-forecasts-worldwide-artificial-intelligence-optimized-iaas-spending-to-grow-96-percent-in-2026)
- **Credibility:** Gartner documents production agentic AI as inference opex that scales without linear headcount. Cap S&M growth at 2% to model that break from revenue-linked hiring.
- **Feeds:** `AI_Operating!E7` (Go to AI_Operating!E7 (payroll Y2+))
- **Secondary:** https://www.ciodive.com/news/AI-spending-soars-enterprise-maturity/827488/

## G&A % of revenue
- **Why reasonable:** General administrative overhead costs. Reasonable for a public-to-private SaaS company, maintaining standard corporate expenses without relying on unrealistic operational magic outside of sales.
- **Source:** [ZoomInfo FY24 G&A $295.3m (company results)](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **Credibility:** ZI FY24 G&A was $295.3m (~24% of revenue). Locking ~18% keeps standard corporate overhead while the thesis only attacks the sales floor—not finance/comp magic.
- **Feeds:** `AI_Operating!D12` (Go to AI_Operating!D12 (G&A))
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## CapEx reduction vs base
- **Why reasonable:** Lowers physical equipment spending. Reasonable because firing 15% of the sales team permanently eliminates the need to purchase and refresh their laptops and hardware.
- **Source:** [In-house SDR all-in includes devices/tooling](https://caliberoutbound.com/blog/build-vs-buy-outbound-the-2026-cost-math)
- **Credibility:** Human SDR all-in budgets include laptops and seats. Cutting outbound humans removes refresh CapEx; −5% vs base CapEx rate is a small, directionally correct haircut.
- **Feeds:** `AI_Operating!D19` (Go to AI_Operating!D19 (CapEx))
- **Secondary:** https://firstsales.io/blog/cost-per-meeting-outbound/

## WC / receivables improvement
- **Why reasonable:** Speeds up cash collection from clients. Reasonable because automated AI billing systems can aggressively follow up on late invoices, freeing up cash for debt paydown.
- **Source:** [ZI receivables ~80 days (SEC-derived ratios)](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc)
- **Credibility:** ZI receivables ~80 days show WC is real. Automated billing follow-ups can shorten DSO; −10% WC intensity vs base is a modest collections gain, not a wipeout.
- **Feeds:** `AI_Operating!D20` (Go to AI_Operating!D20 (ΔNWC))
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## SOFR
- **Why reasonable:** The foundational secured lending interest rate. Reasonable and historically accurate for the macroeconomic environment, properly reflecting actual institutional baseline borrowing costs.
- **Source:** [NY Fed official SOFR reference rate](https://www.newyorkfed.org/markets/reference-rates/sofr)
- **Credibility:** SOFR is the NY Fed’s official overnight Treasury-repo financing benchmark used to price leveraged loans. ~4.3% is a planning level near recent SOFR prints (see FRED).
- **Feeds:** `Debt_Sweep_AI!C12` (Go to Debt_Sweep_AI!C12 (TLA rate uses SOFR))
- **Secondary:** https://fred.stlouisfed.org/series/SOFR

## Term Loan A spread
- **Why reasonable:** The risk premium for senior debt. Reasonable for a standard LBO, reflecting typical syndicated commercial bank pricing for secured corporate borrowing.
- **Source:** [LBO debt structure: TLA illustrative spreads](https://ryanoconnellfinance.com/lbo-debt-structure/)
- **Credibility:** Bank TLA in LBO guides typically prices tighter than TLB (often SOFR+~275–375bps). SOFR+400bps is a conservative senior bank spread independent of AI ops.
- **Feeds:** `Assumptions_Drivers!C27` (Go to Assumptions_Drivers!C27 (TLA rate))
- **Secondary:** https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending

## Term Loan B spread
- **Why reasonable:** The risk premium for riskier debt. Reasonable, as Term Loan B is usually held by institutional investors requiring higher yields for slightly higher risk.
- **Source:** [TLB pricing often SOFR+300–500bps](https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine)
- **Credibility:** Institutional TLB commonly prices SOFR+300–500bps. SOFR+500bps is the upper published band—set by syndicated loan markets, not by internal AI strategy.
- **Feeds:** `Assumptions_Drivers!C28` (Go to Assumptions_Drivers!C28 (TLB rate))
- **Secondary:** https://ryanoconnellfinance.com/lbo-debt-structure/

## Mandatory amortization
- **Why reasonable:** Required annual principal repayment. Highly reasonable and standard market practice for leveraged loans, forcing a tiny minimum debt reduction each year.
- **Source:** [Syndicated TLB ~1% annual amort standard](https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending)
- **Credibility:** Market-standard institutional TLB schedules ~1% annual amortization with a bullet maturity. Loan docs—not AI FCF—set this mandatory floor.
- **Feeds:** `Debt_Sweep_AI!D5` (Go to Debt_Sweep_AI!D5 (Mandatory 1%))
- **Secondary:** https://clearvaluelending.com/glossary/term-loan-b

## Cash sweep %
- **Why reasonable:** Dedicates all excess cash to debt. Reasonable and standard for maximizing LBO returns, ensuring every free dollar generated immediately aggressively reduces outstanding leverage.
- **Source:** [LBO optional prepay / excess cash practice](https://ryanoconnellfinance.com/lbo-debt-structure/)
- **Credibility:** Sponsor LBOs typically sweep excess FCF to optional debt prepay after mandatory amort. 100% sweep maximizes deleveraging so AI margins hit MOIC/IRR.
- **Feeds:** `Debt_Sweep_AI!E5` (Go to Debt_Sweep_AI!E5 (Total paydown))
- **Secondary:** https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine

## Tax rate
- **Why reasonable:** Corporate income tax obligation. Perfectly reasonable because it strictly matches the current standard U.S. federal statutory corporate tax rate without aggressive tax-dodging assumptions.
- **Source:** [PwC Tax Summaries: US federal CIT 21%](https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income)
- **Credibility:** US federal corporate income tax is a flat 21% statutory rate. Operating-margin gains do not change the federal statutory rate used here.
- **Feeds:** `AI_Operating!D18` (Go to AI_Operating!D18 (Net income))
- **Secondary:** https://tradingeconomics.com/united-states/corporate-tax-rate

## Exit EV/EBITDA multiple
- **Why reasonable:** The valuation multiple upon selling. Reasonable to keep it conservative; relying on a 13x exit assumes no market inflation, proving returns come from operational growth.
- **Source:** [PE SaaS buyouts often ~15–22x EBITDA](https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean)
- **Credibility:** PE SaaS deals often underwrite ~15–22x EBITDA. Locking exit at 13x in BOTH cases is deliberately conservative so returns come from ops, not multiple expansion.
- **Feeds:** `LBO!D17` (Go to LBO!D17 (Exit multiple))
- **Secondary:** https://aventis-advisors.com/software-valuation-multiples/

## S&M baseline ($m)
- **Why reasonable:** This is historical assumptions without AI. Total initial sales spend is reasonable, representing roughly 35% of FY24 revenue, perfectly aligning with standard SaaS benchmarks.
- **Source:** [ZI FY24 S&M $414.1m on $1,214.3m revenue](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **Credibility:** Company-reported FY24 S&M = $414.1m (~34% of $1,214.3m). Model ~$425m (~35%) is a rounded baseline consistent with that filing.
- **Feeds:** `AI_Operating!C9` (Go to AI_Operating!C9 (S&M human FY24A))
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## Payroll baseline ($m)
- **Why reasonable:** This is historical assumptions without AI. The fixed salary portion is reasonable, representing 75% of S&M, accurately reflecting that human software sales floors are base-salary heavy.
- **Source:** [SDR OTE mostly base (~70%) — SyncGTM](https://syncgtm.com/blog/how-to-pay-sales-development-rep)
- **Credibility:** With ~70/30 base/variable SDR mix, most people cost is payroll. Payroll baseline = 75% of S&M matches that base-salary-heavy sales-floor structure.
- **Feeds:** `AI_Operating!C7` (Go to AI_Operating!C7 (Payroll FY24A))
- **Secondary:** https://pulserevops.com/knowledge/ra0197

## Commission baseline ($m)
- **Why reasonable:** This is historical assumptions without AI. The variable payout is reasonable, calculating exactly to the remaining 25% of the budget, representing standard quota-based enterprise sales incentives.
- **Source:** [Variable ~30–35% of tech SDR OTE](https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies)
- **Credibility:** Tech SDR variable is typically ~30–35% of OTE. Setting commission pool at 25% of S&M is a slightly conservative carve of total S&M for variable payouts.
- **Feeds:** `AI_Operating!C8` (Go to AI_Operating!C8 (Commissions FY24A))
- **Secondary:** https://syncgtm.com/blog/how-to-pay-sales-development-rep

## CapEx base % of rev
- **Why reasonable:** This is historical assumptions without AI. Normal hardware and capitalized software spending is reasonable for asset-light tech companies investing heavily in code rather than factories.
- **Source:** [ZI FY24 strong FCF / asset-light SaaS profile](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **Credibility:** SaaS is asset-light vs industrials; ZI still printed strong FY24 FCF. 2% of revenue is a simple maintenance CapEx planning rate before the AI −5% haircut.
- **Feeds:** `AI_Operating!D19` (Go to AI_Operating!D19 (CapEx))
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## WC base % of Δrev
- **Why reasonable:** This is historical assumptions without AI. Normal working capital needs are reasonable, showing that as revenue grows, a small percentage of cash gets tied up in daily operations.
- **Source:** [ZI receivables days ~80 (WC intensity)](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc)
- **Credibility:** Receivables ~80 days imply WC scales with growth. Using 2% of Δrevenue as ΔNWC is a simplified SaaS WC plug anchored to that receivables reality.
- **Feeds:** `AI_Operating!D20` (Go to AI_Operating!D20 (ΔNWC))
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## TLA rate (SOFR+4)
- **Why reasonable:** Total Term Loan A interest. Mathematically correct and reasonable for the total borrowing cost demanded by senior lenders in this specific rate environment.
- **Source:** [LBO TLA = SOFR + senior bank spread](https://ryanoconnellfinance.com/lbo-debt-structure/)
- **Credibility:** Computed as SOFR (C14) + TLA spread (C15). All-in ~8.3% matches senior secured bank pricing in this SOFR + 400bps setup.
- **Feeds:** `Debt_Sweep_AI!C12` (Go to Debt_Sweep_AI!C12 (TLA rate))
- **Secondary:** https://www.newyorkfed.org/markets/reference-rates/sofr

## TLB rate (SOFR+5)
- **Why reasonable:** Total Term Loan B interest. Mathematically correct and perfectly reasonable, pricing the institutional debt risk appropriately above the safer Term Loan A.
- **Source:** [LBO TLB = SOFR + institutional spread](https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine)
- **Credibility:** Computed as SOFR (C14) + TLB spread (C16). All-in ~9.3% prices institutional TLB risk appropriately above TLA in this structure.
- **Feeds:** `Debt_Sweep_AI!D12` (Go to Debt_Sweep_AI!D12 (TLB rate))
- **Secondary:** https://www.newyorkfed.org/markets/reference-rates/sofr
