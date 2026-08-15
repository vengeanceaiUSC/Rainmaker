# LBOMODEL2 — Assumptions_Drivers Source Map

Sources live **on the same row** as each driver (`Assumptions_Drivers` columns E–G). There is **no** separate AI_Cost_Sources tab.

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx

**PDF:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_AI_COST_SOURCES.pdf

| Row | Driver | Source | Why reasonable |
|---|---|---|---|
| 5 | Revenue growth 10% | [ZoomInfo FY2024 results (BusinessWire)](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results) | ZI FY24 revenue fell ~2%, so 10% is a forward recovery/LBO underwrite—not trailing actual. Held constant in AI vs Base so IRR uplift comes from cost takeout, not inflated top-line. |
| | | Secondary: https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean | |
| 6 | Gross margin 80% | [Aleph/Benchmarkit SaaS GM median ~80%](https://www.getaleph.com/answers/saas-gross-margin-2026) | 2025 software median gross margin is ~80% (Aleph×Benchmarkit). ZI already prints ~84% GAAP GM; using 80% is a conservative SaaS benchmark, not an aggressive stretch. |
| | | Secondary: https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc | |
| 7 | Sales payroll cut −15% | [AI SDR vs human SDR cost/meeting](https://firstsales.io/blog/cost-per-meeting-outbound/) | Published outbound math: fully loaded human SDRs ~$110–150k vs much cheaper AI stacks. Cutting human sales payroll 15% in Y1 is a modest slice of ZI’s 1,513 S&M headcount. |
| | | Secondary: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm | |
| 8 | Commission cut −20% | [SDR pay mix ~70/30 base/variable](https://syncgtm.com/blog/how-to-pay-sales-development-rep) | SaaS SDR OTE is typically ~70% base / ~30% variable. Replacing outbound humans with AI logically shrinks the commission pool; −20% is a partial, not full, cut of that variable layer. |
| | | Secondary: https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies | |
| 9 | AI infrastructure $15m | [Salesforce Agentforce $2/conv + Flex Credits](https://www.salesforce.com/news/press-releases/2025/05/15/agentforce-flexible-pricing-news/) | SF publishes $2/conversation and ~$0.10/action Flex Credits. At ZI scale (~378 outbound roles at 25% of 1,513 S&M), $5–10k/mo AI agents imply ~$23–45m; $15m is conservative. |
| | | Secondary: https://highticketaisystems.com/blog/ai-sdr-pricing-explained | |
| 10 | S&M growth Y2+ 2% | [Gartner: agentic AI scales via inference opex](https://www.gartner.com/en/newsroom/press-releases/2026-08-10-gartner-forecasts-worldwide-artificial-intelligence-optimized-iaas-spending-to-grow-96-percent-in-2026) | Once agents replace headcount, S&M opex need not grow with revenue. Cap growth at ~2% (near inflation/maintenance) to model infinite agent scale vs linear human hiring. |
| | | Secondary: https://www.ciodive.com/news/AI-spending-soars-enterprise-maturity/827488/ | |
| 11 | G&A % rev 18% | [ZoomInfo FY24 G&A $295.3m (~24% rev)](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results) | ZI FY24 G&A was $295.3m (~24% of $1,214m). We lock ~18% as a leaner but still corporate G&A load; thesis does not touch finance/comp overhead—only the sales floor. |
| | | Secondary: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm | |
| 12 | CapEx reduction −5% | [In-house SDR tooling/laptop stack costs](https://caliberoutbound.com/blog/build-vs-buy-outbound-the-2026-cost-math) | Human SDR all-in budgets include laptops, seats, and tooling. Terminating outbound humans cuts recurring device/refresh CapEx; −5% vs base CapEx rate is a small, directionally correct haircut. |
| | | Secondary: https://firstsales.io/blog/cost-per-meeting-outbound/ | |
| 13 | WC improvement −10% | [ZI receivables ~80 days (SEC ratios)](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc) | ZI receivables turnover ~80 days shows WC is material. Automated billing reminders can shorten DSO; modeling −10% WC intensity vs base is a modest collections improvement, not a wipeout. |
| | | Secondary: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm | |
| 14 | SOFR ~4.3% | [NY Fed official SOFR page](https://www.newyorkfed.org/markets/reference-rates/sofr) | SOFR is the NY Fed’s official overnight Treasury-repo financing benchmark used to price leveraged loans. ~4.3% is a planning level near recent SOFR prints (see also FRED SOFR series). |
| | | Secondary: https://fred.stlouisfed.org/series/SOFR | |
| 15 | TLA spread +4% | [TLA typical SOFR+275–375bps (illustrative)](https://ryanoconnellfinance.com/lbo-debt-structure/) | Bank Term Loan A in LBO structures typically prices tighter than TLB. SOFR+400bps sits at the wider end of published TLA ranges—conservative senior pricing independent of AI ops. |
| | | Secondary: https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending | |
| 16 | TLB spread +5% | [TLB often SOFR+300–500bps; 1%/yr amort](https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine) | Institutional TLB commonly prices SOFR+300–500bps. SOFR+500bps is the upper published band—credit-market priced, not changed by internal AI software strategy. |
| | | Secondary: https://ryanoconnellfinance.com/lbo-debt-structure/ | |
| 17 | Mandatory amort 1% | [Syndicated TLB ~1% annual amort standard](https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending) | Market-standard institutional TLB schedules ~1% annual amortization with a bullet at maturity. Kept fixed because loan docs—not AI FCF—set the mandatory floor. |
| | | Secondary: https://clearvaluelending.com/glossary/term-loan-b | |
| 18 | Cash sweep 100% | [LBO excess cash / optional prepay practice](https://ryanoconnellfinance.com/lbo-debt-structure/) | Sponsor LBOs typically sweep excess free cash to optional debt prepay after mandatory amort. 100% sweep maximizes deleveraging so AI margin gains show up in MOIC/IRR, not idle cash. |
| | | Secondary: https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine | |
| 19 | Tax rate 21% | [PwC: US federal corporate rate 21%](https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income) | US federal corporate income tax is a flat 21% statutory rate (TCJA). Operating-margin gains do not change the federal statutory rate used in the model. |
| | | Secondary: https://tradingeconomics.com/united-states/corporate-tax-rate | |
| 20 | Exit multiple 13x | [PE SaaS buyouts often ~15–22x EBITDA](https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean) | PE SaaS deals often underwrite ~15–22x EBITDA. Locking exit at 13x in BOTH AI and Base cases is deliberately conservative so equity creation is attributed to ops, not multiple expansion. |
| | | Secondary: https://aventis-advisors.com/software-valuation-multiples/ | |
| 21 | S&M baseline $425m | [ZI FY24 S&M $414.1m on $1,214.3m rev](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results) | Company-reported FY24 S&M = $414.1m (~34% of $1,214.3m revenue). Model uses ~$425m (~35%) as a rounded baseline consistent with that filing. |
| | | Secondary: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm | |
| 22 | Payroll baseline $318.8m | [SDR OTE mostly base (~70%)](https://syncgtm.com/blog/how-to-pay-sales-development-rep) | With ~70/30 base/variable SDR mix, most S&M people cost is payroll/base. Payroll baseline = S&M − commission pool (~75% of S&M) matches that pay-mix structure. |
| | | Secondary: https://pulserevops.com/knowledge/ra0197 | |
| 23 | Commission baseline $106.2m | [Variable ~30–35% of SDR OTE](https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies) | Tech SDR variable is typically ~30–35% of OTE. Setting commission pool at 25% of S&M is a slightly conservative carve of total S&M for variable payouts. |
| | | Secondary: https://syncgtm.com/blog/how-to-pay-sales-development-rep | |
| 24 | CapEx base 2% rev | [ZI FY24 unlevered FCF / light CapEx SaaS](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results) | SaaS models run low CapEx vs industrials; ZI still printed strong FY24 FCF. 2% of revenue is a simple maintenance CapEx planning rate before the AI −5% haircut. |
| | | Secondary: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm | |
| 25 | WC base 2% Δrev | [ZI receivables days ~80 (WC intensity)](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc) | Receivables ~80 days imply WC scales with growth. Using 2% of Δrevenue as ΔNWC is a simplified SaaS WC plug anchored to that receivables reality. |
| | | Secondary: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm | |

Educational / research use only — not investment advice.
