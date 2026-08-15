# SOURCE MATH → DRIVER

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx

**PDF:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_DRIVER_SOURCES.pdf

## Revenue growth = 0.1
- **FIND:** [FIND: ZI FY24 Revenue $1,214.3m (−2% YoY)](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **MATH:** SOURCE MATH → DRIVER: Filing shows FY24 rev $1,214.3m / FY23 $1,239.5m − 1 = −2.0% (not +10%). Driver C5=10% is a forward LBO underwrite held equal in AI vs Base so IRR delta is cost-driven. Math tie: baseline scale from source; growth rate is thesis, not the −2% print.
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## Gross margin = 0.8
- **FIND:** [FIND: Aleph/Benchmarkit software median GM = 80%](https://www.getaleph.com/answers/saas-gross-margin-2026)
- **MATH:** SOURCE MATH → DRIVER: Published 2025 software median GM = 80%. Driver C6=80% ⇒ COGS% = 1 − 0.80 = 20%; COGS_t = Rev_t × 0.20. (ZI GAAP GM ~84% in SEC comps would imply COGS 16%; we deliberately use the 80% median, not ZI’s higher print.)
- **Secondary:** https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc

## Sales payroll cut (Y1) = 0.15
- **FIND:** [FIND: US B2B SaaS SDR headcount −~18% YoY (junior −31%)](https://resources.rework.com/pt/news/sales-tech/ai-sdr-worth-it-2026-hybrid-model-sales-leader)
- **MATH:** SOURCE MATH → DRIVER: Published SDR headcount −~18% YoY. Driver C7=15% ≈ that cut. Payroll_Y1 = Payroll0 × (1 − 0.15) = 318.8 × 0.85 = $270.98m. Source supports ~15–18% outbound labor reduction; 15% is the model switch near the −18% evidence (not a ZI 10-K line).
- **Secondary:** https://firstsales.io/blog/cost-per-meeting-outbound/

## Variable commission cut = 0.2
- **FIND:** [FIND: SDR OTE ~70% base / ~30% variable](https://syncgtm.com/blog/how-to-pay-sales-development-rep)
- **MATH:** SOURCE MATH → DRIVER: Source variable share ≈ 30% of OTE. Driver C8=20% cut on commission pool: Commission_AI = 106.2 × (1 − 0.20) = $84.96m. Logic: AI doesn’t take commissions ⇒ shrink the ~30% variable layer; −20% is a partial cut of that published variable share.
- **Secondary:** https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies

## AI infrastructure ($m) = 34
- **FIND:** [FIND: SF $2/conversation; Flex Credits ~$0.10/action; ZI S&M HC 1,513](https://www.salesforce.com/news/press-releases/2025/05/15/agentforce-flexible-pricing-news/)
- **MATH:** SOURCE MATH → DRIVER: SF list $2/conv and ~$0.10/action. Outbound roles ≈ 25% × 1,513 S&M (10-K) = 378. At $5–10k/mo/agent: 378 × $5k × 12 = $22.7m; 378 × $10k × 12 = $45.4m. Driver C9 = avg(23,45) = $34m. $34m is derived — SF does not quote “$34m.”
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## S&M expense growth Y2+ = 0.02
- **FIND:** [FIND: Gartner — agentic/inference AI scales as compute opex, not headcount-linear](https://www.gartner.com/en/newsroom/press-releases/2026-08-10-gartner-forecasts-worldwide-artificial-intelligence-optimized-iaas-spending-to-grow-96-percent-in-2026)
- **MATH:** SOURCE MATH → DRIVER: Source: production agents scale via inference opex (not 1:1 hiring). Driver C10=2%: for t≥2, Payroll_t = Payroll_{t−1} × 1.02 (same for commissions). Example: 270.98 × 1.02 = $276.4m in Y2. 2% ≈ maintenance/inflation growth after severing revenue-linked hiring (thesis rate; Gartner doesn’t print “2%”).
- **Secondary:** https://www.ciodive.com/news/AI-spending-soars-enterprise-maturity/827488/

## G&A % of revenue = 0.18
- **FIND:** [FIND: ZI FY24 G&A $295.3m on Revenue $1,214.3m](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **MATH:** SOURCE MATH → DRIVER: Filing G&A/Rev = 295.3 / 1214.3 = 24.3%. Driver C11=18% is a leaner thesis rate. G&A_t = Rev_t × 0.18 (e.g. Y1 rev 1335.7 × 0.18 ≈ $240.4m). Source gives the actual 24.3% anchor; model deliberately uses 18% below that print.
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## CapEx reduction vs base = 0.05
- **FIND:** [FIND: human SDR all-in includes tooling/devices (~$12k tooling in stack examples)](https://caliberoutbound.com/blog/build-vs-buy-outbound-the-2026-cost-math)
- **MATH:** SOURCE MATH → DRIVER: Source shows devices/tooling inside SDR all-in cost. Driver C12=5% CapEx cut: CapEx_AI = Rev × CapEx_base% × (1 − 0.05) = Rev × 0.02 × 0.95 = Rev × 1.9%. Cutting ~15% outbound seats (C7) removes laptop refresh — 5% is the modeled CapEx haircut tied to that labor cut.
- **Secondary:** https://firstsales.io/blog/cost-per-meeting-outbound/

## WC / receivables improvement = 0.1
- **FIND:** [FIND: ZI receivables ~80 days (SEC-derived)](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc)
- **MATH:** SOURCE MATH → DRIVER: ~80 DSO ⇒ WC scales with growth. Driver C13=10% improvement: ΔNWC_AI = ΔRev × WC_base% × (1 − 0.10) = ΔRev × 0.02 × 0.90 = ΔRev × 1.8%. Source proves receivables intensity; −10% is the thesis collections gain from automated billing.
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## SOFR = 0.043
- **FIND:** [FIND live SOFR on FRED (e.g. ~3.62% on 2026-08-13); NY Fed publishes the rate](https://fred.stlouisfed.org/series/SOFR)
- **MATH:** SOURCE MATH → DRIVER: Official SOFR feeds leveraged-loan coupons. Driver C14=4.30% is a planning cushion above recent ~3.6% prints (+~70 bps buffer). TLA/TLB interest use C14 directly: rate = SOFR + spread. Open FRED for the exact daily print.
- **Secondary:** https://www.newyorkfed.org/markets/reference-rates/sofr

## Term Loan A spread = 0.04
- **FIND:** [FIND: TLA illustrative ~SOFR+275–375 bps](https://ryanoconnellfinance.com/lbo-debt-structure/)
- **MATH:** SOURCE MATH → DRIVER: Guide band SOFR+275–375bps. Driver C15=+400bps (0.04) is ~25–125bps above that band (conservative). TLA_rate = C14 + C15 = 4.30% + 4.00% = 8.30% (see C27).
- **Secondary:** https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending

## Term Loan B spread = 0.05
- **FIND:** [FIND: TLB typically SOFR+300–500 bps](https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine)
- **MATH:** SOURCE MATH → DRIVER: Published TLB band SOFR+300–500bps. Driver C16=+500bps = top of that band. TLB_rate = C14 + C16 = 4.30% + 5.00% = 9.30% (see C28).
- **Secondary:** https://ryanoconnellfinance.com/lbo-debt-structure/

## Mandatory amortization = 0.01
- **FIND:** [FIND: TLB scheduled amort typically ~1%/year](https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending)
- **MATH:** SOURCE MATH → DRIVER: Source standard = ~1% of original principal / year. Driver C17=1%. Mandatory_t = Initial_new_debt × 0.01 (e.g. 915.5 × 0.01 ≈ $9.2m/yr before optional sweep).
- **Secondary:** https://clearvaluelending.com/glossary/term-loan-b

## Cash sweep % = 1
- **FIND:** [FIND: LBO design uses excess cash / optional prepay to delever](https://ryanoconnellfinance.com/lbo-debt-structure/)
- **MATH:** SOURCE MATH → DRIVER: Source mechanism = optional prepay after mandatory amort. Driver C18=100%: Optional_sweep = 1.0 × max(0, FCF − Mandatory); Debt_end = Debt_beg − Mandatory − Optional. 100% is the switch setting maximizing that standard delever path.
- **Secondary:** https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine

## Tax rate = 0.21
- **FIND:** [FIND: US federal corporate tax rate = 21%](https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income)
- **MATH:** SOURCE MATH → DRIVER: Statutory federal CIT = 21%. Driver C19=21%. Tax_t = max(0, EBT_t) × 0.21; NI_t = EBT_t − Tax_t. Direct 1:1 match to the published statutory rate.
- **Secondary:** https://tradingeconomics.com/united-states/corporate-tax-rate

## Exit EV/EBITDA multiple = 13
- **FIND:** [FIND: PE SaaS buyouts often ~15–22x EBITDA](https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean)
- **MATH:** SOURCE MATH → DRIVER: Source band ~15–22x. Driver C20=13.0x is deliberately ~2–9 turns below that band. Exit_EV = Y5_EBITDA × 13. Locked equal in AI vs Base so MOIC/IRR uplift comes from ops, not multiple expansion.
- **Secondary:** https://aventis-advisors.com/software-valuation-multiples/

## S&M baseline ($m) = 425
- **FIND:** [FIND: ZI FY24 S&M $414.1m; Revenue $1,214.3m](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **MATH:** SOURCE MATH → DRIVER: Filing S&M/Rev = 414.1 / 1214.3 = 34.1%. Driver C21=$425m ⇒ 425 / 1214.3 = 35.0% (rounded near the filing). KeyBanc <10% growth cohort S&M ~31% also sits in the same ~30–35% neighborhood.
- **Secondary:** https://www.key.com/content/dam/kco/documents/businesses___institutions/2024_kbcm_sapphire_saas_survey.pdf

## Payroll baseline ($m) = 318.8
- **FIND:** [FIND: SDR OTE ~70% base / ~30% variable](https://syncgtm.com/blog/how-to-pay-sales-development-rep)
- **MATH:** SOURCE MATH → DRIVER: Source base share ~70% of OTE. Driver C22 = 75% × S&M baseline = 0.75 × 425 = $318.8m payroll pool. Allocation set near the published base-heavy mix (slightly above 70% because total S&M also includes non-OTE costs).
- **Secondary:** https://pulserevops.com/knowledge/ra0197

## Commission baseline ($m) = 106.2
- **FIND:** [FIND: tech SDR variable ~30–35% of OTE](https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies)
- **MATH:** SOURCE MATH → DRIVER: Source variable ~30–35% of OTE. Driver C23 = 25% × 425 = $106.2m (and 318.8 + 106.2 = 425). 25% is a slightly conservative carve of total S&M vs the ~30% OTE variable share.
- **Secondary:** https://syncgtm.com/blog/how-to-pay-sales-development-rep

## CapEx base % of rev = 0.02
- **FIND:** [FIND: ZI FY24 Unlevered FCF $446.9m on $1,214.3m rev (asset-light)](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- **MATH:** SOURCE MATH → DRIVER: uFCF/Rev = 446.9 / 1214.3 ≈ 36.8% ⇒ low capital intensity. Driver C24=2%: CapEx_base = Rev × 0.02 (before AI −5% in C12). 2% is a maintenance CapEx planning plug consistent with asset-light FCF, not a filing CapEx% line.
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## WC base % of Δrev = 0.02
- **FIND:** [FIND: ZI receivables ~80 days](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc)
- **MATH:** SOURCE MATH → DRIVER: 80/365 ≈ 21.9% of annual sales theoretically tied in AR at a point in time; growth still needs incremental WC. Driver C25=2%: ΔNWC = ΔRev × 0.02 (before AI −10% in C13). Simplified annual plug anchored to that receivables reality.
- **Secondary:** https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm

## TLA rate (SOFR+4) = =C14+C15
- **FIND:** [FIND: TLA coupon = SOFR + senior bank spread](https://ryanoconnellfinance.com/lbo-debt-structure/)
- **MATH:** SOURCE MATH → DRIVER: C27 = C14 + C15 = 4.30% + 4.00% = 8.30%. Components sourced on rows 14–15 (official SOFR + TLA spread band). Interest_TLA = TLA_beg × 8.30%.
- **Secondary:** https://fred.stlouisfed.org/series/SOFR

## TLB rate (SOFR+5) = =C14+C16
- **FIND:** [FIND: TLB coupon = SOFR + institutional spread](https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine)
- **MATH:** SOURCE MATH → DRIVER: C28 = C14 + C16 = 4.30% + 5.00% = 9.30%. Components sourced on rows 14 & 16 (official SOFR + TLB SOFR+300–500bps band). Interest_TLB = TLB_beg × 9.30%.
- **Secondary:** https://fred.stlouisfed.org/series/SOFR
