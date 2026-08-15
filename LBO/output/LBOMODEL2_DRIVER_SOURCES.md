# CLOSED MATH → TOTAL $

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx

## Revenue growth = 0.1
- FIND: [FIND: FY24 Rev $1,214.3m / FY23 $1,239.5m](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- CLOSED MATH: Source growth = 1214.3/1239.5 − 1 = −2.03%. Driver C5 is NOT that print — C5=10% is a thesis forward rate. Total $ path from C5: Y1 Rev = 1214.3×1.10 = $1,335.73m; Y5 Rev = 1214.3×1.10^5 = $1,955.0m. Source supplies the $1,214.3m base only.

## Gross margin = 0.8
- FIND: [FIND: software median GM = 80%](https://www.getaleph.com/answers/saas-gross-margin-2026)
- CLOSED MATH: Source median GM = 80% ⇒ C6=0.80. Total $: COGS_t = Rev_t × (1−0.80) = Rev_t × 0.20. At Y1 Rev $1,335.73m → COGS = $267.15m; Gross profit = $1,068.58m. Direct 1:1 copy of the published 80% median into C6.

## Sales payroll cut (Y1) = 0.15
- FIND: [FIND: US B2B SaaS SDR HC −~18% YoY](https://resources.rework.com/pt/news/sales-tech/ai-sdr-worth-it-2026-hybrid-model-sales-leader)
- CLOSED MATH: Source cut ≈ 18%. Set C7=15% (near 18%). Total $: Payroll_Y1 = 318.8 × (1−0.15) = 318.8 × 0.85 = $270.98m. Savings vs baseline = 318.8 − 270.98 = $47.82m. Source gives the ~15–18% cut rate; multiply by our $318.8m payroll pool to get the dollar total.

## Variable commission cut = 0.2
- FIND: [FIND: SDR OTE ~70% base / ~30% variable](https://syncgtm.com/blog/how-to-pay-sales-development-rep)
- CLOSED MATH: Source variable share ≈ 30%. Set C8=20% cut of commission pool. Total $: Commission_AI = 106.2 × (1−0.20) = $84.96m. Savings = 106.2 − 84.96 = $21.24m. Unit fact (30% variable) → pool size $106.2m → ×20% cut → dollar savings (not just a %).

## AI infrastructure ($m) = 34
- FIND: [FIND: SF $2/conv & ~$0.10/action; ZI S&M HC 1,513; AI agents ~$5–10k/mo](https://www.salesforce.com/news/press-releases/2025/05/15/agentforce-flexible-pricing-news/)
- CLOSED MATH (TOTAL $): Roles = 0.25 × 1,513 = 378. Low = 378 × $5,000/mo × 12 = $22.7m. High = 378 × $10,000/mo × 12 = $45.4m. C9 = ($22.7+$45.4)/2 = $34.0m. SF unit prices alone are NOT the total — they are a cross-check: $34m/$2 = 17.0m conversations/yr, or $34m/$0.10 = 340m actions/yr. Primary total comes from roles × $/mo × 12.

## S&M expense growth Y2+ = 0.02
- FIND: [FIND: agentic AI scales via inference opex (not 1:1 headcount)](https://www.gartner.com/en/newsroom/press-releases/2026-08-10-gartner-forecasts-worldwide-artificial-intelligence-optimized-iaas-spending-to-grow-96-percent-in-2026)
- CLOSED MATH: Thesis growth C10=2%. Total $ path: Payroll_Y2 = 270.98 × 1.02 = $276.40m; Payroll_Y5 = 270.98 × 1.02^4 = $293.45m. Commission same factor. Source justifies WHY growth can be low (compute scales); 2% × prior-year $ balance is what produces the dollar totals.

## G&A % of revenue = 0.18
- FIND: [FIND: G&A $295.3m / Rev $1,214.3m = 24.3%](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- CLOSED MATH: Source rate = 295.3/1214.3 = 24.32%. Driver C11=18% (thesis below filing). Total $: G&A_t = Rev_t × 0.18. Y1: 1335.73 × 0.18 = $240.43m (vs filing-like 1335.73 × 0.2432 ≈ $324.9m). Source gives the $295.3m and % anchor; × Rev produces the model dollar total at 18%.

## CapEx reduction vs base = 0.05
- FIND: [FIND: SDR all-in includes devices/tooling dollars](https://caliberoutbound.com/blog/build-vs-buy-outbound-the-2026-cost-math)
- CLOSED MATH: CapEx_AI = Rev × 0.02 × (1−0.05) = Rev × 0.019. Y1: 1335.73 × 0.019 = $25.38m vs base 1335.73 × 0.02 = $26.71m → saves $1.34m. Source shows device $ inside SDR cost; C7’s labor cut × C12’s 5% rate turns that into a CapEx $ total.

## WC / receivables improvement = 0.1
- FIND: [FIND: receivables ~80 days](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc)
- CLOSED MATH: ΔNWC_AI = ΔRev × 0.02 × (1−0.10) = ΔRev × 0.018. Y1 ΔRev = 1335.73−1214.3 = 121.43 → ΔNWC = 121.43 × 0.018 = $2.19m (vs base 121.43 × 0.02 = $2.43m). Source’s ~80 DSO motivates a WC $ plug; rate × ΔRev is the total.

## SOFR = 0.043
- FIND: [FIND: official daily SOFR (recent ~3.62%)](https://fred.stlouisfed.org/series/SOFR)
- CLOSED MATH: C14=4.30% (planning). Dollar interest uses balances × rate — e.g. TLA interest ≈ TLA_beg × (C14+C15). If TLA_beg ≈ $457.8m and all-in 8.30%, Interest_TLA ≈ 457.8 × 0.083 = $38.0m. Source is the rate series; $ total = rate × debt principal.

## Term Loan A spread = 0.04
- FIND: [FIND: TLA ~SOFR+275–375bps band](https://ryanoconnellfinance.com/lbo-debt-structure/)
- CLOSED MATH: C15=+400bps. All-in TLA = 4.30%+4.00% = 8.30%. $ total: Interest_TLA = TLA_beg × 0.0830 (e.g. ~$457.8m × 0.083 ≈ $38.0m). Source gives the spread band; spread + SOFR + beginning balance = dollar interest.

## Term Loan B spread = 0.05
- FIND: [FIND: TLB SOFR+300–500bps](https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine)
- CLOSED MATH: C16=+500bps (= top of band). All-in TLB = 4.30%+5.00% = 9.30%. $ total: Interest_TLB = TLB_beg × 0.0930 (e.g. ~$274.7m × 0.093 ≈ $25.5m). Source spread × SOFR × principal → dollars.

## Mandatory amortization = 0.01
- FIND: [FIND: ~1%/yr scheduled amort](https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending)
- CLOSED MATH: C17=1%. Initial new debt ≈ 183.1×5.0 = $915.5m. Mandatory = 915.5 × 0.01 = $9.155m per year. Source % × initial principal = dollar mandatory paydown.

## Cash sweep % = 1
- FIND: [FIND: excess cash / optional prepay to delever](https://ryanoconnellfinance.com/lbo-debt-structure/)
- CLOSED MATH: C18=100%. Optional_sweep_$ = 1.00 × max(0, FCF_$ − Mandatory_$). Example if FCF=$80m and Mandatory=$9.2m → sweep = $70.8m of debt paydown. Source is the mechanism; FCF $ × 100% is the dollar total applied to debt.

## Tax rate = 0.21
- FIND: [FIND: federal CIT = 21%](https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income)
- CLOSED MATH: C19=21%. Tax_$ = max(0, EBT_$) × 0.21. If EBT=$100m → Tax=$21m; NI=$79m. Source rate × taxable $ = tax dollars (1:1 statutory match).

## Exit EV/EBITDA multiple = 13
- FIND: [FIND: PE SaaS often ~15–22x EBITDA](https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean)
- CLOSED MATH: C20=13.0x (below band). Exit_EV_$ = Y5_EBITDA_$ × 13. If Y5 EBITDA ≈ $470m → Exit EV ≈ $6,110m. Source gives the multiple neighborhood; EBITDA $ × 13 is the enterprise-value total.

## S&M baseline ($m) = 425
- FIND: [FIND: S&M $414.1m; Rev $1,214.3m](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- CLOSED MATH: Source total = $414.1m (34.1% of rev). Driver C21=$425.0m ≈ 414.1 × (0.35/0.341) ≈ rounded to 35%×1214.3 = 0.35×1214.3 = $425.0m. Filing dollars → % → rounded model $ total.

## Payroll baseline ($m) = 318.8
- FIND: [FIND: ~70% of SDR OTE is base](https://syncgtm.com/blog/how-to-pay-sales-development-rep)
- CLOSED MATH: C22 = 0.75 × 425 = $318.8m. Check: 318.8/425 = 75% (near source ~70% base). Dollar total is the allocation of the $425m S&M pool — not a % left hanging.

## Commission baseline ($m) = 106.2
- FIND: [FIND: variable ~30–35% of OTE](https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies)
- CLOSED MATH: C23 = 0.25 × 425 = $106.2m. Check: 318.8 + 106.2 = 425.0 exactly. Source variable ~30% informs the 25% carve; dollars = % × $425m S&M pool.

## CapEx base % of rev = 0.02
- FIND: [FIND: uFCF $446.9m on $1,214.3m rev](https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results)
- CLOSED MATH: Source implies low CapEx intensity (large FCF). C24=2% → CapEx_$ = Rev × 0.02. Y1 base CapEx = 1335.73 × 0.02 = $26.71m. Rate × revenue $ = CapEx dollars.

## WC base % of Δrev = 0.02
- FIND: [FIND: ~80 receivable days](https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc)
- CLOSED MATH: C25=2% → ΔNWC_$ = ΔRev_$ × 0.02. Y1: ΔRev $121.43m × 0.02 = $2.43m WC use. Source DSO motivates a WC need; ΔRev × rate is the dollar total.

## TLA rate (SOFR+4) = =C14+C15
- FIND: [FIND: TLA = SOFR + spread](https://ryanoconnellfinance.com/lbo-debt-structure/)
- CLOSED MATH: C27 = C14+C15 = 4.30%+4.00% = 8.30%. $ total = TLA_beg × 8.30% (e.g. $457.8m × 0.083 ≈ $38.0m interest). Formula rate × principal balance.

## TLB rate (SOFR+5) = =C14+C16
- FIND: [FIND: TLB = SOFR + spread](https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine)
- CLOSED MATH: C28 = C14+C16 = 4.30%+5.00% = 9.30%. $ total = TLB_beg × 9.30% (e.g. $274.7m × 0.093 ≈ $25.5m interest). Formula rate × principal balance.
