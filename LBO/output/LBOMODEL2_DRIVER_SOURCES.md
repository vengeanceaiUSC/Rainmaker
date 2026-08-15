# Assumptions_Drivers — Verbatim / Ctrl+F map

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx
**PDF:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_DRIVER_SOURCES.pdf

## Value changes for Task A (so yellow = findable text)
- Payroll cut **15% → 18%** (Ctrl+F `down 18%` on Digital Applied)
- Exit multiple **13 → 12.9x** (Ctrl+F `12.9x` on Aventis)
- Recomputed: AI IRR **27.4%**, Base **3.1%**, Y2 margin uplift **441 bps**

## CapEx 1–3% (C24) — why Ctrl+F may fail
Page uses an **en-dash**: `1–3%`. Searching hyphen `1-3%` can miss.
Use Ctrl+F: **`CapEx of 1`** → `Most pure-play SaaS companies report CapEx of 1–3% of revenue`
Yellow **2%** = MODEL midpoint of that band.

## Per-row Ctrl+F (column I in Excel)
### Row 5
- Source: https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results
- Ctrl+F: $1,214.3  →  "GAAP Revenue of $1,214.3 million, a decrease of 2% year-over-year."  |  10% growth: N/A (MODEL CONST)

### Row 6
- Source: https://www.getaleph.com/answers/saas-gross-margin-2026
- Ctrl+F: median software gross margin is 80  →  "The 2025 median software gross margin is 80%"

### Row 7
- Source: https://www.digitalapplied.com/blog/ai-sdr-statistics-2026-outbound-sales-data-points
- Ctrl+F: down 18%  →  "Net SDR headcount in US B2B SaaS companies is down 18% YoY in 2026 per Bridge Group"

### Row 8
- Source: https://syncgtm.com/blog/how-to-pay-sales-development-rep
- Ctrl+F: 70/30  →  "Standard SDR OTE in 2026 is $85,000 on a 70/30 base/variable split."  |  20% cut: N/A (MODEL CONST)

### Row 9
- Source: https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026
- Ctrl+F: $5,000-10,000  →  "11x.ai (Alice) $5,000-10,000+"  |  Ctrl+F: $10,000-$15,000  →  "Enterprise $10,000-$15,000+"  |  Ctrl+F: 1,513 in sales  →  "1,513 in sales and marketing"  |  25% share: N/A (MODEL CONST)

### Row 10
- Source: https://www.federalreserve.gov/faqs/economy_14400.htm
- Ctrl+F: 2 percent  →  "the FOMC judges that inflation of 2 percent over the longer run"

### Row 11
- Source: https://blossomstreetventures.medium.com/what-should-saas-spend-on-cogs-s-m-r-d-and-g-a-the-data-from-57-public-cos-8e88eabe99ca
- Ctrl+F: G&A was 18%  →  "G&A was 18%, 20%, 21, and 23% of revenue in 2025, 2024, 2023, and 2022."  |  Alt table Ctrl+F: MEDIAN then G&A col 18%

### Row 12
- Source: (none — MODEL CONST)
- 5% CapEx cut: N/A — MODEL CONST (no Ctrl+F string; no source link).

### Row 13
- Source: (none — MODEL CONST)
- 10% WC improvement: N/A — MODEL CONST (no Ctrl+F string; no source link).

### Row 14
- Source: https://www.global-rates.com/en/interest-rates/sofr/historical/2024/
- Ctrl+F: 4.30  →  table row "Lowest … 4.30 %" (SOFR 2024 First/Last/Highest/Lowest/Average)

### Row 15
- Source: https://ctacquisitions.com/acquisition-financing/
- Ctrl+F: 400-500  →  "Senior term loan A | … | SOFR + 400-500 bps"

### Row 16
- Source: https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine
- Ctrl+F: 300-500  →  "TLB pricing is typically SOFR + 300-500 basis points"  |  table Ctrl+F: Term Loan B

### Row 17
- Source: https://clearvaluelending.com/glossary/term-loan-b
- Ctrl+F: 1% annual  →  "minimal amortization (1% annually)"  |  Alt: "1% per year (nominal)"

### Row 18
- Source: https://ryanoconnellfinance.com/lbo-model-fundamentals/
- Ctrl+F: 50-100%  →  "apply 50-100% of annual excess free cash flow"  |  Ctrl+F: modeled at 100%  →  "typically modeled at 100%"

### Row 19
- Source: https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income
- Ctrl+F: flat rate of 21  →  "taxes resident corporations at a flat rate of 21%."

### Row 20
- Source: https://aventis-advisors.com/saas-valuation-multiples-how-much-is-my-saas-business-worth/
- Ctrl+F: 12.9x  →  table "EV/EBITDA | … | 1st quartile | 12.9x"  |  Alt Ctrl+F: 12.8x  →  "lower quartile are closer to 12.8x"

### Row 21
- Source: https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results
- Ctrl+F: $1,214.3  and  Ctrl+F: 414.1  →  Revenue $1,214.3 / Sales and marketing $414.1  |  425: MODEL ROUND of 35%×1214.3

### Row 22
- Source: https://syncgtm.com/blog/how-to-pay-sales-development-rep
- Ctrl+F: 70/30  →  "70/30 base/variable split"  |  318.8 / 75%: MODEL carve (not printed as $318.8)

### Row 23
- Source: https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies
- Ctrl+F: 30–35%  →  "Commission accounts for 30–35% of total pay"  |  106.2 / 25%: MODEL carve

### Row 24
- Source: https://saasdb.app/learn/financials/fcf-margin/
- Ctrl+F: CapEx of 1  →  "Most pure-play SaaS companies report CapEx of 1–3% of revenue"  |  Tip: search CapEx of 1 (en-dash); plain 1-3% may miss.  |  2%: MODEL midpoint of that band

### Row 25
- Source: (none — MODEL CONST)
- 2% WC plug: N/A — MODEL CONST (no Ctrl+F string; no source link).

### Row 27
- Source: (none — MODEL CONST)
- No separate third-party quote — sum of sourced C14 and C15.

### Row 28
- Source: (none — MODEL CONST)
- No separate third-party quote — sum of sourced C14 and C16.
