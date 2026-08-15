# Sources only when IN the equation

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx

## AI $34m — sources that feed the math
1. **$5,000–$10,000+/mo** enterprise AI SDR (11x band): https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026
2. **1,513 S&M employees**: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm
3. **MODEL CONST:** 25% outbound share
4. Cross-check floor (not used as $5k): 11x Growth **$3,750/mo** https://www.11x.ai/products/alice/pricing

EQ: `0.25×1513=378` → `378×$5k×12=$22.7m` / `378×$10k×12=$45.4m` → avg **$34m**

## Other drivers
### Revenue growth = 0.1
- IN EQ: FY24 Rev=$1,214.3m (ZI). NOT in EQ: −2% growth (filing). C5=10% = MODEL CONST.
- EQ: Y1_Rev = 1214.3 × (1+0.10) = $1,335.73m. SOURCED in EQ: 1214.3 from ZI FY24 results. MODEL CONST in EQ: 0.10 (filing growth −2.03% is NOT an input). Removed as if it drove 10%: nothing — we do not cite a growth source for 10% because we do not have one.

### Gross margin = 0.8
- IN EQ: software median GM=80% → C6=0.80 (Aleph/Benchmarkit)
- EQ: COGS = Rev × (1−0.80). Y1: 1335.73×0.20=$267.15m. SOURCED in EQ: 0.80 = published software median GM. Source number = driver.

### Sales payroll cut (Y1) = 0.15
- IN EQ: ~18% SDR HC decline (set C7=15%≈18%); payroll pool $318.8m
- EQ: Payroll_Y1 = 318.8×(1−0.15)=$270.98m. SOURCED rate evidence: ~18% YoY US B2B SaaS SDR headcount decline → model switch 15%≈18%. SOURCED $: 318.8 from row22 equation. If exact 18%: 318.8×0.82=$261.4m (we use 15%).

### Variable commission cut = 0.2
- IN EQ: ~30% variable OTE (explains commission pool); pool $106.2m; C8=20% MODEL CONST
- EQ: Commission_AI = 106.2×(1−0.20)=$84.96m. SOURCED: ~30% variable share (why a commission pool exists). SOURCED $: 106.2=0.25×425. MODEL CONST: 0.20 cut (partial take of ~30% variable layer).

### AI infrastructure ($m) = 34
- IN EQ: AI SDR $5k–$10k/mo (White Space / 11x enterprise band) — click
- EQ: Roles=0.25×1,513=378; Low=378×$5,000×12=$22.7m; High=378×$10,000×12=$45.4m; C9=avg=$34.0m. SOURCED $/mo: White Space AI SDR guide lists 11x.ai (Alice) at $5,000–10,000+/mo (enterprise). SOURCED HC: 1,513 S&M employees from ZI 10-K (secondary link). MODEL CONST: 0.25 outbound share of S&M. Cross-check: 11x official Growth starts $3,750/mo (11x.ai Alice pricing) — below our $5k low; we use the published $5–10k enterprise band that feeds this EQ. NOT in EQ / removed: Salesforce $/conversation pricing.

### S&M expense growth Y2+ = 0.02
- IN EQ: C10=2% = Fed longer-run inflation target 2% (FOMC)
- EQ: Payroll_t = Payroll_{t−1}×(1.02). Y2=270.98×1.02=$276.40m. SOURCED in EQ: 0.02 from Fed 2% longer-run inflation goal. SOURCED $ start: 270.98 from row7. Removed: Gartner qualitative agent note (supplied no numeric 2%).

### G&A % of revenue = 0.18
- Filing G&A rate=295.3/1214.3=24.3% NOT used. C11=18% = MODEL CONST (no source equals 18%).
- EQ USED: G&A=Rev×0.18 → Y1 1335.73×0.18=$240.43m. MODEL CONST: 0.18 (no publication in our pack prints 18% G&A for ZI). SOURCED BUT NOT IN EQ: filing 24.3% (would be Y1 $324.9m). Honesty: do not treat the filing as the source of 18%.

### CapEx reduction vs base = 0.05
- C12=5% CapEx cut = MODEL CONST (no sourced 5%). Device $ in SDR stacks is qualitative only — removed as EQ source.
- EQ: CapEx_AI=Rev×0.02×(1−0.05)=Rev×0.019 → Y1=$25.38m. MODEL CONSTS: 0.05 and (with row24) 0.02. No source listed: prior laptop-stack articles do not print 5%, so they do not drive C12.

### WC / receivables improvement = 0.1
- Source ~80 DSO motivates WC but is NOT equal to C13=10%. 10% = MODEL CONST.
- EQ: ΔNWC_AI=ΔRev×0.02×0.90 → Y1 121.43×0.018=$2.19m. MODEL CONST: 0.10 improvement. SOURCED but not algebraic in 10%: ~80 DSO (shows WC exists). 80/365≈21.9% AR/sales stock ≠ 10% improvement rate.

### SOFR = 0.043
- FRED/NY Fed define SOFR series. C14=4.30% = MODEL PLANNING (≠ latest ~3.62% print).
- EQ: rates use C14 directly (TLA=C14+C15). SOURCED: SOFR is the official series (FRED). MODEL CONST value: 4.30% planning cushion above recent ~3.62%. To match source exactly, set C14 to the FRED print; currently not equal.

### Term Loan A spread = 0.04
- Source TLA band ~SOFR+275–375bps. C15=+400bps is ABOVE band → MODEL CONST (25bps over).
- EQ: TLA_rate=4.30%+4.00%=8.30%. SOURCED band: 275–375bps (not 400). MODEL CONST: 400bps. Honesty: source does not equal 400.

### Term Loan B spread = 0.05
- IN EQ: TLB SOFR+300–500bps → C16=+500bps = upper bound of source
- EQ: TLB_rate=4.30%+5.00%=9.30%. SOURCED in EQ: 500bps = top of published SOFR+300–500 band. Source upper bound = driver.

### Mandatory amortization = 0.01
- IN EQ: ~1%/yr TLB amort → C17=1%; principal≈$915.5m
- EQ: Mandatory=915.5×0.01=$9.155m/yr. SOURCED: 1% amort standard. Source number = driver.

### Cash sweep % = 1
- C18=100% sweep = MODEL CONST. LBO guides describe optional prepay mechanism but do not print 100%.
- EQ: Sweep=1.00×max(0,FCF−Mandatory). MODEL CONST: 1.00. Removed as EQ source: qualitative LBO delever articles (no 100% figure).

### Tax rate = 0.21
- IN EQ: US federal CIT=21% → C19=0.21 (PwC)
- EQ: Tax=EBT×0.21. SOURCED: 21% statutory. Source number = driver.

### Exit EV/EBITDA multiple = 13
- Source PE SaaS ~15–22x. C20=13x is BELOW band → MODEL CONST (not equal to source).
- EQ: Exit_EV=Y5_EBITDA×13. MODEL CONST: 13.0x. SOURCED band 15–22x is NOT used as 13. Honesty: source does not equal driver.

### S&M baseline ($m) = 425
- IN EQ: Rev $1,214.3m; choose 35%≈filing S&M 34.1% ($414.1m) → C21=$425m
- EQ: 0.35×1214.3=$425.0m=C21. SOURCED: 1214.3 rev and 414.1 S&M (34.1%) from ZI FY24. MODEL ROUND: 35% near 34.1%.

### Payroll baseline ($m) = 318.8
- IN EQ: ~70% base OTE → set 75%×$425m=$318.8m
- EQ: C22=0.75×425=$318.8m. SOURCED: ~70% base (SyncGTM). MODEL: 75% of S&M pool near that mix. SOURCED $: 425 from row21.

### Commission baseline ($m) = 106.2
- IN EQ: residual 25%×$425m=$106.2m (near sourced ~30–35% variable)
- EQ: C23=0.25×425=$106.2m; 318.8+106.2=425. SOURCED: variable ~30–35% OTE. MODEL: 25% of S&M as commission carve.

### CapEx base % of rev = 0.02
- IN EQ: SaaS CapEx typically 1–3% of rev → C24=2% = midpoint of sourced band
- EQ: CapEx=Rev×0.02 → Y1=$26.71m. SOURCED: pure-play SaaS CapEx 1–3% of revenue (SaaSDB). 2% = midpoint of 1–3%. Source band drives driver.

### WC base % of Δrev = 0.02
- C25=2% ΔNWC/ΔRev = MODEL CONST. ~80 DSO is not equal to 2% — removed as if it produced 2%.
- EQ: ΔNWC=ΔRev×0.02 → Y1=$2.43m. MODEL CONST: 0.02. No source equals 2% WC plug (80 DSO would imply ~22% AR/sales stock, a different metric).

### TLA rate (SOFR+4) = =C14+C15
- EQ: C27=C14+C15 only (inputs from rows 14–15)
- EQ: 4.30%+4.00%=8.30%. Interest_TLA=TLA_beg×0.083. No third-party source beyond C14/C15 component rows.

### TLB rate (SOFR+5) = =C14+C16
- EQ: C28=C14+C16 only; C16=+500bps sourced as TLB upper bound on row16
- EQ: 4.30%+5.00%=9.30%. Interest_TLB=TLB_beg×0.093. Component sources on rows 14 & 16.
