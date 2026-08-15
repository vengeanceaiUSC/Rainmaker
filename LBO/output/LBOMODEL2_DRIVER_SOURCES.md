# Sources only if used in the equation

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx

AI $34m = `0.25×1513×($5–10k)×12` midpoint.
**Not a source:** Salesforce $2/conv (not in equation).

## Revenue growth = 0.1
- EQ uses: FY24 Rev $1,214.3m (ZI results). Growth 10% = MODEL CONST (not in filing; filing is −2%).
- EQUATION: Y1_Rev = 1214.3 × (1 + 0.10) = $1,335.73m. SOURCED input: 1214.3 from ZI FY24 results. MODEL CONST (unsourced rate): 0.10 — filing growth = 1214.3/1239.5−1 = −2.03% is NOT used. Only list ZI results because it supplies the $1,214.3m base in the equation.

## Gross margin = 0.8
- EQ uses: software median GM = 80% (Aleph/Benchmarkit) → C6=0.80
- EQUATION: COGS_t = Rev_t × (1 − 0.80) = Rev_t × 0.20. Y1 example: 1335.73 × 0.20 = $267.15m. SOURCED input: 0.80 from Aleph/Benchmarkit median. No other inputs. Source number = driver number.

## Sales payroll cut (Y1) = 0.15
- EQ uses: published SDR HC cut ~18% → model sets C7=15% (near 18%); payroll pool $318.8m
- EQUATION: Payroll_Y1 = 318.8 × (1 − 0.15) = $270.98m; savings = $47.82m. SOURCED rate evidence: ~18% US B2B SaaS SDR headcount decline (we set 15% ≈ 18%). SOURCED $ pool: 318.8 = 0.75 × 425 from pay-mix rows. If we used 18% exactly: 318.8 × 0.82 = $261.4m — we chose 15% as the nearby model switch.

## Variable commission cut = 0.2
- EQ uses: ~30% variable share of OTE (SyncGTM) → informs cutting commission pool; pool $106.2m; C8=20%
- EQUATION: Commission_AI = 106.2 × (1 − 0.20) = $84.96m; savings = $21.24m. SOURCED: ~30% variable (SyncGTM) explains why a commission pool exists to cut. MODEL CONST: 20% cut rate (partial take of the ~30% variable layer). SOURCED $ pool: 106.2 = 0.25 × 425. Equation dollars come from pool × cut — not from an unstated unit price.

## AI infrastructure ($m) = 34
- EQ uses: ZI S&M HC 1,513 (10-K) × 25% roles × $5–10k/mo agents × 12 → avg $34m. NOT SF $2/conv.
- EQUATION (only these inputs): Roles = 0.25 × 1,513 = 378. Low = 378 × $5,000 × 12 = $22.7m. High = 378 × $10,000 × 12 = $45.4m. C9 = ($22.7+$45.4)/2 = $34.0m. SOURCED: 1,513 from ZI 10-K Human Capital. SOURCED: $5k–$10k/mo enterprise AI SDR agent range (secondary link). MODEL CONST: 0.25 outbound share of S&M. REMOVED: SF $2/conv and $0.10/action — they do not appear in this equation, so they are not sources for C9.

## S&M expense growth Y2+ = 0.02
- EQ uses: prior-year S&M $ × (1+C10); C10=2% = MODEL CONST. Gartner supports non-linear headcount scaling (qualitative).
- EQUATION: Payroll_t = Payroll_{t−1} × (1.02). Y2 = 270.98 × 1.02 = $276.40m. SOURCED $ start: 270.98 from row-7 equation. MODEL CONST: 0.02 growth. Gartner is listed only as qualitative support that agent opex need not track headcount 1:1 — it does NOT supply the 2%. If you require every rate to be numeric-from-source, treat 2% as pure model const (Gartner optional).

## G&A % of revenue = 0.18
- EQ uses: Rev_t × C11; C11=18% = MODEL CONST. Filing G&A $295.3m/1214.3=24.3% is the sourced alternative rate (not used).
- EQUATION USED: G&A_t = Rev_t × 0.18 → Y1 1335.73 × 0.18 = $240.43m. SOURCED Rev: from growth path / ZI base. MODEL CONST: 0.18. SOURCED BUT NOT USED IN EQ: filing G&A rate 295.3/1214.3 = 24.32% (would give Y1 G&A $324.9m). Honesty: 18% is not taken from the filing number.

## CapEx reduction vs base = 0.05
- EQ uses: Rev × 2% CapEx_base × (1−0.05). 5% = MODEL CONST; CapEx base 2% = MODEL CONST. Source shows device $ exist in SDR all-in (qualitative).
- EQUATION: CapEx_AI = Rev × 0.02 × 0.95 = Rev × 0.019 → Y1 1335.73 × 0.019 = $25.38m. MODEL CONSTS in EQ: 0.02 and 0.05. Source is qualitative (SDR stacks include hardware $) — it does not print 5% or 2%. Labeled honestly: device-cost existence from source; rates are model consts.

## WC / receivables improvement = 0.1
- EQ uses: ΔRev × 2% × (1−0.10). 10% & 2% = MODEL CONST. Source supplies ~80 DSO (motivates WC, not the 10%).
- EQUATION: ΔNWC_AI = ΔRev × 0.02 × 0.90 = ΔRev × 0.018 → Y1 121.43 × 0.018 = $2.19m. SOURCED: ~80 receivable days (WC is real). MODEL CONSTS: 0.02 WC rate, 0.10 improvement. 80 DSO is not algebraically equal to 10% — improvement rate is model const.

## SOFR = 0.043
- EQ uses: C14 as SOFR in TLA/TLB rates. C14=4.30% = MODEL PLANNING (FRED recent ~3.62% — not equal).
- EQUATION: TLA_rate = C14 + C15; Interest_TLA ≈ TLA_beg × (C14+C15). SOURCED series: FRED/NY Fed SOFR. MODEL CONST: 4.30% cushion (≠ latest ~3.62% print). To be strictly equal to source, set C14 to the FRED print; currently a planning cushion above it.

## Term Loan A spread = 0.04
- EQ uses: C15=+400bps. Source band SOFR+275–375bps — 400 is ABOVE band (conservative MODEL choice).
- EQUATION: TLA_rate = 4.30% + 4.00% = 8.30%; Interest_TLA = TLA_beg × 0.083. SOURCED band: ~275–375bps. MODEL CONST used: 400bps (not inside the cited midpoint). Honesty: we are outside the guide’s upper 375bps by 25bps.

## Term Loan B spread = 0.05
- EQ uses: C16=+500bps = TOP of sourced TLB band SOFR+300–500bps
- EQUATION: TLB_rate = 4.30% + 5.00% = 9.30%; Interest_TLB = TLB_beg × 0.093. SOURCED: 500bps is the published upper bound of SOFR+300–500. Source number (upper bound) = driver number.

## Mandatory amortization = 0.01
- EQ uses: C17=1% = sourced TLB amort standard; principal ≈ $915.5m
- EQUATION: Mandatory = 915.5 × 0.01 = $9.155m/yr. SOURCED: 1% amort from TLB market standard. SOURCED structure: 915.5 = 183.1 EBITDA × 5.0x leverage (model leverage turns). 1% from source × principal = dollar mandatory.

## Cash sweep % = 1
- EQ uses: C18=100% × max(0, FCF − Mandatory). 100% = MODEL CONST (mechanism sourced, not the 100%).
- EQUATION: Sweep_$ = 1.00 × max(0, FCF_$ − Mandatory_$). SOURCED: optional prepay / excess-cash delever mechanism. MODEL CONST: 100% sweep fraction. Dollar total = FCF path × 100% — rate is model switch.

## Tax rate = 0.21
- EQ uses: C19=21% = sourced US federal CIT
- EQUATION: Tax_$ = max(0, EBT_$) × 0.21. SOURCED: 21% statutory rate (PwC). Source number = driver number. If EBT=$100m → Tax=$21m.

## Exit EV/EBITDA multiple = 13
- EQ uses: C20=13x. Source band 15–22x — 13 is BELOW band (MODEL CONST conservative).
- EQUATION: Exit_EV = Y5_EBITDA × 13. SOURCED band: ~15–22x (not used as 13). MODEL CONST: 13.0x. Honesty: 13 does not equal a number printed in the source; source only bounds a higher range we under-cut.

## S&M baseline ($m) = 425
- EQ uses: target ≈35%×$1,214.3m. Filing S&M $414.1m (34.1%) → round to C21=$425m
- EQUATION: 0.35 × 1214.3 = $425.005 ≈ C21=$425.0m. SOURCED: Rev 1214.3 and S&M 414.1 (34.1%) from ZI FY24 results. MODEL ROUND: choose 35% near 34.1% → $425m. Filing dollars drive the equation; 35% is the rounding choice.

## Payroll baseline ($m) = 318.8
- EQ uses: ~70% base share (SyncGTM) → set 75% × $425m = $318.8m
- EQUATION: C22 = 0.75 × 425 = $318.8m. SOURCED: ~70% base of OTE (SyncGTM) → we use 75% of S&M pool. SOURCED $: 425 from row 21. 75% is a model allocation near the sourced ~70%.

## Commission baseline ($m) = 106.2
- EQ uses: residual 25% × $425m = $106.2m (near sourced ~30–35% variable)
- EQUATION: C23 = 0.25 × 425 = $106.2m; check 318.8+106.2=425. SOURCED: variable ~30–35% of OTE. MODEL: 25% of S&M as commission pool (residual after 75% payroll). Dollars from 425 × 0.25 — residual identity, informed by sourced variable share.

## CapEx base % of rev = 0.02
- EQ uses: CapEx = Rev × 0.02. 2% = MODEL CONST. Source uFCF $446.9m shows low CapEx intensity (qualitative).
- EQUATION: CapEx_base_$ = Rev × 0.02 → Y1 1335.73 × 0.02 = $26.71m. MODEL CONST: 0.02. SOURCED qualitative: uFCF 446.9 on 1214.3 rev supports asset-light profile — does not equal 2%.

## WC base % of Δrev = 0.02
- EQ uses: ΔNWC = ΔRev × 0.02. 2% = MODEL CONST. Source ~80 DSO motivates WC (not equal to 2%).
- EQUATION: ΔNWC_$ = ΔRev × 0.02 → Y1 121.43 × 0.02 = $2.43m. MODEL CONST: 0.02. SOURCED: ~80 DSO. Not algebraically 2% — WC plug is model const.

## TLA rate (SOFR+4) = =C14+C15
- EQ uses: C27 = C14 + C15 only (SOFR + TLA spread). No other sources.
- EQUATION: C27 = 4.30% + 4.00% = 8.30%. Interest_TLA = TLA_beg × 0.083. Inputs solely from C14 and C15 (their sources/notes on those rows). No SF or unrelated links.

## TLB rate (SOFR+5) = =C14+C16
- EQ uses: C28 = C14 + C16 only (SOFR + TLB spread).
- EQUATION: C28 = 4.30% + 5.00% = 9.30%. Interest_TLB = TLB_beg × 0.093. Inputs solely from C14 and C16. TLB +500bps sourced on row 16 as top of SOFR+300–500 band.
