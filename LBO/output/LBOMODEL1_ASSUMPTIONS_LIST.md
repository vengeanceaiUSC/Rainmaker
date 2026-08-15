# vengeanceaiUSC-LBOMODEL1 — Assumptions List

**Excel:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx

**PDF:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.pdf

Primary engine = LBO / AI_Operating / Debt_Sweep_AI (not DCF).

## 1. Revenue Growth — 10% per year (unchanged by AI)
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C5`
- **Math location:** `AI_Operating!D5:H5`
- **Math:** Revenue_t = Revenue_{t-1} × (1 + Assumptions_Drivers!C5). Same growth in AI and Base cases.
- **Commentary:** Annual growth remains stable at ten percent. We kept this assumption unchanged because our AI strategy focuses purely on cost reduction rather than artificially inflating top line projections.

## 2. Gross Margin — 80%
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C6`
- **Math location:** `AI_Operating!D6:H6 (COGS)`
- **Math:** COGS_t = Revenue_t × (1 − Assumptions_Drivers!C6); Gross Profit_t = Revenue_t × C6.
- **Commentary:** Expansion reached precisely eighty percent. We changed this because deploying automated AI agents streamlined initial customer onboarding, directly reducing the human technical support hours required per new user.

## 3. Sales Payroll — −15% in Year 1
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C7`
- **Math location:** `AI_Operating!D7 (Sales payroll Y1)`
- **Math:** Payroll_Y1 = Assumptions_Drivers!C22 × (1 − C7). Baseline payroll = C22.
- **Commentary:** Baseline costs dropped by fifteen percent immediately. We changed this to reflect the aggressive termination of outbound human development representatives, instantly eliminating their recurring annual base salary expenses.

## 4. Variable Commissions — −20% of baseline commission pool from Y1
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C8`
- **Math location:** `AI_Operating!D8:H8 (Commissions)`
- **Math:** Commissions_AI = Assumptions_Drivers!C23 × (1 − C8); Base keeps full C23 pool scaled with revenue.
- **Commentary:** Rates decreased significantly starting in year one. Replacing human representatives with AI meant fewer outbound deals required traditional payouts, allowing the company to retain higher profit per sale.

## 5. AI Infrastructure — $15 million fixed / year
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C9`
- **Math location:** `AI_Operating!D10:H10 (AI infrastructure)`
- **Math:** AI_cost_t = Assumptions_Drivers!C9 each forecast year (AI case only). EBITDA reduces by this amount.
- **Commentary:** A new fixed fifteen million dollar expense was added. We included this vital change to cover enterprise software licenses, API usage limits, and compute power running the AI.

## 6. Sales Growth (S&M expense growth) — 2% per year from Year 2
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C10`
- **Math location:** `AI_Operating!E7:H7 / E8:H8 (payroll & commissions Y2–Y5)`
- **Math:** For t≥2: Payroll_t = Payroll_{t-1} × (1 + Assumptions_Drivers!C10); same for commissions.
- **Commentary:** Expense growth was permanently capped at two percent. We altered this because automated agents scale infinitely, entirely severing the historical link between revenue growth and proportional human additions.

## 7. Administrative Costs (G&A) — Locked at baseline % of revenue
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C11`
- **Math location:** `AI_Operating!D12:H12 (G&A)`
- **Math:** G&A_t = Revenue_t × Assumptions_Drivers!C11. Identical % in AI and Base cases.
- **Commentary:** General overhead remained locked at baseline percentages. We avoided changing this because the value creation strategy strictly impacts the active sales floor without altering corporate finance or compensation.

## 8. Capital Expenditures — −5% vs baseline CapEx rate
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C12`
- **Math location:** `AI_Operating!D19:H19 (CapEx)`
- **Math:** CapEx_AI = Revenue_t × Assumptions_Drivers!C24 × (1 − C12); Base uses C24 only.
- **Commentary:** Hardware spending was reduced by five percent. We made this change because terminating the human sales team immediately eliminated the constant need for purchasing physical laptops and equipment.

## 9. Working Capital — Receivables / WC requirement −10% vs base
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C13`
- **Math location:** `AI_Operating!D20:H20 (ΔNWC)`
- **Math:** ΔNWC_AI = ΔRevenue_t × Assumptions_Drivers!C25 × (1 − C13).
- **Commentary:** Receivables collection periods improved slightly overall. Implementing automated AI billing reminders accelerated cash collections from late customers, lowering total working capital requirements and freeing extra cash for debt.

## 10. Term Loan A interest — SOFR + 4.00%
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C14 / C15`
- **Math location:** `Debt_Sweep_AI!C12 (TLA rate); Interest on AI_Operating!D17:H17`
- **Math:** TLA_rate = Assumptions_Drivers!C14 + C15 (=C27). Interest_TLA = TLA_beg × TLA_rate.
- **Commentary:** The interest rate remained static at SOFR plus four. This was unchanged because the macroeconomic lending environment and the base company credit profile strictly dictate senior pricing.

## 11. Term Loan B interest — SOFR + 5.00%
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C14 / C16`
- **Math location:** `Debt_Sweep_AI!D12 (TLB rate); Interest on AI_Operating!D17:H17`
- **Math:** TLB_rate = Assumptions_Drivers!C14 + C16 (=C28). Interest_TLB = TLB_beg × TLB_rate.
- **Commentary:** Pricing was maintained exactly at SOFR plus five. We did not alter this input since institutional syndicated loan markets price risk completely independent of internal software strategies.

## 12. Mandatory Amortization — 1% of initial principal / year
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C17`
- **Math location:** `Debt_Sweep_AI!D5:D9 (Mandatory 1%)`
- **Math:** Mandatory_t = Initial_Debt × Assumptions_Drivers!C17 (floor paydown before optional sweep).
- **Commentary:** Principal repayment schedules were held constant at one percent. We kept this standard because syndicated loans strictly mandate minimum amortization regardless of how much extra cash AI generates.

## 13. Cash Sweep — 100% of excess cash to debt
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C18`
- **Math location:** `Debt_Sweep_AI!E5:E9 (Total paydown)`
- **Math:** Optional_sweep_t = C18 × max(0, FCF_t − Mandatory_t); Debt_end = Debt_beg − Mandatory − Optional_sweep.
- **Commentary:** One hundred percent of excess cash targets debt paydown. This remained unchanged to ensure all newly generated margins from AI savings aggressively reduce leverage and maximize sponsor returns.

## 14. Tax Rate — 21%
- **Engine:** LBO (not DCF)
- **Input:** `Assumptions_Drivers!C19`
- **Math location:** `AI_Operating!D18:H18 (Net income)`
- **Math:** Tax_t = max(0, EBIT_t − Interest_t) × Assumptions_Drivers!C19; NI_t = EBT_t − Tax_t.
- **Commentary:** The corporate tax rate remained steady at twenty one percent. We did not change this assumption because federal statutory tax obligations remain unaffected by internal operating margin improvements.

## 15. Exit Multiple — 13.0x EV / EBITDA
- **Engine:** LBO (Ex LBO!D17; not DCF)
- **Input:** `Assumptions_Drivers!C20`
- **Math location:** `LBO!D17 (also Strategy_Summary!D10)`
- **Math:** Exit_EV = Y5_EBITDA × Assumptions_Drivers!C20; Exit_Equity = Exit_EV − Debt_end; MOIC / IRR from that.
- **Commentary:** The valuation multiple was securely locked at thirteen. We kept this exact multiple consistent across scenarios to prove equity value creation stemmed purely from AI operational improvements alone.
