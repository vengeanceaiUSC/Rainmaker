#!/usr/bin/env python3
"""
Driver–source converse reasoning for LBOMODEL2.

Resolve contradictions:
- Keep driver only if source-vs-driver gap is justified.
- Else set driver to the sourced figure (or drop unsourced cuts to 0)
  and refresh Assumptions_Drivers A–K + recompute ops / strategy.
"""

from __future__ import annotations

import math
import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL2.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL2.xlsx"

# --- Adjusted drivers (post converse reasoning) ---
REV_G = 0.05  # mature CAGR CY2022→CY2025 ≈4.4%≈5%; FY24 −2% was outlier (only decline)
GM = 0.80  # Aleph median 80% — KEEP (matches cited source; ZI GAAP ~84% noted as not used)
PAYROLL_CUT = 0.18  # Digital Applied — KEEP
COMM_CUT = 0.18  # was 20% MODEL CONST → align to same 18% SDR decline source
# AI roles = same 18% of S&M HC (was unjustified 25%)
SM_HC = 1513
AI_ROLE_SHARE = 0.18
AI_LOW = SM_HC * AI_ROLE_SHARE * 5000 * 12 / 1e6
AI_HIGH = SM_HC * AI_ROLE_SHARE * 10000 * 12 / 1e6
AI_COST = round((AI_LOW + AI_HIGH) / 2, 1)  # ~24.5
SM_G = 0.02  # Fed 2% restored — valid again with positive forward revenue growth
GA_PCT = 0.18  # Blossom median — KEEP as post-LBO target (ZI ~24% not used)
CAPEX_CUT = 0.0  # was 5% MODEL CONST — no source → no cut
WC_IMPROVE = 0.0  # was 10% MODEL CONST — no source → no improve
SOFR = 0.043
TLA_SP = 0.04
TLB_SP = 0.05
MAND = 0.01
SWEEP = 1.0
TAX = 0.21
EXIT_MULT = 12.9
SM0 = 414.1  # was 425; exact ZI FY24 S&M
PAY0 = round(SM0 * 0.70, 1)  # SyncGTM 70% base → 289.9
COMM0 = round(SM0 * 0.30, 1)  # SyncGTM 30% variable → 124.2
CAPEX_BASE = 0.02  # SaaSDB midpoint — KEEP
WC_BASE = 0.0  # was 2% MODEL CONST — no source → 0 plug
REV0 = 1214.3
RD0 = 196.1
DA0 = 85.7
EBITDA0 = 183.1

ZI_FY25 = "https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm"
ZI_URL = (
    "https://markets.financialcontent.com/stocks/article/"
    "bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results"
)
ZI_10K = "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm"
WHITE = "https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026"
MINI = "https://www.miniloop.ai/blog/11x-pricing"
DIGITAL = "https://www.digitalapplied.com/blog/ai-sdr-statistics-2026-outbound-sales-data-points"
SYNC_PAY = "https://syncgtm.com/blog/how-to-pay-sales-development-rep"
SYNC_COMM = "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies"
ALEPH = "https://www.getaleph.com/answers/saas-gross-margin-2026"
FED = "https://www.federalreserve.gov/faqs/economy_14400.htm"
BLOSSOM = (
    "https://blossomstreetventures.medium.com/"
    "what-should-saas-spend-on-cogs-s-m-r-d-and-g-a-the-data-from-57-public-cos-8e88eabe99ca"
)
SAASDB = "https://saasdb.app/learn/financials/fcf-margin/"
GLOBAL_RATES = "https://www.global-rates.com/en/interest-rates/sofr/historical/2024/"
CT = "https://ctacquisitions.com/acquisition-financing/"
LBO_DEBT = (
    "https://ibinterviewquestions.com/guides/valuation-investment-banking/"
    "lbo-debt-structures-senior-subordinated-mezzanine"
)
CLEAR = "https://clearvaluelending.com/glossary/term-loan-b"
RYAN = "https://ryanoconnellfinance.com/lbo-model-fundamentals/"
PWC = "https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income"
AVENTIS = (
    "https://aventis-advisors.com/saas-valuation-multiples-how-much-is-my-saas-business-worth/"
)

YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="C6EFCE")
WRAP = Alignment(wrap_text=True, vertical="top")
FONT = Font(name="Calibri", size=9)
LINK_F = Font(name="Calibri", color="0563C1", underline="single", size=9)
THIN = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)


def hyperlink(url: str, label: str) -> str:
    # Escape quotes in label for Excel formula
    lab = label.replace('"', '""')
    return f'=HYPERLINK("{url}","{lab}")'


def style_row(ws, r: int, fill):
    for c in range(1, 12):
        cell = ws.cell(r, c)
        if cell.value is not None or c <= 6:
            cell.border = THIN
            cell.alignment = WRAP
            if c not in (5, 8) or not (isinstance(cell.value, str) and cell.value.startswith("=HYPERLINK")):
                if c != 3:
                    cell.font = FONT
        if c == 3 and fill is not None:
            cell.fill = fill


def set_driver_row(ws, r, *, baseline, driver, value, why, e, f, feeds, h, i, j, k, fill):
    ws.cell(r, 1, baseline)
    ws.cell(r, 2, driver)
    ws.cell(r, 3, value)
    ws.cell(r, 4, why)
    ws.cell(r, 5, e)
    if isinstance(e, str) and e.startswith("=HYPERLINK"):
        ws.cell(r, 5).font = LINK_F
    ws.cell(r, 6, f)
    ws.cell(r, 7, feeds)
    ws.cell(r, 8, h)
    if isinstance(h, str) and h.startswith("=HYPERLINK"):
        ws.cell(r, 8).font = LINK_F
    ws.cell(r, 9, i)
    ws.cell(r, 10, j)
    ws.cell(r, 11, k)
    if isinstance(value, float) and abs(value) <= 1.5 and r not in (9, 20, 21, 22, 23):
        ws.cell(r, 3).number_format = "0.0%"
    elif r == 20:
        ws.cell(r, 3).number_format = "0.0"
    elif r == 9:
        ws.cell(r, 3).number_format = "0.0"
    elif r in (21, 22, 23):
        ws.cell(r, 3).number_format = "#,##0.0"
    style_row(ws, r, fill)


def irr(cashflows, guess=0.1):
    r = guess
    for _ in range(100):
        f = sum(cf / ((1 + r) ** i) for i, cf in enumerate(cashflows))
        df = sum(-i * cf / ((1 + r) ** (i + 1)) for i, cf in enumerate(cashflows) if i > 0)
        if abs(df) < 1e-12:
            break
        r2 = r - f / df
        if abs(r2 - r) < 1e-10:
            return r2
        r = r2
    return r


def project(ai: bool):
    rd_pct = RD0 / REV0
    da_pct = DA0 / REV0
    tla0 = EBITDA0 * 2.5
    tlb0 = EBITDA0 * 1.5
    notes0 = EBITDA0 * 1.0
    new_debt = tla0 + tlb0 + notes0
    years = []
    rev = REV0
    payroll = PAY0
    commission = COMM0
    tla, tlb, notes = tla0, tlb0, notes0
    for i in range(5):
        rev = rev * (1 + REV_G)
        if ai:
            if i == 0:
                payroll = PAY0 * (1 - PAYROLL_CUT)
                commission = COMM0 * (1 - COMM_CUT)
            else:
                payroll = payroll * (1 + SM_G)
                commission = commission * (1 + SM_G)
            ai_c = AI_COST
            capex = rev * CAPEX_BASE * (1 - CAPEX_CUT)
            dnwc = (rev - (REV0 if i == 0 else years[i - 1]["rev"])) * WC_BASE * (1 - WC_IMPROVE)
        else:
            payroll = rev * (PAY0 / REV0)
            commission = rev * (COMM0 / REV0)
            ai_c = 0.0
            capex = rev * CAPEX_BASE
            dnwc = (rev - (REV0 if i == 0 else years[i - 1]["rev"])) * WC_BASE
        cogs = rev * (1 - GM)
        rd = rev * rd_pct
        ga = rev * GA_PCT
        da = rev * da_pct
        sm = payroll + commission
        ebitda = rev - cogs - sm - ai_c - rd - ga
        ebit = ebitda - da
        interest = tla * (SOFR + TLA_SP) + tlb * (SOFR + TLB_SP) + notes * (SOFR + 0.035)
        ebt = ebit - interest
        tax_amt = max(0, ebt) * TAX
        ni = ebt - tax_amt
        fcf = ni + da - capex - dnwc
        debt_beg = tla + tlb + notes
        mandatory = min(debt_beg, new_debt * MAND)
        optional = SWEEP * max(0, fcf - mandatory)
        paydown = min(debt_beg, mandatory + optional)
        rem = paydown
        for bucket in ("tla", "tlb", "notes"):
            bal = {"tla": tla, "tlb": tlb, "notes": notes}[bucket]
            pay = min(bal, rem)
            rem -= pay
            if bucket == "tla":
                tla -= pay
            elif bucket == "tlb":
                tlb -= pay
            else:
                notes -= pay
        years.append(
            dict(
                rev=rev,
                cogs=cogs,
                payroll=payroll,
                commission=commission,
                sm=sm,
                ai=ai_c,
                rd=rd,
                ga=ga,
                da=da,
                ebitda=ebitda,
                ebit=ebit,
                interest=interest,
                ni=ni,
                fcf=fcf,
                mandatory=mandatory,
                paydown=paydown,
                debt_end=tla + tlb + notes,
                ebitda_m=ebitda / rev if rev else 0,
                capex=capex,
                dnwc=dnwc,
            )
        )
    return years, new_debt, tla0 + tlb0 + notes0


def write_ops(ws, rows, ai: bool):
    mapping = {
        5: "rev",
        6: "cogs",
        7: "payroll",
        8: "commission",
        9: "sm",
        10: "ai",
        11: "rd",
        12: "ga",
        13: "ebitda",
        14: "ebitda_m",
        15: "da",
        16: "ebit",
        17: "interest",
        18: "ni",
        19: "capex",
        20: "dnwc",
        21: "fcf",
        22: "mandatory",
        23: "paydown",
        24: "debt_end",
    }
    year_cols = [4, 5, 6, 7, 8]
    for i, col in enumerate(year_cols):
        y = rows[i]
        for r, key in mapping.items():
            if key == "ai" and not ai:
                ws.cell(r, col).value = 0
            else:
                ws.cell(r, col).value = y[key]
    ws["C7"] = PAY0
    ws["C8"] = COMM0
    ws["C9"] = SM0
    if ai:
        for col in year_cols:
            ws.cell(10, col).value = AI_COST


def patch_drivers(ws):
    ws["B2"] = (
        "Assumptions Drivers — Col A Baseline Yes/No. Col E/H = source page link. "
        "Cols I/J/K = plain Verbatim Ctrl+F strings. Converse pass: driver must match "
        "sourced figure or a justified gap; else driver was reset to the source."
    )
    ws["B3"] = (
        "How to verify: click Source (E) → open page → copy Verbatim from I/J/K → Ctrl+F → paste. "
        f"Overpessimism fix: rev +5% (mature CAGR; FY24 −2% outlier). S&M +2% Fed restored. "
        f"S&M $414.1; 70/30; cuts 18%/18%; AI ${AI_COST:.1f}m; CapEx/WC cuts 0."
    )

    # Row 5 — Revenue growth: ADJUST −2% → +5% (FY24 outlier; mature multi-year CAGR)
    set_driver_row(
        ws,
        5,
        baseline="NO",
        driver="Revenue growth",
        value=REV_G,
        why=(
            "FY24 −2% was the only revenue decline in CY2019–CY2025 (outlier). FY25 recovered +3% "
            "($1,249.5m). Mature CAGR CY2022 $1,098m → CY2025 $1,249.5m ≈ 4.4%, rounded to 5% forward."
        ),
        e=hyperlink(
            ZI_FY25,
            "IN EQ: C5=+5% ≈ mature CAGR CY2022→CY2025 (~4.4%). FY24 −2% = outlier; FY25 +3% recovery.",
        ),
        f=(
            "FY24 −2% was an outlier (only decline). FY25 rose 3% to $1,249.5m. "
            "CAGR 2022–2025 ≈4.4%; we use 5% forward."
        ),
        feeds="Go to AI_Operating!D5 (Revenue)",
        h=hyperlink(ZI_URL, "FY24 −2% outlier: decrease of 2% YoY"),
        i="Ctrl+F: increase of $35.2 million, or 3%",
        j="Ctrl+F: revenue of $1,249.5 million",
        k="Ctrl+F: a decrease of 2% year-over-year (FY24 outlier)",
        fill=GREEN,
    )

    # Row 6 — GM: KEEP 80%
    set_driver_row(
        ws,
        6,
        baseline="NO",
        driver="Gross margin",
        value=GM,
        why=(
            "Uses the 2025 software median gross margin of 80%. ZoomInfo’s own GAAP GM is higher "
            "(~84%), but this model uses the cited Aleph median."
        ),
        e=hyperlink(ALEPH, "IN EQ: software median GM=80% → C6=0.80 (Aleph × Benchmarkit)"),
        f="Aleph says the 2025 software median gross margin is 80%. We use that 80% in the COGS math.",
        feeds="Go to AI_Operating!D6 (COGS)",
        h=None,
        i="Ctrl+F: The 2025 median software gross margin is 80%",
        j=None,
        k="ZI GAAP GM ~84% not used — Aleph median is the driver",
        fill=GREEN,
    )

    # Row 7 — payroll cut: KEEP 18%
    set_driver_row(
        ws,
        7,
        baseline="NO",
        driver="Sales payroll cut (Y1)",
        value=PAYROLL_CUT,
        why="Matches published US B2B SaaS net SDR headcount decline of 18% YoY.",
        e=hyperlink(DIGITAL, "IN EQ: C7=18% = US B2B SaaS net SDR headcount down 18% YoY"),
        f=(
            "Digital Applied says U.S. B2B SaaS SDR headcount fell 18% YoY. "
            "We use that 18% as the Year 1 payroll cut."
        ),
        feeds="Go to AI_Operating!D7 (Sales payroll)",
        h=None,
        i="Ctrl+F: down 18% YoY",
        j=None,
        k=None,
        fill=GREEN,
    )

    # Row 8 — commission cut: ADJUST 20% → 18%
    set_driver_row(
        ws,
        8,
        baseline="NO",
        driver="Variable commission cut",
        value=COMM_CUT,
        why=(
            "Was an unsourced 20% cut. Now 18% to match the same SDR headcount decline used for payroll."
        ),
        e=hyperlink(DIGITAL, "IN EQ: C8=18% cut = same SDR HC down 18% YoY (Digital Applied)"),
        f=(
            "Digital Applied says SDR headcount fell 18% YoY. We use that same 18% as the commission cut."
        ),
        feeds="Go to AI_Operating!D8 (Commissions)",
        h=hyperlink(SYNC_PAY, "Pay mix context only: 70/30 base/variable"),
        i="Ctrl+F: down 18% YoY",
        j="Ctrl+F: 70/30 base/variable split",
        k="Prior 20% cut removed — no source equaled 20%",
        fill=GREEN,
    )

    # Row 9 — AI cost: ADJUST 34 → 24.5 via 18% roles
    set_driver_row(
        ws,
        9,
        baseline="NO",
        driver="AI infrastructure ($m)",
        value=AI_COST,
        why=(
            f"AI ${AI_COST:.1f}m = {AI_ROLE_SHARE:.0%}×{SM_HC:,} roles × $5–10k/mo × 12, midpoint. "
            "Role share now uses the same sourced 18% SDR decline (was unsourced 25%)."
        ),
        e=hyperlink(
            WHITE,
            f"IN EQ: $5k/$10k/mo + 1,513 HC; roles={AI_ROLE_SHARE:.0%} (same 18% as C7) → ${AI_COST:.1f}m",
        ),
        f=(
            f"White Space/Miniloop say ~$5k–$10k+/mo. ZoomInfo’s 10-K says 1,513 S&M staff. "
            f"Roles use the sourced 18% (not 25%). Those build the ${AI_COST:.1f}m cost."
        ),
        feeds="Go to AI_Operating!D10 (AI infrastructure)",
        h=hyperlink(MINI, "Miniloop Enterprise $10,000+"),
        i="Ctrl+F: $5,000-10,000+",
        j="Ctrl+F: $10,000-$15,000+",
        k="Ctrl+F: 1,513 in sales and marketing",
        fill=GREEN,
    )

    # Row 10 — S&M growth: restore Fed +2% (valid with positive rev growth)
    set_driver_row(
        ws,
        10,
        baseline="NO",
        driver="S&M expense growth Y2+",
        value=SM_G,
        why=(
            "After the Y1 headcount cut, AI keeps S&M growing only at ~inflation (2%), not with revenue. "
            "Restored now that forward revenue is positive again."
        ),
        e=hyperlink(FED, "IN EQ: C10=2% = Fed longer-run inflation goal of 2 percent"),
        f="The Fed says longer-run inflation is 2%. We use that 2% to grow S&M costs after Year 1.",
        feeds="Go to AI_Operating!E7 (payroll Y2+)",
        h="FOMC Statement on Longer-Run Goals (PDF)",
        i="Ctrl+F: inflation of 2 percent over the longer run",
        j=None,
        k="Valid with +5% revenue (row 5); not used when rev was −2%",
        fill=GREEN,
    )

    # Row 11 — G&A: KEEP with justification
    set_driver_row(
        ws,
        11,
        baseline="NO",
        driver="G&A % of revenue",
        value=GA_PCT,
        why=(
            "Post-LBO target at the public SaaS median (18%). ZoomInfo’s own FY24 G&A is higher (~24%) "
            "and is not used."
        ),
        e=hyperlink(BLOSSOM, "IN EQ: C11=18% = Blossom Street 2025 public SaaS median G&A"),
        f=(
            "Blossom Street says 2025 public SaaS median G&A is 18% of sales. "
            "We use that 18% in the G&A math."
        ),
        feeds="Go to AI_Operating!D12 (G&A)",
        h=None,
        i="Ctrl+F: G&A was 18%",
        j="Ctrl+F: MEDIAN",
        k="ZI own ~24% G&A not used — median is the driver",
        fill=GREEN,
    )

    # Row 12 — CapEx cut: ADJUST 5% → 0%
    set_driver_row(
        ws,
        12,
        baseline="NO",
        driver="CapEx reduction vs base",
        value=CAPEX_CUT,
        why="No source printed a 5% CapEx cut, so the cut is set to 0% (no AI CapEx discount).",
        e="C12=0% CapEx cut — no source printed a cut, so none is applied.",
        f="No source says a CapEx cut. The cut is 0%, so CapEx uses the base rate only.",
        feeds="Go to AI_Operating!D19 (CapEx)",
        h=None,
        i="0% CapEx cut: no source claimed a cut",
        j=None,
        k=None,
        fill=GREEN,
    )

    # Row 13 — WC improve: ADJUST 10% → 0%
    set_driver_row(
        ws,
        13,
        baseline="NO",
        driver="WC / receivables improvement",
        value=WC_IMPROVE,
        why="No source printed a 10% WC improvement, so the improvement is set to 0%.",
        e="C13=0% WC improve — no source printed 10%, so none is applied.",
        f="No source says a WC improvement. The improvement is 0%, so WC uses the base plug only.",
        feeds="Go to AI_Operating!D20 (ΔNWC)",
        h=None,
        i="0% WC improve: no source claimed 10%",
        j=None,
        k=None,
        fill=GREEN,
    )

    # Rows 14–20 mostly KEEP — refresh F clarity only where needed
    set_driver_row(
        ws,
        14,
        baseline="NO",
        driver="SOFR",
        value=SOFR,
        why="Uses the sourced 2024 calendar-year lowest SOFR print of 4.30%.",
        e=hyperlink(GLOBAL_RATES, "IN EQ: C14=4.30% = SOFR 2024 calendar-year Lowest"),
        f="Global-Rates shows the lowest 2024 SOFR print was 4.30%. We use that 4.30% in the interest math.",
        feeds="Go to Debt_Sweep_AI!C12 (TLA rate uses SOFR)",
        h="NY Fed SOFR publisher",
        i="Ctrl+F: 4.30",
        j=None,
        k="Uses Lowest print (sourced), not Average",
        fill=GREEN,
    )
    set_driver_row(
        ws,
        15,
        baseline="NO",
        driver="Term Loan A spread",
        value=TLA_SP,
        why="Floor of the sourced senior Term Loan A SOFR+400–500 bps band.",
        e=hyperlink(CT, "IN EQ: C15=+400bps = floor of senior TLA SOFR+400-500 bps"),
        f=(
            "CT Acquisitions lists senior Term Loan A at SOFR + 400–500 bps. "
            "We use +400 bps from that range in the rate math."
        ),
        feeds="Go to Assumptions_Drivers!C27 (TLA rate)",
        h=None,
        i="Ctrl+F: SOFR + 400-500 bps",
        j=None,
        k=None,
        fill=GREEN,
    )
    set_driver_row(
        ws,
        16,
        baseline="NO",
        driver="Term Loan B spread",
        value=TLB_SP,
        why="Upper bound of the sourced TLB SOFR+300–500 bps band (conservative).",
        e=hyperlink(LBO_DEBT, "IN EQ: TLB SOFR+300-500 bps → C16=+500bps = upper bound"),
        f=(
            "LBO debt guides list Term Loan B at SOFR + 300–500 bps. "
            "We use +500 bps from that range in the rate math."
        ),
        feeds="Go to Assumptions_Drivers!C28 (TLB rate)",
        h=None,
        i="Ctrl+F: SOFR + 300-500 basis points",
        j=None,
        k=None,
        fill=GREEN,
    )
    set_driver_row(
        ws,
        17,
        baseline="NO",
        driver="Mandatory amortization",
        value=MAND,
        why="Standard ~1%/year TLB amortization.",
        e=hyperlink(CLEAR, "IN EQ: ~1%/yr TLB amort → C17=1%"),
        f="ClearValue says Term Loan B amortizes about 1% per year. We use that 1% in the debt paydown math.",
        feeds="Go to Debt_Sweep_AI!D5 (Mandatory 1%)",
        h=None,
        i="Ctrl+F: minimal amortization (1% annually)",
        j=None,
        k=None,
        fill=GREEN,
    )
    set_driver_row(
        ws,
        18,
        baseline="NO",
        driver="Cash sweep %",
        value=SWEEP,
        why="Uses the upper end of the sourced 50–100% excess-cash sweep range.",
        e=hyperlink(RYAN, "IN EQ: C18=100% = upper bound of ECF sweep 50-100%"),
        f=(
            "Ryan O’Connell says excess-cash sweeps are often 50–100% and models often use 100%. "
            "We use that 100% in the sweep math."
        ),
        feeds="Go to Debt_Sweep_AI!E5 (Total paydown)",
        h=None,
        i="Ctrl+F: 50-100% of annual excess free cash flow",
        j="Ctrl+F: typically modeled at 100%",
        k=None,
        fill=GREEN,
    )
    set_driver_row(
        ws,
        19,
        baseline="NO",
        driver="Tax rate",
        value=TAX,
        why="Matches the U.S. federal corporate tax rate of 21%.",
        e=hyperlink(PWC, "IN EQ: US federal CIT=21% → C19=0.21 (PwC)"),
        f="PwC says the U.S. federal corporate tax rate is a flat 21%. We use that 21% in the tax math.",
        feeds="Go to AI_Operating!D18 (Net income)",
        h=None,
        i="Ctrl+F: flat rate of 21%",
        j=None,
        k=None,
        fill=GREEN,
    )
    set_driver_row(
        ws,
        20,
        baseline="NO",
        driver="Exit EV/EBITDA multiple",
        value=EXIT_MULT,
        why="Matches Aventis private SaaS EV/EBITDA first quartile of 12.9x.",
        e=hyperlink(AVENTIS, "IN EQ: C20=12.9x = Aventis private SaaS EV/EBITDA 1st quartile"),
        f="Aventis shows private SaaS EV/EBITDA first quartile at 12.9x. We use that 12.9x as the exit multiple.",
        feeds="Go to LBO!D17 (Exit multiple)",
        h=None,
        i="Ctrl+F: 12.9x",
        j=None,
        k=None,
        fill=GREEN,
    )

    # Row 21 — S&M baseline: ADJUST 425 → 414.1
    set_driver_row(
        ws,
        21,
        baseline="YES",
        driver="S&M baseline ($m)",
        value=SM0,
        why="Was a rounded $425m (~35% of sales). Now the exact ZoomInfo FY24 S&M of $414.1m.",
        e=hyperlink(ZI_URL, "IN EQ: C21=$414.1m = ZI FY24 Sales and marketing"),
        f="ZoomInfo says FY24 S&M was $414.1m. We use that $414.1m as the S&M baseline in the math.",
        feeds="Go to AI_Operating!C9 (S&M human FY24A)",
        h=hyperlink(ZI_10K, "SEC 10-K"),
        i="Ctrl+F: 414.1",
        j="Ctrl+F: $414.1",
        k="Ctrl+F: Sales and marketing",
        fill=GREEN,
    )

    # Row 22 — Payroll: ADJUST to 70% of 414.1
    set_driver_row(
        ws,
        22,
        baseline="YES",
        driver="Payroll baseline ($m)",
        value=PAY0,
        why=f"Was 75% of $425m. Now 70% of $414.1m = ${PAY0:.1f}m to match SyncGTM’s ~70% base mix.",
        e=hyperlink(SYNC_PAY, f"IN EQ: ~70% base × $414.1m S&M → C22=${PAY0:.1f}m"),
        f=(
            f"SyncGTM says pay is about 70% base. We use 70% of $414.1m = ${PAY0:.1f}m as payroll in the math."
        ),
        feeds="Go to AI_Operating!C7 (Payroll FY24A)",
        h=None,
        i="Ctrl+F: 70/30 base/variable split",
        j=f"Ctrl+F: 414.1 (on ZI results for S&M pool)",
        k=f"${PAY0:.1f}m = 70% × $414.1m",
        fill=GREEN,
    )

    # Row 23 — Commission: ADJUST to 30% of 414.1
    set_driver_row(
        ws,
        23,
        baseline="YES",
        driver="Commission baseline ($m)",
        value=COMM0,
        why=f"Was 25% of $425m. Now 30% of $414.1m = ${COMM0:.1f}m to match SyncGTM’s ~30% variable mix.",
        e=hyperlink(SYNC_COMM, f"IN EQ: ~30% variable × $414.1m → C23=${COMM0:.1f}m"),
        f=(
            f"SyncGTM says commission is about 30–35% of pay. We use 30% of $414.1m = ${COMM0:.1f}m in the math."
        ),
        feeds="Go to AI_Operating!C8 (Commissions FY24A)",
        h=hyperlink(SYNC_PAY, "70/30 pay article"),
        i="Ctrl+F: 30–35% of total pay",
        j="Ctrl+F: 70/30 base/variable split",
        k=f"${COMM0:.1f}m = 30% × $414.1m",
        fill=GREEN,
    )

    # Row 24 — CapEx base: KEEP
    set_driver_row(
        ws,
        24,
        baseline="NO",
        driver="CapEx base % of rev",
        value=CAPEX_BASE,
        why="Midpoint of the sourced pure-play SaaS CapEx band of 1–3% of revenue.",
        e=hyperlink(SAASDB, "IN EQ band: SaaS CapEx 1–3% of rev → C24=2% midpoint"),
        f=(
            "SaaSDB says pure-play SaaS CapEx is often 1–3% of sales. "
            "We use 2%, the midpoint of that band, in the CapEx math."
        ),
        feeds="Go to AI_Operating!D19 (CapEx)",
        h=None,
        i="Ctrl+F: CapEx of 1",
        j="Tip: page uses en-dash 1–3%; or Ctrl+F: Most pure-play SaaS companies report CapEx of",
        k="2%: MODEL midpoint of 1–3% band",
        fill=GREEN,
    )

    # Row 25 — WC base: ADJUST 2% → 0%
    set_driver_row(
        ws,
        25,
        baseline="NO",
        driver="WC base % of Δrev",
        value=WC_BASE,
        why="Was an unsourced 2% ΔNWC/ΔRev plug. No page equals 2%, so the plug is set to 0%.",
        e="C25=0% ΔNWC/ΔRev — no source printed 2%, so the plug is zero.",
        f="No source says a 2% WC plug. The plug is 0%, so ΔNWC is zero in the forecast.",
        feeds="Go to AI_Operating!D20 (ΔNWC)",
        h=None,
        i="0% WC plug: no source equaled 2%",
        j=None,
        k=None,
        fill=GREEN,
    )

    # Computed rates
    ws["C27"] = "=C14+C15"
    ws["C28"] = "=C14+C16"
    ws["F27"] = (
        "No new source. This cell only adds SOFR (row 14) and the TLA spread (row 15). "
        "Both already go into the rate math."
    )
    ws["F28"] = (
        "No new source. This cell only adds SOFR (row 14) and the TLB spread (row 16). "
        "Both already go into the rate math."
    )

    # Footer
    ws["B29"] = f"AI ${AI_COST:.1f}m IN-EQ sources + Ctrl+F:"
    # Unmerge if needed then write
    try:
        ws.unmerge_cells("C29:F29")
    except Exception:
        pass
    ws["C29"] = (
        f"White Space Ctrl+F $5,000-10,000+ | Miniloop Ctrl+F $10,000-$15,000+ | "
        f"SEC Ctrl+F 1,513 in sales | roles={AI_ROLE_SHARE:.0%} of HC (same 18% as payroll cut)"
    )
    ws["B30"] = "ZI 10-K S&M HC 1,513:"
    ws["C30"] = ZI_10K


def update_debt_sweep(wb, ai_rows, base_rows):
    if "Debt_Sweep_AI" not in wb.sheetnames:
        return
    dws = wb["Debt_Sweep_AI"]
    dws["C12"] = SOFR + TLA_SP
    dws["D12"] = SOFR + TLB_SP
    # Years 2025-2029 rows 5-9
    for i, r in enumerate(range(5, 10)):
        dws.cell(r, 3).value = ai_rows[i]["fcf"]
        dws.cell(r, 4).value = ai_rows[i]["mandatory"]
        dws.cell(r, 5).value = ai_rows[i]["paydown"]
        dws.cell(r, 6).value = ai_rows[i]["debt_end"]
        dws.cell(r, 7).value = base_rows[i]["debt_end"]
        dws.cell(r, 8).value = base_rows[i]["debt_end"] - ai_rows[i]["debt_end"]


def update_strategy(wb, ai_rows, base_rows):
    price, shares, premium = 4.05, 307.294, 0.25
    offer = price * (1 + premium)
    equity_buy = offer * shares
    cash0, min_cash = 179.9, 50.0
    fees = 0.02 * equity_buy
    uses = equity_buy + 1329.9 + fees
    excess = max(0, cash0 - min_cash)
    new_debt = EBITDA0 * 5.0
    sponsor = uses - excess - new_debt

    def exit_equity(rows):
        return rows[4]["ebitda"] * EXIT_MULT - rows[4]["debt_end"]

    ai_exit_eq = exit_equity(ai_rows)
    base_exit_eq = exit_equity(base_rows)
    ai_irr = irr([-sponsor, 0, 0, 0, 0, ai_exit_eq])
    base_irr = irr([-sponsor, 0, 0, 0, 0, base_exit_eq])
    uplift = ai_irr - base_irr
    y2_ai_m = ai_rows[1]["ebitda_m"]
    y2_base_m = base_rows[1]["ebitda_m"]
    fy24_m = EBITDA0 / REV0
    bps = (y2_ai_m - fy24_m) * 10000
    ai_moic = ai_exit_eq / sponsor if sponsor else 0
    base_moic = base_exit_eq / sponsor if sponsor else 0

    if "LBO" in wb.sheetnames:
        wb["LBO"]["D17"] = EXIT_MULT

    if "Strategy_Summary" not in wb.sheetnames:
        return dict(ai_irr=ai_irr, base_irr=base_irr, bps=bps, sponsor=sponsor)

    ss = wb["Strategy_Summary"]
    ss["C11"] = f"{1024.5} / 84.4%"
    ss["C16"] = SM0
    ss["B16"] = "S&M baseline (ZI FY24 S&M $414.1m)"
    ss["C22"] = EXIT_MULT
    ss["B22"] = "Exit multiple"
    ss["D22"] = "Aventis private SaaS EV/EBITDA 1st quartile 12.9x"
    ss["B25"] = (
        f"Y1: cut payroll & commissions {PAYROLL_CUT:.0%} vs baseline ${SM0:.1f}m S&M; "
        f"add ${AI_COST:.1f}m AI software/compute cost"
    )
    ss["B26"] = "Y2–Y5: S&M grows only 2%/yr after Y1 cut (AI scales without headcount; Fed inflation)"
    ss["C31"] = y2_base_m
    ss["D31"] = y2_ai_m
    ss["E31"] = y2_ai_m - y2_base_m
    ss["C32"] = (y2_base_m - fy24_m) * 10000
    ss["D32"] = bps
    ss["E32"] = bps - (y2_base_m - fy24_m) * 10000
    ss["C33"] = base_rows[4]["ebitda"]
    ss["D33"] = ai_rows[4]["ebitda"]
    ss["E33"] = ai_rows[4]["ebitda"] - base_rows[4]["ebitda"]
    ss["C34"] = base_rows[4]["ebitda"] * EXIT_MULT
    ss["D34"] = ai_rows[4]["ebitda"] * EXIT_MULT
    ss["E34"] = ss["D34"].value - ss["C34"].value
    ss["C35"] = base_rows[4]["debt_end"]
    ss["D35"] = ai_rows[4]["debt_end"]
    ss["E35"] = ai_rows[4]["debt_end"] - base_rows[4]["debt_end"]
    ss["F35"] = (
        f"Remaining debt at exit; AI case ${ai_rows[4]['debt_end']:.1f}m vs Base ${base_rows[4]['debt_end']:.1f}m."
    )
    ss["C36"] = base_exit_eq
    ss["D36"] = ai_exit_eq
    ss["E36"] = ai_exit_eq - base_exit_eq
    ss["F36"] = (
        f"Total equity payout; AI strategy changed final payout by ${ai_exit_eq-base_exit_eq:,.1f}m "
        f"to ${ai_exit_eq:,.1f}m (exit {EXIT_MULT}x)."
    )
    ss["C37"] = base_moic
    ss["D37"] = ai_moic
    ss["E37"] = ai_moic - base_moic
    ss["F37"] = (
        f"Multiple on Invested Capital; AI improved gross cash return from {base_moic:.2f}x to {ai_moic:.2f}x."
    )
    ss["C38"] = base_irr
    ss["D38"] = ai_irr
    ss["E38"] = uplift
    for coord in ("C38", "D38", "E38"):
        ss[coord].number_format = "0.0%"
    ss["F38"] = (
        f"5-year IRR; AI boosted return by {uplift*100:.1f} percentage points to {ai_irr*100:.1f}% "
        f"(AI infra ${AI_COST:.1f}m; payroll/commission cut {PAYROLL_CUT:.0%}; rev {REV_G:.0%}; exit {EXIT_MULT}x)."
    )
    ss["B41"] = (
        f"AI case IRR {ai_irr*100:.1f}% − Base IRR {base_irr*100:.1f}% = {uplift*100:.1f} percentage points of IRR"
    )
    ss["B42"] = (
        f"Y2 AI EBITDA margin {y2_ai_m*100:.1f}% vs FY24 baseline {fy24_m*100:.1f}% = {bps:.0f} bps "
        f"(target ≥500: {'YES' if bps >= 500 else 'NO'})"
    )
    return dict(
        ai_irr=ai_irr,
        base_irr=base_irr,
        uplift=uplift,
        bps=bps,
        sponsor=sponsor,
        ai_exit_eq=ai_exit_eq,
        base_exit_eq=base_exit_eq,
        ai_cost=AI_COST,
        rev_g=REV_G,
        sm0=SM0,
        pay0=PAY0,
        comm0=COMM0,
    )


def sync_assumptions_list(wb, stats):
    if "00_Assumptions_List" not in wb.sheetnames:
        return
    ws = wb["00_Assumptions_List"]
    # Update value column for known rows
    updates = {
        7: f"{REV_G:.0%} per year (unchanged by AI)",
        8: f"{GM:.0%}",
        9: f"−{PAYROLL_CUT:.0%} in Year 1",
        10: f"−{COMM_CUT:.0%} of baseline commission pool from Y1",
        11: f"${AI_COST:.1f} million fixed / year",
        12: f"{SM_G:.0%} per year from Year 2",
        13: f"{GA_PCT:.0%} of revenue",
        14: f"−{CAPEX_CUT:.0%} vs baseline CapEx rate (no sourced cut)",
        15: f"WC improve {WC_IMPROVE:.0%} vs base (no sourced improve)",
        16: f"SOFR {SOFR:.2%} + {TLA_SP:.2%} (= {(SOFR+TLA_SP):.2%})",
        17: f"SOFR {SOFR:.2%} + {TLB_SP:.2%} (= {(SOFR+TLB_SP):.2%})",
        21: f"{EXIT_MULT}x EV / EBITDA",
        23: f"${SM0:.1f}m",
        24: f"${PAY0:.1f}m",
        25: f"${COMM0:.1f}m",
        26: f"{CAPEX_BASE:.1%}",
        27: f"{WC_BASE:.1%}",
    }
    for r, val in updates.items():
        if ws.cell(r, 4).value is not None:
            ws.cell(r, 4).value = val
    ws["B30"] = (
        f"This list must match Assumptions_Drivers: rev {REV_G:.0%}, payroll/commission cut "
        f"−{PAYROLL_CUT:.0%}, AI ${AI_COST:.1f}m, S&M ${SM0:.1f}m, exit {EXIT_MULT}x, "
        f"CapEx/WC cuts 0% (unsourced removed)."
    )


def main():
    assert abs(PAY0 + COMM0 - SM0) < 0.05, (PAY0, COMM0, SM0)
    print(f"AI_COST={AI_COST} (low={AI_LOW:.2f} high={AI_HIGH:.2f})")
    print(f"SM0={SM0} PAY0={PAY0} COMM0={COMM0} REV_G={REV_G} COMM_CUT={COMM_CUT}")

    wb = load_workbook(XLSX)
    patch_drivers(wb["Assumptions_Drivers"])

    ai_rows, new_debt, debt0 = project(True)
    base_rows, _, _ = project(False)
    write_ops(wb["AI_Operating"], ai_rows, True)
    write_ops(wb["Base_Operating"], base_rows, False)
    update_debt_sweep(wb, ai_rows, base_rows)
    stats = update_strategy(wb, ai_rows, base_rows)
    sync_assumptions_list(wb, stats)

    wb.save(XLSX)
    shutil.copy2(XLSX, XLSX2)

    print("Y1 AI rev", ai_rows[0]["rev"], "ebitda_m", ai_rows[0]["ebitda_m"])
    print("Y2 AI ebitda_m", ai_rows[1]["ebitda_m"], "bps", stats["bps"])
    print("IRR AI/Base", stats["ai_irr"], stats["base_irr"])
    print("Saved", XLSX)


if __name__ == "__main__":
    main()
