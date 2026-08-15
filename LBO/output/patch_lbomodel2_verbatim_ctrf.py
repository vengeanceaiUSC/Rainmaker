#!/usr/bin/env python3
"""
Add Verbatim (Ctrl+F) column to Assumptions_Drivers.
Every sourced figure must have a findable string + quote.
Replace sources / yellow values that are not Ctrl+F-able.
Recompute ops when C7 (→18%) or C20 (→12.9x) change.
"""

from __future__ import annotations

import math
import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL2.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL2.xlsx"
MD = REPO / "LBO" / "output" / "LBOMODEL2_DRIVER_SOURCES.md"
PDF = REPO / "LBO" / "output" / "LBOMODEL2_DRIVER_SOURCES.pdf"
NOTES = REPO / "LBO" / "output" / "LBOMODEL2_NOTES.md"

RAW_XLSX = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx"
RAW_PDF = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_DRIVER_SOURCES.pdf"

LINK = Font(name="Calibri", color="0563C1", underline="single", size=9)
BLACK = Font(name="Calibri", size=9)
HDR = PatternFill("solid", fgColor="1F4E79")
HDR_F = Font(name="Calibri", color="FFFFFF", bold=True, size=9)
GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

# Drivers after Task A alignment
REV_G = 0.10
GM = 0.80
PAYROLL_CUT = 0.18  # was 0.15 midpoint — now exact findable 18%
COMM_CUT = 0.20
AI_COST = 34.0
SM_G = 0.02
GA_PCT = 0.18
CAPEX_CUT = 0.05
WC_IMPROVE = 0.10
SOFR = 0.043
TLA_SP = 0.04
TLB_SP = 0.05
MAND = 0.01
SWEEP = 1.0
TAX = 0.21
EXIT_MULT = 12.9  # was 13 round — now exact Aventis Q1 12.9x
SM0 = 425.0
PAY0 = 318.8
COMM0 = 106.2
CAPEX_BASE = 0.02
WC_BASE = 0.02
REV0 = 1214.3
RD0 = 196.1
DA0 = 85.7
EBITDA0 = 183.1

# Row specs: E, F, H, I(verbatim), E_url, H_url, fill, C_value_optional
ROWS = {
    5: dict(
        C=REV_G,
        E="IN EQ $ base only: FY24 Rev=$1,214.3m (ZI). C5=10% = MODEL CONST (10% not in filing).",
        F="EQ: Y1_Rev=1214.3×(1+0.10)=$1,335.73m. SOURCED $: 1,214.3. MODEL CONST: 0.10 (filing prints −2%, not +10%).",
        H="SEC 10-K same $1,214.3m",
        I='Ctrl+F: $1,214.3  →  "GAAP Revenue of $1,214.3 million, a decrease of 2% year-over-year."  |  10% growth: N/A (MODEL CONST)',
        E_url="https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
        H_url="https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
        fill="yellow",
    ),
    6: dict(
        C=GM,
        E="IN EQ: software median GM=80% → C6=0.80 (Aleph × Benchmarkit)",
        F="EQ: COGS=Rev×(1−0.80). Y1: 1335.73×0.20=$267.15m. SOURCED: 80% = 2025 median software gross margin.",
        H=None,
        I='Ctrl+F: median software gross margin is 80  →  "The 2025 median software gross margin is 80%"',
        E_url="https://www.getaleph.com/answers/saas-gross-margin-2026",
        H_url=None,
        fill="green",
    ),
    7: dict(
        C=PAYROLL_CUT,
        E="IN EQ: C7=18% = US B2B SaaS net SDR headcount down 18% YoY (Digital Applied / Bridge Group)",
        F="EQ: Payroll_Y1=318.8×(1−0.18)=$261.416m. SOURCED: 18% YoY SDR HC decline. SOURCED $: 318.8 from row22. (Was 15% midpoint — replaced so yellow equals findable text.)",
        H=None,
        I='Ctrl+F: down 18%  →  "Net SDR headcount in US B2B SaaS companies is down 18% YoY in 2026 per Bridge Group"',
        E_url="https://www.digitalapplied.com/blog/ai-sdr-statistics-2026-outbound-sales-data-points",
        H_url=None,
        fill="green",
    ),
    8: dict(
        C=COMM_CUT,
        E="Pool mix sourced (~30% variable); C8=20% cut = MODEL CONST (20% cut not in source).",
        F="EQ: Commission_AI=106.2×(1−0.20)=$84.96m. SOURCED mix: 70/30 base/variable. MODEL CONST: 0.20 cut.",
        H="Also: commission 30–35% of pay",
        I='Ctrl+F: 70/30  →  "Standard SDR OTE in 2026 is $85,000 on a 70/30 base/variable split."  |  20% cut: N/A (MODEL CONST)',
        E_url="https://syncgtm.com/blog/how-to-pay-sales-development-rep",
        H_url="https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
        fill="yellow",
    ),
    9: dict(
        C=AI_COST,
        E="IN EQ: $5,000 and $10,000/mo AI SDR (White Space + Miniloop) + 1,513 S&M HC (SEC)",
        F="EQ: Roles=0.25×1,513=378; Low=378×$5,000×12=$22.68m; High=378×$10,000×12=$45.36m; C9=avg=$34.0m. MODEL CONST: 0.25 outbound share. NOT in EQ: 11x Growth $3,750.",
        H="Miniloop Entry $5,000 / Enterprise $10,000+",
        I='Ctrl+F: $5,000-10,000  →  "11x.ai (Alice) $5,000-10,000+"  |  Ctrl+F: $10,000-$15,000  →  "Enterprise $10,000-$15,000+"  |  Ctrl+F: 1,513 in sales  →  "1,513 in sales and marketing"  |  25% share: N/A (MODEL CONST)',
        E_url="https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026",
        H_url="https://www.miniloop.ai/blog/11x-pricing",
        fill="green",
    ),
    10: dict(
        C=SM_G,
        E="IN EQ: C10=2% = Fed longer-run inflation goal of 2 percent (FOMC FAQ)",
        F="EQ: Payroll_t=Payroll_{t−1}×(1.02). SOURCED: 2 percent longer-run inflation target.",
        H="FOMC Statement on Longer-Run Goals (PDF)",
        I='Ctrl+F: 2 percent  →  "the FOMC judges that inflation of 2 percent over the longer run"',
        E_url="https://www.federalreserve.gov/faqs/economy_14400.htm",
        H_url="https://www.federalreserve.gov/monetarypolicy/files/FOMC_LongerRunGoals.pdf",
        fill="green",
    ),
    11: dict(
        C=GA_PCT,
        E="IN EQ: C11=18% = Blossom Street 2025 public SaaS median G&A 18% of revenue",
        F="EQ: G&A=Rev×0.18. SOURCED: G&A was 18% in 2025 (medians). NOT in EQ: ZI filing ~24.3%.",
        H="Table MEDIAN G&A as % of Revenue = 18%",
        I='Ctrl+F: G&A was 18%  →  "G&A was 18%, 20%, 21, and 23% of revenue in 2025, 2024, 2023, and 2022."  |  Alt table Ctrl+F: MEDIAN then G&A col 18%',
        E_url="https://blossomstreetventures.medium.com/what-should-saas-spend-on-cogs-s-m-r-d-and-g-a-the-data-from-57-public-cos-8e88eabe99ca",
        H_url="https://www.blossomstreetventures.com/saas-metric/2025-annual-percent-of-revenue",
        fill="green",
    ),
    12: dict(
        C=CAPEX_CUT,
        E="C12=5% CapEx cut = MODEL CONST (no source prints 5%).",
        F="EQ: CapEx_AI=Rev×0.02×(1−0.05)=Rev×0.019. MODEL CONST: 0.05.",
        H=None,
        I="5% CapEx cut: N/A — MODEL CONST (no Ctrl+F string; no source link).",
        E_url=None,
        H_url=None,
        fill="yellow",
    ),
    13: dict(
        C=WC_IMPROVE,
        E="C13=10% WC improvement = MODEL CONST (no source equals 10%).",
        F="EQ: ΔNWC_AI=ΔRev×0.02×0.90. MODEL CONST: 0.10. Removed prior ~80 DSO link (did not equal 10%).",
        H=None,
        I="10% WC improvement: N/A — MODEL CONST (no Ctrl+F string; no source link).",
        E_url=None,
        H_url=None,
        fill="yellow",
    ),
    14: dict(
        C=SOFR,
        E="IN EQ: C14=4.30% = SOFR 2024 calendar-year Lowest (Global-Rates summary of NY Fed)",
        F="EQ: rates use C14. SOURCED: Lowest SOFR in 2024 = 4.30%. NOT used: latest FRED ~3.62%.",
        H="NY Fed SOFR publisher",
        I='Ctrl+F: 4.30  →  table row "Lowest … 4.30 %" (SOFR 2024 First/Last/Highest/Lowest/Average)',
        E_url="https://www.global-rates.com/en/interest-rates/sofr/historical/2024/",
        H_url="https://www.newyorkfed.org/markets/reference-rates/sofr",
        fill="green",
    ),
    15: dict(
        C=TLA_SP,
        E="IN EQ: C15=+400bps = floor of senior Term Loan A SOFR+400-500 bps (CT Acquisitions)",
        F="EQ: TLA_rate=4.30%+4.00%=8.30%. SOURCED: SOFR + 400-500 bps on Senior term loan A.",
        H="Also: senior debt SOFR plus 400 to 500",
        I='Ctrl+F: 400-500  →  "Senior term loan A | … | SOFR + 400-500 bps"',
        E_url="https://ctacquisitions.com/acquisition-financing/",
        H_url="https://ctacquisitions.com/debt-financing/",
        fill="green",
    ),
    16: dict(
        C=TLB_SP,
        E="IN EQ: TLB SOFR+300-500 bps → C16=+500bps = upper bound",
        F="EQ: TLB_rate=4.30%+5.00%=9.30%. SOURCED: SOFR + 300-500 basis points for TLB.",
        H="Table: Term Loan B | SOFR + 300-500 bps | 1% annual",
        I='Ctrl+F: 300-500  →  "TLB pricing is typically SOFR + 300-500 basis points"  |  table Ctrl+F: Term Loan B',
        E_url="https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine",
        H_url="https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine",
        fill="green",
    ),
    17: dict(
        C=MAND,
        E="IN EQ: ~1%/yr TLB amort → C17=1%",
        F="EQ: Mandatory≈915.5×0.01=$9.155m/yr. SOURCED: 1% annual amortization.",
        H='ClearValue: "minimal amortization (1% annually)"',
        I='Ctrl+F: 1% annual  →  "minimal amortization (1% annually)"  |  Alt: "1% per year (nominal)"',
        E_url="https://clearvaluelending.com/glossary/term-loan-b",
        H_url="https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending",
        fill="green",
    ),
    18: dict(
        C=SWEEP,
        E="IN EQ: C18=100% = upper bound of ECF sweep 50-100%; models often use 100%",
        F="EQ: Sweep=1.00×max(0,FCF−Mandatory). SOURCED: typically 50-100%; typically modeled at 100%.",
        H=None,
        I='Ctrl+F: 50-100%  →  "apply 50-100% of annual excess free cash flow"  |  Ctrl+F: modeled at 100%  →  "typically modeled at 100%"',
        E_url="https://ryanoconnellfinance.com/lbo-model-fundamentals/",
        H_url=None,
        fill="green",
    ),
    19: dict(
        C=TAX,
        E="IN EQ: US federal CIT=21% → C19=0.21 (PwC)",
        F="EQ: Tax=EBT×0.21. SOURCED: flat rate of 21%.",
        H="Trading Economics: 21 percent",
        I='Ctrl+F: flat rate of 21  →  "taxes resident corporations at a flat rate of 21%."',
        E_url="https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income",
        H_url="https://tradingeconomics.com/united-states/corporate-tax-rate",
        fill="green",
    ),
    20: dict(
        C=EXIT_MULT,
        E="IN EQ: C20=12.9x = Aventis private SaaS EV/EBITDA 1st quartile 12.9x (was 13 round)",
        F="EQ: Exit_EV=Y5_EBITDA×12.9. SOURCED: table 1st quartile EV/EBITDA = 12.9x. Text also ~12.8x.",
        H=None,
        I='Ctrl+F: 12.9x  →  table "EV/EBITDA | … | 1st quartile | 12.9x"  |  Alt Ctrl+F: 12.8x  →  "lower quartile are closer to 12.8x"',
        E_url="https://aventis-advisors.com/saas-valuation-multiples-how-much-is-my-saas-business-worth/",
        H_url=None,
        fill="green",
    ),
    21: dict(
        C=SM0,
        E="IN EQ: Rev $1,214.3m; S&M $414.1m (34.1%) → model 35%→$425m",
        F="EQ: 0.35×1214.3=$425.0m=C21. SOURCED: 1,214.3 and 414.1. MODEL ROUND: 35% near 34.1%.",
        H="SEC 10-K Sales and marketing $414.1",
        I='Ctrl+F: $1,214.3  and  Ctrl+F: 414.1  →  Revenue $1,214.3 / Sales and marketing $414.1  |  425: MODEL ROUND of 35%×1214.3',
        E_url="https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
        H_url="https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
        fill="green",
    ),
    22: dict(
        C=PAY0,
        E="IN EQ: ~70% base OTE → model 75%×$425m=$318.8m",
        F="EQ: C22=0.75×425=$318.8m. SOURCED mix: 70/30. MODEL: 75% of S&M pool.",
        H=None,
        I='Ctrl+F: 70/30  →  "70/30 base/variable split"  |  318.8 / 75%: MODEL carve (not printed as $318.8)',
        E_url="https://syncgtm.com/blog/how-to-pay-sales-development-rep",
        H_url=None,
        fill="yellow",
    ),
    23: dict(
        C=COMM0,
        E="IN EQ: residual 25%×$425m=$106.2m (near sourced ~30–35% variable)",
        F="EQ: C23=0.25×425=$106.2m. SOURCED: 30–35% variable. MODEL: 25% of S&M as commission carve.",
        H="70/30 pay article",
        I='Ctrl+F: 30–35%  →  "Commission accounts for 30–35% of total pay"  |  106.2 / 25%: MODEL carve',
        E_url="https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
        H_url="https://syncgtm.com/blog/how-to-pay-sales-development-rep",
        fill="yellow",
    ),
    24: dict(
        C=CAPEX_BASE,
        E="IN EQ band: SaaS CapEx 1–3% of rev (SaaSDB) → C24=2% = midpoint of that band",
        F="EQ: CapEx=Rev×0.02. SOURCED band: CapEx of 1–3% of revenue. MODEL midpoint: 2%. NOTE: page uses en-dash 1–3% (Ctrl+F '1-3%' with hyphen may miss).",
        H=None,
        I='Ctrl+F: CapEx of 1  →  "Most pure-play SaaS companies report CapEx of 1–3% of revenue"  |  Tip: search CapEx of 1 (en-dash); plain 1-3% may miss.  |  2%: MODEL midpoint of that band',
        E_url="https://saasdb.app/learn/financials/fcf-margin/",
        H_url=None,
        fill="green",
    ),
    25: dict(
        C=WC_BASE,
        E="C25=2% ΔNWC/ΔRev = MODEL CONST (no source equals 2%).",
        F="EQ: ΔNWC=ΔRev×0.02. MODEL CONST: 0.02.",
        H=None,
        I="2% WC plug: N/A — MODEL CONST (no Ctrl+F string; no source link).",
        E_url=None,
        H_url=None,
        fill="yellow",
    ),
    27: dict(
        E="EQ only: C27=C14+C15 (see rows 14–15 Verbatim)",
        F="EQ: 4.30%+4.00%=8.30%. Component Ctrl+F strings on rows 14 and 15.",
        H=None,
        I="No separate third-party quote — sum of sourced C14 and C15.",
        E_url=None,
        H_url=None,
        fill="green",
    ),
    28: dict(
        E="EQ only: C28=C14+C16 (see rows 14 & 16 Verbatim)",
        F="EQ: 4.30%+5.00%=9.30%. Component Ctrl+F strings on rows 14 and 16.",
        H=None,
        I="No separate third-party quote — sum of sourced C14 and C16.",
        E_url=None,
        H_url=None,
        fill="green",
    ),
}


def set_link(cell, url, text):
    cell.value = text
    cell.alignment = WRAP
    if url:
        cell.font = LINK
        cell.hyperlink = url
    else:
        cell.font = BLACK
        if cell.hyperlink:
            cell.hyperlink = None


def npv(rate, cashflows):
    return sum(cf / ((1 + rate) ** (i + 1)) for i, cf in enumerate(cashflows))


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
        prev = rev
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
        tla_pay = min(tla, rem)
        rem -= tla_pay
        tla -= tla_pay
        tlb_pay = min(tlb, rem)
        rem -= tlb_pay
        tlb -= tlb_pay
        notes_pay = min(notes, rem)
        notes -= notes_pay
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
                ebitda_m=ebitda / rev,
                capex=capex,
                dnwc=dnwc,
            )
        )
    return years, new_debt, tla0 + tlb0 + notes0


def write_ops(ws, rows, ai: bool):
    # FY24A baselines in col C already mostly set; update Y1-Y5 D-H
    # Row map from earlier dump
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
    cols = [4, 5, 6, 7, 8]  # D..H = Y1..Y5
    # Also Y5 might be col H only through Y4? Dump showed D..G for Y1-Y4 and maybe H for Y5
    # From dump: C=FY24, D=Y1, E=Y2, F=Y3, G=Y4 — only 4 years in D-G?
    # Check: G5 = 1777... that's Y4. Need H for Y5.
    max_col = ws.max_column
    # Use D-H if H exists with headers
    year_cols = []
    for col in range(4, 10):
        hdr = ws.cell(4, col).value
        if hdr and ("Y" in str(hdr) or "202" in str(hdr)):
            year_cols.append(col)
    if len(year_cols) < 5:
        # force D-H
        year_cols = [4, 5, 6, 7, 8]
    for i, col in enumerate(year_cols[:5]):
        y = rows[i]
        for r, key in mapping.items():
            if key == "ai" and not ai:
                ws.cell(r, col).value = 0
            else:
                ws.cell(r, col).value = y[key]
    # FY24 payroll/commission/sm stay
    ws["C7"] = PAY0
    ws["C8"] = COMM0
    ws["C9"] = SM0
    if ai:
        for col in year_cols[:5]:
            ws.cell(10, col).value = AI_COST


def patch_drivers(ws):
    # Insert / set header for Verbatim in column I
    ws["I4"] = "Verbatim — Ctrl+F this string on the linked page"
    ws["I4"].fill = HDR
    ws["I4"].font = HDR_F
    ws["I4"].alignment = WRAP
    ws.column_dimensions["I"].width = 55
    ws.column_dimensions["E"].width = 42
    ws.column_dimensions["F"].width = 48

    ws["B2"] = (
        "Assumptions Drivers — every sourced figure has a Ctrl+F string in column I. "
        "If yellow ≠ findable text, it is MODEL CONST or a labeled midpoint."
    )
    ws["B3"] = (
        "Task A: CapEx Ctrl+F CapEx of 1 (en-dash 1–3%). Payroll cut now 18% (was 15%). "
        "Exit multiple now 12.9x (was 13). MODEL CONST rows have no fake links."
    )

    for r, cfg in ROWS.items():
        if "C" in cfg and cfg["C"] is not None and r not in (27, 28):
            ws[f"C{r}"] = cfg["C"]
            if r in (5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 24, 25):
                ws[f"C{r}"].number_format = "0.0%" if cfg["C"] <= 1 else "0.0"
            if r == 20:
                ws[f"C{r}"].number_format = "0.0"
            if r == 9:
                ws[f"C{r}"].number_format = "0.0"
            if r in (21, 22, 23):
                ws[f"C{r}"].number_format = "#,##0.0"
            ws[f"C{r}"].fill = YELLOW if cfg["fill"] == "yellow" else GREEN

        fill = GREEN if cfg["fill"] == "green" else YELLOW
        set_link(ws[f"E{r}"], cfg.get("E_url"), cfg["E"])
        ws[f"E{r}"].fill = fill
        ws[f"F{r}"] = cfg["F"]
        ws[f"F{r}"].font = BLACK
        ws[f"F{r}"].alignment = WRAP
        if cfg.get("H"):
            set_link(ws[f"H{r}"], cfg.get("H_url"), cfg["H"])
        else:
            ws[f"H{r}"] = None
            if ws[f"H{r}"].hyperlink:
                ws[f"H{r}"].hyperlink = None
        ws[f"I{r}"] = cfg["I"]
        ws[f"I{r}"].font = BLACK
        ws[f"I{r}"].alignment = WRAP
        ws[f"I{r}"].fill = fill
        for col in ("E", "F", "H", "I"):
            ws[f"{col}{r}"].border = THIN

    # Footer
    ws["B29"] = "AI $34m IN-EQ sources + Ctrl+F:"
    ws["C29"] = (
        "White Space Ctrl+F $5,000-10,000 | Miniloop Ctrl+F $10,000-$15,000 | "
        "SEC Ctrl+F 1,513 in sales | MODEL CONST 25% outbound | NOT in EQ: 11x $3,750"
    )
    ws["B30"] = "ZI 10-K S&M HC 1,513:"
    set_link(
        ws["C30"],
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
    )
    ws["B31"] = "11x Growth $3,750 NOT in EQ:"
    set_link(ws["C31"], "https://www.11x.ai/products/alice/pricing", "https://www.11x.ai/products/alice/pricing")
    ws["B32"] = "PDF map:"
    set_link(ws["C32"], RAW_PDF, RAW_PDF)


def update_lbo_and_strategy(wb, ai_rows, base_rows, debt0):
    # Sponsor equity implied by prior MOIC/exit (≈$1,871.3m)
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
    # If debt fully swept, debt_end may be ~0
    ai_irr = irr([-sponsor, 0, 0, 0, 0, ai_exit_eq])
    base_irr = irr([-sponsor, 0, 0, 0, 0, base_exit_eq])
    uplift = ai_irr - base_irr
    y2_ai_m = ai_rows[1]["ebitda_m"]
    y2_base_m = base_rows[1]["ebitda_m"]
    fy24_m = EBITDA0 / REV0  # ~15.1% — Strategy_Summary bps vs FY24 baseline
    bps = (y2_ai_m - fy24_m) * 10000
    ai_moic = ai_exit_eq / sponsor
    base_moic = base_exit_eq / sponsor

    if "LBO" in wb.sheetnames:
        wb["LBO"]["D17"] = EXIT_MULT

    if "Strategy_Summary" in wb.sheetnames:
        ss = wb["Strategy_Summary"]
        # C=Base, D=AI, E=delta (rows 35-38)
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
            f"(AI infra $34m; payroll cut 18%; exit {EXIT_MULT}x)."
        )
        ss["B41"] = (
            f"AI case IRR {ai_irr*100:.1f}% − Base IRR {base_irr*100:.1f}% = {uplift*100:.1f} percentage points of IRR"
        )
        ss["B42"] = (
            f"Y2 AI EBITDA margin {y2_ai_m*100:.1f}% vs FY24 baseline {fy24_m*100:.1f}% = {bps:.0f} bps "
            f"(target ≥500: {'YES' if bps >= 500 else 'NO'})"
        )
        for r in range(1, 80):
            for c in range(2, 8):
                v = ss.cell(r, c).value
                if isinstance(v, str):
                    nv = (
                        v.replace("13x", f"{EXIT_MULT}x")
                        .replace("13×", f"{EXIT_MULT}×")
                        .replace("−15%", "−18%")
                        .replace("-15%", "-18%")
                    )
                    if nv != v:
                        ss.cell(r, c).value = nv

    return dict(
        sponsor=sponsor,
        ai_irr=ai_irr,
        base_irr=base_irr,
        uplift=uplift,
        bps=bps,
        ai_exit_eq=ai_exit_eq,
        base_exit_eq=base_exit_eq,
        y2_ai_m=y2_ai_m,
        y2_base_m=y2_base_m,
        ai_y5_ebitda=ai_rows[4]["ebitda"],
        ai_y5_debt=ai_rows[4]["debt_end"],
        ai_y1_payroll=ai_rows[0]["payroll"],
        ai_moic=ai_moic,
        base_moic=base_moic,
    )


def write_md(stats):
    lines = [
        "# Assumptions_Drivers — Verbatim / Ctrl+F map",
        "",
        f"**Excel:** {RAW_XLSX}",
        f"**PDF:** {RAW_PDF}",
        "",
        "## Value changes for Task A (so yellow = findable text)",
        f"- Payroll cut **15% → 18%** (Ctrl+F `down 18%` on Digital Applied)",
        f"- Exit multiple **13 → 12.9x** (Ctrl+F `12.9x` on Aventis)",
        f"- Recomputed: AI IRR **{stats['ai_irr']*100:.1f}%**, Base **{stats['base_irr']*100:.1f}%**, Y2 margin uplift **{stats['bps']:.0f} bps**",
        "",
        "## CapEx 1–3% (C24) — why Ctrl+F may fail",
        "Page uses an **en-dash**: `1–3%`. Searching hyphen `1-3%` can miss.",
        "Use Ctrl+F: **`CapEx of 1`** → `Most pure-play SaaS companies report CapEx of 1–3% of revenue`",
        "Yellow **2%** = MODEL midpoint of that band.",
        "",
        "## Per-row Ctrl+F (column I in Excel)",
    ]
    for r, cfg in sorted(ROWS.items()):
        lines.append(f"### Row {r}")
        lines.append(f"- Source: {cfg.get('E_url') or '(none — MODEL CONST)'}")
        lines.append(f"- {cfg['I']}")
        lines.append("")
    MD.write_text("\n".join(lines), encoding="utf-8")


def write_pdf(stats):
    styles = getSampleStyleSheet()
    title = ParagraphStyle("t", parent=styles["Heading1"], fontSize=13)
    body = ParagraphStyle("b", parent=styles["Normal"], fontSize=8, leading=10)
    story = [
        Paragraph("LBOMODEL2 Assumptions — Verbatim / Ctrl+F", title),
        Paragraph(
            f"Payroll cut now 18%; exit 12.9x. AI IRR {stats['ai_irr']*100:.1f}% / Base {stats['base_irr']*100:.1f}%. "
            f"CapEx: Ctrl+F CapEx of 1 (en-dash 1–3%).",
            body,
        ),
        Spacer(1, 8),
    ]
    data = [[Paragraph("<b>Row</b>", body), Paragraph("<b>Ctrl+F / Verbatim</b>", body)]]
    for r in [5, 6, 7, 9, 10, 11, 14, 15, 16, 17, 18, 19, 20, 21, 24]:
        data.append([Paragraph(str(r), body), Paragraph(ROWS[r]["I"].replace("|", "<br/>"), body)])
    t = Table(data, colWidths=[0.6 * inch, 6.5 * inch])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(t)
    doc = SimpleDocTemplate(str(PDF), pagesize=letter, leftMargin=0.5 * inch, rightMargin=0.5 * inch)
    doc.build(story)


def main():
    wb = load_workbook(XLSX)
    ws = wb["Assumptions_Drivers"]
    patch_drivers(ws)

    ai_rows, new_debt, debt0 = project(True)
    base_rows, _, _ = project(False)

    write_ops(wb["AI_Operating"], ai_rows, True)
    write_ops(wb["Base_Operating"], base_rows, False)

    # Debt sweep sheets — update interest/paydown if present similarly
    for name, rows in [("Debt_Sweep_AI", ai_rows), ("Debt_Sweep_Base", base_rows)]:
        if name not in wb.sheetnames:
            continue
        dws = wb[name]
        # best-effort: update rates
        for r in range(1, 30):
            for c in range(1, 10):
                v = dws.cell(r, c).value
                if isinstance(v, str) and "SOFR" in v:
                    pass
        # set TLA/TLB rate cells if labeled
        dws["C12"] = SOFR + TLA_SP if dws["C12"].value is not None or True else SOFR + TLA_SP

    stats = update_lbo_and_strategy(wb, ai_rows, base_rows, debt0)

    # Update AI_Operating Y1 payroll note via values already written
    print("Y1 AI payroll", ai_rows[0]["payroll"], "was ~270.98 at 15%")
    print("Exit mult", EXIT_MULT, "AI exit equity", stats["ai_exit_eq"])
    print("IRR AI/Base", stats["ai_irr"], stats["base_irr"], "bps", stats["bps"])

    wb.save(XLSX)
    shutil.copy2(XLSX, XLSX2)
    write_md(stats)
    write_pdf(stats)

    note = (
        "\n\n## Verbatim / Ctrl+F pass\n"
        "- Added column **I Verbatim** with Ctrl+F strings for every sourced figure.\n"
        "- CapEx: Ctrl+F `CapEx of 1` (en-dash 1–3%); hyphen search can miss.\n"
        f"- Payroll cut **18%** (was 15%); exit **{EXIT_MULT}x** (was 13).\n"
        f"- AI IRR ~{stats['ai_irr']*100:.1f}%, Base ~{stats['base_irr']*100:.1f}%, Y2 bps ~{stats['bps']:.0f}.\n"
    )
    if NOTES.exists():
        t = NOTES.read_text(encoding="utf-8")
        if "## Verbatim / Ctrl+F pass" in t:
            start = t.index("## Verbatim / Ctrl+F pass")
            t = t[:start] + note.strip() + "\n"
        else:
            t = t.rstrip() + note
        NOTES.write_text(t, encoding="utf-8")
    print("Done")


if __name__ == "__main__":
    main()
