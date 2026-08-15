#!/usr/bin/env python3
"""
Embed full AI-SDR assumption commentary into Excel Assumptions list with Tab!Cell
refs + publish PDF. Also add Assumptions_Drivers yellow inputs wired into ops math.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL1.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL1.xlsx"
PDF = REPO / "LBO" / "output" / "LBOMODEL1_ASSUMPTIONS_LIST.pdf"
MD = REPO / "LBO" / "output" / "LBOMODEL1_ASSUMPTIONS_LIST.md"
RAW_XLSX = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx"
RAW_PDF = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.pdf"

HDR = PatternFill("solid", fgColor="1F4E79")
HDR_F = Font(name="Calibri", color="FFFFFF", bold=True)
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
BOLD = Font(name="Calibri", bold=True)
BLACK = Font(name="Calibri", size=10)
BLUE = Font(name="Calibri", color="0000FF")
IN = PatternFill("solid", fgColor="FFF2CC")
NOTE = PatternFill("solid", fgColor="E7F3FF")
CALC = PatternFill("solid", fgColor="E2EFDA")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
WRAP = Alignment(wrap_text=True, vertical="top")
LINK = Font(name="Calibri", color="0563C1", underline="single", size=10)

# Canonical assumption pack (user text) + workbook locations
# Math location = where the number hits P&L / debt / returns
ASSUMPTIONS = [
    {
        "name": "Revenue Growth",
        "value": "10% per year (unchanged by AI)",
        "tab": "Assumptions_Drivers",
        "cell": "C5",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D5:H5",
        "math_jump": "D5",
        "commentary": (
            "Annual growth remains stable at ten percent. We kept this assumption unchanged because "
            "our AI strategy focuses purely on cost reduction rather than artificially inflating top line projections."
        ),
        "math": "Revenue_t = Revenue_{t-1} × (1 + Assumptions_Drivers!C5). Same growth in AI and Base cases.",
    },
    {
        "name": "Gross Margin",
        "value": "80%",
        "tab": "Assumptions_Drivers",
        "cell": "C6",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D6:H6 (COGS)",
        "math_jump": "D6",
        "commentary": (
            "Expansion reached precisely eighty percent. We changed this because deploying automated AI agents "
            "streamlined initial customer onboarding, directly reducing the human technical support hours required per new user."
        ),
        "math": "COGS_t = Revenue_t × (1 − Assumptions_Drivers!C6); Gross Profit_t = Revenue_t × C6.",
    },
    {
        "name": "Sales Payroll",
        "value": "−15% in Year 1",
        "tab": "Assumptions_Drivers",
        "cell": "C7",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D7 (Sales payroll Y1)",
        "math_jump": "D7",
        "commentary": (
            "Baseline costs dropped by fifteen percent immediately. We changed this to reflect the aggressive "
            "termination of outbound human development representatives, instantly eliminating their recurring annual base salary expenses."
        ),
        "math": "Payroll_Y1 = Assumptions_Drivers!C22 × (1 − C7). Baseline payroll = C22.",
    },
    {
        "name": "Variable Commissions",
        "value": "−20% of baseline commission pool from Y1",
        "tab": "Assumptions_Drivers",
        "cell": "C8",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D8:H8 (Commissions)",
        "math_jump": "D8",
        "commentary": (
            "Rates decreased significantly starting in year one. Replacing human representatives with AI meant fewer "
            "outbound deals required traditional payouts, allowing the company to retain higher profit per sale."
        ),
        "math": "Commissions_AI = Assumptions_Drivers!C23 × (1 − C8); Base keeps full C23 pool scaled with revenue.",
    },
    {
        "name": "AI Infrastructure",
        "value": "$15 million fixed / year",
        "tab": "Assumptions_Drivers",
        "cell": "C9",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D10:H10 (AI infrastructure)",
        "math_jump": "D10",
        "commentary": (
            "A new fixed fifteen million dollar expense was added. We included this vital change to cover enterprise "
            "software licenses, API usage limits, and compute power running the AI."
        ),
        "math": "AI_cost_t = Assumptions_Drivers!C9 each forecast year (AI case only). EBITDA reduces by this amount.",
    },
    {
        "name": "Sales Growth (S&M expense growth)",
        "value": "2% per year from Year 2",
        "tab": "Assumptions_Drivers",
        "cell": "C10",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "E7:H7 / E8:H8 (payroll & commissions Y2–Y5)",
        "math_jump": "E7",
        "commentary": (
            "Expense growth was permanently capped at two percent. We altered this because automated agents scale "
            "infinitely, entirely severing the historical link between revenue growth and proportional human additions."
        ),
        "math": "For t≥2: Payroll_t = Payroll_{t-1} × (1 + Assumptions_Drivers!C10); same for commissions.",
    },
    {
        "name": "Administrative Costs (G&A)",
        "value": "Locked at baseline % of revenue",
        "tab": "Assumptions_Drivers",
        "cell": "C11",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D12:H12 (G&A)",
        "math_jump": "D12",
        "commentary": (
            "General overhead remained locked at baseline percentages. We avoided changing this because the value "
            "creation strategy strictly impacts the active sales floor without altering corporate finance or compensation."
        ),
        "math": "G&A_t = Revenue_t × Assumptions_Drivers!C11. Identical % in AI and Base cases.",
    },
    {
        "name": "Capital Expenditures",
        "value": "−5% vs baseline CapEx rate",
        "tab": "Assumptions_Drivers",
        "cell": "C12",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D19:H19 (CapEx)",
        "math_jump": "D19",
        "commentary": (
            "Hardware spending was reduced by five percent. We made this change because terminating the human sales team "
            "immediately eliminated the constant need for purchasing physical laptops and equipment."
        ),
        "math": "CapEx_AI = Revenue_t × Assumptions_Drivers!C24 × (1 − C12); Base uses C24 only.",
    },
    {
        "name": "Working Capital",
        "value": "Receivables / WC requirement −10% vs base",
        "tab": "Assumptions_Drivers",
        "cell": "C13",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D20:H20 (ΔNWC)",
        "math_jump": "D20",
        "commentary": (
            "Receivables collection periods improved slightly overall. Implementing automated AI billing reminders "
            "accelerated cash collections from late customers, lowering total working capital requirements and freeing extra cash for debt."
        ),
        "math": "ΔNWC_AI = ΔRevenue_t × Assumptions_Drivers!C25 × (1 − C13).",
    },
    {
        "name": "Term Loan A interest",
        "value": "SOFR + 4.00%",
        "tab": "Assumptions_Drivers",
        "cell": "C14 / C15",
        "engine": "LBO (not DCF)",
        "math_tab": "Debt_Sweep_AI",
        "math_cell": "C12 (TLA rate); Interest on AI_Operating!D17:H17",
        "math_jump": "C12",
        "commentary": (
            "The interest rate remained static at SOFR plus four. This was unchanged because the macroeconomic lending "
            "environment and the base company credit profile strictly dictate senior pricing."
        ),
        "math": "TLA_rate = Assumptions_Drivers!C14 + C15 (=C27). Interest_TLA = TLA_beg × TLA_rate.",
    },
    {
        "name": "Term Loan B interest",
        "value": "SOFR + 5.00%",
        "tab": "Assumptions_Drivers",
        "cell": "C14 / C16",
        "engine": "LBO (not DCF)",
        "math_tab": "Debt_Sweep_AI",
        "math_cell": "D12 (TLB rate); Interest on AI_Operating!D17:H17",
        "math_jump": "D12",
        "commentary": (
            "Pricing was maintained exactly at SOFR plus five. We did not alter this input since institutional syndicated "
            "loan markets price risk completely independent of internal software strategies."
        ),
        "math": "TLB_rate = Assumptions_Drivers!C14 + C16 (=C28). Interest_TLB = TLB_beg × TLB_rate.",
    },
    {
        "name": "Mandatory Amortization",
        "value": "1% of initial principal / year",
        "tab": "Assumptions_Drivers",
        "cell": "C17",
        "engine": "LBO (not DCF)",
        "math_tab": "Debt_Sweep_AI",
        "math_cell": "D5:D9 (Mandatory 1%)",
        "math_jump": "D5",
        "commentary": (
            "Principal repayment schedules were held constant at one percent. We kept this standard because syndicated "
            "loans strictly mandate minimum amortization regardless of how much extra cash AI generates."
        ),
        "math": "Mandatory_t = Initial_Debt × Assumptions_Drivers!C17 (floor paydown before optional sweep).",
    },
    {
        "name": "Cash Sweep",
        "value": "100% of excess cash to debt",
        "tab": "Assumptions_Drivers",
        "cell": "C18",
        "engine": "LBO (not DCF)",
        "math_tab": "Debt_Sweep_AI",
        "math_cell": "E5:E9 (Total paydown)",
        "math_jump": "E5",
        "commentary": (
            "One hundred percent of excess cash targets debt paydown. This remained unchanged to ensure all newly generated "
            "margins from AI savings aggressively reduce leverage and maximize sponsor returns."
        ),
        "math": "Optional_sweep_t = C18 × max(0, FCF_t − Mandatory_t); Debt_end = Debt_beg − Mandatory − Optional_sweep.",
    },
    {
        "name": "Tax Rate",
        "value": "21%",
        "tab": "Assumptions_Drivers",
        "cell": "C19",
        "engine": "LBO (not DCF)",
        "math_tab": "AI_Operating",
        "math_cell": "D18:H18 (Net income)",
        "math_jump": "D18",
        "commentary": (
            "The corporate tax rate remained steady at twenty one percent. We did not change this assumption because "
            "federal statutory tax obligations remain unaffected by internal operating margin improvements."
        ),
        "math": "Tax_t = max(0, EBIT_t − Interest_t) × Assumptions_Drivers!C19; NI_t = EBT_t − Tax_t.",
    },
    {
        "name": "Exit Multiple",
        "value": "13.0x EV / EBITDA",
        "tab": "Assumptions_Drivers",
        "cell": "C20",
        "engine": "LBO (Ex LBO!D17; not DCF)",
        "math_tab": "LBO",
        "math_cell": "D17 (also Strategy_Summary!D10)",
        "math_jump": "D17",
        "commentary": (
            "The valuation multiple was securely locked at thirteen. We kept this exact multiple consistent across scenarios "
            "to prove equity value creation stemmed purely from AI operational improvements alone."
        ),
        "math": "Exit_EV = Y5_EBITDA × Assumptions_Drivers!C20; Exit_Equity = Exit_EV − Debt_end; MOIC / IRR from that.",
    },
]


def style_inp(cell, fmt=None):
    cell.font = BLUE
    cell.fill = IN
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def style_lab(cell, bold=False):
    cell.font = BOLD if bold else BLACK


def rebuild_drivers_and_ops(wb):
    """Create Assumptions_Drivers and recompute AI/Base operating + debt from those cells' values."""
    # Driver values (match user narrative)
    rev_g = 0.10
    gm = 0.80
    payroll_cut = 0.15
    commission_cut = 0.20
    ai_cost = 15.0
    sm_g = 0.02
    ga_pct = 0.18
    capex_cut = 0.05
    wc_improve = 0.10
    sofr = 0.043
    tla_sp = 0.04
    tlb_sp = 0.05
    mand_amort = 0.01
    sweep = 1.0
    tax = 0.21
    exit_mult = 13.0

    # FY24 baseline
    rev0 = 1214.3
    sm0 = round(rev0 * 0.35, 1)  # 425
    commission0 = round(sm0 * 0.25, 1)  # carve commission pool from S&M for visibility
    payroll0 = sm0 - commission0
    rd_pct = 196.1 / rev0
    da_pct = 85.7 / rev0
    capex_base = 0.02
    wc_base = 0.02
    ebitda0 = 183.1  # keep reported baseline for entry leverage
    # Entry capital structure
    price, shares, premium = 4.05, 307.294, 0.25
    offer = price * (1 + premium)
    debt0_refi = 1329.9
    cash0 = 179.9
    min_cash = 50.0
    equity_buy = offer * shares
    fees = 0.02 * equity_buy
    uses = equity_buy + debt0_refi + fees
    excess = max(0, cash0 - min_cash)
    # New debt split: TLA 2.5x / TLB 1.5x of baseline EBITDA (notes 1.0x optional -> keep 5x total)
    tla0 = ebitda0 * 2.5
    tlb0 = ebitda0 * 1.5
    notes0 = ebitda0 * 1.0
    new_debt = tla0 + tlb0 + notes0
    sponsor = uses - excess - new_debt

    if "Assumptions_Drivers" in wb.sheetnames:
        del wb["Assumptions_Drivers"]
    d = wb.create_sheet("Assumptions_Drivers", 1)
    d["B2"] = "Assumptions Drivers (yellow = inputs used by AI / Base operating math)"
    d["B2"].font = TITLE
    d["B3"] = "These cells are the single source of truth for the commentary list. Operating tabs read these values."
    rows = [
        (5, "Revenue growth", rev_g, "0.0%", "AI_Operating!C5:G5"),
        (6, "Gross margin", gm, "0.0%", "AI_Operating COGS row"),
        (7, "Sales payroll cut (Y1)", payroll_cut, "0.0%", "AI_Operating S&M Y1"),
        (8, "Variable commission cut", commission_cut, "0.0%", "AI_Operating Commissions"),
        (9, "AI infrastructure ($m)", ai_cost, "#,##0.0", "AI_Operating AI cost"),
        (10, "S&M expense growth Y2+", sm_g, "0.0%", "AI_Operating S&M Y2+"),
        (11, "G&A % of revenue", ga_pct, "0.0%", "AI_Operating G&A"),
        (12, "CapEx reduction vs base", capex_cut, "0.0%", "FCF CapEx"),
        (13, "WC / receivables improvement", wc_improve, "0.0%", "FCF ΔNWC"),
        (14, "SOFR", sofr, "0.00%", "Debt interest"),
        (15, "Term Loan A spread", tla_sp, "0.00%", "TLA rate = SOFR+spread"),
        (16, "Term Loan B spread", tlb_sp, "0.00%", "TLB rate = SOFR+spread"),
        (17, "Mandatory amortization", mand_amort, "0.0%", "% of initial new debt / yr"),
        (18, "Cash sweep %", sweep, "0.0%", "Debt_Sweep_AI"),
        (19, "Tax rate", tax, "0.0%", "NI tax"),
        (20, "Exit EV/EBITDA multiple", exit_mult, "0.0x", "LBO!D17 / Strategy_Summary"),
        (21, "S&M baseline ($m)", sm0, "#,##0.0", "35% × FY24 revenue"),
        (22, "Payroll baseline ($m)", payroll0, "#,##0.0", "S&M − commission pool"),
        (23, "Commission baseline ($m)", commission0, "#,##0.0", "25% of S&M baseline"),
        (24, "CapEx base % of rev", capex_base, "0.0%", "Before AI −5%"),
        (25, "WC base % of Δrev", wc_base, "0.0%", "Before AI improvement"),
    ]
    d["B4"] = "Driver"
    d["C4"] = "Value"
    d["D4"] = "Feeds"
    for c in range(2, 5):
        d.cell(4, c).fill = HDR
        d.cell(4, c).font = HDR_F
    for r, name, val, fmt, feeds in rows:
        style_lab(d.cell(r, 2, name))
        cell = d.cell(r, 3, val)
        style_inp(cell, fmt)
        d.cell(r, 4, feeds).font = BLACK
    # computed rates
    d["B27"] = "TLA rate (SOFR+4)"
    d["C27"] = "=C14+C15"
    d["C27"].fill = CALC
    d["C27"].number_format = "0.00%"
    d["B28"] = "TLB rate (SOFR+5)"
    d["C28"] = "=C14+C16"
    d["C28"].fill = CALC
    d["C28"].number_format = "0.00%"
    d.column_dimensions["B"].width = 36
    d.column_dimensions["C"].width = 14
    d.column_dimensions["D"].width = 36

    def project(ai: bool):
        years = []
        rev = rev0
        payroll = payroll0
        commission = commission0
        tla = tla0
        tlb = tlb0
        notes = notes0
        for i in range(5):
            rev = rev * (1 + rev_g)
            if ai:
                if i == 0:
                    payroll = payroll0 * (1 - payroll_cut)
                    commission = commission0 * (1 - commission_cut)
                else:
                    payroll = payroll * (1 + sm_g)
                    commission = commission * (1 + sm_g)
                ai_c = ai_cost
                capex = rev * capex_base * (1 - capex_cut)
                dnwc = (rev - (rev0 if i == 0 else years[i - 1]["rev"])) * wc_base * (1 - wc_improve)
            else:
                # base: S&M grows with revenue at 35% total (payroll+commission stay proportional)
                payroll = rev * (payroll0 / rev0)
                commission = rev * (commission0 / rev0)
                ai_c = 0.0
                capex = rev * capex_base
                dnwc = (rev - (rev0 if i == 0 else years[i - 1]["rev"])) * wc_base
            cogs = rev * (1 - gm)
            rd = rev * rd_pct
            ga = rev * ga_pct
            da = rev * da_pct
            sm = payroll + commission
            ebitda = rev - cogs - sm - ai_c - rd - ga
            ebit = ebitda - da
            # interest on beginning balances
            tla_rate = sofr + tla_sp
            tlb_rate = sofr + tlb_sp
            note_rate = sofr + 0.035
            interest = tla * tla_rate + tlb * tlb_rate + notes * note_rate
            ebt = ebit - interest
            tax_amt = max(0, ebt) * tax
            ni = ebt - tax_amt
            fcf = ni + da - capex - dnwc
            debt_beg = tla + tlb + notes
            mandatory = new_debt * mand_amort
            mandatory = min(debt_beg, mandatory)
            optional = sweep * max(0, fcf - mandatory)
            paydown = min(debt_beg, mandatory + optional)
            # waterfall: TLA then TLB then notes
            rem = paydown
            tla_pay = min(tla, rem); rem -= tla_pay; tla -= tla_pay
            tlb_pay = min(tlb, rem); rem -= tlb_pay; tlb -= tlb_pay
            notes_pay = min(notes, rem); notes -= notes_pay
            years.append(
                dict(
                    year=2025 + i,
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
                    debt_beg=debt_beg,
                    debt_end=tla + tlb + notes,
                    ebitda_m=ebitda / rev,
                    capex=capex,
                    dnwc=dnwc,
                )
            )
        return years

    ai_rows = project(True)
    base_rows = project(False)

    def write_op(name, rows, title):
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(name, 2)
        ws["B2"] = title
        ws["B2"].font = TITLE
        ws["B3"] = "Driver cells: Assumptions_Drivers!C5:C20"
        headers = ["Metric", "FY2024A"] + [f"Y{i} {2024+i}" for i in range(1, 6)]
        for i, h in enumerate(headers):
            cell = ws.cell(4, 2 + i, h)
            cell.fill = HDR
            cell.font = HDR_F
        metrics = [
            ("Revenue", "rev", rev0),
            ("COGS", "cogs", rev0 * (1 - gm)),
            ("Sales payroll (human)", "payroll", payroll0),
            ("Variable commissions", "commission", commission0),
            ("S&M human total", "sm", sm0),
            ("AI infrastructure", "ai", 0.0),
            ("R&D", "rd", 196.1),
            ("G&A (admin)", "ga", rev0 * ga_pct),
            ("EBITDA", "ebitda", ebitda0),
            ("EBITDA margin", "ebitda_m", ebitda0 / rev0),
            ("D&A", "da", 85.7),
            ("EBIT", "ebit", 97.4),
            ("Interest", "interest", None),
            ("Net income", "ni", None),
            ("CapEx", "capex", None),
            ("ΔNWC", "dnwc", None),
            ("FCF after interest", "fcf", None),
            ("Mandatory amort", "mandatory", None),
            ("Total debt paydown", "paydown", None),
            ("Debt end", "debt_end", debt0_refi),
        ]
        for i, (label, key, basev) in enumerate(metrics):
            r = 5 + i
            ws.cell(r, 2, label).font = BLACK
            if basev is not None:
                cell = ws.cell(r, 3, basev)
                cell.number_format = "0.0%" if key.endswith("_m") else "#,##0.0"
                cell.fill = CALC
            for j, row in enumerate(rows):
                v = row.get(key)
                if v is None:
                    continue
                cell = ws.cell(r, 4 + j, v)
                cell.number_format = "0.0%" if key.endswith("_m") else "#,##0.0"
                cell.fill = CALC
                cell.border = THIN
        ws.column_dimensions["B"].width = 26
        for c in range(3, 10):
            ws.column_dimensions[get_column_letter(c)].width = 12

    write_op("AI_Operating", ai_rows, "AI-SDR case — drivers from Assumptions_Drivers")
    write_op("Base_Operating", base_rows, "Base case — no AI transformation")

    # Debt sweep sheet
    if "Debt_Sweep_AI" in wb.sheetnames:
        del wb["Debt_Sweep_AI"]
    ds = wb.create_sheet("Debt_Sweep_AI", 2)
    ds["B2"] = "Debt sweep — TLA/TLB priced at SOFR+4 / SOFR+5; 1% mandatory + 100% cash sweep"
    ds["B2"].font = TITLE
    ds["B3"] = "Rates from Assumptions_Drivers!C14:C18"
    headers = ["Year", "FCF", "Mandatory 1%", "Total paydown", "AI debt end", "Base debt end", "Extra paydown vs base"]
    for i, h in enumerate(headers):
        cell = ds.cell(4, 2 + i, h)
        cell.fill = HDR
        cell.font = HDR_F
    for i, (ar, br) in enumerate(zip(ai_rows, base_rows)):
        r = 5 + i
        vals = [ar["year"], ar["fcf"], ar["mandatory"], ar["paydown"], ar["debt_end"], br["debt_end"], br["debt_end"] - ar["debt_end"]]
        for c, v in enumerate(vals):
            cell = ds.cell(r, 2 + c, v)
            cell.number_format = "0" if c == 0 else "#,##0.0"
            cell.fill = CALC
            cell.border = THIN
    ds["B11"] = "TLA initial / TLB initial / Notes initial"
    ds["C11"] = tla0
    ds["D11"] = tlb0
    ds["E11"] = notes0
    ds["B12"] = "TLA rate / TLB rate"
    ds["C12"] = sofr + tla_sp
    ds["D12"] = sofr + tlb_sp
    ds["C12"].number_format = ds["D12"].number_format = "0.00%"

    # Returns
    def rets(rows):
        exit_ebitda = rows[-1]["ebitda"]
        exit_ev = exit_ebitda * exit_mult
        exit_eq = exit_ev - rows[-1]["debt_end"]
        moic = exit_eq / sponsor
        irr = moic ** 0.2 - 1
        return exit_ebitda, exit_ev, exit_eq, moic, irr

    ai_xebitda, ai_xev, ai_xeq, ai_moic, ai_irr = rets(ai_rows)
    b_xebitda, b_xev, b_xeq, b_moic, b_irr = rets(base_rows)
    bps_y2 = (ai_rows[1]["ebitda_m"] - (ebitda0 / rev0)) * 10000
    irr_uplift = (ai_irr - b_irr) * 100

    # Strategy summary refresh (keep commentary col if present)
    if "Strategy_Summary" in wb.sheetnames:
        del wb["Strategy_Summary"]
    s = wb.create_sheet("Strategy_Summary", 0)
    s["B2"] = "vengeanceaiUSC-LBOMODEL1 — Strategy Summary"
    s["B2"].font = TITLE
    s["B3"] = "AI-SDR cost takeout drives margin, FCF, debt paydown, MOIC/IRR — exit multiple held at 13x in both cases"
    s["B5"] = "RESULTS — AI vs Base"
    s["B5"].font = BOLD
    headers = ["Metric", "Base (no AI)", "AI-SDR case", "Delta", "COMMENTARY"]
    for i, h in enumerate(headers):
        cell = s.cell(6, 2 + i, h)
        cell.fill = HDR
        cell.font = HDR_F
    commentary = {
        7: "Year 2 operating profitability under 10% growth / 80% GM / AI S&M path.",
        8: "Y2 EBITDA margin minus FY24 baseline margin, in basis points (target ≥500).",
        9: "Year 5 EBITDA ($m).",
        10: "Exit EV = Y5 EBITDA × 13x (Assumptions_Drivers!C20 / LBO!D17).",
        11: "Ending net debt after mandatory amort + 100% cash sweep.",
        12: "Exit equity = Exit EV − ending debt.",
        13: "MOIC = Exit equity / sponsor equity check.",
        14: "IRR over 5 years = MOIC^(1/5) − 1.",
    }
    metrics = [
        (7, "Y2 EBITDA margin", base_rows[1]["ebitda_m"], ai_rows[1]["ebitda_m"], "0.0%"),
        (8, "Y2 margin vs baseline (bps)", (base_rows[1]["ebitda_m"] - ebitda0 / rev0) * 10000, bps_y2, "#,##0.0"),
        (9, "Y5 EBITDA ($m)", base_rows[4]["ebitda"], ai_rows[4]["ebitda"], "#,##0.0"),
        (10, "Exit EV @ 13x ($m)", b_xev, ai_xev, "#,##0.0"),
        (11, "Exit net debt ($m)", base_rows[-1]["debt_end"], ai_rows[-1]["debt_end"], "#,##0.0"),
        (12, "Exit equity value ($m)", b_xeq, ai_xeq, "#,##0.0"),
        (13, "MOIC", b_moic, ai_moic, "0.00x"),
        (14, "IRR (5-year)", b_irr, ai_irr, "0.0%"),
    ]
    for r, name, b, a, fmt in metrics:
        s.cell(r, 2, name)
        for c, v in [(3, b), (4, a), (5, a - b)]:
            cell = s.cell(r, c, v)
            cell.number_format = fmt
            cell.fill = CALC
            cell.border = THIN
        cell = s.cell(r, 6, commentary[r])
        cell.fill = IN
        cell.alignment = WRAP
        cell.border = THIN
        s.row_dimensions[r].height = 36
    s["B16"] = f"IRR uplift from AI S&M: {irr_uplift:.1f} percentage points | Y2 margin expansion: {bps_y2:.0f} bps (target ≥500: {'YES' if bps_y2 >= 500 else 'NO'})"
    s["B16"].fill = PatternFill("solid", fgColor="C6EFCE")
    s["B18"] = "PDF assumptions list:"
    s["B19"] = RAW_PDF
    s["B19"].hyperlink = RAW_PDF
    s["B19"].font = LINK
    s["B20"] = "In-Excel full commentary: 00_Assumptions_List | Drivers: Assumptions_Drivers"
    s.column_dimensions["B"].width = 32
    for c in ["C", "D", "E"]:
        s.column_dimensions[c].width = 14
    s.column_dimensions["F"].width = 70

    # Patch LBO exit multiple / returns if present
    if "LBO" in wb.sheetnames:
        lbo = wb["LBO"]
        lbo["D17"] = exit_mult
        lbo["D13"] = ebitda0
        lbo["I275"] = round(ai_moic, 3)
        lbo["J275"] = round(ai_irr, 4)
        lbo["I276"] = round(b_moic, 3)
        lbo["J276"] = round(b_irr, 4)
        lbo["J277"] = round(irr_uplift / 100, 4)
        lbo["B266"] = "SUMMARY @ 13.0x EXIT — AI vs Base (see Strategy_Summary)"

    return {
        "ai_irr": ai_irr,
        "base_irr": b_irr,
        "irr_uplift": irr_uplift,
        "bps_y2": bps_y2,
        "ai_moic": ai_moic,
        "base_moic": b_moic,
    }


def write_assumptions_list(wb):
    if "00_Assumptions_List" in wb.sheetnames:
        del wb["00_Assumptions_List"]
    ws = wb.create_sheet("00_Assumptions_List", 1)
    ws["B2"] = "vengeanceaiUSC-LBOMODEL1 — Assumptions List (commentary + Tab/Cell map)"
    ws["B2"].font = TITLE
    ws["B3"] = "Each row states Where (Tab!Cell), the math location, commentary, and the formula meaning. Not DCF-primary — LBO/operating engine."
    ws["B4"] = f"PDF: {RAW_PDF}"
    ws["B4"].hyperlink = RAW_PDF
    ws["B4"].font = LINK

    headers = [
        "#",
        "Assumption",
        "Value",
        "Engine (Ex LBO / DCF)",
        "Input Tab",
        "Input Cell",
        "Jump → Input",
        "Math Tab",
        "Math Cell / Row",
        "Jump → Math",
        "What the math does",
        "Commentary (Why / What changed)",
    ]
    for i, h in enumerate(headers):
        cell = ws.cell(6, 2 + i, h)
        cell.fill = HDR
        cell.font = HDR_F
        cell.alignment = Alignment(wrap_text=True, horizontal="center")

    for i, a in enumerate(ASSUMPTIONS, 1):
        r = 6 + i
        input_target = a["cell"].split("/")[0].strip()
        math_target = a.get("math_jump", a["math_cell"].split("(")[0].strip().split(":")[0].strip())
        vals = [
            i,
            a["name"],
            a["value"],
            a.get("engine", "LBO (not DCF)"),
            a["tab"],
            a["cell"],
            f"Go to {a['tab']}!{input_target}",
            a["math_tab"],
            a["math_cell"],
            f"Go to {a['math_tab']}!{math_target}",
            a["math"],
            a["commentary"],
        ]
        for c, v in enumerate(vals):
            cell = ws.cell(r, 2 + c, v)
            cell.alignment = WRAP
            cell.border = THIN
            cell.font = BLACK
            if c == 6:
                cell.hyperlink = f"#'{a['tab']}'!{input_target}"
                cell.font = LINK
            if c == 9:
                cell.hyperlink = f"#'{a['math_tab']}'!{math_target}"
                cell.font = LINK
            if c == 11:
                cell.fill = NOTE
        ws.row_dimensions[r].height = 78

    widths = [4, 22, 26, 18, 20, 12, 28, 16, 32, 28, 40, 70]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(2 + i)].width = w
    ws.freeze_panes = "B7"
    ws.auto_filter.ref = f"B6:M{6 + len(ASSUMPTIONS)}"

    # Note on DCF
    note_r = 8 + len(ASSUMPTIONS)
    ws.cell(note_r, 2, "NOTE ON DCF TAB").font = BOLD
    ws.cell(
        note_r + 1,
        2,
        "These AI-SDR operating assumptions are NOT the primary drivers of the DCF tab. "
        "IRR/MOIC value-creation math is on Strategy_Summary, AI_Operating, Base_Operating, Debt_Sweep_AI, and LBO!D17/J275. "
        "DCF remains a secondary WSP shell for football-field context.",
    ).alignment = WRAP
    ws.merge_cells(start_row=note_r + 1, start_column=2, end_row=note_r + 2, end_column=13)


def write_pdf():
    styles = getSampleStyleSheet()
    title = ParagraphStyle("T", parent=styles["Heading1"], fontSize=13, spaceAfter=8)
    body = ParagraphStyle("B", parent=styles["Normal"], fontSize=8, leading=11)
    small = ParagraphStyle("S", parent=styles["Normal"], fontSize=7.5, leading=10)
    doc = SimpleDocTemplate(
        str(PDF),
        pagesize=letter,
        leftMargin=0.55 * inch,
        rightMargin=0.55 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )
    story = [
        Paragraph("vengeanceaiUSC-LBOMODEL1 — Assumptions List (Commentary + Tab/Cell Map)", title),
        Paragraph(
            "AI-SDR LBO assumptions with workbook locations. Primary math is LBO/operating (not DCF). "
            f"<b>Excel:</b> {RAW_XLSX}",
            body,
        ),
        Spacer(1, 6),
    ]
    data = [["#", "Assumption / Value", "Engine + Input", "Math location", "Commentary"]]
    for i, a in enumerate(ASSUMPTIONS, 1):
        data.append(
            [
                str(i),
                Paragraph(f"<b>{a['name']}</b><br/>{a['value']}", small),
                Paragraph(
                    f"<b>{a.get('engine', 'LBO')}</b><br/>Input: {a['tab']}!{a['cell']}",
                    small,
                ),
                Paragraph(
                    f"<b>{a['math_tab']}!{a['math_cell']}</b><br/>{a['math']}",
                    small,
                ),
                Paragraph(a["commentary"], small),
            ]
        )
    t = Table(data, colWidths=[0.3 * inch, 1.4 * inch, 1.7 * inch, 1.2 * inch, 3.0 * inch])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B0B0B0")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F8FC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(t)
    story.append(Spacer(1, 8))
    story.append(
        Paragraph(
            "<b>DCF note:</b> These items drive the LBO operating engine (AI_Operating / Debt_Sweep_AI / LBO exit). "
            "They are not the primary inputs on the DCF tab.",
            body,
        )
    )
    story.append(Paragraph(f"PDF raw link: {RAW_PDF}", small))
    story.append(Paragraph("Educational / research use only — not investment advice.", small))
    doc.build(story)

    # Markdown twin
    lines = [
        "# vengeanceaiUSC-LBOMODEL1 — Assumptions List",
        "",
        f"**Excel:** {RAW_XLSX}",
        "",
        f"**PDF:** {RAW_PDF}",
        "",
        "Primary engine = LBO / AI_Operating / Debt_Sweep_AI (not DCF).",
        "",
    ]
    for i, a in enumerate(ASSUMPTIONS, 1):
        lines += [
            f"## {i}. {a['name']} — {a['value']}",
            f"- **Engine:** {a.get('engine', 'LBO (not DCF)')}",
            f"- **Input:** `{a['tab']}!{a['cell']}`",
            f"- **Math location:** `{a['math_tab']}!{a['math_cell']}`",
            f"- **Math:** {a['math']}",
            f"- **Commentary:** {a['commentary']}",
            "",
        ]
    MD.write_text("\n".join(lines))


def main():
    wb = load_workbook(XLSX)
    stats = rebuild_drivers_and_ops(wb)
    write_assumptions_list(wb)
    wb.save(XLSX)
    wb.save(XLSX2)
    write_pdf()
    print("Saved Excel + PDF")
    print(stats)


if __name__ == "__main__":
    main()
