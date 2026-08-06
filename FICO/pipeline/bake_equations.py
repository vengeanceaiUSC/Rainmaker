"""Bake live CAPM / FCFF / TV equations into the DCF sheet (columns Q–V).

vengeanceaiUSCMODEL3: real Excel formulas + clickable SOURCE hyperlinks
next to every WACC input (Rf, ERP, Beta, Rd, tax, weights).
"""

from __future__ import annotations

from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.cell.cell import Cell

from .model3_assumptions import SOURCE_LINKS
from .named_range_map import SHEET_DCF
from .wacc import WaccInputs

BLUE = Font(name="Calibri", color="0000FF")
BLACK = Font(name="Calibri", color="000000")
BOLD = Font(name="Calibri", bold=True, color="1F4E79")
HDR = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
EQ_FONT = Font(name="Consolas", size=10, color="000000")
LINK_FONT = Font(name="Calibri", color="0563C1", underline="single")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
SECT_FILL = PatternFill("solid", fgColor="D6EAF8")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

# Canonical URLs for each CAPM input
URL_RF = SOURCE_LINKS["FRED DGS10"]
URL_RF_ALT = "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve"
URL_ERP = SOURCE_LINKS["Damodaran ERP"]
URL_ERP_HIST = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histimpl.html"
URL_BETA = SOURCE_LINKS["Yahoo FICO"]
URL_RD = SOURCE_LINKS["8-K 6.250% notes"]
URL_RD_PRICING = "https://www.sec.gov/Archives/edgar/data/814547/000119312526102450/d92566dex991.htm"
URL_TAX = SOURCE_LINKS["10-K FY2025"]
URL_10Q = SOURCE_LINKS["10-Q Q3 FY2026"]
URL_FACTS = SOURCE_LINKS["companyfacts"]


def _input(cell: Cell, value, fmt: str | None = None) -> None:
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def _formula(cell: Cell, formula: str, fmt: str | None = None) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def _label(cell: Cell, text: str, *, bold: bool = False, eq: bool = False) -> None:
    cell.value = text
    cell.font = EQ_FONT if eq else (BOLD if bold else BLACK)
    cell.alignment = Alignment(wrap_text=True, vertical="center")


def _hyperlink(cell: Cell, url: str, display: str) -> None:
    """Clickable Excel hyperlink."""
    cell.value = display
    cell.hyperlink = url
    cell.font = LINK_FONT
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    cell.border = THIN


def bake_equations_into_dcf(wb, *, tax_rate: float | None = None) -> None:
    """Write CAPM stack + equation board + source hyperlinks; wire D6 ← WACC."""
    if SHEET_DCF not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_DCF}")
    ws = wb[SHEET_DCF]
    w = WaccInputs(tax_rate=tax_rate if tax_rate is not None else WaccInputs().tax_rate)

    for col, width in (
        ("Q", 40),
        ("R", 14),
        ("S", 36),
        ("T", 14),
        ("U", 52),  # primary source URL (clickable)
        ("V", 52),  # secondary source URL (clickable)
    ):
        ws.column_dimensions[col].width = width

    # ----- Header -----
    ws["Q2"] = "vengeanceaiUSCMODEL8 — LIVE EQUATIONS + SOURCE LINKS"
    ws["Q2"].font = HDR
    ws["Q2"].fill = HDR_FILL
    ws.merge_cells("Q2:V2")

    ws["Q3"] = (
        "Yellow = inputs (edit). Green = formulas. Blue underlined = clickable sources."
    )
    ws["Q3"].font = Font(name="Calibri", italic=True, color="666666")
    ws.merge_cells("Q3:V3")

    ws["Q4"] = "Input"
    ws["R4"] = "Value"
    ws["S4"] = "Description"
    ws["T4"] = "Role"
    ws["U4"] = "Primary source (click)"
    ws["V4"] = "Secondary source (click)"
    for col in "QRSTUV":
        ws[f"{col}4"].font = BOLD
        ws[f"{col}4"].fill = SECT_FILL
        ws[f"{col}4"].border = THIN

    # ----- CAPM / WACC block -----
    ws["Q5"] = "1) CAPM → WACC   (D6 = R15)"
    ws["Q5"].font = BOLD
    ws["Q5"].fill = SECT_FILL
    ws.merge_cells("Q5:V5")

    # Rf
    _label(ws["Q6"], "Risk-free rate (Rf)")
    _input(ws["R6"], w.risk_free_rate, "0.00%")
    _label(ws["S6"], "US 10Y Treasury yield")
    _label(ws["T6"], "CAPM input")
    _hyperlink(ws["U6"], URL_RF, "FRED DGS10 (US 10Y)")
    _hyperlink(ws["V6"], URL_RF_ALT, "US Treasury Daily Par Yield Curve")

    # ERP
    _label(ws["Q7"], "Equity risk premium (ERP)")
    _input(ws["R7"], w.equity_risk_premium, "0.00%")
    _label(ws["S7"], "Damodaran implied ERP")
    _label(ws["T7"], "CAPM input")
    _hyperlink(ws["U7"], URL_ERP, "Damodaran Online — Implied ERP")
    _hyperlink(ws["V7"], URL_ERP_HIST, "Damodaran histimpl.xls / table")

    # Beta
    _label(ws["Q8"], "Beta (β)")
    _input(ws["R8"], w.beta, "0.00")
    _label(ws["S8"], "Yahoo 5Y monthly beta")
    _label(ws["T8"], "CAPM input")
    _hyperlink(ws["U8"], URL_BETA, "Yahoo Finance FICO Key Statistics")
    _hyperlink(ws["V8"], "https://finance.yahoo.com/quote/FICO/", "Yahoo Finance FICO quote")

    # Ke formula
    _label(ws["Q9"], "Ke = Rf + β × ERP", bold=True, eq=True)
    _formula(ws["R9"], "=R6+R8*R7", "0.00%")
    _label(ws["S9"], "Cost of equity (CAPM)")
    _label(ws["T9"], "→ WACC", eq=True)
    _label(ws["U9"], "Derived: R6 + R8 × R7")
    _label(ws["V9"], "")

    # Rd
    _label(ws["Q10"], "Pre-tax cost of debt (Rd)")
    _input(ws["R10"], w.pre_tax_cost_of_debt, "0.00%")
    _label(ws["S10"], "6.250% Senior Notes due 2034")
    _label(ws["T10"], "WACC input")
    _hyperlink(ws["U10"], URL_RD, "SEC 8-K — Notes closing (6.250%)")
    _hyperlink(ws["V10"], URL_RD_PRICING, "SEC EX-99.1 — Notes pricing press release")

    # Tax — D5 is linked to 3-Statement!J13 (same rate as forecast)
    _label(ws["Q11"], "Tax rate (t)")
    _formula(ws["R11"], "=$D$5", "0.00%")
    _label(ws["S11"], "D5 ← 3S!J13 (FY25 tax/EBT)")
    _label(ws["T11"], "Linked D5←3S")
    _hyperlink(ws["U11"], URL_TAX, "SEC 10-K FY2025 (tax / EBT)")
    _hyperlink(ws["V11"], URL_FACTS, "SEC companyfacts XBRL (CIK 0000814547)")

    # Rd after tax
    _label(ws["Q12"], "Rd_aftertax = Rd × (1 − t)", bold=True, eq=True)
    _formula(ws["R12"], "=R10*(1-R11)", "0.00%")
    _label(ws["S12"], "After-tax cost of debt")
    _label(ws["T12"], "→ WACC", eq=True)
    _label(ws["U12"], "Derived: R10 × (1 − R11)")
    _label(ws["V12"], "")

    # Weights
    _label(ws["Q13"], "Equity weight (We)")
    _input(ws["R13"], w.equity_weight, "0%")
    _label(ws["S13"], "Target ~80% equity at market")
    _label(ws["T13"], "WACC weight")
    _hyperlink(ws["U13"], URL_10Q, "SEC 10-Q Q3 FY2026 (debt / capital)")
    _hyperlink(ws["V13"], URL_BETA, "Yahoo market cap for E/(D+E)")

    _label(ws["Q14"], "Debt weight (Wd)")
    _input(ws["R14"], w.debt_weight, "0%")
    _label(ws["S14"], "Target ~20% debt (We+Wd=100%)")
    _label(ws["T14"], "WACC weight")
    _hyperlink(ws["U14"], URL_10Q, "SEC 10-Q — total debt $5,582,389k")
    _hyperlink(ws["V14"], URL_TAX, "SEC 10-K — senior notes footnote")

    # WACC
    _label(ws["Q15"], "WACC = We×Ke + Wd×Rd_aftertax", bold=True, eq=True)
    _formula(ws["R15"], "=R13*R9+R14*R12", "0.00%")
    _label(ws["S15"], "PRIMARY discount rate → cell D6")
    _label(ws["T15"], "→ D6", eq=True)
    _hyperlink(ws["U15"], URL_TAX, "SEC filings (tax, debt coupons)")
    _hyperlink(ws["V15"], URL_ERP, "Damodaran (ERP methodology)")

    # Wire Discount Rate cell to live WACC + source note under assumptions
    _formula(ws["D6"], "=R15", "0.00%")
    ws["B6"] = "Discount Rate (WACC = We×Ke + Wd×Rd(1−t))"
    ws["C6"] = "Sources →"
    ws["C6"].font = Font(name="Calibri", italic=True, size=9, color="666666")
    # Put a compact source pointer next to D6 (column E area is forecast — use A6 note)
    # Pointers next to the assumption block (columns E–F are free on rows 5–8)
    _hyperlink(ws["E5"], URL_TAX, "Tax source: SEC 10-K")
    _hyperlink(ws["E6"], URL_RF, "WACC sources → scroll to col U (Rf/ERP/β/Rd links)")
    _hyperlink(ws["F6"], URL_ERP, "Damodaran ERP")
    _hyperlink(ws["G6"], URL_RD, "SEC 8-K Rd 6.250%")
    _hyperlink(ws["E8"], URL_RD, "Rd source: SEC 8-K 6.250% notes")
    ws["A6"] = "WACC sources → columns U–V"
    ws["A6"].font = Font(name="Calibri", italic=True, size=9, color="0563C1")

    # ----- FCFF identity -----
    ws["Q17"] = "2) Unlevered Free Cash Flow"
    ws["Q17"].font = BOLD
    ws["Q17"].fill = SECT_FILL
    ws.merge_cells("Q17:V17")

    _label(ws["Q18"], "FCFF = EBIT − EBIT×t + D&A − CapEx − ΔNWC", bold=True, eq=True)
    ws.merge_cells("Q18:S18")
    _hyperlink(ws["U18"], URL_FACTS, "Drivers from SEC XBRL → 3-statement")
    _hyperlink(ws["V18"], URL_TAX, "SEC 10-K historical IS/CF")

    _label(ws["Q19"], "Excel (col E example)")
    _label(ws["R19"], "=E21-E22+E23-E24-E25", eq=True)
    ws.merge_cells("R19:T19")

    _label(ws["Q20"], "Tax on EBIT (unlevered)")
    _label(ws["R20"], "=E21*$D$5", eq=True)
    _label(ws["S20"], "NOT levered IS tax dollars")
    _hyperlink(ws["U20"], URL_TAX, "Tax rate from SEC 10-K effective rate")

    _label(ws["Q21"], "FCFF check (FY5 / col I)")
    _formula(ws["R21"], "=I26", "#,##0.0")
    _label(ws["S21"], "Must equal I21−I22+I23−I24−I25")

    # ----- Terminal value -----
    ws["Q23"] = "3) Terminal Value"
    ws["Q23"].font = BOLD
    ws["Q23"].fill = SECT_FILL
    ws.merge_cells("Q23:V23")

    _label(ws["Q24"], "TV_exit = EBITDA_n × ExitMultiple   [PRIMARY]", bold=True, eq=True)
    ws.merge_cells("Q24:S24")
    _label(ws["Q25"], "Excel")
    _label(ws["R25"], "=($I$21+$I$23)*$D$8", eq=True)
    _formula(ws["T25"], "=J27", "#,##0.0")
    _hyperlink(ws["U25"], URL_BETA, "Exit multiple policy vs Yahoo / market multiples")

    _label(ws["Q26"], "TV_gordon = FCFF_n×(1+g)/(WACC−g)   [CHECK]", bold=True, eq=True)
    ws.merge_cells("Q26:S26")
    _label(ws["Q27"], "Excel")
    _label(ws["R27"], '=IF(($D$6-$D$7)<=0,"err",$I$26*(1+$D$7)/($D$6-$D$7))', eq=True)
    _formula(ws["T27"], "=M19", "#,##0.0")
    _hyperlink(ws["U27"], URL_ERP, "Gordon g vs Damodaran long-run growth framing")

    _label(ws["Q28"], "Implied Gordon exit multiple")
    _formula(ws["R28"], "=IF(($I$21+$I$23)=0,0,T27/($I$21+$I$23))", "0.00x")
    _label(ws["S28"], "Gordon TV / EBITDA_n")

    # ----- EV bridge -----
    ws["Q30"] = "4) Enterprise → Equity → $/share"
    ws["Q30"].font = BOLD
    ws["Q30"].fill = SECT_FILL
    ws.merge_cells("Q30:V30")

    _label(ws["Q31"], "EV = XNPV(WACC, TransactionCF, mid-year dates)", bold=True, eq=True)
    ws.merge_cells("Q31:S31")
    _label(ws["Q32"], "Excel")
    _label(ws["R32"], "=XNPV(D6,D28:J28,D18:J18)", eq=True)
    _formula(ws["T32"], "=D32", "#,##0.0")

    _label(ws["Q33"], "Equity = EV + Cash − Debt", bold=True, eq=True)
    _formula(ws["R33"], "=D35", "#,##0.0")
    _label(ws["S33"], "Net debt from 10-Q bridge")
    _hyperlink(ws["U33"], URL_10Q, "SEC 10-Q — cash, mkt secs, total debt")
    _hyperlink(ws["V33"], URL_TAX, "SEC 10-K — FY debt footnote")

    _label(ws["Q34"], "$/share = Equity / Shares", bold=True, eq=True)
    _formula(ws["R34"], "=D37", "$#,##0.00")
    _label(ws["S34"], "Shares outstanding (Yahoo)")
    _hyperlink(ws["U34"], URL_BETA, "Yahoo Finance — shares outstanding")
    _hyperlink(ws["V34"], "https://finance.yahoo.com/quote/FICO/", "Yahoo FICO quote (price)")

    _label(ws["Q35"], "Upside vs market")
    _formula(ws["R35"], "=D37/D11-1", "0.0%")
    _label(ws["S35"], "Intrinsic / Current Price − 1")

    # ----- NWC -----
    ws["Q37"] = "5) Operating NWC (feeds ΔNWC → FCFF)"
    ws["Q37"].font = BOLD
    ws["Q37"].fill = SECT_FILL
    ws.merge_cells("Q37:V37")

    _label(ws["Q38"], "NWC = Revenue × NWC% (flat 2.5% MODEL8)", bold=True, eq=True)
    ws.merge_cells("Q38:T38")
    _hyperlink(ws["U38"], URL_TAX, "SEC 10-K — AR / AP / deferred revenue")
    _hyperlink(ws["V38"], URL_FACTS, "XBRL DeferredRevenueCurrent")

    _label(ws["Q39"], "ΔNWC_t = NWC_t − NWC_(t−1)", eq=True)
    _label(ws["S39"], "Linked from 3-statement CF row 64 (NWC% driver)")
    _formula(ws["T39"], "=I25", "#,##0.0")

    # ----- Full source index -----
    ws["Q41"] = "6) Full source index (all clickable)"
    ws["Q41"].font = BOLD
    ws["Q41"].fill = SECT_FILL
    ws.merge_cells("Q41:V41")

    index = [
        (42, "SEC 10-K FY2025", URL_TAX, "Tax 18.8%, IS/BS/CF, senior notes"),
        (43, "SEC 10-Q Q3 FY2026", URL_10Q, "Debt $5,582M, cash+mkt secs net-debt bridge"),
        (44, "SEC 8-K 6.250% notes", URL_RD, "Pre-tax Rd coupon 6.250% due 2034"),
        (45, "SEC EX-99.1 pricing", URL_RD_PRICING, "Notes priced at par / offering terms"),
        (46, "SEC companyfacts XBRL", URL_FACTS, "Pipeline data source (CIK 0000814547)"),
        (47, "FRED DGS10", URL_RF, "Risk-free rate (US 10Y)"),
        (48, "US Treasury yield curve", URL_RF_ALT, "Official daily par yields"),
        (49, "Damodaran Implied ERP", URL_ERP, "Equity risk premium"),
        (50, "Damodaran histimpl", URL_ERP_HIST, "Historical implied ERP table"),
        (51, "Yahoo FICO stats", URL_BETA, "Beta 5Y monthly + shares out"),
    ]
    for r, name, url, note in index:
        _hyperlink(ws[f"Q{r}"], url, name)
        ws.merge_cells(f"Q{r}:S{r}")
        _label(ws[f"T{r}"], note)
        ws.merge_cells(f"T{r}:V{r}")

    # Row labels on FCFF block
    ws["C21"] = "EBIT"
    ws["C22"] = "EBIT×t"
    ws["C23"] = "+D&A"
    ws["C24"] = "−CapEx"
    ws["C25"] = "−ΔNWC"
    ws["C26"] = "FCFF=EBIT−tax+DA−CapEx−ΔNWC"
    for r in range(21, 27):
        ws[f"C{r}"].font = Font(name="Consolas", size=8, color="666666")
