"""Bake live CAPM / FCFF / TV equations into the DCF sheet (columns Q–T).

vengeanceaiUSCMODEL3: these are real Excel formulas, not text labels.
Changing Rf / ERP / Beta / Rd / weights updates Ke and WACC (D6) automatically.
"""

from __future__ import annotations

from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .named_range_map import SHEET_DCF
from .wacc import WaccInputs

BLUE = Font(name="Calibri", color="0000FF")
BLACK = Font(name="Calibri", color="000000")
BOLD = Font(name="Calibri", bold=True, color="1F4E79")
HDR = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
EQ_FONT = Font(name="Consolas", size=10, color="000000")
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


def _input(cell, value, fmt: str | None = None) -> None:
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def _formula(cell, formula: str, fmt: str | None = None) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def _label(cell, text: str, *, bold: bool = False, eq: bool = False) -> None:
    cell.value = text
    cell.font = EQ_FONT if eq else (BOLD if bold else BLACK)
    cell.alignment = Alignment(wrap_text=True, vertical="center")


def bake_equations_into_dcf(wb, *, tax_rate: float | None = None) -> None:
    """Write CAPM stack + equation board into DCF!Q:T and wire D6 ← computed WACC."""
    if SHEET_DCF not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_DCF}")
    ws = wb[SHEET_DCF]
    w = WaccInputs(tax_rate=tax_rate if tax_rate is not None else WaccInputs().tax_rate)

    for col, width in (("Q", 38), ("R", 14), ("S", 44), ("T", 16)):
        ws.column_dimensions[col].width = width

    # ----- Header -----
    ws["Q2"] = "vengeanceaiUSCMODEL3 — LIVE EQUATIONS"
    ws["Q2"].font = HDR
    ws["Q2"].fill = HDR_FILL
    ws.merge_cells("Q2:T2")

    ws["Q3"] = "Yellow = inputs (edit these). Green = formulas (auto-update)."
    ws["Q3"].font = Font(name="Calibri", italic=True, color="666666")
    ws.merge_cells("Q3:T3")

    # ----- CAPM / WACC block (Q5:T16) -----
    ws["Q5"] = "1) CAPM → WACC"
    ws["Q5"].font = BOLD
    ws["Q5"].fill = SECT_FILL
    ws.merge_cells("Q5:T5")

    rows = [
        (6, "Risk-free rate (Rf)", w.risk_free_rate, "US 10Y Treasury / FRED DGS10", "0.00%"),
        (7, "Equity risk premium (ERP)", w.equity_risk_premium, "Damodaran implied ERP", "0.00%"),
        (8, "Beta (β)", w.beta, "Yahoo Finance 5Y monthly", "0.00"),
    ]
    for r, label, val, src, fmt in rows:
        _label(ws[f"Q{r}"], label)
        _input(ws[f"R{r}"], val, fmt)
        _label(ws[f"S{r}"], src)
        ws[f"T{r}"] = ""

    _label(ws["Q9"], "Ke = Rf + β × ERP", bold=True, eq=True)
    _formula(ws["R9"], "=R6+R8*R7", "0.00%")
    _label(ws["S9"], "Cost of equity (CAPM)")
    _label(ws["T9"], "→ feeds WACC", eq=True)

    _label(ws["Q10"], "Pre-tax cost of debt (Rd)")
    _input(ws["R10"], w.pre_tax_cost_of_debt, "0.00%")
    _label(ws["S10"], "6.250% Senior Notes due 2034 (SEC 8-K)")

    _label(ws["Q11"], "Tax rate (t)")
    _formula(ws["R11"], "=$D$5", "0.00%")
    _label(ws["S11"], "Linked to DCF tax assumption (FY25 effective)")

    _label(ws["Q12"], "Rd_aftertax = Rd × (1 − t)", bold=True, eq=True)
    _formula(ws["R12"], "=R10*(1-R11)", "0.00%")
    _label(ws["S12"], "After-tax cost of debt")

    _label(ws["Q13"], "Equity weight (We)")
    _input(ws["R13"], w.equity_weight, "0%")
    _label(ws["S13"], "Target market weights")

    _label(ws["Q14"], "Debt weight (Wd)")
    _input(ws["R14"], w.debt_weight, "0%")
    _label(ws["S14"], "We + Wd should = 100%")

    _label(ws["Q15"], "WACC = We×Ke + Wd×Rd_aftertax", bold=True, eq=True)
    _formula(ws["R15"], "=R13*R9+R14*R12", "0.00%")
    _label(ws["S15"], "PRIMARY discount rate — linked into D6")
    _label(ws["T15"], "→ D6", eq=True)

    # Wire Discount Rate cell to live WACC
    _formula(ws["D6"], "=R15", "0.00%")
    ws["B6"] = "Discount Rate (WACC = We×Ke + Wd×Rd(1−t))"

    # ----- FCFF identity -----
    ws["Q17"] = "2) Unlevered Free Cash Flow"
    ws["Q17"].font = BOLD
    ws["Q17"].fill = SECT_FILL
    ws.merge_cells("Q17:T17")

    _label(ws["Q18"], "FCFF = EBIT − EBIT×t + D&A − CapEx − ΔNWC", bold=True, eq=True)
    ws.merge_cells("Q18:S18")
    _label(ws["Q19"], "Excel (col E example)")
    _label(ws["R19"], "=E21-E22+E23-E24-E25", eq=True)
    ws.merge_cells("R19:T19")
    _label(ws["Q20"], "Tax on EBIT (unlevered)")
    _label(ws["R20"], "=E21*$D$5", eq=True)
    _label(ws["S20"], "NOT levered IS tax dollars")

    # Show computed FCFF cross-check for final year
    _label(ws["Q21"], "FCFF check (FY5 / col I)")
    _formula(ws["R21"], "=I26", "#,##0.0")
    _label(ws["S21"], "Must equal I21−I22+I23−I24−I25")

    # ----- Terminal value -----
    ws["Q23"] = "3) Terminal Value"
    ws["Q23"].font = BOLD
    ws["Q23"].fill = SECT_FILL
    ws.merge_cells("Q23:T23")

    _label(ws["Q24"], "TV_exit = EBITDA_n × ExitMultiple   [PRIMARY]", bold=True, eq=True)
    ws.merge_cells("Q24:S24")
    _label(ws["Q25"], "Excel")
    _label(ws["R25"], "=($I$21+$I$23)*$D$8", eq=True)
    _formula(ws["T25"], "=J27", "#,##0.0")

    _label(ws["Q26"], "TV_gordon = FCFF_n×(1+g)/(WACC−g)   [CHECK]", bold=True, eq=True)
    ws.merge_cells("Q26:S26")
    _label(ws["Q27"], "Excel")
    _label(ws["R27"], "=IF(($D$6-$D$7)<=0,\"err\",$I$26*(1+$D$7)/($D$6-$D$7))", eq=True)
    _formula(ws["T27"], "=M19", "#,##0.0")

    _label(ws["Q28"], "Implied Gordon exit multiple")
    _formula(ws["R28"], "=IF(($I$21+$I$23)=0,0,T27/($I$21+$I$23))", "0.00x")
    _label(ws["S28"], "Gordon TV / EBITDA_n")

    # ----- EV bridge -----
    ws["Q30"] = "4) Enterprise → Equity → $/share"
    ws["Q30"].font = BOLD
    ws["Q30"].fill = SECT_FILL
    ws.merge_cells("Q30:T30")

    _label(ws["Q31"], "EV = XNPV(WACC, TransactionCF, mid-year dates)", bold=True, eq=True)
    ws.merge_cells("Q31:S31")
    _label(ws["Q32"], "Excel")
    _label(ws["R32"], "=XNPV(D6,D28:J28,D18:J18)", eq=True)
    _formula(ws["T32"], "=D32", "#,##0.0")

    _label(ws["Q33"], "Equity = EV + Cash − Debt", bold=True, eq=True)
    _formula(ws["R33"], "=D35", "#,##0.0")
    _label(ws["S33"], "D32+D33−D34")

    _label(ws["Q34"], "$/share = Equity / Shares", bold=True, eq=True)
    _formula(ws["R34"], "=D37", "$#,##0.00")
    _label(ws["S34"], "Intrinsic value per share")

    _label(ws["Q35"], "Upside vs market")
    _formula(ws["R35"], "=D37/D11-1", "0.0%")
    _label(ws["S35"], "Intrinsic / Current Price − 1")

    # ----- NWC -----
    ws["Q37"] = "5) Operating NWC (feeds ΔNWC → FCFF)"
    ws["Q37"].font = BOLD
    ws["Q37"].fill = SECT_FILL
    ws.merge_cells("Q37:T37")

    _label(ws["Q38"], "NWC = AR + Inventory − AP − DeferredRevenue", bold=True, eq=True)
    ws.merge_cells("Q38:T38")
    _label(ws["Q39"], "ΔNWC_t = NWC_t − NWC_(t−1)", eq=True)
    _label(ws["S39"], "Linked from 3-statement CF row 64")
    _formula(ws["T39"], "=I25", "#,##0.0")

    # ----- Sources -----
    ws["Q41"] = "Sources (SEC preferred)"
    ws["Q41"].font = BOLD
    ws["Q41"].fill = SECT_FILL
    ws.merge_cells("Q41:T41")
    sources = [
        "10-K: sec.gov/.../fico-20250930.htm",
        "10-Q: sec.gov/.../fico-20260630.htm",
        "8-K 6.250% notes: sec.gov/.../d56220d8k.htm",
        "Damodaran ERP: pages.stern.nyu.edu/adamodar/",
        "FRED DGS10 + Yahoo FICO beta",
    ]
    for i, s in enumerate(sources):
        _label(ws[f"Q{42+i}"], s)
        ws.merge_cells(f"Q{42+i}:T{42+i}")

    # Row labels on FCFF block so equations are visible in-place
    ws["C21"] = "EBIT"
    ws["C22"] = "EBIT×t"
    ws["C23"] = "+D&A"
    ws["C24"] = "−CapEx"
    ws["C25"] = "−ΔNWC"
    ws["C26"] = "FCFF=EBIT−tax+DA−CapEx−ΔNWC"
    for r in range(21, 27):
        ws[f"C{r}"].font = Font(name="Consolas", size=8, color="666666")
