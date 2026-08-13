#!/usr/bin/env python3
"""Build an editable FICO LBO Excel model (formulas preserved for Excel recalc)."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output" / "FICO_LBO_Completed_Model.xlsx"
REPO_ROOT_COPY = ROOT.parent / "FICO_LBO_Completed_Model.xlsx"

BLUE = Font(name="Calibri", color="0000FF")
BLACK = Font(name="Calibri", color="000000")
BOLD = Font(name="Calibri", bold=True)
TITLE = Font(name="Calibri", size=16, bold=True, color="1F4E79")
HDR_FONT = Font(name="Calibri", color="FFFFFF", bold=True)
HDR = PatternFill("solid", fgColor="1F4E79")
INPUT = PatternFill("solid", fgColor="FFF2CC")
CALC = PatternFill("solid", fgColor="E2EFDA")
SECTION = PatternFill("solid", fgColor="D6DCE4")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

# FICO SEC figures ($ millions) — FY2025 10-K / Q3 FY2026 10-Q / market
PRICE = 1046.23
SHARES_M = 21.597635
PREMIUM = 0.25
EBITDA_LTM = 939.802  # OpInc 924.850 + D&A 14.952
REV = {2023: 1513.557, 2024: 1717.526, 2025: 1990.869}
COGS = {2023: 311.053, 2024: 348.206, 2025: 353.722}
RD = {2023: 159.950, 2024: 171.940, 2025: 188.347}
SGA = {2023: 400.565, 2024: 462.834, 2025: 513.028}
DA = {2023: 14.638, 2024: 13.827, 2025: 14.952}
DEBT = 5552.389  # 10-Q total debt less cash+ST investments later via net
CASH = 304.537  # 248.444 + 56.093
GROSS_DEBT = 5582.389
MIN_CASH = 150.0
TAX = 0.188
ENTRY_MULT = None  # derived
EXIT_MULT = 18.0
HOLD_YEARS = 5
GROWTHS = [0.14, 0.12, 0.10, 0.09, 0.08]
EBITDA_MARGINS = [0.48, 0.49, 0.50, 0.505, 0.51]
CAPEX_PCT = [0.012, 0.011, 0.010, 0.009, 0.008]
NWC_PCT = [0.08, 0.075, 0.07, 0.065, 0.06]
# Debt structure (EBITDA turns)
REV_TURNS = 0.5
TLA_TURNS = 2.5
TLB_TURNS = 2.0
NOTE_TURNS = 1.5
SOFR = 0.043
REV_SPR = 0.025
TLA_SPR = 0.030
TLB_SPR = 0.035
NOTE_RATE = 0.075
TLA_AMORT = 0.05  # annual % of initial
FEE_PCT = 0.02


def style_header_row(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = HDR
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")


def input_cell(ws, r, c, value, fmt=None):
    cell = ws.cell(r, c, value)
    cell.font = BLUE
    cell.fill = INPUT
    cell.border = THIN
    if fmt:
        cell.number_format = fmt
    return cell


def calc_cell(ws, r, c, value, fmt=None):
    cell = ws.cell(r, c, value)
    cell.font = BLACK
    cell.fill = CALC
    cell.border = THIN
    if fmt:
        cell.number_format = fmt
    return cell


def label(ws, r, c, text, bold=False):
    cell = ws.cell(r, c, text)
    cell.font = BOLD if bold else BLACK
    return cell


def build() -> Path:
    wb = Workbook()

    # ----- Cover -----
    cover = wb.active
    cover.title = "Cover"
    cover["B2"] = "FICO Leveraged Buyout Model"
    cover["B2"].font = TITLE
    cover["B3"] = "Educational / research — SEC-sourced inputs; open in Excel to recalculate"
    cover["B5"] = "Target"
    cover["C5"] = "Fair Isaac Corporation (FICO)"
    cover["B6"] = "CIK"
    cover["C6"] = "0000814547"
    cover["B7"] = "Units"
    cover["C7"] = "$ millions except per share / multiples / rates"
    cover["B9"] = "Sheets"
    for i, name in enumerate(
        [
            "01_Assumptions — blue cells are editable inputs",
            "02_Sources_Uses — purchase price, debt, sponsor equity",
            "03_Operating — hist FY23–25 + 5-yr forecast",
            "04_Debt_Schedule — interest + mandatory amort + cash sweep",
            "05_Returns — exit equity, MoIC, IRR",
            "06_Sensitivity — exit multiple × entry premium grid",
        ],
        start=10,
    ):
        cover[f"B{i}"] = name
    cover["B17"] = "Primary SEC sources"
    cover["B18"] = "FY2025 10-K"
    cover["C18"] = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"
    cover["C18"].hyperlink = cover["C18"].value
    cover["C18"].font = Font(name="Calibri", color="0563C1", underline="single")
    cover["B19"] = "Q3 FY2026 10-Q"
    cover["C19"] = "https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm"
    cover["C19"].hyperlink = cover["C19"].value
    cover["C19"].font = Font(name="Calibri", color="0563C1", underline="single")
    cover["B20"] = "Companyfacts"
    cover["C20"] = "https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json"
    cover["C20"].hyperlink = cover["C20"].value
    cover["C20"].font = Font(name="Calibri", color="0563C1", underline="single")
    cover.column_dimensions["B"].width = 28
    cover.column_dimensions["C"].width = 95

    # ----- Assumptions -----
    a = wb.create_sheet("01_Assumptions")
    a["B2"] = "Assumptions (edit yellow / blue inputs)"
    a["B2"].font = TITLE
    a["B3"] = "All formulas elsewhere link here — change inputs and open in Excel to recalc"
    label(a, 5, 2, "TRANSACTION", bold=True)
    a["B5"].fill = SECTION
    rows = [
        (6, "Current share price ($)", PRICE, "0.00"),
        (7, "Diluted shares (millions)", SHARES_M, "0.000"),
        (8, "Acquisition premium", PREMIUM, "0.0%"),
        (9, "Offer price / share", "=C6*(1+C8)", "0.00"),
        (10, "Equity purchase price", "=C9*C7", "#,##0.0"),
        (11, "LTM EBITDA", EBITDA_LTM, "#,##0.0"),
        (12, "Gross debt (refinance)", GROSS_DEBT, "#,##0.0"),
        (13, "Cash (latest)", CASH, "#,##0.0"),
        (14, "Minimum cash", MIN_CASH, "#,##0.0"),
        (15, "Excess cash to sources", "=MAX(0,C13-C14)", "#,##0.0"),
        (16, "Transaction fees % of offer equity", FEE_PCT, "0.0%"),
        (17, "Transaction fees $", "=C10*C16", "#,##0.0"),
        (18, "Tax rate", TAX, "0.0%"),
        (19, "Hold period (years)", HOLD_YEARS, "0"),
        (20, "Exit EV / EBITDA", EXIT_MULT, "0.0x"),
    ]
    for r, name, val, fmt in rows:
        label(a, r, 2, name)
        if isinstance(val, str) and val.startswith("="):
            calc_cell(a, r, 3, val, fmt)
        else:
            input_cell(a, r, 3, val, fmt)

    label(a, 22, 2, "DEBT STRUCTURE (EBITDA turns)", bold=True)
    a["B22"].fill = SECTION
    debt_inputs = [
        (23, "Revolver turns", REV_TURNS, "0.00"),
        (24, "Term Loan A turns", TLA_TURNS, "0.00"),
        (25, "Term Loan B turns", TLB_TURNS, "0.00"),
        (26, "Senior Notes turns", NOTE_TURNS, "0.00"),
        (27, "Revolver $", "=C23*$C$11", "#,##0.0"),
        (28, "Term Loan A $", "=C24*$C$11", "#,##0.0"),
        (29, "Term Loan B $", "=C25*$C$11", "#,##0.0"),
        (30, "Senior Notes $", "=C26*$C$11", "#,##0.0"),
        (31, "Total new debt", "=SUM(C27:C30)", "#,##0.0"),
    ]
    for r, name, val, fmt in debt_inputs:
        label(a, r, 2, name)
        if isinstance(val, str) and val.startswith("="):
            calc_cell(a, r, 3, val, fmt)
        else:
            input_cell(a, r, 3, val, fmt)

    label(a, 33, 2, "INTEREST / AMORT", bold=True)
    a["B33"].fill = SECTION
    rates = [
        (34, "SOFR", SOFR, "0.00%"),
        (35, "Revolver spread", REV_SPR, "0.00%"),
        (36, "TLA spread", TLA_SPR, "0.00%"),
        (37, "TLB spread", TLB_SPR, "0.00%"),
        (38, "Senior Notes coupon", NOTE_RATE, "0.00%"),
        (39, "TLA annual amort % of initial", TLA_AMORT, "0.0%"),
        (40, "Revolver rate", "=C34+C35", "0.00%"),
        (41, "TLA rate", "=C34+C36", "0.00%"),
        (42, "TLB rate", "=C34+C37", "0.00%"),
    ]
    for r, name, val, fmt in rates:
        label(a, r, 2, name)
        if isinstance(val, str) and val.startswith("="):
            calc_cell(a, r, 3, val, fmt)
        else:
            input_cell(a, r, 3, val, fmt)

    label(a, 44, 2, "OPERATING FORECAST DRIVERS", bold=True)
    a["B44"].fill = SECTION
    a["B45"] = "Metric"
    years = [2026, 2027, 2028, 2029, 2030]
    for i, y in enumerate(years):
        a.cell(45, 3 + i, y)
    style_header_row(a, 45, 7)
    a["B45"].fill = HDR
    a["B45"].font = HDR_FONT

    label(a, 46, 2, "Revenue growth")
    for i, g in enumerate(GROWTHS):
        input_cell(a, 46, 3 + i, g, "0.0%")
    label(a, 47, 2, "EBITDA margin")
    for i, m in enumerate(EBITDA_MARGINS):
        input_cell(a, 47, 3 + i, m, "0.0%")
    label(a, 48, 2, "CapEx % revenue")
    for i, m in enumerate(CAPEX_PCT):
        input_cell(a, 48, 3 + i, m, "0.0%")
    label(a, 49, 2, "NWC % revenue")
    for i, m in enumerate(NWC_PCT):
        input_cell(a, 49, 3 + i, m, "0.0%")

    label(a, 51, 2, "HISTORICAL ANCHORS ($m)", bold=True)
    a["B51"].fill = SECTION
    for i, y in enumerate([2023, 2024, 2025]):
        a.cell(52, 3 + i, y)
    style_header_row(a, 52, 5)
    a["B52"] = "Line"
    a["B52"].fill = HDR
    a["B52"].font = HDR_FONT
    hist_lines = [
        (53, "Revenue", REV),
        (54, "COGS", COGS),
        (55, "R&D", RD),
        (56, "SG&A", SGA),
        (57, "D&A", DA),
    ]
    for r, name, series in hist_lines:
        label(a, r, 2, name)
        for i, y in enumerate([2023, 2024, 2025]):
            input_cell(a, r, 3 + i, series[y], "#,##0.0")
    label(a, 58, 2, "EBITDA (hist)")
    for i, col in enumerate(["C", "D", "E"]):
        # Rev - COGS - RD - SGA + (already opex) => OpInc+DA approx via Rev-COGS-RD-SGA
        calc_cell(
            a,
            58,
            3 + i,
            f"={col}53-{col}54-{col}55-{col}56",
            "#,##0.0",
        )

    a.column_dimensions["B"].width = 38
    for col in range(3, 8):
        a.column_dimensions[get_column_letter(col)].width = 14

    # ----- Sources & Uses -----
    su = wb.create_sheet("02_Sources_Uses")
    su["B2"] = "Sources & Uses"
    su["B2"].font = TITLE
    su["B4"] = "USES"
    su["B4"].fill = SECTION
    su["B4"].font = BOLD
    su["E4"] = "SOURCES"
    su["E4"].fill = SECTION
    su["E4"].font = BOLD

    uses = [
        (5, "Equity purchase price", "='01_Assumptions'!C10"),
        (6, "Refinance existing debt", "='01_Assumptions'!C12"),
        (7, "Transaction fees", "='01_Assumptions'!C17"),
        (8, "Total Uses", "=SUM(C5:C7)"),
    ]
    for r, name, f in uses:
        label(su, r, 2, name, bold=(r == 8))
        calc_cell(su, r, 3, f, "#,##0.0")

    sources = [
        (5, "Excess cash", "='01_Assumptions'!C15"),
        (6, "Revolver", "='01_Assumptions'!C27"),
        (7, "Term Loan A", "='01_Assumptions'!C28"),
        (8, "Term Loan B", "='01_Assumptions'!C29"),
        (9, "Senior Notes", "='01_Assumptions'!C30"),
        (10, "Sponsor equity (plug)", "=C8-SUM(F5:F9)"),  # wait - Total Uses is C8
    ]
    # Fix sponsor equity plug: Total Uses - other sources
    sources = [
        (5, "Excess cash", "='01_Assumptions'!C15"),
        (6, "Revolver", "='01_Assumptions'!C27"),
        (7, "Term Loan A", "='01_Assumptions'!C28"),
        (8, "Term Loan B", "='01_Assumptions'!C29"),
        (9, "Senior Notes", "='01_Assumptions'!C30"),
        (10, "Sponsor equity (plug)", "=C8-SUM(F5:F9)"),
        (11, "Total Sources", "=SUM(F5:F10)"),
    ]
    for r, name, f in sources:
        label(su, r, 5, name, bold=(r >= 10))
        calc_cell(su, r, 6, f, "#,##0.0")

    label(su, 13, 2, "Entry EV", bold=True)
    calc_cell(su, 13, 3, "='01_Assumptions'!C10+'01_Assumptions'!C12-'01_Assumptions'!C13", "#,##0.0")
    label(su, 14, 2, "Entry EV / LTM EBITDA")
    calc_cell(su, 14, 3, "=C13/'01_Assumptions'!C11", "0.00x")
    label(su, 15, 2, "Sources − Uses check")
    calc_cell(su, 15, 3, "=F11-C8", "#,##0.0")
    label(su, 16, 2, "Debt / EBITDA (entry)")
    calc_cell(su, 16, 3, "='01_Assumptions'!C31/'01_Assumptions'!C11", "0.00x")
    label(su, 17, 2, "Equity contribution % of sources")
    calc_cell(su, 17, 3, "=F10/F11", "0.0%")

    for col, w in [("B", 32), ("C", 14), ("E", 28), ("F", 14)]:
        su.column_dimensions[col].width = w

    # ----- Operating -----
    op = wb.create_sheet("03_Operating")
    op["B2"] = "Operating Model"
    op["B2"].font = TITLE
    # Columns: C=2023 D=2024 E=2025 F=2026 ... J=2030
    hist_yrs = [2023, 2024, 2025]
    fcast_yrs = [2026, 2027, 2028, 2029, 2030]
    all_yrs = hist_yrs + fcast_yrs
    op["B4"] = "Fiscal year"
    for i, y in enumerate(all_yrs):
        op.cell(4, 3 + i, y)
    style_header_row(op, 4, 2 + len(all_yrs))
    op["B4"].fill = HDR
    op["B4"].font = HDR_FONT

    # Revenue
    label(op, 5, 2, "Revenue")
    for i, y in enumerate(hist_yrs):
        calc_cell(op, 5, 3 + i, f"='01_Assumptions'!{get_column_letter(3+i)}53", "#,##0.0")
    # F26 = E25*(1+growth)
    calc_cell(op, 5, 6, "=E5*(1+'01_Assumptions'!C46)", "#,##0.0")
    calc_cell(op, 5, 7, "=F5*(1+'01_Assumptions'!D46)", "#,##0.0")
    calc_cell(op, 5, 8, "=G5*(1+'01_Assumptions'!E46)", "#,##0.0")
    calc_cell(op, 5, 9, "=H5*(1+'01_Assumptions'!F46)", "#,##0.0")
    calc_cell(op, 5, 10, "=I5*(1+'01_Assumptions'!G46)", "#,##0.0")

    label(op, 6, 2, "Revenue growth")
    for i, col in enumerate(["D", "E", "F", "G", "H", "I", "J"]):
        prev = chr(ord(col) - 1)
        calc_cell(op, 6, 4 + i, f"={col}5/{prev}5-1", "0.0%")

    label(op, 7, 2, "EBITDA")
    for i, y in enumerate(hist_yrs):
        calc_cell(op, 7, 3 + i, f"='01_Assumptions'!{get_column_letter(3+i)}58", "#,##0.0")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        gcol = get_column_letter(3 + i)  # C..G on assumptions row 47
        calc_cell(op, 7, 6 + i, f"={col}5*'01_Assumptions'!{gcol}47", "#,##0.0")

    label(op, 8, 2, "EBITDA margin")
    for i, col in enumerate(["C", "D", "E", "F", "G", "H", "I", "J"]):
        calc_cell(op, 8, 3 + i, f"={col}7/{col}5", "0.0%")

    label(op, 9, 2, "D&A")
    for i, y in enumerate(hist_yrs):
        calc_cell(op, 9, 3 + i, f"='01_Assumptions'!{get_column_letter(3+i)}57", "#,##0.0")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(op, 9, 6 + i, f"={col}5*0.008", "#,##0.0")  # ~hist D&A/Rev

    label(op, 10, 2, "EBIT")
    for i, col in enumerate(["C", "D", "E", "F", "G", "H", "I", "J"]):
        calc_cell(op, 10, 3 + i, f"={col}7-{col}9", "#,##0.0")

    label(op, 11, 2, "Interest expense")
    # hist hardcoded from SEC; forecast from debt schedule
    for i, v in enumerate([95.546, 105.638, 133.647]):
        input_cell(op, 11, 3 + i, v, "#,##0.0")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(op, 11, 6 + i, f"='04_Debt_Schedule'!{col}20", "#,##0.0")

    label(op, 12, 2, "EBT")
    for i, col in enumerate(["C", "D", "E", "F", "G", "H", "I", "J"]):
        calc_cell(op, 12, 3 + i, f"={col}10-{col}11", "#,##0.0")

    label(op, 13, 2, "Taxes")
    for i, col in enumerate(["C", "D", "E", "F", "G", "H", "I", "J"]):
        calc_cell(op, 13, 3 + i, f"=MAX(0,{col}12)*'01_Assumptions'!$C$18", "#,##0.0")

    label(op, 14, 2, "Net income")
    for i, col in enumerate(["C", "D", "E", "F", "G", "H", "I", "J"]):
        calc_cell(op, 14, 3 + i, f"={col}12-{col}13", "#,##0.0")

    label(op, 16, 2, "CapEx")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        gcol = get_column_letter(3 + i)
        calc_cell(op, 16, 6 + i, f"={col}5*'01_Assumptions'!{gcol}48", "#,##0.0")

    label(op, 17, 2, "NWC")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        gcol = get_column_letter(3 + i)
        calc_cell(op, 17, 6 + i, f"={col}5*'01_Assumptions'!{gcol}49", "#,##0.0")
    # Opening NWC at close ≈ FY25 Rev * first NWC%
    label(op, 18, 2, "ΔNWC (use of cash)")
    calc_cell(op, 18, 6, "=F17-E5*'01_Assumptions'!C49", "#,##0.0")
    for i, col in enumerate(["G", "H", "I", "J"]):
        prev = chr(ord(col) - 1)
        calc_cell(op, 18, 7 + i, f"={col}17-{prev}17", "#,##0.0")

    label(op, 19, 2, "Free cash flow (pre-debt)", bold=True)
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(op, 19, 6 + i, f"={col}14+{col}9-{col}16-{col}18", "#,##0.0")

    label(op, 20, 2, "Cash available for debt paydown", bold=True)
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        # FCF after interest already in NI; add back non-cash; interest already deducted
        calc_cell(op, 20, 6 + i, f"=MAX(0,{col}19)", "#,##0.0")

    op.column_dimensions["B"].width = 32
    for col in range(3, 11):
        op.column_dimensions[get_column_letter(col)].width = 12

    # ----- Debt Schedule -----
    d = wb.create_sheet("04_Debt_Schedule")
    d["B2"] = "Debt Schedule"
    d["B2"].font = TITLE
    d["B4"] = "Item"
    for i, y in enumerate(fcast_yrs):
        d.cell(4, 6 + i, y)  # F-J
    style_header_row(d, 4, 10)
    d["B4"].fill = HDR
    d["B4"].font = HDR_FONT

    # Opening balances at close (end FY25 / start FY26)
    label(d, 5, 2, "Revolver — opening")
    calc_cell(d, 5, 6, "='01_Assumptions'!C27", "#,##0.0")
    for i, col in enumerate(["G", "H", "I", "J"]):
        prev = chr(ord(col) - 1)
        calc_cell(d, 5, 7 + i, f"={prev}8", "#,##0.0")  # prior ending

    label(d, 6, 2, "TLA — opening")
    calc_cell(d, 6, 6, "='01_Assumptions'!C28", "#,##0.0")
    for i, col in enumerate(["G", "H", "I", "J"]):
        prev = chr(ord(col) - 1)
        calc_cell(d, 6, 7 + i, f"={prev}12", "#,##0.0")

    label(d, 7, 2, "TLB — opening")
    calc_cell(d, 7, 6, "='01_Assumptions'!C29", "#,##0.0")
    for i, col in enumerate(["G", "H", "I", "J"]):
        prev = chr(ord(col) - 1)
        calc_cell(d, 7, 7 + i, f"={prev}15", "#,##0.0")

    label(d, 8, 2, "Notes — opening / ending (bullet)")
    # Will overwrite — structure:
    # Row 5-8 openings already set for F; for revolver ending need sweep rows

    # Rebuild debt schedule more carefully
    # Clear and rewrite a cleaner block
    for r in range(5, 30):
        for c in range(2, 11):
            d.cell(r, c).value = None
            d.cell(r, c).fill = PatternFill()
            d.cell(r, c).font = BLACK

    label(d, 5, 2, "Opening balances", bold=True)
    d["B5"].fill = SECTION
    label(d, 6, 2, "Revolver")
    label(d, 7, 2, "Term Loan A")
    label(d, 8, 2, "Term Loan B")
    label(d, 9, 2, "Senior Notes")
    label(d, 10, 2, "Total debt opening")

    calc_cell(d, 6, 6, "='01_Assumptions'!C27", "#,##0.0")
    calc_cell(d, 7, 6, "='01_Assumptions'!C28", "#,##0.0")
    calc_cell(d, 8, 6, "='01_Assumptions'!C29", "#,##0.0")
    calc_cell(d, 9, 6, "='01_Assumptions'!C30", "#,##0.0")
    calc_cell(d, 10, 6, "=SUM(F6:F9)", "#,##0.0")
    for i, col in enumerate(["G", "H", "I", "J"]):
        prev = chr(ord(col) - 1)
        calc_cell(d, 6, 7 + i, f"={prev}16", "#,##0.0")  # ending revolver
        calc_cell(d, 7, 7 + i, f"={prev}17", "#,##0.0")
        calc_cell(d, 8, 7 + i, f"={prev}18", "#,##0.0")
        calc_cell(d, 9, 7 + i, f"={prev}19", "#,##0.0")
        calc_cell(d, 10, 7 + i, f"=SUM({col}6:{col}9)", "#,##0.0")

    label(d, 12, 2, "Mandatory amortization", bold=True)
    d["B12"].fill = SECTION
    label(d, 13, 2, "TLA mandatory")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 13, 6 + i, f"=MIN({col}7,'01_Assumptions'!$C$28*'01_Assumptions'!$C$39)", "#,##0.0")
    label(d, 14, 2, "Notes / TLB mandatory")
    for i in range(5):
        input_cell(d, 14, 6 + i, 0, "#,##0.0")

    label(d, 15, 2, "Cash interest", bold=True)
    d["B15"].fill = SECTION
    # Interest on average opening (simplified: opening * rate)
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        label(d, 16, 2, "Interest expense")
        calc_cell(
            d,
            16,
            6 + i,
            f"={col}6*'01_Assumptions'!$C$40+{col}7*'01_Assumptions'!$C$41+{col}8*'01_Assumptions'!$C$42+{col}9*'01_Assumptions'!$C$38",
            "#,##0.0",
        )
    # Row 20 used by operating sheet — put total interest there
    label(d, 20, 2, "Total interest (link to Op)", bold=True)
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 20, 6 + i, f"={col}16", "#,##0.0")

    label(d, 22, 2, "Cash flow sweep (optional paydown)", bold=True)
    d["B22"].fill = SECTION
    label(d, 23, 2, "Cash for paydown (from Op)")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 23, 6 + i, f"='03_Operating'!{col}20", "#,##0.0")
    label(d, 24, 2, "After mandatory")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 24, 6 + i, f"=MAX(0,{col}23-{col}13-{col}14)", "#,##0.0")
    # Sweep order: Revolver → TLA → TLB → Notes
    label(d, 25, 2, "Sweep to Revolver")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 25, 6 + i, f"=MIN({col}6,{col}24)", "#,##0.0")
    label(d, 26, 2, "Sweep to TLA")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 26, 6 + i, f"=MIN(MAX(0,{col}7-{col}13),MAX(0,{col}24-{col}25))", "#,##0.0")
    label(d, 27, 2, "Sweep to TLB")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 27, 6 + i, f"=MIN({col}8,MAX(0,{col}24-{col}25-{col}26))", "#,##0.0")
    label(d, 28, 2, "Sweep to Notes")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 28, 6 + i, f"=MIN({col}9,MAX(0,{col}24-{col}25-{col}26-{col}27))", "#,##0.0")

    label(d, 30, 2, "Ending balances", bold=True)
    d["B30"].fill = SECTION
    # Put endings on rows 16-19? Conflict with interest on 16.
    # Use rows 32-36 for endings; link openings G-J from these via earlier formulas.
    # Fix: change opening links to rows 32-35 endings

    label(d, 32, 2, "Revolver ending")
    label(d, 33, 2, "TLA ending")
    label(d, 34, 2, "TLB ending")
    label(d, 35, 2, "Notes ending")
    label(d, 36, 2, "Total debt ending")
    for i, col in enumerate(["F", "G", "H", "I", "J"]):
        calc_cell(d, 32, 6 + i, f"=MAX(0,{col}6-{col}25)", "#,##0.0")
        calc_cell(d, 33, 6 + i, f"=MAX(0,{col}7-{col}13-{col}26)", "#,##0.0")
        calc_cell(d, 34, 6 + i, f"=MAX(0,{col}8-{col}27)", "#,##0.0")
        calc_cell(d, 35, 6 + i, f"=MAX(0,{col}9-{col}28)", "#,##0.0")
        calc_cell(d, 36, 6 + i, f"=SUM({col}32:{col}35)", "#,##0.0")

    # Fix opening roll-forward to use ending rows 32-35
    for i, col in enumerate(["G", "H", "I", "J"]):
        prev = chr(ord(col) - 1)
        d.cell(6, 7 + i).value = f"={prev}32"
        d.cell(7, 7 + i).value = f"={prev}33"
        d.cell(8, 7 + i).value = f"={prev}34"
        d.cell(9, 7 + i).value = f"={prev}35"

    d.column_dimensions["B"].width = 34
    for col in range(6, 11):
        d.column_dimensions[get_column_letter(col)].width = 12

    # ----- Returns -----
    r = wb.create_sheet("05_Returns")
    r["B2"] = "Returns Analysis"
    r["B2"].font = TITLE
    label(r, 4, 2, "Exit year EBITDA (Y5)")
    calc_cell(r, 4, 3, "='03_Operating'!J7", "#,##0.0")
    label(r, 5, 2, "Exit EV / EBITDA")
    calc_cell(r, 5, 3, "='01_Assumptions'!C20", "0.0x")
    label(r, 6, 2, "Exit enterprise value")
    calc_cell(r, 6, 3, "=C4*C5", "#,##0.0")
    label(r, 7, 2, "Exit net debt (debt ending − min cash)")
    calc_cell(r, 7, 3, "='04_Debt_Schedule'!J36-'01_Assumptions'!C14", "#,##0.0")
    label(r, 8, 2, "Exit equity value", bold=True)
    calc_cell(r, 8, 3, "=C6-C7", "#,##0.0")
    label(r, 10, 2, "Sponsor equity at entry (check)")
    calc_cell(r, 10, 3, "='02_Sources_Uses'!F10", "#,##0.0")
    label(r, 11, 2, "MoIC", bold=True)
    calc_cell(r, 11, 3, "=IF(C10>0,C8/C10,0)", "0.00x")
    label(r, 12, 2, "IRR (annual)", bold=True)
    calc_cell(r, 12, 3, "=IF(C10>0,(C8/C10)^(1/'01_Assumptions'!C19)-1,0)", "0.0%")

    label(r, 14, 2, "Cash flow to sponsor")
    r["B15"] = "Year 0 (entry)"
    calc_cell(r, 15, 3, "=-C10", "#,##0.0")
    r["B16"] = "Year 5 (exit)"
    calc_cell(r, 16, 3, "=C8", "#,##0.0")
    r["B17"] = "Excel IRR (same as above if no interim CF)"
    calc_cell(r, 17, 3, "=C12", "0.0%")

    label(r, 19, 2, "How to use", bold=True)
    r["B20"] = "1. Edit yellow cells on 01_Assumptions (premium, leverage turns, exit multiple, growth)."
    r["B21"] = "2. Open this file in Excel / Google Sheets so formulas calculate."
    r["B22"] = "3. Read MoIC and IRR on this sheet; stress on 06_Sensitivity."

    r.column_dimensions["B"].width = 48
    r.column_dimensions["C"].width = 16

    # ----- Sensitivity -----
    s = wb.create_sheet("06_Sensitivity")
    s["B2"] = "Sensitivity — Exit MoIC"
    s["B2"].font = TITLE
    s["B3"] = "Rows = exit EV/EBITDA; Cols = acquisition premium"
    s["B5"] = "MoIC"
    premiums = [0.10, 0.15, 0.20, 0.25, 0.30, 0.40]
    exits = [14, 16, 18, 20, 22, 25]
    for i, p in enumerate(premiums):
        input_cell(s, 5, 3 + i, p, "0%")
    for i, e in enumerate(exits):
        input_cell(s, 6 + i, 2, e, "0")
    # Approximate sensitivity using algebraic MoIC vs base:
    # MoIC ≈ base_MoIC * (exit_mult/base_exit) / ((1+prem)/(1+base_prem))  — rough
    # Better: show formula guidance + computed grid from linked base equity scaling
    label(s, 13, 2, "Base case MoIC (from Returns)")
    calc_cell(s, 13, 3, "='05_Returns'!C11", "0.00x")
    label(s, 14, 2, "Base premium")
    calc_cell(s, 14, 3, "='01_Assumptions'!C8", "0.0%")
    label(s, 15, 2, "Base exit multiple")
    calc_cell(s, 15, 3, "='01_Assumptions'!C20", "0.0x")

    s["B6"] = "Exit \\ Prem"
    for i, e in enumerate(exits):
        for j, p in enumerate(premiums):
            # MoIC_sens ≈ MoIC_base * (e/exit_base) * (1+prem_base)/(1+p)
            # Equity entry scales with (1+p); exit EV scales with e — net debt approx constant
            calc_cell(
                s,
                6 + i,
                3 + j,
                f"=$C$13*({e}/$C$15)*((1+$C$14)/(1+{p}))",
                "0.00x",
            )

    label(s, 17, 2, "IRR grid (same axes)", bold=True)
    s["B18"] = "IRR"
    for i, p in enumerate(premiums):
        input_cell(s, 18, 3 + i, p, "0%")
    for i, e in enumerate(exits):
        input_cell(s, 19 + i, 2, e, "0")
        for j, p in enumerate(premiums):
            # IRR = MoIC^(1/n)-1
            calc_cell(
                s,
                19 + i,
                3 + j,
                f"=({get_column_letter(3+j)}{6+i})^(1/'01_Assumptions'!$C$19)-1",
                "0.0%",
            )

    s.column_dimensions["B"].width = 16
    for col in range(3, 9):
        s.column_dimensions[get_column_letter(col)].width = 10

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    wb.save(REPO_ROOT_COPY)
    print(f"Wrote {OUT}")
    print(f"Wrote {REPO_ROOT_COPY}")
    return OUT


if __name__ == "__main__":
    build()
