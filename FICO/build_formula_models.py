#!/usr/bin/env python3
"""
Build FICO 3-statement + DCF where ONLY inputs are typed numbers.
All calculated line items are Excel formulas that depend on other cells.

Inputs (yellow/blue): assumptions, historical hardcodes, forecast drivers.
Everything else: = formulas (taxes, UFCF, TV, XNPV, revenue roll-forward, etc.).
"""

from __future__ import annotations

import json
import shutil
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parents[1]
FICO_DIR = Path(__file__).resolve().parent
ORIG_3 = ROOT / "DCF resources" / "originals" / "CFI_3-Statement-Model-Complete.xlsx"
ORIG_DCF = ROOT / "DCF resources" / "originals" / "CFI_DCF-Model.xlsx"
CACHE = ROOT / "fico-valuation" / "data" / "fico_companyfacts.json"
OUT_3S = FICO_DIR / "3S_FICO.xlsx"
OUT_DCF = FICO_DIR / "DCF_FICO.xlsx"
OUT_JOINED = FICO_DIR / "FICO_3S_DCF_Joined.xlsx"
OUT_MATH = FICO_DIR / "FORMULA_MATH.md"

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
NOTE = Font(name="Calibri", size=10, italic=True, color="C00000")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True)
LINK_FONT = Font(name="Calibri", color="0563C1", underline="single")

HIST_COLS = list("EFGHI")
HIST_YEARS = [2021, 2022, 2023, 2024, 2025]
FC_COLS = list("JKLMN")
DCF_COLS = list("EFGHI")
GROWTHS = [0.12, 0.10, 0.09, 0.08, 0.07]

FY25_OPINC = 924_850
FY25_DA = 14_952
FY25_CAPEX = 39_407
FY25_SGA = 513_028
FY25_RD = 188_347
FY25_REV = 1_990_869
FY25_COGS = 353_722
FY25_AR = 529_148
FY25_AP = 32_315
FY25_DEBT = 3_055_691
FY25_INTEREST = 133_647
FY25_TAX = 150_649
FY25_EBT = 802_595
TAX_RATE = round(FY25_TAX / FY25_EBT, 4)
COGS_PCT = round(FY25_COGS / FY25_REV, 4)
WACC = 0.096
G_PERP = 0.03
EV_EBITDA = 22.0
PRICE = 1046.23
SHARES = 21_597.635
DEBT_10Q = 5_582_389
CASH_10Q = 304_537
CAPEX_GROWTH = 0.08
FILING = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"


def is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def set_input(ws, coord, value, fmt=None):
    """Typed input only — never used for calculated lines."""
    cell = ws[coord]
    if is_formula(cell.value):
        raise RuntimeError(f"Refusing to replace formula at {coord} with input: {cell.value}")
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL
    if fmt:
        cell.number_format = fmt


def set_formula(ws, coord, formula, *, linked=False, fmt=None):
    """Calculated line — Excel dependency."""
    cell = ws[coord]
    cell.value = formula
    cell.fill = LINK_FILL if linked else PatternFill()
    cell.font = Font(name="Calibri", color="000000")
    if fmt:
        cell.number_format = fmt


def widen(ws, width=15):
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 12, 36)
    for col in range(4, 16):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 10, width)


def annual(facts, concept):
    usgaap = facts["facts"].get("us-gaap", {})
    if concept not in usgaap or "USD" not in usgaap[concept]["units"]:
        return {}
    best = {}
    for it in usgaap[concept]["units"]["USD"]:
        if it.get("form") not in ("10-K", "10-K/A") or it.get("fp") != "FY":
            continue
        end = it.get("end") or ""
        if not end.endswith("-09-30"):
            continue
        y = int(end[:4])
        if 2020 <= y <= 2025:
            prev = best.get(y)
            if prev is None or it.get("filed", "") >= prev["filed"]:
                best[y] = {"val": it["val"], "filed": it.get("filed", "")}
    return {y: best[y]["val"] / 1000.0 for y in best}


def load_series():
    facts = json.loads(CACHE.read_text())
    s = {k: annual(facts, c) for k, c in [
        ("revenue", "RevenueFromContractWithCustomerExcludingAssessedTax"),
        ("cogs", "CostOfRevenue"),
        ("sga", "SellingGeneralAndAdministrativeExpense"),
        ("rd", "ResearchAndDevelopmentExpense"),
        ("opinc", "OperatingIncomeLoss"),
        ("interest", "InterestExpense"),
        ("tax", "IncomeTaxExpenseBenefit"),
        ("ni", "NetIncomeLoss"),
        ("cash", "CashAndCashEquivalentsAtCarryingValue"),
        ("ar", "AccountsReceivableNetCurrent"),
        ("ppe", "PropertyPlantAndEquipmentNet"),
        ("ap", "AccountsPayableCurrent"),
        ("re", "RetainedEarningsAccumulatedDeficit"),
        ("da", "DepreciationDepletionAndAmortization"),
        ("debt_lt", "LongTermDebtNoncurrent"),
        ("debt_cur", "LongTermDebtCurrent"),
        ("capex_ppe", "PaymentsToAcquirePropertyPlantAndEquipment"),
    ]}
    s["interest"].update({2023: 95_546, 2024: 105_638, 2025: 133_647})
    debt = {y: s["debt_lt"].get(y, 0) + s["debt_cur"].get(y, 0) for y in range(2020, 2026)}
    debt[2024] = 2_209_021
    debt[2025] = 3_055_691
    s["debt"] = debt
    s["capex"] = dict(s["capex_ppe"])
    s["capex"].update({2023: 4_237, 2024: 25_551, 2025: 39_407})
    return s


def inject_three_statement(ws, s):
    """Historical inputs typed; forecast drivers typed; statement mechanics remain formulas."""
    # Year start input; F2:N2 stay as =prior+1 formulas (dependency)
    set_input(ws, "E2", HIST_YEARS[0], fmt="0")

    for col, y in zip(HIST_COLS, HIST_YEARS):
        set_input(ws, f"{col}24", round(s["revenue"][y], 1))
        set_input(ws, f"{col}25", round(s["cogs"][y], 1))
        set_input(ws, f"{col}28", round(s["sga"][y], 1))
        set_input(ws, f"{col}29", round(s["rd"][y], 1))
        set_input(ws, f"{col}30", round(s["da"][y], 1))
        set_input(ws, f"{col}31", round(s["interest"][y], 1))
        set_input(ws, f"{col}35", round(s["tax"][y], 1))
        set_input(ws, f"{col}41", round(s["cash"][y], 1))
        set_input(ws, f"{col}42", round(s["ar"][y], 1))
        set_input(ws, f"{col}43", 0)
        set_input(ws, f"{col}44", round(s["ppe"][y], 1))
        set_input(ws, f"{col}48", round(s["ap"][y], 1))
        set_input(ws, f"{col}49", round(s["debt"][y], 1))
        assets_simple = s["cash"][y] + s["ar"][y] + s["ppe"][y]
        liab_simple = s["ap"][y] + s["debt"][y]
        equity_plug = assets_simple - liab_simple
        re = s["re"][y]
        set_input(ws, f"{col}53", round(re, 1))
        set_input(ws, f"{col}52", round(equity_plug - re, 1))

        if y == HIST_YEARS[0]:
            prev_nwc = s["ar"][y - 1] - s["ap"][y - 1]
            prev_debt = s["debt"].get(y - 1, 739_435)
            prev_cash = s["cash"][y - 1]
            open_ppe = s["ppe"][y - 1]
        else:
            prev_nwc = s["ar"][y - 1] - s["ap"][y - 1]
            prev_debt = s["debt"][y - 1]
            prev_cash = s["cash"][y - 1]
            open_ppe = s["ppe"][y - 1]
        nwc = s["ar"][y] - s["ap"][y]
        dnwc = nwc - prev_nwc
        debt_issue = s["debt"][y] - prev_debt
        capex = abs(s["capex"][y])
        set_input(ws, f"{col}64", round(dnwc, 1))
        set_input(ws, f"{col}68", round(capex, 1))
        set_input(ws, f"{col}72", round(debt_issue, 1))
        cfo_proxy = s["ni"][y] + s["da"][y] - dnwc
        equity_issue = (s["cash"][y] - prev_cash) - (cfo_proxy - capex + debt_issue)
        set_input(ws, f"{col}73", round(equity_issue, 1))
        if col == "E":
            set_input(ws, "D78", round(s["cash"][y - 1], 1))
        else:
            set_input(ws, f"{col}77", round(prev_cash, 1))
        set_input(ws, f"{col}85", round(s["ar"][y], 1))
        set_input(ws, f"{col}86", 0)
        set_input(ws, f"{col}87", round(s["ap"][y], 1))
        da = s["da"][y]
        set_input(ws, f"{col}92", round(open_ppe, 1))
        set_input(ws, f"{col}93", round(s["ppe"][y] - open_ppe + da, 1))
        set_input(ws, f"{col}94", round(da, 1))
        set_input(ws, f"{col}98", round(prev_debt, 1))
        set_input(ws, f"{col}99", round(debt_issue, 1))
        set_input(ws, f"{col}101", round(s["interest"][y], 1))

    # Forecast DRIVER inputs only — IS/BS/CF forecast rows stay native CFI formulas
    int_rate = round(FY25_INTEREST / FY25_DEBT, 4)
    da_pct = min(round(FY25_DA / max(s["ppe"][2024], 1), 4), 0.35)
    ar_days = round(FY25_AR / FY25_REV * 365, 1)
    ap_days = round(FY25_AP / FY25_COGS * 365, 1)
    for i, col in enumerate(FC_COLS):
        g = GROWTHS[i]
        set_input(ws, f"{col}7", g, fmt="0.0%")
        set_input(ws, f"{col}8", COGS_PCT, fmt="0.00%")
        # SG&A / R&D absolute drivers — feed Salaries/Rent formulas (=J9, =J10)
        set_input(ws, f"{col}9", round(FY25_SGA * ((1 + g) ** (i + 1)), 1))
        set_input(ws, f"{col}10", round(FY25_RD * ((1 + g) ** (i + 1)), 1))
        set_input(ws, f"{col}11", da_pct, fmt="0.00%")
        set_input(ws, f"{col}12", int_rate, fmt="0.00%")
        set_input(ws, f"{col}13", TAX_RATE, fmt="0.00%")
        set_input(ws, f"{col}15", ar_days)
        set_input(ws, f"{col}16", 0)
        set_input(ws, f"{col}17", ap_days)
        set_input(ws, f"{col}18", round(FY25_CAPEX * ((1 + CAPEX_GROWTH) ** (i + 1)), 1))
        set_input(ws, f"{col}19", 0)
        set_input(ws, f"{col}20", 0)

    # Clarify mapping labels (text only)
    ws["B24"] = "Revenue (FICO: Revenues)"
    ws["B25"] = "COGS (FICO: Cost of revenues)"
    ws["B28"] = "SG&A ← Salaries (depends on row 9)"
    ws["B29"] = "R&D ← Rent/Overhead (depends on row 10)"
    ws["B131"] = (
        "Yellow = inputs. All Gross Profit / totals / forecast Revenue(=prior*(1+g)) / "
        "COGS(=Rev*COGS%) / Capex(=row18) / ΔNWC(=schedule) are Excel formulas."
    )
    ws["B131"].font = NOTE
    widen(ws)

    # Ensure key forecast formulas still present (template should already have them)
    assert is_formula(ws["J24"].value)  # =I24*(1+J7)
    assert is_formula(ws["J25"].value)  # =J24*J8
    assert is_formula(ws["J26"].value)  # =J24-J25
    assert is_formula(ws["J28"].value)  # =J9
    assert is_formula(ws["I26"].value)


def setup_dcf_assumptions(ws):
    set_input(ws, "D5", TAX_RATE, fmt="0.00%")
    set_input(ws, "D6", WACC, fmt="0.0%")
    set_input(ws, "D7", G_PERP, fmt="0.0%")
    set_input(ws, "D8", EV_EBITDA, fmt="0.0")
    set_input(ws, "D9", datetime(2025, 9, 30), fmt="yyyy-mm-dd")
    set_input(ws, "D10", datetime(2025, 9, 30), fmt="yyyy-mm-dd")
    set_input(ws, "D11", PRICE, fmt="#,##0.00")
    set_input(ws, "D12", SHARES, fmt="#,##0.000")
    set_input(ws, "D13", DEBT_10Q, fmt="#,##0")
    set_input(ws, "D14", CASH_10Q, fmt="#,##0")
    set_input(ws, "D15", FY25_CAPEX, fmt="#,##0")

    # Extra drivers used by formulas
    ws["B16"] = "Capex growth"
    set_input(ws, "D16", CAPEX_GROWTH, fmt="0.0%")
    ws["B3"] = "FY25 OpInc (base)"
    set_input(ws, "D3", FY25_OPINC, fmt="#,##0")
    ws["C3"] = "FY25 D&A"
    set_input(ws, "E3", FY25_DA, fmt="#,##0")


def restore_dcf_structure_formulas(ws):
    """Put back native CFI dependency formulas (September FYE, period starts at 1)."""
    # Time periods: E19 input=1; F19:I19 = prior+1
    set_input(ws, "E19", 1, fmt="0")
    set_formula(ws, "F19", "=E19+1")
    set_formula(ws, "G19", "=F19+1")
    set_formula(ws, "H19", "=G19+1")
    set_formula(ws, "I19", "=H19+1")

    set_formula(ws, "D18", "=D9", fmt="yyyy-mm-dd")
    for col in DCF_COLS:
        set_formula(ws, f"{col}18", f"=DATE(YEAR($D$10)+{col}19,9,30)", fmt="yyyy-mm-dd")
        set_formula(ws, f"{col}17", f"=YEAR({col}18)", fmt="0")
    set_formula(ws, "J18", "=I18", fmt="yyyy-mm-dd")

    set_formula(ws, "E20", "=YEARFRAC(D18,E18)", fmt="0.00")
    set_formula(ws, "F20", "=YEARFRAC(E18,F18)", fmt="0.00")
    set_formula(ws, "G20", "=YEARFRAC(F18,G18)", fmt="0.00")
    set_formula(ws, "H20", "=YEARFRAC(G18,H18)", fmt="0.00")
    set_formula(ws, "I20", "=YEARFRAC(H18,I18)", fmt="0.00")

    # Taxes, Capex (grows with D16 & period), UFCF — all formulas
    for col in DCF_COLS:
        set_formula(ws, f"{col}22", f"={col}21*$D$5", fmt="#,##0")
        set_formula(ws, f"{col}24", f"=$D$15*(1+$D$16)^{col}19", fmt="#,##0")
        set_formula(ws, f"{col}26", f"={col}21-{col}22+{col}23-{col}24-{col}25", fmt="#,##0")
        set_formula(ws, f"{col}28", f"=({col}27+{col}26)*{col}20", fmt="#,##0")
        set_formula(ws, f"{col}29", f"=({col}27+{col}26)*{col}20", fmt="#,##0")

    set_formula(ws, "D27", "=-I35", fmt="#,##0")
    set_formula(ws, "J27", "=N20", fmt="#,##0")
    set_input(ws, "D28", 0)  # template uses 0 on entry CF row 28
    set_formula(ws, "J28", "=J27+J26", fmt="#,##0")
    set_formula(ws, "D29", "=D27+D26", fmt="#,##0")
    set_formula(ws, "J29", "=J27", fmt="#,##0")

    # Terminal value formulas
    set_formula(ws, "N18", "=(I26*(1+D7))/(D6-D7)", fmt="#,##0")
    set_formula(ws, "N19", "=D8*(I21+I23)", fmt="#,##0")
    set_formula(ws, "N20", "=AVERAGE(N18:N19)", fmt="#,##0")

    # Intrinsic / market / returns — formulas
    set_formula(ws, "D32", "=XNPV(D6,D28:J28,D18:J18)", fmt="#,##0")
    set_formula(ws, "D33", "=+D14", fmt="#,##0")
    set_formula(ws, "D34", "=+D13", fmt="#,##0")
    set_formula(ws, "D35", "=D32+D33-D34", fmt="#,##0")
    set_formula(ws, "D37", "=D35/D12", fmt="#,##0.00")

    set_formula(ws, "I32", "=D12*D11", fmt="#,##0")
    set_formula(ws, "I33", "=D13", fmt="#,##0")
    set_formula(ws, "I34", "=+D14", fmt="#,##0")
    set_formula(ws, "I35", "=I32+I33-I34", fmt="#,##0")
    set_formula(ws, "I37", "=D11", fmt="#,##0.00")

    set_formula(ws, "N32", "=D37/I37-1", fmt="0.0%")
    set_formula(ws, "N33", "=XIRR(D29:J29,D18:J18)", fmt="0.0%")
    set_formula(ws, "N36", "=I37", fmt="#,##0.00")
    set_formula(ws, "N37", "=D37-I37", fmt="#,##0.00")
    set_formula(ws, "N38", "=SUM(N36:N37)", fmt="#,##0.00")

    ws["L17"] = "Terminal Value"
    ws["L18"] = "Perpetual Growth"
    ws["L19"] = "EV/EBITDA"
    ws["L20"] = "Average"


def dcf_standalone_drivers(ws):
    """
    EBIT / D&A / ΔNWC as formulas from base + growth assumptions.
    Growth rates typed in a driver row; EBIT depends on them.
    """
    # Growth inputs for explicit period (row 4 under assumptions area use columns E-I)
    ws["B4"] = "EBIT growth (explicit)"
    for i, col in enumerate(DCF_COLS):
        set_input(ws, f"{col}4", GROWTHS[i], fmt="0.0%")

    # EBIT_t = prior*(1+g). E21 depends on D3 base and E4 growth; F21 depends on E21, etc.
    set_formula(ws, "E21", "=$D$3*(1+E4)", fmt="#,##0", linked=True)
    set_formula(ws, "F21", "=E21*(1+F4)", fmt="#,##0", linked=True)
    set_formula(ws, "G21", "=F21*(1+G4)", fmt="#,##0", linked=True)
    set_formula(ws, "H21", "=G21*(1+H4)", fmt="#,##0", linked=True)
    set_formula(ws, "I21", "=H21*(1+I4)", fmt="#,##0", linked=True)

    # D&A scales with EBIT vs FY25 OpInc: DA_t = FY25_DA * EBIT_t / FY25_OpInc
    for col in DCF_COLS:
        set_formula(ws, f"{col}23", f"=$E$3*{col}21/$D$3", fmt="#,##0", linked=True)

    # ΔNWC ≈ 5% of YoY EBIT growth proxy (simple WC investment formula)
    set_formula(ws, "E25", "=MAX(0,E21-$D$3)*0.05", fmt="#,##0", linked=True)
    set_formula(ws, "F25", "=MAX(0,F21-E21)*0.05", fmt="#,##0", linked=True)
    set_formula(ws, "G25", "=MAX(0,G21-F21)*0.05", fmt="#,##0", linked=True)
    set_formula(ws, "H25", "=MAX(0,H21-G21)*0.05", fmt="#,##0", linked=True)
    set_formula(ws, "I25", "=MAX(0,I21-H21)*0.05", fmt="#,##0", linked=True)

    ws["B21"] = "EBIT (formula from D3 & growth row 4)"
    ws["B22"] = "Less: Cash Taxes (=EBIT × tax rate)"
    ws["B23"] = "Plus: D&A (scales with EBIT)"
    ws["B24"] = "Less: Capex (=D15×(1+D16)^period)"
    ws["B25"] = "Less: ΔNWC (formula on EBIT change)"
    ws["B26"] = "Unlevered FCF (=EBIT-Tax+D&A-Capex-ΔNWC)"


def dcf_link_to_three_statement(ws):
    """Green linked formulas — DCF depends on 03_Three_Statement forecast."""
    set_formula(ws, "D5", "='03_Three_Statement'!J13", linked=True, fmt="0.00%")
    for dcol, scol in zip(DCF_COLS, FC_COLS):
        # EBIT ≈ GP - SG&A - R&D - D&A
        set_formula(
            ws,
            f"{dcol}21",
            (
                f"='03_Three_Statement'!{scol}26-'03_Three_Statement'!{scol}28"
                f"-'03_Three_Statement'!{scol}29-'03_Three_Statement'!{scol}30"
            ),
            linked=True,
            fmt="#,##0",
        )
        set_formula(ws, f"{dcol}23", f"='03_Three_Statement'!{scol}30", linked=True, fmt="#,##0")
        set_formula(ws, f"{dcol}24", f"='03_Three_Statement'!{scol}68", linked=True, fmt="#,##0")
        set_formula(ws, f"{dcol}25", f"='03_Three_Statement'!{scol}89", linked=True, fmt="#,##0")
        # taxes & UFCF already set as formulas depending on EBIT/tax/DA/capex/nwc

    ws["B21"] = "EBIT (linked 3-stmt GP−SG&A−R&D−D&A)"
    ws["B22"] = "Less: Cash Taxes (=EBIT×D5)"
    ws["B23"] = "Plus: D&A (linked 3-stmt)"
    ws["B24"] = "Less: Capex (linked 3-stmt row 68)"
    ws["B25"] = "Less: ΔNWC (linked 3-stmt row 89)"
    ws["B26"] = "Unlevered FCF (formula)"


def add_dcf_chart(ws):
    ws._charts = []
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Unlevered FCF (formula-driven)"
    chart.y_axis.title = "USD thousands"
    data = Reference(ws, min_col=5, min_row=26, max_col=9, max_row=26)
    cats = Reference(ws, min_col=5, min_row=17, max_col=9, max_row=17)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.width = 16
    chart.height = 9
    ws.add_chart(chart, "L3")


def add_formula_legend(wb, joined=False):
    if "Formula_Map" in wb.sheetnames:
        del wb["Formula_Map"]
    ws = wb.create_sheet("Formula_Map", 0)
    ws["A1"] = "Updated math — Excel formula dependencies (not typed results)"
    ws["A1"].font = TITLE
    ws["A2"] = "Yellow cells = inputs. Green/white formula cells = depend on other rows/sheets."
    headers = ["Line item", "Cell", "Excel formula / dependency", "Meaning"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(4, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    rows = [
        ("Tax rate", "D5", "input (or ='03_Three_Statement'!J13 if joined)", "FY25 ETR"),
        ("Years", "E17:I17", "=YEAR(E18) …", "Depends on date row"),
        ("Dates", "E18:I18", "=DATE(YEAR($D$10)+E19,9,30)", "Depends on FYE + period"),
        ("Periods", "F19:I19", "=E19+1 …", "Depends on prior period"),
        ("Year fraction", "E20:I20", "=YEARFRAC(prior_date, date)", "Discount stub"),
        ("EBIT", "E21:I21", "Standalone: =$D$3*(1+E4) then rolls; Joined: 3-stmt GP−opex−DA", "Operating income path"),
        ("Cash Taxes", "E22:I22", "=E21*$D$5", "Depends on EBIT and tax rate"),
        ("D&A", "E23:I23", "Standalone: =$E$3*E21/$D$3; Joined: ='03_Three_Statement'!J30", "Depends on EBIT or 3-stmt"),
        ("Capex", "E24:I24", "Standalone: =$D$15*(1+$D$16)^E19; Joined: 3-stmt!J68", "Depends on base capex/growth or 3-stmt"),
        ("ΔNWC", "E25:I25", "Standalone: =MAX(0,ΔEBIT)*5%; Joined: 3-stmt!J89", "WC investment"),
        ("Unlevered FCF", "E26:I26", "=EBIT − Tax + D&A − Capex − ΔNWC", "Core FCF identity"),
        ("TV (PG)", "N18", "=(I26*(1+D7))/(D6-D7)", "Gordon growth on final UFCF"),
        ("TV (EV/EBITDA)", "N19", "=D8*(I21+I23)", "Exit multiple on final EBITDA"),
        ("TV Average", "N20", "=AVERAGE(N18:N19)", "Blended exit"),
        ("Enterprise Value", "D32", "=XNPV(D6,D28:J28,D18:J18)", "PV of transaction CFs"),
        ("Equity Value", "D35", "=D32+D33−D34", "EV + cash − debt"),
        ("Equity / share", "D37", "=D35/D12", "Depends on equity & shares"),
        ("3S Revenue fcst", "J24", "=I24*(1+J7)", "Grows from prior year"),
        ("3S COGS fcst", "J25", "=J24*J8", "Depends on revenue & COGS%"),
        ("3S Gross Profit", "J26", "=J24-J25", "Depends on Rev & COGS"),
        ("3S SG&A fcst", "J28", "=J9", "Depends on salaries driver"),
        ("3S Capex fcst", "J68", "=J18", "Depends on capex assumption"),
        ("3S ΔNWC fcst", "J89", "=J88-I88", "Depends on NWC schedule"),
    ]
    for i, row in enumerate(rows, start=5):
        for j, v in enumerate(row, start=1):
            ws.cell(i, j, v)
            ws.cell(i, j).alignment = Alignment(wrap_text=True, vertical="top")
    ws["A30"] = "10-K"
    ws["B30"] = FILING
    ws["B30"].hyperlink = FILING
    ws["B30"].font = LINK_FONT
    if joined:
        ws["A32"] = "Joined roll-forward: 04_DCF EBIT/D&A/Capex/ΔNWC are formulas into 03_Three_Statement J:N."
    for i, w in enumerate([22, 18, 70, 45], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_math_md():
    OUT_MATH.write_text(
        f"""# Updated DCF / 3-Statement math (Excel formulas)

Only **yellow input cells** are typed numbers. Calculated lines are Excel formulas.

## DCF identities

| Line | Formula |
|---|---|
| Cash Taxes | `=EBIT × TaxRate` → `E22 = E21*$D$5` |
| Capex (standalone) | `=Capex0 × (1+g_capex)^period` → `E24 = $D$15*(1+$D$16)^E19` |
| Capex (joined) | `='03_Three_Statement'!J68` |
| Unlevered FCF | `=EBIT − CashTaxes + D&A − Capex − ΔNWC` |
| TV (perpetuity) | `=UFCF_t × (1+g) / (WACC − g)` → `N18` |
| TV (exit multiple) | `=EV/EBITDA × (EBIT_t + D&A_t)` → `N19` |
| TV used | `=AVERAGE(N18:N19)` → `N20` |
| Enterprise Value | `=XNPV(WACC, TransactionCFs, Dates)` → `D32` |
| Equity Value | `=EV + Cash − Debt` → `D35 = D32+D33-D34` |
| Equity / share | `=Equity / Shares` → `D37 = D35/D12` |

## EBIT path

- **Standalone:** `E21 = $D$3*(1+E4)` then `F21 = E21*(1+F4)` … (D3 = FY25 OpInc {FY25_OPINC:,})
- **Joined:** `E21 = '03_Three_Statement'!J26 - J28 - J29 - J30` (GP − SG&A − R&D − D&A)

## 3-Statement forecast dependencies

| Line | Formula |
|---|---|
| Revenue | `J24 = I24*(1+J7)` |
| COGS | `J25 = J24*J8` |
| Gross Profit | `J26 = J24-J25` |
| SG&A | `J28 = J9` |
| R&D | `J29 = J10` |
| Capex | `J68 = J18` |
| ΔNWC | `J89 = J88-I88` |

## Assumptions used

- Tax rate `{TAX_RATE:.2%}` · WACC `{WACC:.1%}` · g `{G_PERP:.0%}` · EV/EBITDA `{EV_EBITDA:.0f}x`
- Capex0 `{FY25_CAPEX:,}` · Capex growth `{CAPEX_GROWTH:.0%}`
- FY25 OpInc `{FY25_OPINC:,}` · D&A `{FY25_DA:,}`

Source: {FILING}
"""
    )


def copy_sheet(src, dst):
    for row in src.iter_rows():
        for cell in row:
            n = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                n.font = copy(cell.font)
                n.fill = copy(cell.fill)
                n.border = copy(cell.border)
                n.number_format = cell.number_format
                n.alignment = copy(cell.alignment)
    for letter, dim in src.column_dimensions.items():
        dst.column_dimensions[letter].width = dim.width


def build():
    s = load_series()
    write_math_md()

    # --- 3-statement standalone ---
    shutil.copy2(ORIG_3, OUT_3S)
    wb3 = load_workbook(OUT_3S)
    for ws in wb3.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []
    inject_three_statement(wb3["3 Statement Model"], s)
    add_formula_legend(wb3, joined=False)
    ws3 = wb3["3 Statement Model"]
    ws3._charts = []
    ch = LineChart()
    ch.title = "Revenue (formula forecast J24:N24)"
    ch.add_data(Reference(ws3, min_col=5, min_row=24, max_col=14, max_row=24), from_rows=True)
    ch.set_categories(Reference(ws3, min_col=5, min_row=2, max_col=14, max_row=2))
    ch.width = 18
    ch.height = 9
    ws3.add_chart(ch, "P3")
    wb3.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    wb3.save(OUT_3S)
    print("Saved", OUT_3S)

    # --- DCF standalone (formula drivers) ---
    shutil.copy2(ORIG_DCF, OUT_DCF)
    wbd = load_workbook(OUT_DCF)
    for ws in wbd.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []
    dcf = wbd["DCF Model"]
    dcf["B2"] = "FICO DCF (formula-driven)"
    setup_dcf_assumptions(dcf)
    restore_dcf_structure_formulas(dcf)
    dcf_standalone_drivers(dcf)
    add_dcf_chart(dcf)
    dcf["D2"] = "Yellow=inputs. All taxes/UFCF/TV/EV/dates are Excel formulas."
    dcf["D2"].font = NOTE
    dcf["B50"] = "Open Formula_Map sheet for the dependency list. Press F9 if needed."
    dcf["B50"].font = NOTE
    widen(dcf)
    add_formula_legend(wbd, joined=False)
    wbd.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    wbd.save(OUT_DCF)
    print("Saved", OUT_DCF)

    # --- Joined ---
    tmp3 = FICO_DIR / "_tmp3.xlsx"
    tmpd = FICO_DIR / "_tmpd.xlsx"
    shutil.copy2(ORIG_3, tmp3)
    shutil.copy2(ORIG_DCF, tmpd)
    wb3t = load_workbook(tmp3)
    wbdt = load_workbook(tmpd)
    for w in (wb3t, wbdt):
        for ws in w.worksheets:
            if getattr(ws, "_images", None):
                ws._images = []
    inject_three_statement(wb3t["3 Statement Model"], s)
    dcfj = wbdt["DCF Model"]
    dcfj["B2"] = "FICO DCF (linked formulas → 03_Three_Statement)"
    setup_dcf_assumptions(dcfj)
    restore_dcf_structure_formulas(dcfj)
    dcf_link_to_three_statement(dcfj)
    add_dcf_chart(dcfj)
    widen(dcfj)

    out = Workbook()
    out.remove(out.active)
    add_formula_legend(out, joined=True)
    ins = out.create_sheet("01_Instructions", 0)
    ins["A1"] = "FICO Joined Model — formula dependencies"
    ins["A1"].font = TITLE
    for i, line in enumerate(
        [
            "",
            "Yellow cells are the ONLY typed inputs.",
            "Every total, tax, UFCF, TV, EV, revenue roll-forward is an Excel formula.",
            "04_DCF!E21:I25 are formulas into 03_Three_Statement (green).",
            "04_DCF!E22/E26/N18:N20/D32:D37 remain native math formulas.",
            "See Formula_Map and FORMULA_MATH.md.",
            FILING,
        ],
        start=2,
    ):
        ins[f"A{i}"] = line
    ins.column_dimensions["A"].width = 100

    ws3o = out.create_sheet("03_Three_Statement")
    copy_sheet(wb3t["3 Statement Model"], ws3o)
    ws3o._charts = []
    ch2 = LineChart()
    ch2.title = "Revenue"
    ch2.add_data(Reference(ws3o, min_col=5, min_row=24, max_col=14, max_row=24), from_rows=True)
    ch2.set_categories(Reference(ws3o, min_col=5, min_row=2, max_col=14, max_row=2))
    ch2.width = 18
    ch2.height = 9
    ws3o.add_chart(ch2, "P3")

    wsdo = out.create_sheet("04_DCF")
    copy_sheet(dcfj, wsdo)
    # Re-apply structure + cross-sheet links (copy already brought inputs)
    # D5 becomes a formula to 3-statement — do not re-run setup_dcf_assumptions.
    restore_dcf_structure_formulas(wsdo)
    dcf_link_to_three_statement(wsdo)
    add_dcf_chart(wsdo)
    widen(wsdo)

    order = ["01_Instructions", "Formula_Map", "03_Three_Statement", "04_DCF"]
    for i, name in enumerate(order):
        out.move_sheet(name, offset=i - out.sheetnames.index(name))
    out.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    out.save(OUT_JOINED)
    tmp3.unlink(missing_ok=True)
    tmpd.unlink(missing_ok=True)
    print("Saved", OUT_JOINED)


def verify():
    # Standalone DCF: calculated lines must be formulas
    wb = load_workbook(OUT_DCF)
    d = wb["DCF Model"]
    assert is_formula(d["E22"].value) and "E21" in d["E22"].value
    assert is_formula(d["E24"].value) and "$D$15" in d["E24"].value
    assert is_formula(d["E26"].value)
    assert is_formula(d["E21"].value) and "$D$3" in d["E21"].value
    assert is_formula(d["F21"].value) and "E21" in d["F21"].value
    assert is_formula(d["N18"].value) and "I26" in d["N18"].value
    assert is_formula(d["N20"].value)
    assert is_formula(d["D32"].value) and "XNPV" in d["D32"].value
    assert is_formula(d["D35"].value)
    assert is_formula(d["E18"].value) and "DATE" in d["E18"].value
    assert d["E19"].value == 1 and is_formula(d["F19"].value)
    print("DCF standalone formulas OK:", d["E26"].value, d["N18"].value)

    wb3 = load_workbook(OUT_3S)
    s = wb3["3 Statement Model"]
    assert is_formula(s["J24"].value) and "I24" in s["J24"].value
    assert is_formula(s["J25"].value)
    assert is_formula(s["J26"].value)
    assert is_formula(s["I26"].value)
    assert not is_formula(s["J9"].value)  # driver input
    print("3S formulas OK:", s["J24"].value, s["J28"].value)

    wj = load_workbook(OUT_JOINED)
    dj = wj["04_DCF"]
    assert "03_Three_Statement" in dj["E21"].value
    assert is_formula(dj["E22"].value)
    assert "03_Three_Statement" in dj["E24"].value
    assert is_formula(dj["E26"].value)
    assert is_formula(dj["D32"].value)
    print("Joined link formulas OK:", dj["E21"].value)


if __name__ == "__main__":
    build()
    verify()
