#!/usr/bin/env python3
"""
Fill EVERY blank DCF cell with real computed numbers + working charts.

Viewers that don't calculate Excel formulas were showing blank Dates, Cash Taxes,
Capex, UFCF, Terminal Value, and Intrinsic Value. This writes the evaluated
results as values so the sheet displays immediately.
"""

from __future__ import annotations

import math
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parents[1]
FICO_DIR = Path(__file__).resolve().parent
ORIG_DCF = ROOT / "DCF resources" / "originals" / "CFI_DCF-Model.xlsx"
OUT_DCF = FICO_DIR / "DCF_FICO.xlsx"
OUT_JOINED = FICO_DIR / "FICO_3S_DCF_Joined.xlsx"

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
RESULT_FILL = PatternFill("solid", fgColor="DEEBF7")
NOTE = Font(name="Calibri", size=10, italic=True, color="C00000")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")

# FICO assumptions (matches prior pack / 10-K + market bridge)
TAX = 150_649 / 802_595  # ~18.77%
WACC = 0.096
G = 0.03
EV_EBITDA = 25.0
TXN = datetime(2025, 9, 30)
PRICE = 1046.23
SHARES = 21_597.635  # thousands
DEBT = 5_582_389
CASH = 304_537
CAPEX = 39_407  # PPE purchases + capitalized software FY25

# Explicit forecast (USD thousands) — same path previously projected from FY25
# Years shown: FY2026 .. FY2030
YEARS = [2026, 2027, 2028, 2029, 2030]
EBIT = [1_048_623.4, 1_284_433.8, 1_505_458.6, 1_717_399.8, 1_922_965.4]
DA = [18_156.7, 20_335.5, 22_369.1, 24_382.3, 26_332.9]
DNWC = [13_936.1, 13_617.5, 12_709.7, 12_582.6, 12_191.2]
COLS = ["E", "F", "G", "H", "I"]


def xnpv(rate: float, cashflows: list[float], dates: list[datetime]) -> float:
    d0 = dates[0]
    total = 0.0
    for cf, dt in zip(cashflows, dates):
        days = (dt - d0).days
        total += cf / ((1.0 + rate) ** (days / 365.0))
    return total


def xirr(cashflows: list[float], dates: list[datetime], guess: float = 0.1) -> float | None:
    """Newton / bisection XIRR; returns None if no sign change."""
    if not any(c > 0 for c in cashflows) or not any(c < 0 for c in cashflows):
        return None

    def f(r: float) -> float:
        return xnpv(r, cashflows, dates)

    lo, hi = -0.99, 5.0
    flo, fhi = f(lo), f(hi)
    if flo * fhi > 0:
        # try expand
        return None
    for _ in range(100):
        mid = (lo + hi) / 2
        fm = f(mid)
        if abs(fm) < 1e-6:
            return mid
        if flo * fm < 0:
            hi, fhi = mid, fm
        else:
            lo, flo = mid, fm
    return (lo + hi) / 2


def put(ws, coord, value, *, input_cell=False, result=False, fmt=None):
    cell = ws[coord]
    cell.value = value
    if input_cell:
        cell.font = BLUE
        cell.fill = INPUT_FILL
    elif result:
        cell.fill = RESULT_FILL
    if fmt:
        cell.number_format = fmt


def widen(ws):
    ws.column_dimensions["B"].width = 34
    for col in range(4, 15):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 10, 14)


def compute() -> dict:
    taxes = [e * TAX for e in EBIT]
    ufcf = [e - t + d - CAPEX - n for e, t, d, n in zip(EBIT, taxes, DA, DNWC)]
    dates = [datetime(y, 9, 30) for y in YEARS]
    # year fractions from entry / prior date
    yf = []
    prev = TXN
    for dt in dates:
        yf.append(max((dt - prev).days / 365.0, 0.0))
        prev = dt

    tv_pg = ufcf[-1] * (1 + G) / (WACC - G)
    tv_ebitda = EV_EBITDA * (EBIT[-1] + DA[-1])
    tv_avg = (tv_pg + tv_ebitda) / 2

    market_cap = SHARES * PRICE
    market_ev = market_cap + DEBT - CASH

    # Transaction CFs: entry = -market EV at t0; explicit UFCF; exit = TV at last date
    # Match CFI layout: D28 uses row28; D27=-I35 (market EV); J27=TV avg
    entry = -market_ev
    txn_dates = [TXN] + dates  # D18, E18..I18, and J18=I18 for exit same day as last CF
    # CFI: J18 = I18 (exit on last forecast date), J27 = TV, J28 = J27+J26 (J26 empty→TV)
    # Row 28 cashflows for XNPV: D28=0 in template sample, but D29=D27+D26 for IRR
    # Intrinsic XNPV uses D28:J28 where D28=0, E28:I28=UFCF*yf, J28=TV
    cf28 = [0.0]  # D28
    for u, f in zip(ufcf, yf):
        cf28.append(u * f)  # no entry/exit in E-I
    cf28.append(tv_avg)  # J28 = TV (+ J26 empty)

    dates28 = [TXN] + dates + [dates[-1]]  # J18 = I18

    # IRR row 29: D29 = D27+D26 = entry; E29:I29 = UFCF*yf; J29 = J27 = TV
    cf29 = [entry] + [u * f for u, f in zip(ufcf, yf)] + [tv_avg]

    ev_intrinsic = xnpv(WACC, cf28, dates28)
    # Alternative economically cleaner EV: PV of UFCFs + PV of TV (no zero at t0)
    cf_clean = [u * f for u, f in zip(ufcf, yf)] + [tv_avg]
    dates_clean = dates + [dates[-1]]
    # Actually standard: discount each UFCF and TV from entry
    ev_std = 0.0
    for u, dt in zip(ufcf, dates):
        years = (dt - TXN).days / 365.0
        ev_std += u / ((1 + WACC) ** years)
    years_exit = (dates[-1] - TXN).days / 365.0
    ev_std += tv_avg / ((1 + WACC) ** years_exit)

    equity_intrinsic = ev_std + CASH - DEBT
    equity_ps = equity_intrinsic / SHARES
    irr = xirr(cf29, dates28)
    upside_pct = equity_ps / PRICE - 1
    upside_abs = equity_ps - PRICE

    return {
        "taxes": taxes,
        "ufcf": ufcf,
        "yf": yf,
        "dates": dates,
        "tv_pg": tv_pg,
        "tv_ebitda": tv_ebitda,
        "tv_avg": tv_avg,
        "market_cap": market_cap,
        "market_ev": market_ev,
        "entry": entry,
        "cf28": cf28,
        "cf29": cf29,
        "dates28": dates28,
        "ev_intrinsic": ev_std,
        "equity_intrinsic": equity_intrinsic,
        "equity_ps": equity_ps,
        "irr": irr,
        "upside_pct": upside_pct,
        "upside_abs": upside_abs,
    }


def fill_dcf_sheet(ws, calc: dict, title: str) -> None:
    # Clear old charts/images
    ws._charts = []
    ws._images = []

    ws["B2"] = title

    # Assumptions
    put(ws, "D5", round(TAX, 4), input_cell=True, fmt="0.00%")
    put(ws, "D6", WACC, input_cell=True, fmt="0.0%")
    put(ws, "D7", G, input_cell=True, fmt="0.0%")
    put(ws, "D8", EV_EBITDA, input_cell=True, fmt="0.0")
    put(ws, "D9", TXN, input_cell=True, fmt="yyyy-mm-dd")
    put(ws, "D10", TXN, input_cell=True, fmt="yyyy-mm-dd")
    put(ws, "D11", PRICE, input_cell=True, fmt="#,##0.00")
    put(ws, "D12", SHARES, input_cell=True, fmt="#,##0.000")
    put(ws, "D13", DEBT, input_cell=True, fmt="#,##0")
    put(ws, "D14", CASH, input_cell=True, fmt="#,##0")
    put(ws, "D15", CAPEX, input_cell=True, fmt="#,##0")

    # Headers / dates / periods / year fraction — ALL VALUES
    put(ws, "D17", "Entry")
    put(ws, "J17", "Exit")
    put(ws, "D18", TXN, result=True, fmt="yyyy-mm-dd")
    put(ws, "J18", calc["dates"][-1], result=True, fmt="yyyy-mm-dd")

    for i, col in enumerate(COLS):
        put(ws, f"{col}17", YEARS[i], result=True, fmt="0")
        put(ws, f"{col}18", calc["dates"][i], result=True, fmt="yyyy-mm-dd")
        put(ws, f"{col}19", i + 1, result=True, fmt="0")
        put(ws, f"{col}20", round(calc["yf"][i], 4), result=True, fmt="0.00")
        put(ws, f"{col}21", round(EBIT[i], 1), input_cell=True, fmt="#,##0")
        put(ws, f"{col}22", round(calc["taxes"][i], 1), result=True, fmt="#,##0")
        put(ws, f"{col}23", round(DA[i], 1), input_cell=True, fmt="#,##0")
        put(ws, f"{col}24", CAPEX, result=True, fmt="#,##0")
        put(ws, f"{col}25", round(DNWC[i], 1), input_cell=True, fmt="#,##0")
        put(ws, f"{col}26", round(calc["ufcf"][i], 1), result=True, fmt="#,##0")
        # Transaction CF (explicit period)
        put(ws, f"{col}28", round(calc["ufcf"][i] * calc["yf"][i], 1), result=True, fmt="#,##0")
        put(ws, f"{col}29", round(calc["ufcf"][i] * calc["yf"][i], 1), result=True, fmt="#,##0")

    # Entry / Exit
    put(ws, "D27", round(calc["entry"], 1), result=True, fmt="#,##0")
    put(ws, "J27", round(calc["tv_avg"], 1), result=True, fmt="#,##0")
    put(ws, "D28", 0, result=True, fmt="#,##0")
    put(ws, "J28", round(calc["tv_avg"], 1), result=True, fmt="#,##0")
    put(ws, "D29", round(calc["entry"], 1), result=True, fmt="#,##0")
    put(ws, "J29", round(calc["tv_avg"], 1), result=True, fmt="#,##0")

    # Terminal value block (multiples)
    put(ws, "N18", round(calc["tv_pg"], 1), result=True, fmt="#,##0")
    put(ws, "N19", round(calc["tv_ebitda"], 1), result=True, fmt="#,##0")
    put(ws, "N20", round(calc["tv_avg"], 1), result=True, fmt="#,##0")

    # Intrinsic value
    put(ws, "D32", round(calc["ev_intrinsic"], 1), result=True, fmt="#,##0")
    put(ws, "D33", CASH, result=True, fmt="#,##0")
    put(ws, "D34", DEBT, result=True, fmt="#,##0")
    put(ws, "D35", round(calc["equity_intrinsic"], 1), result=True, fmt="#,##0")
    put(ws, "D37", round(calc["equity_ps"], 2), result=True, fmt="#,##0.00")

    # Market value
    put(ws, "I32", round(calc["market_cap"], 1), result=True, fmt="#,##0")
    put(ws, "I33", DEBT, result=True, fmt="#,##0")
    put(ws, "I34", CASH, result=True, fmt="#,##0")
    put(ws, "I35", round(calc["market_ev"], 1), result=True, fmt="#,##0")
    put(ws, "I37", PRICE, result=True, fmt="#,##0.00")

    # Rate of return / upside
    put(ws, "N32", round(calc["upside_pct"], 4), result=True, fmt="0.0%")
    if calc["irr"] is not None:
        put(ws, "N33", round(calc["irr"], 4), result=True, fmt="0.0%")
    else:
        put(ws, "N33", "n/a", result=True)
    put(ws, "N36", PRICE, result=True, fmt="#,##0.00")
    put(ws, "N37", round(calc["upside_abs"], 2), result=True, fmt="#,##0.00")
    put(ws, "N38", round(calc["equity_ps"], 2), result=True, fmt="#,##0.00")

    # Labels clarity
    ws["B21"] = "EBIT"
    ws["B22"] = "Less: Cash Taxes"
    ws["B23"] = "Plus: D&A"
    ws["B24"] = "Less: Capex"
    ws["B25"] = "Less: Changes in NWC"
    ws["B26"] = "Unlevered FCF"
    ws["L17"] = "Terminal Value"
    ws["L18"] = "Perpetual Growth"
    ws["L19"] = "EV/EBITDA"
    ws["L20"] = "Average"

    ws["B50"] = (
        "Blue/yellow = inputs. Blue-gray = computed (dates, taxes, Capex, UFCF, TV, EV). "
        "Capex = 10-K PPE purchases + capitalized software. Tax = FY25 ETR."
    )
    ws["B50"].font = NOTE
    ws["D2"] = "All DCF lines filled — open 10K_Sources for SEC links"
    ws["D2"].font = NOTE

    # Working chart from UFCF values
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Unlevered Free Cash Flow"
    chart.y_axis.title = "USD thousands"
    chart.x_axis.title = "Fiscal year"
    data = Reference(ws, min_col=5, min_row=26, max_col=9, max_row=26)
    cats = Reference(ws, min_col=5, min_row=17, max_col=9, max_row=17)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 16
    chart.height = 9
    ws.add_chart(chart, "L3")

    widen(ws)


def ensure_sources(wb):
    if "10K_Sources" in wb.sheetnames:
        return
    ws = wb.create_sheet("10K_Sources", 0)
    ws["A1"] = "FICO DCF — 10-K source links"
    ws["A1"].font = TITLE
    ws["A3"] = "FY2025 10-K"
    ws["B3"] = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"
    ws["B3"].hyperlink = ws["B3"].value
    ws["A4"] = "Capex"
    ws["B4"] = "Purchases of PPE 8,922 + Capitalized internal-use software 30,485 = 39,407"
    ws["A5"] = "OpEx"
    ws["B5"] = "SG&A + R&D on income statement (no single OpEx line)"
    ws["A6"] = "EBIT"
    ws["B6"] = "Operating income (projected from FY25 base)"
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 90


def build_standalone(calc: dict) -> None:
    shutil.copy2(ORIG_DCF, OUT_DCF)
    wb = load_workbook(OUT_DCF)
    for ws in wb.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []
    fill_dcf_sheet(wb["DCF Model"], calc, "FICO DCF Model (values filled)")
    ensure_sources(wb)
    if "Cover Page" in wb.sheetnames:
        wb["Cover Page"]["C12"] = "FICO — DCF Model (complete values + chart)"
    wb.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    wb.save(OUT_DCF)
    print(f"Saved {OUT_DCF} ({OUT_DCF.stat().st_size} bytes)")


def patch_joined(calc: dict) -> None:
    if not OUT_JOINED.exists():
        print("Joined workbook missing — skip patch")
        return
    wb = load_workbook(OUT_JOINED)
    if "04_DCF" not in wb.sheetnames:
        print("No 04_DCF sheet — skip")
        return
    fill_dcf_sheet(wb["04_DCF"], calc, "FICO DCF (joined — values filled)")
    # Keep instruction note about 3-statement on Instructions sheet
    if "01_Instructions" in wb.sheetnames:
        ws = wb["01_Instructions"]
        ws["A40"] = (
            "DCF display values are fully computed (dates/taxes/Capex/UFCF/TV/EV) "
            "so nothing appears blank. Yellow cells remain editable inputs."
        )
        ws["A40"].font = NOTE
    wb.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    wb.save(OUT_JOINED)
    print(f"Patched {OUT_JOINED} ({OUT_JOINED.stat().st_size} bytes)")


def verify():
    for path in (OUT_DCF, OUT_JOINED):
        wb = load_workbook(path)
        name = "DCF Model" if "DCF Model" in wb.sheetnames else "04_DCF"
        ws = wb[name]
        assert ws["E17"].value == 2026
        assert ws["I17"].value == 2030
        assert isinstance(ws["E18"].value, datetime)
        assert isinstance(ws["E22"].value, (int, float)) and ws["E22"].value > 0
        assert ws["E24"].value == CAPEX
        assert isinstance(ws["E26"].value, (int, float)) and ws["E26"].value > 0
        assert isinstance(ws["N18"].value, (int, float)) and ws["N18"].value > 0
        assert isinstance(ws["N19"].value, (int, float)) and ws["N19"].value > 0
        assert isinstance(ws["N20"].value, (int, float)) and ws["N20"].value > 0
        assert isinstance(ws["D32"].value, (int, float))
        assert isinstance(ws["D37"].value, (int, float))
        assert isinstance(ws["I35"].value, (int, float))
        assert len(ws._charts) >= 1
        print(
            f"OK {path.name}: years={ws['E17'].value}-{ws['I17'].value} "
            f"UFCF_E={ws['E26'].value:,.0f} TV_avg={ws['N20'].value:,.0f} "
            f"Eq/sh={ws['D37'].value:,.2f} charts={len(ws._charts)}"
        )


def main():
    calc = compute()
    print("Computed UFCF:", [round(x) for x in calc["ufcf"]])
    print("TV pg/ebitda/avg:", round(calc["tv_pg"]), round(calc["tv_ebitda"]), round(calc["tv_avg"]))
    print("Intrinsic EV / Eq/sh:", round(calc["ev_intrinsic"]), round(calc["equity_ps"], 2))
    print("Market EV / upside%:", round(calc["market_ev"]), f"{calc['upside_pct']*100:.1f}%")
    build_standalone(calc)
    patch_joined(calc)
    verify()


if __name__ == "__main__":
    main()
