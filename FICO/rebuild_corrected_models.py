#!/usr/bin/env python3
"""
Rebuild corrected FICO 3-statement + DCF after audit.

Documented errors fixed
----------------------
3-Statement (3S_FICO.xlsx):
  1. Forecast drivers still CFI sample defaults (COGS 37%, SG&A $25k, R&D $10k,
     tax 28%, AR days 18, inv days 73, capex ~15k) — replaced with FICO drivers.
  2. Year headers were formulas (=E2+1) that show blank before calc — write 2021–2030.
  3. Missing FICO-scale forecast entries for growth, opex, capex, tax, WC days.
  4. Interest rate assumption 3% vs FICO ~4–5% on opening debt — set from FY25.

DCF (DCF_FICO.xlsx / joined):
  1. Year-1 EBIT started at ~1,048,623 (> FY25 Operating income 924,850) — re-anchor.
  2. D&A path used ~0.8% of revenue (~18k) vs FY25 D&A 14,952 — re-anchor.
  3. Explicit years skipped a clear FY2025 base bridge — columns = FY2026–2030 with
     FY2025 actuals shown in assumptions / base row.
  4. Capex frozen flat — grow with revenue.
  5. EV/EBITDA exit 25x produced TV ~2x perpetuity; keep both but also show mid.
  6. All previously blank formula rows kept as computed values for display.
"""

from __future__ import annotations

import json
import shutil
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
ERRORS_MD = FICO_DIR / "MODEL_ERROR_FIXES.md"

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
RESULT_FILL = PatternFill("solid", fgColor="DEEBF7")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
NOTE = Font(name="Calibri", size=10, italic=True, color="C00000")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True)
LINK_FONT = Font(name="Calibri", color="0563C1", underline="single")

HIST_COLS = list("EFGHI")
HIST_YEARS = [2021, 2022, 2023, 2024, 2025]
FC_COLS = list("JKLMN")
FC_YEARS = [2026, 2027, 2028, 2029, 2030]
GROWTHS = [0.12, 0.10, 0.09, 0.08, 0.07]  # FY26..FY30

# FY25 audited anchors ($000s)
FY25_REV = 1_990_869
FY25_COGS = 353_722
FY25_SGA = 513_028
FY25_RD = 188_347
FY25_OPINC = 924_850
FY25_DA = 14_952
FY25_INTEREST = 133_647
FY25_TAX = 150_649
FY25_EBT = 802_595
FY25_NI = 651_946
FY25_CAPEX = 39_407  # 8,922 + 30,485
FY25_CASH = 134_136
FY25_AR = 529_148
FY25_PPE = 67_713
FY25_AP = 32_315
FY25_DEBT = 3_055_691

TAX_RATE = FY25_TAX / FY25_EBT  # 0.1877
COGS_PCT = FY25_COGS / FY25_REV  # ~0.1777
WACC = 0.096
G_PERP = 0.03
EV_EBITDA = 22.0  # still rich; less extreme than 25x for exit
PRICE = 1046.23
SHARES = 21_597.635
DEBT_10Q = 5_582_389
CASH_10Q = 304_537
TXN = datetime(2025, 9, 30)
FILING_2025 = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"


def is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def put(ws, coord, value, *, kind="input", fmt=None, overwrite_formula=False):
    cell = ws[coord]
    if is_formula(cell.value) and not overwrite_formula:
        raise RuntimeError(f"formula guard {coord}: {cell.value}")
    cell.value = value
    if kind == "input":
        cell.font = BLUE
        cell.fill = INPUT_FILL
    elif kind == "result":
        cell.fill = RESULT_FILL
    elif kind == "link":
        cell.fill = LINK_FILL
    if fmt:
        cell.number_format = fmt


def widen(ws, width=15):
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 12, 34)
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
        if y < 2020 or y > 2025:
            continue
        prev = best.get(y)
        if prev is None or it.get("filed", "") >= prev["filed"]:
            best[y] = {"val": it["val"], "filed": it.get("filed", "")}
    return {y: best[y]["val"] / 1000.0 for y in best}


def load_series():
    facts = json.loads(CACHE.read_text())
    s = {
        "revenue": annual(facts, "RevenueFromContractWithCustomerExcludingAssessedTax"),
        "cogs": annual(facts, "CostOfRevenue"),
        "sga": annual(facts, "SellingGeneralAndAdministrativeExpense"),
        "rd": annual(facts, "ResearchAndDevelopmentExpense"),
        "opinc": annual(facts, "OperatingIncomeLoss"),
        "interest": annual(facts, "InterestExpense"),
        "tax": annual(facts, "IncomeTaxExpenseBenefit"),
        "ni": annual(facts, "NetIncomeLoss"),
        "cash": annual(facts, "CashAndCashEquivalentsAtCarryingValue"),
        "ar": annual(facts, "AccountsReceivableNetCurrent"),
        "ppe": annual(facts, "PropertyPlantAndEquipmentNet"),
        "ap": annual(facts, "AccountsPayableCurrent"),
        "re": annual(facts, "RetainedEarningsAccumulatedDeficit"),
        "da": annual(facts, "DepreciationDepletionAndAmortization"),
        "debt_lt": annual(facts, "LongTermDebtNoncurrent"),
        "debt_cur": annual(facts, "LongTermDebtCurrent"),
        "capex_ppe": annual(facts, "PaymentsToAcquirePropertyPlantAndEquipment"),
    }
    s["interest"].update({2023: 95_546, 2024: 105_638, 2025: 133_647})
    debt = {y: s["debt_lt"].get(y, 0) + s["debt_cur"].get(y, 0) for y in range(2020, 2026)}
    debt[2024] = 2_209_021
    debt[2025] = 3_055_691
    s["debt"] = debt
    s["capex"] = dict(s["capex_ppe"])
    s["capex"].update({2023: 4_237, 2024: 25_551, 2025: 39_407})
    return s


def forecast_path():
    """Build FY26–30 operating path anchored to FY25 actuals."""
    rev = FY25_REV
    sga = FY25_SGA
    rd = FY25_RD
    rows = []
    prev_rev = FY25_REV
    prev_nwc = FY25_AR - FY25_AP
    for i, g in enumerate(GROWTHS):
        rev = prev_rev * (1 + g)
        cogs = rev * COGS_PCT
        sga = FY25_SGA * ((1 + g) ** (i + 1))
        rd = FY25_RD * ((1 + g) ** (i + 1))
        da = FY25_DA * (rev / FY25_REV)  # scale with revenue
        ebit = rev - cogs - sga - rd - da
        # approx restructuring already in FY25 opinc; ignore going forward
        tax = ebit * TAX_RATE
        capex = FY25_CAPEX * ((1 + 0.08) ** (i + 1))
        ar = rev * (FY25_AR / FY25_REV)
        ap = cogs * (FY25_AP / FY25_COGS)
        nwc = ar - ap
        dnwc = nwc - prev_nwc
        ufcf = ebit - tax + da - capex - dnwc
        rows.append(
            {
                "year": FC_YEARS[i],
                "g": g,
                "rev": rev,
                "cogs": cogs,
                "cogs_pct": COGS_PCT,
                "sga": sga,
                "rd": rd,
                "da": da,
                "ebit": ebit,
                "tax": tax,
                "capex": capex,
                "dnwc": dnwc,
                "ufcf": ufcf,
                "ar_days": FY25_AR / FY25_REV * 365,
                "ap_days": FY25_AP / FY25_COGS * 365,
            }
        )
        prev_rev = rev
        prev_nwc = nwc
    return rows


def xnpv(rate, cfs, dates):
    d0 = dates[0]
    return sum(cf / ((1 + rate) ** ((dt - d0).days / 365.0)) for cf, dt in zip(cfs, dates))


def xirr(cfs, dates):
    if not (any(c > 0 for c in cfs) and any(c < 0 for c in cfs)):
        return None

    def f(r):
        return xnpv(r, cfs, dates)

    lo, hi = -0.9, 5.0
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(80):
        mid = (lo + hi) / 2
        if abs(f(mid)) < 1e-5:
            return mid
        if f(lo) * f(mid) < 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def inject_3s(ws, s, fc):
    # Concrete year headers so years never appear missing
    for col, y in zip(HIST_COLS, HIST_YEARS):
        put(ws, f"{col}2", y, kind="result", overwrite_formula=True, fmt="0")
    for col, y in zip(FC_COLS, FC_YEARS):
        put(ws, f"{col}2", y, kind="result", overwrite_formula=True, fmt="0")

    for col, y in zip(HIST_COLS, HIST_YEARS):
        put(ws, f"{col}24", round(s["revenue"][y], 1))
        put(ws, f"{col}25", round(s["cogs"][y], 1))
        put(ws, f"{col}28", round(s["sga"][y], 1))
        put(ws, f"{col}29", round(s["rd"][y], 1))
        put(ws, f"{col}30", round(s["da"][y], 1))
        put(ws, f"{col}31", round(s["interest"][y], 1))
        put(ws, f"{col}35", round(s["tax"][y], 1))
        put(ws, f"{col}41", round(s["cash"][y], 1))
        put(ws, f"{col}42", round(s["ar"][y], 1))
        put(ws, f"{col}43", 0)
        put(ws, f"{col}44", round(s["ppe"][y], 1))
        put(ws, f"{col}48", round(s["ap"][y], 1))
        put(ws, f"{col}49", round(s["debt"][y], 1))
        assets_simple = s["cash"][y] + s["ar"][y] + s["ppe"][y]
        liab_simple = s["ap"][y] + s["debt"][y]
        equity_plug = assets_simple - liab_simple
        re = s["re"][y]
        put(ws, f"{col}53", round(re, 1))
        put(ws, f"{col}52", round(equity_plug - re, 1))

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
        put(ws, f"{col}64", round(dnwc, 1))
        put(ws, f"{col}68", round(capex, 1))
        put(ws, f"{col}72", round(debt_issue, 1))
        cfo_proxy = s["ni"][y] + s["da"][y] - dnwc
        equity_issue = (s["cash"][y] - prev_cash) - (cfo_proxy - capex + debt_issue)
        put(ws, f"{col}73", round(equity_issue, 1))
        if col == "E":
            put(ws, "D78", round(s["cash"][y - 1], 1))
        else:
            put(ws, f"{col}77", round(prev_cash, 1))
        put(ws, f"{col}85", round(s["ar"][y], 1))
        put(ws, f"{col}86", 0)
        put(ws, f"{col}87", round(s["ap"][y], 1))
        da = s["da"][y]
        put(ws, f"{col}92", round(open_ppe, 1))
        put(ws, f"{col}93", round(s["ppe"][y] - open_ppe + da, 1))
        put(ws, f"{col}94", round(da, 1))
        put(ws, f"{col}98", round(prev_debt, 1))
        put(ws, f"{col}99", round(debt_issue, 1))
        put(ws, f"{col}101", round(s["interest"][y], 1))

    # FICO forecast assumptions (the missing/wrong entries)
    int_rate = FY25_INTEREST / FY25_DEBT
    da_pct_ppe = FY25_DA / max(s["ppe"][2024], 1)  # vs opening PPE
    for i, col in enumerate(FC_COLS):
        row = fc[i]
        put(ws, f"{col}7", row["g"])
        put(ws, f"{col}8", round(COGS_PCT, 4))
        put(ws, f"{col}9", round(row["sga"], 1))
        put(ws, f"{col}10", round(row["rd"], 1))
        put(ws, f"{col}11", round(min(da_pct_ppe, 0.35), 4))
        put(ws, f"{col}12", round(int_rate, 4))
        put(ws, f"{col}13", round(TAX_RATE, 4))
        put(ws, f"{col}15", round(row["ar_days"], 1))
        put(ws, f"{col}16", 0)  # no inventory
        put(ws, f"{col}17", round(row["ap_days"], 1))
        put(ws, f"{col}18", round(row["capex"], 1))
        put(ws, f"{col}19", 0)
        put(ws, f"{col}20", 0)

    # Labels clarifying FICO mapping on key rows
    ws["B24"] = "Revenue (FICO: Revenues)"
    ws["B25"] = "COGS (FICO: Cost of revenues)"
    ws["B28"] = "SG&A (CFI Salaries ← FICO SG&A)"
    ws["B29"] = "R&D (CFI Rent/Overhead ← FICO R&D)"
    ws["B131"] = (
        "FIXED: forecast drivers now FICO-scale (COGS%~17.8%, SG&A/R&D from FY25, "
        "tax~18.8%, AR days~97, Inv days=0, Capex from 10-K). Years 2021–2030 hardcoded."
    )
    ws["B131"].font = NOTE
    widen(ws)


def build_dcf_calc(fc):
    ebit = [r["ebit"] for r in fc]
    da = [r["da"] for r in fc]
    capex = [r["capex"] for r in fc]
    dnwc = [r["dnwc"] for r in fc]
    taxes = [r["tax"] for r in fc]
    ufcf = [r["ufcf"] for r in fc]
    years = [r["year"] for r in fc]
    dates = [datetime(y, 9, 30) for y in years]
    yf = []
    prev = TXN
    for dt in dates:
        yf.append(max((dt - prev).days / 365.0, 0.0))
        prev = dt

    tv_pg = ufcf[-1] * (1 + G_PERP) / (WACC - G_PERP)
    tv_mult = EV_EBITDA * (ebit[-1] + da[-1])
    tv_avg = (tv_pg + tv_mult) / 2

    # Standard intrinsic EV = PV(UFCF) + PV(TV)
    ev = 0.0
    for u, dt in zip(ufcf, dates):
        ev += u / ((1 + WACC) ** ((dt - TXN).days / 365.0))
    ev += tv_avg / ((1 + WACC) ** ((dates[-1] - TXN).days / 365.0))
    equity = ev + CASH_10Q - DEBT_10Q
    eq_ps = equity / SHARES

    market_cap = SHARES * PRICE
    market_ev = market_cap + DEBT_10Q - CASH_10Q
    entry = -market_ev
    cf_irr = [entry] + ufcf + [tv_avg]
    dates_irr = [TXN] + dates + [dates[-1]]
    irr = xirr(cf_irr, dates_irr)

    return {
        "ebit": ebit,
        "da": da,
        "capex": capex,
        "dnwc": dnwc,
        "taxes": taxes,
        "ufcf": ufcf,
        "years": years,
        "dates": dates,
        "yf": yf,
        "tv_pg": tv_pg,
        "tv_mult": tv_mult,
        "tv_avg": tv_avg,
        "ev": ev,
        "equity": equity,
        "eq_ps": eq_ps,
        "market_cap": market_cap,
        "market_ev": market_ev,
        "entry": entry,
        "irr": irr,
        "upside_pct": eq_ps / PRICE - 1,
        "upside_abs": eq_ps - PRICE,
        # FY25 base bridge shown above forecast
        "fy25_ebit": FY25_OPINC,
        "fy25_da": FY25_DA,
        "fy25_capex": FY25_CAPEX,
        "fy25_tax": FY25_TAX,
    }


def fill_dcf(ws, calc, title):
    ws._charts = []
    ws._images = []
    ws["B2"] = title

    put(ws, "D5", round(TAX_RATE, 4), kind="input", fmt="0.00%")
    put(ws, "D6", WACC, kind="input", fmt="0.0%")
    put(ws, "D7", G_PERP, kind="input", fmt="0.0%")
    put(ws, "D8", EV_EBITDA, kind="input", fmt="0.0")
    put(ws, "D9", TXN, kind="input", fmt="yyyy-mm-dd")
    put(ws, "D10", TXN, kind="input", fmt="yyyy-mm-dd")
    put(ws, "D11", PRICE, kind="input", fmt="#,##0.00")
    put(ws, "D12", SHARES, kind="input", fmt="#,##0.000")
    put(ws, "D13", DEBT_10Q, kind="input", fmt="#,##0")
    put(ws, "D14", CASH_10Q, kind="input", fmt="#,##0")
    put(ws, "D15", FY25_CAPEX, kind="input", fmt="#,##0")

    # FY2025 actual bridge (was missing)
    ws["B16"] = "FY2025 actuals (10-K base)"
    put(ws, "D16", "EBIT/DA/Capex/Tax", kind="result")
    put(ws, "E16", calc["fy25_ebit"], kind="result", fmt="#,##0")
    put(ws, "F16", calc["fy25_da"], kind="result", fmt="#,##0")
    put(ws, "G16", calc["fy25_capex"], kind="result", fmt="#,##0")
    put(ws, "H16", calc["fy25_tax"], kind="result", fmt="#,##0")
    ws["E15"] = "EBIT"
    ws["F15"] = "D&A"
    ws["G15"] = "Capex"
    ws["H15"] = "Tax"

    of = True  # overwrite template formulas with computed display values
    put(ws, "D17", "Entry", overwrite_formula=of)
    put(ws, "J17", "Exit", overwrite_formula=of)
    put(ws, "D18", TXN, kind="result", fmt="yyyy-mm-dd", overwrite_formula=of)
    put(ws, "J18", calc["dates"][-1], kind="result", fmt="yyyy-mm-dd", overwrite_formula=of)

    for i, col in enumerate(list("EFGHI")):
        put(ws, f"{col}17", calc["years"][i], kind="result", fmt="0", overwrite_formula=of)
        put(ws, f"{col}18", calc["dates"][i], kind="result", fmt="yyyy-mm-dd", overwrite_formula=of)
        put(ws, f"{col}19", i + 1, kind="result", fmt="0", overwrite_formula=of)
        put(ws, f"{col}20", round(calc["yf"][i], 4), kind="result", fmt="0.00", overwrite_formula=of)
        put(ws, f"{col}21", round(calc["ebit"][i], 1), kind="input", fmt="#,##0", overwrite_formula=of)
        put(ws, f"{col}22", round(calc["taxes"][i], 1), kind="result", fmt="#,##0", overwrite_formula=of)
        put(ws, f"{col}23", round(calc["da"][i], 1), kind="input", fmt="#,##0", overwrite_formula=of)
        put(ws, f"{col}24", round(calc["capex"][i], 1), kind="result", fmt="#,##0", overwrite_formula=of)
        put(ws, f"{col}25", round(calc["dnwc"][i], 1), kind="input", fmt="#,##0", overwrite_formula=of)
        put(ws, f"{col}26", round(calc["ufcf"][i], 1), kind="result", fmt="#,##0", overwrite_formula=of)
        put(ws, f"{col}28", round(calc["ufcf"][i] * calc["yf"][i], 1), kind="result", fmt="#,##0", overwrite_formula=of)
        put(ws, f"{col}29", round(calc["ufcf"][i] * calc["yf"][i], 1), kind="result", fmt="#,##0", overwrite_formula=of)

    put(ws, "D27", round(calc["entry"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "J27", round(calc["tv_avg"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "D28", 0, kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "J28", round(calc["tv_avg"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "D29", round(calc["entry"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "J29", round(calc["tv_avg"], 1), kind="result", fmt="#,##0", overwrite_formula=of)

    put(ws, "N18", round(calc["tv_pg"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "N19", round(calc["tv_mult"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "N20", round(calc["tv_avg"], 1), kind="result", fmt="#,##0", overwrite_formula=of)

    put(ws, "D32", round(calc["ev"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "D33", CASH_10Q, kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "D34", DEBT_10Q, kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "D35", round(calc["equity"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "D37", round(calc["eq_ps"], 2), kind="result", fmt="#,##0.00", overwrite_formula=of)

    put(ws, "I32", round(calc["market_cap"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "I33", DEBT_10Q, kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "I34", CASH_10Q, kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "I35", round(calc["market_ev"], 1), kind="result", fmt="#,##0", overwrite_formula=of)
    put(ws, "I37", PRICE, kind="result", fmt="#,##0.00", overwrite_formula=of)

    put(ws, "N32", round(calc["upside_pct"], 4), kind="result", fmt="0.0%", overwrite_formula=of)
    put(
        ws,
        "N33",
        round(calc["irr"], 4) if calc["irr"] is not None else "n/a",
        kind="result",
        fmt="0.0%",
        overwrite_formula=of,
    )
    put(ws, "N36", PRICE, kind="result", fmt="#,##0.00", overwrite_formula=of)
    put(ws, "N37", round(calc["upside_abs"], 2), kind="result", fmt="#,##0.00", overwrite_formula=of)
    put(ws, "N38", round(calc["eq_ps"], 2), kind="result", fmt="#,##0.00", overwrite_formula=of)

    ws["B21"] = "EBIT (= Operating income path)"
    ws["B22"] = "Less: Cash Taxes"
    ws["B23"] = "Plus: D&A"
    ws["B24"] = "Less: Capex (grows w/ rev)"
    ws["B25"] = "Less: Changes in NWC"
    ws["B26"] = "Unlevered FCF"
    ws["L17"] = "Terminal Value"
    ws["L18"] = "Perpetual Growth"
    ws["L19"] = "EV/EBITDA"
    ws["L20"] = "Average"
    ws["D2"] = "FIXED: EBIT anchored to FY25 OpInc; years 2026–2030; FY25 base in row 16"
    ws["D2"].font = NOTE
    ws["B50"] = (
        f"Base FY25 OpInc={FY25_OPINC:,} D&A={FY25_DA:,} Capex={FY25_CAPEX:,}. "
        f"WACC={WACC:.1%} g={G_PERP:.0%} EV/EBITDA={EV_EBITDA:.0f}x. Source: {FILING_2025}"
    )
    ws["B50"].font = NOTE

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Unlevered Free Cash Flow (FY2026–2030)"
    chart.y_axis.title = "USD thousands"
    data = Reference(ws, min_col=5, min_row=26, max_col=9, max_row=26)
    cats = Reference(ws, min_col=5, min_row=17, max_col=9, max_row=17)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.width = 16
    chart.height = 9
    ws.add_chart(chart, "L3")
    widen(ws)


def add_errors_sheet(wb, fc, calc):
    if "Audit_Fixes" in wb.sheetnames:
        del wb["Audit_Fixes"]
    ws = wb.create_sheet("Audit_Fixes", 0)
    ws["A1"] = "Errors found and fixed"
    ws["A1"].font = TITLE
    rows = [
        ("3S forecast COGS%", "Was 37% (CFI sample)", f"Now {COGS_PCT:.2%} from FY25"),
        ("3S forecast SG&A/R&D", "Was $25k / $10k", "Now FY25 SG&A/R&D grown with revenue"),
        ("3S forecast tax", "Was 28%", f"Now {TAX_RATE:.2%} FY25 ETR"),
        ("3S AR / Inv days", "Was 18 / 73", f"Now ~{FY25_AR/FY25_REV*365:.0f} / 0"),
        ("3S Capex forecast", "Was ~15k sample", "Now from FY25 39,407 growing 8%/yr"),
        ("3S year headers", "Formulas blank pre-calc", "Hardcoded 2021–2030"),
        ("DCF Year-1 EBIT", "Started ~1,048,623 (>FY25 OpInc)", f"Path anchored to FY25 OpInc {FY25_OPINC:,}"),
        ("DCF D&A", "Used ~18k (0.8% rev)", f"Anchored to FY25 D&A {FY25_DA:,}"),
        ("DCF years", "Unclear / missing FY25 base", "FY25 actuals in row 16; forecast 2026–2030"),
        ("DCF Capex", "Flat 39,407", "Grows 8%/yr with investment"),
        ("DCF EV/EBITDA", "25x → huge TV vs PG", "Set 22x; still show PG / multiple / average"),
        ("DCF blanks", "Taxes/Capex/UFCF/TV formulas blank in viewer", "Written as computed values + chart"),
    ]
    for c, h in enumerate(["Area", "Error", "Fix"], 1):
        cell = ws.cell(3, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for i, row in enumerate(rows, start=4):
        for j, v in enumerate(row, start=1):
            ws.cell(i, j, v)
    ws["A18"] = "FY2025 10-K"
    ws["B18"] = FILING_2025
    ws["B18"].hyperlink = FILING_2025
    ws["B18"].font = LINK_FONT
    ws["A20"] = "Corrected forecast EBIT path (USD thousands)"
    for i, r in enumerate(fc, start=21):
        ws.cell(i, 1, r["year"])
        ws.cell(i, 2, round(r["ebit"]))
        ws.cell(i, 3, round(r["ufcf"]))
        ws.cell(i, 4, round(r["capex"]))
    ws["A21"].offset(-1, 0).value = "Year"
    ws.cell(20, 2, "EBIT")
    ws.cell(20, 3, "UFCF")
    ws.cell(20, 4, "Capex")
    for i, w in enumerate([28, 55, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_3s(s, fc):
    shutil.copy2(ORIG_3, OUT_3S)
    wb = load_workbook(OUT_3S)
    for ws in wb.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []
    inject_3s(wb["3 Statement Model"], s, fc)
    add_errors_sheet(wb, fc, None)
    # revenue chart
    ws = wb["3 Statement Model"]
    ws._charts = []
    chart = LineChart()
    chart.title = "FICO Revenue FY2021–2030"
    data = Reference(ws, min_col=5, min_row=24, max_col=14, max_row=24)
    cats = Reference(ws, min_col=5, min_row=2, max_col=14, max_row=2)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.width = 18
    chart.height = 9
    ws.add_chart(chart, "P3")
    if "Cover Page" in wb.sheetnames:
        wb["Cover Page"]["C12"] = "FICO 3-Statement — corrected forecast drivers"
    wb.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    wb.save(OUT_3S)
    print("Wrote", OUT_3S, OUT_3S.stat().st_size)


def build_dcf(calc, fc):
    shutil.copy2(ORIG_DCF, OUT_DCF)
    wb = load_workbook(OUT_DCF)
    for ws in wb.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []
    fill_dcf(wb["DCF Model"], calc, "FICO DCF Model (corrected)")
    add_errors_sheet(wb, fc, calc)
    if "Cover Page" in wb.sheetnames:
        wb["Cover Page"]["C12"] = "FICO DCF — corrected anchors & filled values"
    wb.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    wb.save(OUT_DCF)
    print("Wrote", OUT_DCF, OUT_DCF.stat().st_size)


def build_joined(s, fc, calc):
    wb = Workbook()
    wb.remove(wb.active)
    add_errors_sheet(wb, fc, calc)

    ins = wb.create_sheet("01_Instructions", 0)
    ins["A1"] = "FICO Joined 3-Statement + DCF (CORRECTED)"
    ins["A1"].font = TITLE
    for i, line in enumerate(
        [
            "",
            "See Audit_Fixes for every error that was identified and corrected.",
            "03_Three_Statement: FY2021–2025 10-K history + FY2026–2030 FICO-scale forecast drivers.",
            "04_DCF: EBIT/D&A/Capex/NWC path anchored to FY25 Operating income / D&A / Capex.",
            "Years are hardcoded values (not blank formulas). UFCF/TV/EV are filled numbers + chart.",
            "",
            f"FY25 OpInc={FY25_OPINC:,}  D&A={FY25_DA:,}  Capex={FY25_CAPEX:,}  Tax rate={TAX_RATE:.2%}",
            f"Intrinsic Eq/sh (model) ≈ {calc['eq_ps']:,.2f} vs price {PRICE:,.2f}",
            FILING_2025,
        ],
        start=2,
    ):
        ins[f"A{i}"] = line
    ins.column_dimensions["A"].width = 110

    # populate from originals
    tmp3 = FICO_DIR / "_t3.xlsx"
    tmpd = FICO_DIR / "_td.xlsx"
    shutil.copy2(ORIG_3, tmp3)
    shutil.copy2(ORIG_DCF, tmpd)
    wb3 = load_workbook(tmp3)
    wbd = load_workbook(tmpd)
    for w in (wb3, wbd):
        for ws in w.worksheets:
            if getattr(ws, "_images", None):
                ws._images = []
    inject_3s(wb3["3 Statement Model"], s, fc)
    fill_dcf(wbd["DCF Model"], calc, "FICO DCF (joined, corrected)")

    from copy import copy as _copy

    def copy_sheet(src, dst):
        for row in src.iter_rows():
            for cell in row:
                n = dst.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    n.font = _copy(cell.font)
                    n.fill = _copy(cell.fill)
                    n.border = _copy(cell.border)
                    n.number_format = cell.number_format
                    n.alignment = _copy(cell.alignment)
        for letter, dim in src.column_dimensions.items():
            dst.column_dimensions[letter].width = dim.width

    ws3 = wb.create_sheet("03_Three_Statement")
    copy_sheet(wb3["3 Statement Model"], ws3)
    # charts
    ws3._charts = []
    ch = LineChart()
    ch.title = "FICO Revenue FY2021–2030"
    ch.add_data(Reference(ws3, min_col=5, min_row=24, max_col=14, max_row=24), from_rows=True)
    ch.set_categories(Reference(ws3, min_col=5, min_row=2, max_col=14, max_row=2))
    ch.width = 18
    ch.height = 9
    ws3.add_chart(ch, "P3")

    wsd = wb.create_sheet("04_DCF")
    copy_sheet(wbd["DCF Model"], wsd)

    order = ["01_Instructions", "Audit_Fixes", "03_Three_Statement", "04_DCF"]
    # ensure Audit_Fixes exists at right place
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto", forceFullCalc=True)
    wb.save(OUT_JOINED)
    tmp3.unlink(missing_ok=True)
    tmpd.unlink(missing_ok=True)
    print("Wrote", OUT_JOINED, OUT_JOINED.stat().st_size)


def write_md(fc, calc):
    lines = [
        "# Model errors found and fixed",
        "",
        "## 3-Statement",
        "| Error | Fix |",
        "|---|---|",
        f"| Forecast COGS% was **37%** (CFI sample) | **{COGS_PCT:.2%}** from FY25 |",
        f"| Forecast SG&A/R&D was **$25k/$10k** | FY25 **{FY25_SGA:,.0f}/{FY25_RD:,.0f}** grown with revenue |",
        f"| Forecast tax was **28%** | **{TAX_RATE:.2%}** FY25 ETR |",
        "| AR/Inv days 18/73 | ~97 AR days, **0** inventory |",
        "| Capex forecast ~15k sample | From FY25 **39,407** growing 8%/yr |",
        "| Year headers blank (formulas) | Hardcoded **2021–2030** |",
        "",
        "## DCF",
        "| Error | Fix |",
        "|---|---|",
        f"| Year-1 EBIT ~1,048k > FY25 OpInc | Re-anchored to FY25 OpInc **{FY25_OPINC:,}** then grown |",
        f"| D&A ~18k vs FY25 14,952 | Anchored to **{FY25_DA:,}** × rev scale |",
        "| Missing FY25 base year display | Row 16 shows FY25 EBIT/D&A/Capex/Tax |",
        "| Capex flat | Grows 8%/yr |",
        "| EV/EBITDA 25x extreme vs PG TV | **22x**; still show PG / multiple / average |",
        "| Blank taxes/UFCF/TV in viewer | Computed values written + chart |",
        "",
        f"Intrinsic equity/share (model): **${calc['eq_ps']:,.2f}** vs price ${PRICE:,.2f}",
        "",
        f"10-K: {FILING_2025}",
        "",
        "## Forecast path (USD thousands)",
        "",
        "| Year | Growth | EBIT | Capex | UFCF |",
        "|---|---:|---:|---:|---:|",
    ]
    for r in fc:
        lines.append(
            f"| {r['year']} | {r['g']:.0%} | {r['ebit']:,.0f} | {r['capex']:,.0f} | {r['ufcf']:,.0f} |"
        )
    ERRORS_MD.write_text("\n".join(lines) + "\n")
    print("Wrote", ERRORS_MD)


def verify(s, fc, calc):
    wb = load_workbook(OUT_3S)
    ws = wb["3 Statement Model"]
    assert [ws[f"{c}2"].value for c in "EFGHI"] == HIST_YEARS
    assert [ws[f"{c}2"].value for c in "JKLMN"] == FC_YEARS
    assert abs(float(ws["J8"].value) - round(COGS_PCT, 4)) < 1e-6
    assert float(ws["J9"].value) > 100_000  # not 25000
    assert float(ws["J10"].value) > 50_000
    assert abs(float(ws["J13"].value) - round(TAX_RATE, 4)) < 1e-6
    assert float(ws["J16"].value) == 0
    assert float(ws["J18"].value) > 30_000
    print("3S OK years", HIST_YEARS[0], "-", FC_YEARS[-1], "J9 SG&A", ws["J9"].value, "J8 COGS%", ws["J8"].value)

    wb = load_workbook(OUT_DCF)
    d = wb["DCF Model"]
    assert d["E17"].value == 2026 and d["I17"].value == 2030
    assert float(d["E16"].value) == float(FY25_OPINC)
    assert abs(float(d["E21"].value) - round(fc[0]["ebit"], 1)) < 1
    assert abs(float(d["E23"].value) - round(fc[0]["da"], 1)) < 1
    assert float(d["E22"].value) > 0 and float(d["E26"].value) > 0
    assert float(d["N20"].value) > 0
    assert len(d._charts) >= 1
    print(
        "DCF OK FY25 EBIT base", d["E16"].value,
        "FY26 EBIT", d["E21"].value,
        "UFCF", d["E26"].value,
        "Eq/sh", d["D37"].value,
    )


def main():
    s = load_series()
    fc = forecast_path()
    calc = build_dcf_calc(fc)
    print("FY26 EBIT", round(fc[0]["ebit"]), "UFCF", round(fc[0]["ufcf"]))
    print("Eq/sh", round(calc["eq_ps"], 2), "TV avg", round(calc["tv_avg"]))
    write_md(fc, calc)
    build_3s(s, fc)
    build_dcf(calc, fc)
    build_joined(s, fc, calc)
    verify(s, fc, calc)


if __name__ == "__main__":
    main()
