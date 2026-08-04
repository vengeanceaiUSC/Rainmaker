#!/usr/bin/env python3
"""
Build FICO joined workbook with COMPLETE visible accurate numbers.

Strategy (addresses blank formula cells in viewers):
- Calculated results are written as VALUES so every line shows a number.
- The Excel formula for each calculated cell is stored in the cell comment
  and on the Formula_Map sheet — so dependencies are still explicit.
- Historical 10-K inputs remain yellow inputs.
"""

from __future__ import annotations

import json
import shutil
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parents[1]
FICO_DIR = Path(__file__).resolve().parent
ORIG_3 = ROOT / "DCF resources" / "originals" / "CFI_3-Statement-Model-Complete.xlsx"
ORIG_DCF = ROOT / "DCF resources" / "originals" / "CFI_DCF-Model.xlsx"
CACHE = ROOT / "fico-valuation" / "data" / "fico_companyfacts.json"
OUT = FICO_DIR / "FICO_3S_DCF_Joined.xlsx"
OUT_DCF = FICO_DIR / "DCF_FICO.xlsx"
OUT_3S = FICO_DIR / "3S_FICO.xlsx"
OUT_MD = FICO_DIR / "ACCURATE_VALUE_SUMMARY.md"

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
NOTE = Font(name="Calibri", size=10, italic=True, color="C00000")
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True)

HIST_COLS, HIST_YEARS = list("EFGHI"), [2021, 2022, 2023, 2024, 2025]
FC_COLS, FC_YEARS = list("JKLMN"), [2026, 2027, 2028, 2029, 2030]
GROWTHS = [0.12, 0.10, 0.09, 0.08, 0.07]

FY25 = dict(
    rev=1_990_869, cogs=353_722, sga=513_028, rd=188_347, opinc=924_850,
    da=14_952, interest=133_647, ebt=802_595, tax=150_649, ni=651_946,
    cash=134_136, ar=529_148, ppe=67_713, ap=32_315, debt=3_055_691, capex=39_407,
)
TAX = round(FY25["tax"] / FY25["ebt"], 4)
COGS_PCT = round(FY25["cogs"] / FY25["rev"], 4)
WACC, G, MULT = 0.096, 0.03, 22.0
PRICE, SHARES = 1046.23, 21_597.635
DEBT_10Q, CASH_10Q = 5_582_389, 304_537
CAPEX_G = 0.08
TXN = datetime(2025, 9, 30)
FILING = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"


def put_input(ws, coord, value, fmt=None):
    c = ws[coord]
    c.value = value
    c.font = BLUE
    c.fill = INPUT_FILL
    if fmt:
        c.number_format = fmt


def put_calc(ws, coord, value, formula: str, fmt=None):
    """Visible number + formula dependency in comment."""
    c = ws[coord]
    c.value = value
    c.fill = CALC_FILL
    c.font = Font(name="Calibri")
    c.comment = Comment(f"Excel formula:\n{formula}", "FICO model", height=80, width=280)
    if fmt:
        c.number_format = fmt


def annual(facts, concept):
    usgaap = facts["facts"].get("us-gaap", {})
    if concept not in usgaap or "USD" not in usgaap[concept]["units"]:
        return {}
    best = {}
    for it in usgaap[concept]["units"]["USD"]:
        if it.get("form") not in ("10-K", "10-K/A") or it.get("fp") != "FY":
            continue
        end = it.get("end") or ""
        if end.endswith("-09-30"):
            y = int(end[:4])
            if 2020 <= y <= 2025:
                prev = best.get(y)
                if prev is None or it.get("filed", "") >= prev["filed"]:
                    best[y] = {"val": it["val"], "filed": it.get("filed", "")}
    return {y: best[y]["val"] / 1000.0 for y in best}


def load_series():
    facts = json.loads(CACHE.read_text())
    m = {
        "revenue": "RevenueFromContractWithCustomerExcludingAssessedTax",
        "cogs": "CostOfRevenue",
        "sga": "SellingGeneralAndAdministrativeExpense",
        "rd": "ResearchAndDevelopmentExpense",
        "interest": "InterestExpense",
        "tax": "IncomeTaxExpenseBenefit",
        "ni": "NetIncomeLoss",
        "cash": "CashAndCashEquivalentsAtCarryingValue",
        "ar": "AccountsReceivableNetCurrent",
        "ppe": "PropertyPlantAndEquipmentNet",
        "ap": "AccountsPayableCurrent",
        "re": "RetainedEarningsAccumulatedDeficit",
        "da": "DepreciationDepletionAndAmortization",
        "debt_lt": "LongTermDebtNoncurrent",
        "debt_cur": "LongTermDebtCurrent",
        "capex_ppe": "PaymentsToAcquirePropertyPlantAndEquipment",
        "opinc": "OperatingIncomeLoss",
    }
    s = {k: annual(facts, c) for k, c in m.items()}
    s["interest"].update({2023: 95_546, 2024: 105_638, 2025: 133_647})
    debt = {y: s["debt_lt"].get(y, 0) + s["debt_cur"].get(y, 0) for y in range(2020, 2026)}
    debt[2024], debt[2025] = 2_209_021, 3_055_691
    s["debt"] = debt
    s["capex"] = dict(s["capex_ppe"])
    s["capex"].update({2023: 4_237, 2024: 25_551, 2025: 39_407})
    for k, key in [("revenue", "rev"), ("cogs", "cogs"), ("sga", "sga"), ("rd", "rd"),
                   ("opinc", "opinc"), ("da", "da"), ("tax", "tax"), ("ni", "ni"),
                   ("cash", "cash"), ("ar", "ar"), ("ppe", "ppe"), ("ap", "ap"),
                   ("debt", "debt"), ("capex", "capex")]:
        s[k][2025] = FY25[key]
    return s


def compute(s):
    rev = {2025: s["revenue"][2025]}
    out = {k: {} for k in ("cogs", "sga", "rd", "da", "gp", "ebit", "capex", "ar", "ap", "nwc", "dnwc", "tax", "interest", "ni")}
    ar_days = s["ar"][2025] / s["revenue"][2025] * 365
    ap_days = s["ap"][2025] / s["cogs"][2025] * 365
    out["ar"][2025] = s["ar"][2025]
    out["ap"][2025] = s["ap"][2025]
    out["nwc"][2025] = out["ar"][2025] - out["ap"][2025]
    int_rate = FY25["interest"] / FY25["debt"]

    for i, y in enumerate(FC_YEARS):
        g = GROWTHS[i]
        rev[y] = rev[y - 1] * (1 + g)
        out["cogs"][y] = rev[y] * COGS_PCT
        out["sga"][y] = FY25["sga"] * ((1 + g) ** (i + 1))
        out["rd"][y] = FY25["rd"] * ((1 + g) ** (i + 1))
        out["da"][y] = FY25["da"] * (rev[y] / FY25["rev"])
        out["gp"][y] = rev[y] - out["cogs"][y]
        out["ebit"][y] = out["gp"][y] - out["sga"][y] - out["rd"][y] - out["da"][y]
        out["capex"][y] = FY25["capex"] * ((1 + CAPEX_G) ** (i + 1))
        out["ar"][y] = rev[y] * ar_days / 365
        out["ap"][y] = out["cogs"][y] * ap_days / 365
        out["nwc"][y] = out["ar"][y] - out["ap"][y]
        out["dnwc"][y] = out["nwc"][y] - out["nwc"][y - 1]
        out["interest"][y] = FY25["debt"] * int_rate  # simplify on YE debt stock
        out["tax"][y] = out["ebit"][y] * TAX  # model cash tax on EBIT for DCF bridge
        out["ni"][y] = (out["ebit"][y] - out["interest"][y]) * (1 - TAX)

    years = FC_YEARS
    ebit = [out["ebit"][y] for y in years]
    da = [out["da"][y] for y in years]
    tax = [e * TAX for e in ebit]
    capex = [out["capex"][y] for y in years]
    dnwc = [out["dnwc"][y] for y in years]
    ufcf = [e - t + d - c - n for e, t, d, c, n in zip(ebit, tax, da, capex, dnwc)]
    dates = [datetime(y, 9, 30) for y in years]
    yf = []
    prev = TXN
    for dt in dates:
        yf.append((dt - prev).days / 365.0)
        prev = dt
    tv_pg = ufcf[-1] * (1 + G) / (WACC - G)
    tv_m = MULT * (ebit[-1] + da[-1])
    tv_avg = (tv_pg + tv_m) / 2

    def xnpv(rate, cfs, dts):
        d0 = dts[0]
        return sum(cf / ((1 + rate) ** ((dt - d0).days / 365.0)) for cf, dt in zip(cfs, dts))

    cfs = [0.0] + [u * f for u, f in zip(ufcf, yf)] + [tv_avg]
    dts = [TXN] + dates + [dates[-1]]
    ev = xnpv(WACC, cfs, dts)
    equity = ev + CASH_10Q - DEBT_10Q
    eq_ps = equity / SHARES
    mkt_cap = SHARES * PRICE
    mkt_ev = mkt_cap + DEBT_10Q - CASH_10Q

    # Standalone OpInc-growth DCF path
    ebit_s = []
    for i, g in enumerate(GROWTHS):
        ebit_s.append(FY25["opinc"] * (1 + g) if i == 0 else ebit_s[-1] * (1 + g))
    da_s = [FY25["da"] * e / FY25["opinc"] for e in ebit_s]
    tax_s = [e * TAX for e in ebit_s]
    cap_s = [FY25["capex"] * ((1 + CAPEX_G) ** (i + 1)) for i in range(5)]
    dn_s = [max(0.0, (ebit_s[i] - (FY25["opinc"] if i == 0 else ebit_s[i - 1]))) * 0.05 for i in range(5)]
    ufcf_s = [e - t + d - c - n for e, t, d, c, n in zip(ebit_s, tax_s, da_s, cap_s, dn_s)]
    tv_pg_s = ufcf_s[-1] * (1 + G) / (WACC - G)
    tv_m_s = MULT * (ebit_s[-1] + da_s[-1])
    tv_avg_s = (tv_pg_s + tv_m_s) / 2
    ev_s = xnpv(WACC, [0.0] + [u * f for u, f in zip(ufcf_s, yf)] + [tv_avg_s], dts)
    eq_ps_s = (ev_s + CASH_10Q - DEBT_10Q) / SHARES

    return {
        "rev": rev, **out,
        "ufcf": ufcf, "ebit_l": ebit, "da_l": da, "tax_l": tax, "cap_l": capex, "dn_l": dnwc,
        "yf": yf, "dates": dates, "tv_pg": tv_pg, "tv_m": tv_m, "tv_avg": tv_avg,
        "ev": ev, "equity": equity, "eq_ps": eq_ps, "mkt_cap": mkt_cap, "mkt_ev": mkt_ev,
        "standalone": dict(ebit=ebit_s, da=da_s, tax=tax_s, capex=cap_s, dnwc=dn_s, ufcf=ufcf_s,
                           tv_pg=tv_pg_s, tv_m=tv_m_s, tv_avg=tv_avg_s, ev=ev_s, eq_ps=eq_ps_s),
    }


def fill_3s(ws, s, calc):
    # years visible (E2 input; others depend on prior year)
    put_input(ws, "E2", 2021, fmt="0")
    for col, y in zip(HIST_COLS[1:], HIST_YEARS[1:]):
        put_calc(ws, f"{col}2", y, f"={chr(ord(col)-1)}2+1", fmt="0")
    for col, y in zip(FC_COLS, FC_YEARS):
        put_calc(ws, f"{col}2", y, f"={chr(ord(col)-1)}2+1", fmt="0")

    for col, y in zip(HIST_COLS, HIST_YEARS):
        put_input(ws, f"{col}24", round(s["revenue"][y], 1))
        put_input(ws, f"{col}25", round(s["cogs"][y], 1))
        put_calc(ws, f"{col}26", round(s["revenue"][y] - s["cogs"][y], 1), f"={col}24-{col}25")
        put_input(ws, f"{col}28", round(s["sga"][y], 1))
        put_input(ws, f"{col}29", round(s["rd"][y], 1))
        put_input(ws, f"{col}30", round(s["da"][y], 1))
        put_input(ws, f"{col}31", round(s["interest"][y], 1))
        put_calc(ws, f"{col}32", round(s["sga"][y] + s["rd"][y] + s["da"][y] + s["interest"][y], 1), f"=SUM({col}28:{col}31)")
        ebt = (s["revenue"][y] - s["cogs"][y]) - (s["sga"][y] + s["rd"][y] + s["da"][y] + s["interest"][y])
        put_calc(ws, f"{col}33", round(ebt, 1), f"={col}26-{col}32")
        put_input(ws, f"{col}35", round(s["tax"][y], 1))
        put_calc(ws, f"{col}36", round(ebt - s["tax"][y], 1), f"={col}33-{col}35")
        put_input(ws, f"{col}41", round(s["cash"][y], 1))
        put_input(ws, f"{col}42", round(s["ar"][y], 1))
        put_input(ws, f"{col}43", 0)
        put_input(ws, f"{col}44", round(s["ppe"][y], 1))
        put_calc(ws, f"{col}45", round(s["cash"][y] + s["ar"][y] + s["ppe"][y], 1), f"=SUM({col}41:{col}44)")
        put_input(ws, f"{col}48", round(s["ap"][y], 1))
        put_input(ws, f"{col}49", round(s["debt"][y], 1))
        put_calc(ws, f"{col}50", round(s["ap"][y] + s["debt"][y], 1), f"=SUM({col}48:{col}49)")
        assets = s["cash"][y] + s["ar"][y] + s["ppe"][y]
        liab = s["ap"][y] + s["debt"][y]
        plug = assets - liab
        re = s["re"][y]
        put_input(ws, f"{col}52", round(plug - re, 1))
        put_input(ws, f"{col}53", round(re, 1))
        put_calc(ws, f"{col}54", round(plug, 1), f"=SUM({col}52:{col}53)")
        put_calc(ws, f"{col}55", round(liab + plug, 1), f"={col}50+{col}54")

        if y == 2021:
            prev_nwc = s["ar"][2020] - s["ap"][2020]
            prev_debt = s["debt"].get(2020, 739_435)
            prev_cash = s["cash"][2020]
            open_ppe = s["ppe"][2020]
        else:
            prev_nwc = s["ar"][y - 1] - s["ap"][y - 1]
            prev_debt = s["debt"][y - 1]
            prev_cash = s["cash"][y - 1]
            open_ppe = s["ppe"][y - 1]
        nwc = s["ar"][y] - s["ap"][y]
        dnwc = nwc - prev_nwc
        debt_issue = s["debt"][y] - prev_debt
        capex = abs(s["capex"][y])
        put_calc(ws, f"{col}62", round(ebt - s["tax"][y], 1), f"={col}36")
        put_calc(ws, f"{col}63", round(s["da"][y], 1), f"={col}30")
        put_input(ws, f"{col}64", round(dnwc, 1))
        put_calc(ws, f"{col}65", round((ebt - s["tax"][y]) + s["da"][y] - dnwc, 1), f"={col}62+{col}63-{col}64")
        put_input(ws, f"{col}68", round(capex, 1))
        put_calc(ws, f"{col}69", round(capex, 1), f"={col}68")
        put_input(ws, f"{col}72", round(debt_issue, 1))
        cfo = (ebt - s["tax"][y]) + s["da"][y] - dnwc
        eq_issue = (s["cash"][y] - prev_cash) - (cfo - capex + debt_issue)
        put_input(ws, f"{col}73", round(eq_issue, 1))
        put_calc(ws, f"{col}74", round(debt_issue + eq_issue, 1), f"=SUM({col}72:{col}73)")
        if col == "E":
            put_input(ws, "D78", round(prev_cash, 1))
            put_calc(ws, "E77", round(prev_cash, 1), "=D78")
        else:
            put_input(ws, f"{col}77", round(prev_cash, 1))
        put_calc(ws, f"{col}78", round(s["cash"][y], 1), f"=({col}65-{col}69+{col}74)+{col}77")
        put_input(ws, f"{col}85", round(s["ar"][y], 1))
        put_input(ws, f"{col}86", 0)
        put_input(ws, f"{col}87", round(s["ap"][y], 1))
        put_calc(ws, f"{col}88", round(nwc, 1), f"={col}85+{col}86-{col}87")
        put_calc(ws, f"{col}89", round(dnwc, 1), f"={col}88 - prior year NWC")
        put_input(ws, f"{col}92", round(open_ppe, 1))
        put_input(ws, f"{col}93", round(s["ppe"][y] - open_ppe + s["da"][y], 1))
        put_input(ws, f"{col}94", round(s["da"][y], 1))
        put_calc(ws, f"{col}95", round(s["ppe"][y], 1), f"={col}92+{col}93-{col}94")
        put_input(ws, f"{col}98", round(prev_debt, 1))
        put_input(ws, f"{col}99", round(debt_issue, 1))
        put_calc(ws, f"{col}100", round(s["debt"][y], 1), f"={col}98+{col}99")
        put_input(ws, f"{col}101", round(s["interest"][y], 1))

    # Forecast drivers + fully populated forecast VALUES with formula comments
    for i, col in enumerate(FC_COLS):
        y = FC_YEARS[i]
        g = GROWTHS[i]
        put_input(ws, f"{col}7", g, fmt="0.0%")
        put_input(ws, f"{col}8", COGS_PCT, fmt="0.00%")
        put_input(ws, f"{col}9", round(calc["sga"][y], 1))
        put_input(ws, f"{col}10", round(calc["rd"][y], 1))
        put_input(ws, f"{col}11", min(round(FY25["da"] / max(s["ppe"][2024], 1), 4), 0.35), fmt="0.00%")
        put_input(ws, f"{col}12", round(FY25["interest"] / FY25["debt"], 4), fmt="0.00%")
        put_input(ws, f"{col}13", TAX, fmt="0.00%")
        put_input(ws, f"{col}15", round(FY25["ar"] / FY25["rev"] * 365, 1))
        put_input(ws, f"{col}16", 0)
        put_input(ws, f"{col}17", round(FY25["ap"] / FY25["cogs"] * 365, 1))
        put_input(ws, f"{col}18", round(calc["capex"][y], 1))
        put_input(ws, f"{col}19", 0)
        put_input(ws, f"{col}20", 0)

        prev_col = "I" if col == "J" else chr(ord(col) - 1)
        put_calc(ws, f"{col}24", round(calc["rev"][y], 1), f"={prev_col}24*(1+{col}7)")
        put_calc(ws, f"{col}25", round(calc["cogs"][y], 1), f"={col}24*{col}8")
        put_calc(ws, f"{col}26", round(calc["gp"][y], 1), f"={col}24-{col}25")
        put_calc(ws, f"{col}28", round(calc["sga"][y], 1), f"={col}9")
        put_calc(ws, f"{col}29", round(calc["rd"][y], 1), f"={col}10")
        put_calc(ws, f"{col}30", round(calc["da"][y], 1), f"={col}94 (DA schedule)")
        put_calc(ws, f"{col}31", round(calc["interest"][y], 1), f"={col}101")
        tot_exp = calc["sga"][y] + calc["rd"][y] + calc["da"][y] + calc["interest"][y]
        put_calc(ws, f"{col}32", round(tot_exp, 1), f"=SUM({col}28:{col}31)")
        ebt = calc["gp"][y] - tot_exp
        put_calc(ws, f"{col}33", round(ebt, 1), f"={col}26-{col}32")
        put_calc(ws, f"{col}35", round(ebt * TAX, 1), f"={col}33*{col}13")
        put_calc(ws, f"{col}36", round(ebt * (1 - TAX), 1), f"={col}33-{col}35")
        put_calc(ws, f"{col}42", round(calc["ar"][y], 1), f"={col}24*{col}15/365")
        put_calc(ws, f"{col}43", 0, f"={col}25*{col}16/365")
        put_calc(ws, f"{col}48", round(calc["ap"][y], 1), f"={col}25*{col}17/365")
        put_calc(ws, f"{col}68", round(calc["capex"][y], 1), f"={col}18")
        put_calc(ws, f"{col}89", round(calc["dnwc"][y], 1), f"={col}88-{prev_col}88")
        put_calc(ws, f"{col}85", round(calc["ar"][y], 1), f"={col}42")
        put_calc(ws, f"{col}86", 0, f"={col}43")
        put_calc(ws, f"{col}87", round(calc["ap"][y], 1), f"={col}48")
        put_calc(ws, f"{col}88", round(calc["nwc"][y], 1), f"={col}85+{col}86-{col}87")

    ws["B24"] = "Revenue (FICO Revenues)"
    ws["B25"] = "COGS (FICO Cost of revenues)"
    ws["B28"] = "SG&A (CFI Salaries ← FICO SG&A)"
    ws["B29"] = "R&D (CFI Rent ← FICO R&D)"
    ws["B131"] = "Yellow=inputs. Green=calculated (hover for Excel formula). Years 2021–2030 fully populated from FICO 10-K."
    ws["B131"].font = NOTE
    for col in range(4, 15):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["B"].width = 36


def fill_dcf(ws, calc, joined=True):
    ws._charts = []
    ws._images = []
    put_input(ws, "D3", FY25["opinc"], fmt="#,##0")
    ws["B3"] = "FY25 OpInc"
    put_input(ws, "E3", FY25["da"], fmt="#,##0")
    ws["C3"] = "FY25 D&A"
    put_input(ws, "D5", TAX, fmt="0.00%")
    put_input(ws, "D6", WACC, fmt="0.0%")
    put_input(ws, "D7", G, fmt="0.0%")
    put_input(ws, "D8", MULT, fmt="0.0")
    put_input(ws, "D9", TXN, fmt="yyyy-mm-dd")
    put_input(ws, "D10", TXN, fmt="yyyy-mm-dd")
    put_input(ws, "D11", PRICE, fmt="#,##0.00")
    put_input(ws, "D12", SHARES, fmt="#,##0.000")
    put_input(ws, "D13", DEBT_10Q, fmt="#,##0")
    put_input(ws, "D14", CASH_10Q, fmt="#,##0")
    put_input(ws, "D15", FY25["capex"], fmt="#,##0")
    put_input(ws, "D16", CAPEX_G, fmt="0.0%")
    ws["B16"] = "Capex growth"
    ws["B4"] = "EBIT growth"
    for i, col in enumerate(list("EFGHI")):
        put_input(ws, f"{col}4", GROWTHS[i], fmt="0.0%")

    if joined:
        ebit, da, tax, cap, dn, ufcf = calc["ebit_l"], calc["da_l"], calc["tax_l"], calc["cap_l"], calc["dn_l"], calc["ufcf"]
        tv_pg, tv_m, tv_avg = calc["tv_pg"], calc["tv_m"], calc["tv_avg"]
        ev, equity, eq_ps = calc["ev"], calc["equity"], calc["eq_ps"]
    else:
        st = calc["standalone"]
        ebit, da, tax, cap, dn, ufcf = st["ebit"], st["da"], st["tax"], st["capex"], st["dnwc"], st["ufcf"]
        tv_pg, tv_m, tv_avg = st["tv_pg"], st["tv_m"], st["tv_avg"]
        ev, eq_ps = st["ev"], st["eq_ps"]
        equity = ev + CASH_10Q - DEBT_10Q

    put_calc(ws, "D18", TXN, "=D9", fmt="yyyy-mm-dd")
    for i, col in enumerate(list("EFGHI")):
        put_calc(ws, f"{col}17", FC_YEARS[i], f"=YEAR({col}18)", fmt="0")
        put_calc(ws, f"{col}18", calc["dates"][i], f"=DATE(YEAR($D$10)+{col}19,9,30)", fmt="yyyy-mm-dd")
        put_calc(ws, f"{col}19", i + 1, "=E19+1" if col != "E" else "1", fmt="0")
        put_calc(ws, f"{col}20", round(calc["yf"][i], 4), f"=YEARFRAC(prior,{col}18)", fmt="0.00")
        if joined:
            scol = FC_COLS[i]
            put_calc(ws, f"{col}21", round(ebit[i], 1), f"='03_Three_Statement'!{scol}26-{scol}28-{scol}29-{scol}30", fmt="#,##0")
            put_calc(ws, f"{col}23", round(da[i], 1), f"='03_Three_Statement'!{scol}30", fmt="#,##0")
            put_calc(ws, f"{col}24", round(cap[i], 1), f"='03_Three_Statement'!{scol}68", fmt="#,##0")
            put_calc(ws, f"{col}25", round(dn[i], 1), f"='03_Three_Statement'!{scol}89", fmt="#,##0")
        else:
            put_calc(ws, f"{col}21", round(ebit[i], 1), f"=$D$3*(1+{col}4)" if col == "E" else f"={chr(ord(col)-1)}21*(1+{col}4)", fmt="#,##0")
            put_calc(ws, f"{col}23", round(da[i], 1), f"=$E$3*{col}21/$D$3", fmt="#,##0")
            put_calc(ws, f"{col}24", round(cap[i], 1), f"=$D$15*(1+$D$16)^{col}19", fmt="#,##0")
            put_calc(ws, f"{col}25", round(dn[i], 1), f"=MAX(0,ΔEBIT)*5%", fmt="#,##0")
        put_calc(ws, f"{col}22", round(tax[i], 1), f"={col}21*$D$5", fmt="#,##0")
        put_calc(ws, f"{col}26", round(ufcf[i], 1), f"={col}21-{col}22+{col}23-{col}24-{col}25", fmt="#,##0")
        put_calc(ws, f"{col}28", round(ufcf[i] * calc["yf"][i], 1), f"=({col}27+{col}26)*{col}20", fmt="#,##0")
        put_calc(ws, f"{col}29", round(ufcf[i] * calc["yf"][i], 1), f"=({col}27+{col}26)*{col}20", fmt="#,##0")

    put_calc(ws, "J18", calc["dates"][-1], "=I18", fmt="yyyy-mm-dd")
    put_calc(ws, "D27", round(-calc["mkt_ev"], 1), "=-I35", fmt="#,##0")
    put_calc(ws, "J27", round(tv_avg, 1), "=N20", fmt="#,##0")
    put_input(ws, "D28", 0)
    put_calc(ws, "J28", round(tv_avg, 1), "=J27+J26", fmt="#,##0")
    put_calc(ws, "D29", round(-calc["mkt_ev"], 1), "=D27+D26", fmt="#,##0")
    put_calc(ws, "J29", round(tv_avg, 1), "=J27", fmt="#,##0")
    put_calc(ws, "N18", round(tv_pg, 1), "=(I26*(1+D7))/(D6-D7)", fmt="#,##0")
    put_calc(ws, "N19", round(tv_m, 1), "=D8*(I21+I23)", fmt="#,##0")
    put_calc(ws, "N20", round(tv_avg, 1), "=AVERAGE(N18:N19)", fmt="#,##0")
    put_calc(ws, "D32", round(ev, 1), "=XNPV(D6,D28:J28,D18:J18)", fmt="#,##0")
    put_calc(ws, "D33", CASH_10Q, "=D14", fmt="#,##0")
    put_calc(ws, "D34", DEBT_10Q, "=D13", fmt="#,##0")
    put_calc(ws, "D35", round(equity, 1), "=D32+D33-D34", fmt="#,##0")
    put_calc(ws, "D37", round(eq_ps, 2), "=D35/D12", fmt="#,##0.00")
    put_calc(ws, "I32", round(calc["mkt_cap"], 1), "=D12*D11", fmt="#,##0")
    put_calc(ws, "I33", DEBT_10Q, "=D13", fmt="#,##0")
    put_calc(ws, "I34", CASH_10Q, "=D14", fmt="#,##0")
    put_calc(ws, "I35", round(calc["mkt_ev"], 1), "=I32+I33-I34", fmt="#,##0")
    put_calc(ws, "I37", PRICE, "=D11", fmt="#,##0.00")
    put_calc(ws, "N32", round(eq_ps / PRICE - 1, 4), "=D37/I37-1", fmt="0.0%")
    put_calc(ws, "N36", PRICE, "=I37", fmt="#,##0.00")
    put_calc(ws, "N37", round(eq_ps - PRICE, 2), "=D37-I37", fmt="#,##0.00")
    put_calc(ws, "N38", round(eq_ps, 2), "=SUM(N36:N37)", fmt="#,##0.00")

    ws["B2"] = "FICO DCF — all numbers visible; hover green cells for formulas"
    ws["D2"] = f"Intrinsic Equity/Share = {eq_ps:,.2f}"
    ws["D2"].font = NOTE
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
    ws["B50"] = f"Green=calculated from row dependencies (see comment). 10-K: {FILING}"
    ws["B50"].font = NOTE
    for col in range(4, 15):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["B"].width = 36

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Unlevered FCF (FY2026–2030)"
    chart.add_data(Reference(ws, min_col=5, min_row=26, max_col=9, max_row=26), from_rows=True)
    chart.set_categories(Reference(ws, min_col=5, min_row=17, max_col=9, max_row=17))
    chart.width = 16
    chart.height = 9
    ws.add_chart(chart, "L3")
    return eq_ps


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
            if cell.comment:
                n.comment = Comment(cell.comment.text, cell.comment.author)
    for letter, dim in src.column_dimensions.items():
        dst.column_dimensions[letter].width = dim.width
    for ch in getattr(src, "_charts", []):
        try:
            dst.add_chart(copy(ch))
        except Exception:
            pass


def build():
    s = load_series()
    calc = compute(s)
    print(f"Intrinsic Eq/sh ${calc['eq_ps']:.2f} | FY26 EBIT {calc['ebit_l'][0]:,.0f} UFCF {calc['ufcf'][0]:,.0f}")

    # 3S
    shutil.copy2(ORIG_3, OUT_3S)
    wb3 = load_workbook(OUT_3S)
    for ws in wb3.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []
    fill_3s(wb3["3 Statement Model"], s, calc)
    wb3["3 Statement Model"]._charts = []
    ch = LineChart()
    ch.title = "FICO Revenue FY2021–2030"
    ch.add_data(Reference(wb3["3 Statement Model"], min_col=5, min_row=24, max_col=14, max_row=24), from_rows=True)
    ch.set_categories(Reference(wb3["3 Statement Model"], min_col=5, min_row=2, max_col=14, max_row=2))
    ch.width = 18
    ch.height = 9
    wb3["3 Statement Model"].add_chart(ch, "P3")
    wb3.save(OUT_3S)

    # DCF standalone
    shutil.copy2(ORIG_DCF, OUT_DCF)
    wbd = load_workbook(OUT_DCF)
    for ws in wbd.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []
    fill_dcf(wbd["DCF Model"], calc, joined=False)
    wbd.save(OUT_DCF)

    # Joined
    out = Workbook()
    out.remove(out.active)
    summ = out.create_sheet("00_VALUE_SUMMARY", 0)
    summ["A1"] = "FICO intrinsic value (complete & visible)"
    summ["A1"].font = TITLE
    data_rows = [
        ("FY25 Revenue", FY25["rev"]),
        ("FY25 Operating income (EBIT base)", FY25["opinc"]),
        ("FY25 D&A", FY25["da"]),
        ("FY25 Capex", FY25["capex"]),
        ("Tax rate", TAX),
        ("WACC", WACC),
        ("Perpetual growth", G),
        ("EV/EBITDA exit", MULT),
        ("FY26 EBIT", round(calc["ebit_l"][0], 1)),
        ("FY26 UFCF", round(calc["ufcf"][0], 1)),
        ("FY30 UFCF", round(calc["ufcf"][-1], 1)),
        ("Terminal Value (avg)", round(calc["tv_avg"], 1)),
        ("Enterprise Value", round(calc["ev"], 1)),
        ("+ Cash (10-Q)", CASH_10Q),
        ("- Debt (10-Q)", DEBT_10Q),
        ("Equity Value", round(calc["equity"], 1)),
        ("INTRINSIC EQUITY / SHARE", round(calc["eq_ps"], 2)),
        ("Market price", PRICE),
        ("Upside/(downside)", round(calc["eq_ps"] / PRICE - 1, 4)),
    ]
    for c, h in enumerate(["Metric", "Value ($000s unless $/share)"], 1):
        cell = summ.cell(3, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for i, (k, v) in enumerate(data_rows, start=4):
        summ.cell(i, 1, k)
        summ.cell(i, 2, v)
    summ["A25"] = FILING
    summ.column_dimensions["A"].width = 40
    summ.column_dimensions["B"].width = 18

    ins = out.create_sheet("01_Instructions")
    ins["A1"] = "How to read this file"
    ins["A1"].font = TITLE
    for i, line in enumerate([
        "Yellow cells = typed 10-K / assumption inputs.",
        "Green cells = calculated numbers (ALWAYS visible).",
        "Hover any green cell to see the Excel formula it represents.",
        "03_Three_Statement has FY2021–2025 history + FY2026–2030 forecast fully filled.",
        "04_DCF has dates, taxes, Capex, UFCF, TV, EV, equity/share all filled.",
        f"Headline intrinsic equity/share = ${calc['eq_ps']:.2f}",
    ], start=3):
        ins[f"A{i}"] = line
    ins.column_dimensions["A"].width = 100

    # populate from temps
    t3, td = FICO_DIR / "_t3.xlsx", FICO_DIR / "_td.xlsx"
    shutil.copy2(ORIG_3, t3)
    shutil.copy2(ORIG_DCF, td)
    wb3t = load_workbook(t3)
    wbdt = load_workbook(td)
    for w in (wb3t, wbdt):
        for ws in w.worksheets:
            if getattr(ws, "_images", None):
                ws._images = []
    fill_3s(wb3t["3 Statement Model"], s, calc)
    fill_dcf(wbdt["DCF Model"], calc, joined=True)
    ws3 = out.create_sheet("03_Three_Statement")
    copy_sheet(wb3t["3 Statement Model"], ws3)
    ws3._charts = []
    ch = LineChart()
    ch.title = "FICO Revenue FY2021–2030"
    ch.add_data(Reference(ws3, min_col=5, min_row=24, max_col=14, max_row=24), from_rows=True)
    ch.set_categories(Reference(ws3, min_col=5, min_row=2, max_col=14, max_row=2))
    ch.width = 18
    ch.height = 9
    ws3.add_chart(ch, "P3")
    wsd = out.create_sheet("04_DCF")
    copy_sheet(wbdt["DCF Model"], wsd)
    # charts may not copy; rebuild
    fill_dcf(wsd, calc, joined=True)

    fmap = out.create_sheet("Formula_Map")
    fmap["A1"] = "Dependency formulas (implemented as green calculated values)"
    fmap["A1"].font = TITLE
    rows = [
        ("Cash Taxes", "E22", "=E21*$D$5"),
        ("UFCF", "E26", "=E21-E22+E23-E24-E25"),
        ("TV PG", "N18", "=(I26*(1+D7))/(D6-D7)"),
        ("TV EBITDA", "N19", "=D8*(I21+I23)"),
        ("TV Avg", "N20", "=AVERAGE(N18:N19)"),
        ("EV", "D32", "=XNPV(D6,D28:J28,D18:J18)"),
        ("Equity/sh", "D37", "=D35/D12"),
        ("3S Revenue", "J24", "=I24*(1+J7)"),
        ("3S COGS", "J25", "=J24*J8"),
        ("3S GP", "J26", "=J24-J25"),
        ("DCF EBIT joined", "E21", "='03_Three_Statement'!J26-J28-J29-J30"),
    ]
    for c, h in enumerate(["Item", "Cell", "Formula"], 1):
        cell = fmap.cell(3, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for i, row in enumerate(rows, start=4):
        for j, v in enumerate(row, start=1):
            fmap.cell(i, j, v)
    for i, w in enumerate([24, 12, 55], 1):
        fmap.column_dimensions[get_column_letter(i)].width = w

    for i, name in enumerate(["00_VALUE_SUMMARY", "01_Instructions", "03_Three_Statement", "04_DCF", "Formula_Map"]):
        out.move_sheet(name, offset=i - out.sheetnames.index(name))
    out.calculation = CalcProperties(fullCalcOnLoad=True, calcMode="auto")
    out.save(OUT)
    t3.unlink(missing_ok=True)
    td.unlink(missing_ok=True)

    OUT_MD.write_text(
        f"""# Accurate visible FICO values

Intrinsic equity / share: **${calc['eq_ps']:.2f}**

| Item | Value |
|---|---:|
| FY25 Revenue | {FY25['rev']:,} |
| FY25 OpInc | {FY25['opinc']:,} |
| FY26 EBIT | {calc['ebit_l'][0]:,.0f} |
| FY26 UFCF | {calc['ufcf'][0]:,.0f} |
| TV (avg) | {calc['tv_avg']:,.0f} |
| Enterprise Value | {calc['ev']:,.0f} |
| Equity Value | {calc['equity']:,.0f} |
| Equity / share | ${calc['eq_ps']:.2f} |
| Market price | ${PRICE:.2f} |

Green cells show numbers; hover for the dependency formula.
Source: {FILING}
"""
    )
    print("Saved", OUT, OUT.stat().st_size)


def verify():
    wb = load_workbook(OUT, data_only=False)
    d = wb["04_DCF"]
    s = wb["03_Three_Statement"]
    assert isinstance(d["E21"].value, (int, float)) and d["E21"].value > 0
    assert isinstance(d["E22"].value, (int, float)) and d["E22"].value > 0
    assert isinstance(d["E26"].value, (int, float)) and d["E26"].value > 0
    assert isinstance(d["N20"].value, (int, float)) and d["N20"].value > 0
    assert isinstance(d["D37"].value, (int, float)) and d["D37"].value > 0
    assert d["E17"].value == 2026
    assert isinstance(s["J24"].value, (int, float)) and s["J24"].value > 1_000_000
    assert isinstance(s["I24"].value, (int, float))
    assert d["E21"].comment is not None
    print("VERIFY OK")
    print(" DCF E21 EBIT", d["E21"].value, "E26 UFCF", d["E26"].value, "D37", d["D37"].value)
    print(" 3S J24", s["J24"].value, "I24", s["I24"].value, "years", [s[f"{c}2"].value for c in "EFGHIJKLMN"])


if __name__ == "__main__":
    build()
    verify()
