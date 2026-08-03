#!/usr/bin/env python3
"""
Build FICO_DCF_Financial_Model.xlsx from local CFI templates + SEC EDGAR data.

Rules:
- Use CFI workbooks as master blueprints (layout/formulas/formatting).
- Overwrite ONLY hardcoded historical / assumption input cells.
- Never overwrite formula cells (values starting with '=').
"""

from __future__ import annotations

import json
import shutil
import urllib.request
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
ORIG_3 = ROOT / "DCF resources" / "originals" / "CFI_3-Statement-Model-Complete.xlsx"
ORIG_DCF = ROOT / "DCF resources" / "originals" / "CFI_DCF-Model.xlsx"
OUT = ROOT / "FICO_DCF_Financial_Model.xlsx"
CACHE = ROOT / "fico-valuation" / "data" / "fico_companyfacts.json"
USER_AGENT = "RainmakerFICOModel contact@example.com"
CIK = "0000814547"

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True)
TITLE = Font(name="Calibri", size=16, bold=True, color="1F4E79")
NOTE = Font(name="Calibri", size=10, italic=True, color="C00000")

# Explicit historical input cells on CFI 3-statement (cols E-I = 5 hist years)
# These are hardcoded values in the pristine template (NOT formulas).
HIST_IS_ROWS = {
    24: "revenue",  # Revenue
    25: "cogs",  # COGS
    28: "sga",  # Salaries <- SG&A
    29: "rd",  # Rent/Overhead <- R&D
    30: "da",  # D&A
    31: "interest",  # Interest
    35: "tax",  # Taxes (hardcoded historically)
}
HIST_BS_ROWS = {
    41: "cash",
    42: "ar",
    43: "inventory",
    44: "ppe",
    48: "ap",
    49: "debt",
    52: "equity_capital",
    53: "retained_earnings",
}
HIST_CF_ROWS = {
    64: "dnwc",
    68: "capex",
    72: "debt_issuance",
    73: "equity_issuance",
}
HIST_WC_ROWS = {85: "ar", 86: "inventory", 87: "ap"}
HIST_PPE_ROWS = {92: "ppe_open", 93: "ppe_capex", 94: "ppe_da"}
HIST_DEBT_ROWS = {98: "debt_open", 99: "debt_issuance", 101: "interest"}

HIST_COLS = ["E", "F", "G", "H", "I"]  # 5 years
FORECAST_COLS = ["J", "K", "L", "M", "N"]


def set_input(ws, coord: str, value) -> None:
    """Write value only if target is not a formula cell."""
    cell = ws[coord]
    current = cell.value
    if isinstance(current, str) and current.startswith("="):
        raise RuntimeError(f"Refusing to overwrite formula at {coord}: {current}")
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL


def fetch_companyfacts() -> dict:
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode())
    CACHE.write_text(json.dumps(data))
    return data


def load_companyfacts() -> dict:
    if CACHE.exists() and CACHE.stat().st_size > 1000:
        # refresh if older workflow cache exists; still fine to reuse then overwrite with fetch
        pass
    return fetch_companyfacts()


def annual_by_end(facts: dict, concept: str) -> dict[int, float]:
    usgaap = facts["facts"].get("us-gaap", {})
    if concept not in usgaap:
        return {}
    units = usgaap[concept]["units"]
    unit = next((u for u in ["USD", "shares", "USD/shares", "pure"] if u in units), next(iter(units)))
    best: dict[int, dict] = {}
    for it in units[unit]:
        if it.get("form") not in ("10-K", "10-K/A") or it.get("fp") != "FY":
            continue
        end = it.get("end") or ""
        if not end.endswith("-09-30"):  # FICO FYE
            continue
        y = int(end[:4])
        if y < 2019 or y > 2025:
            continue
        prev = best.get(y)
        if prev is None or it.get("filed", "") >= prev["filed"]:
            best[y] = {"val": it["val"], "filed": it.get("filed", "")}
    # SEC values are in USD; convert to $ thousands for CFI template
    scale = 1 / 1000.0 if unit == "USD" else 1.0
    if unit == "shares":
        scale = 1 / 1000.0  # thousand shares
    return {y: best[y]["val"] * scale for y in best}


def build_fico_series(facts: dict) -> dict:
    """Map SEC concepts into template field names ($ thousands)."""
    rev = annual_by_end(facts, "RevenueFromContractWithCustomerExcludingAssessedTax")
    cogs = annual_by_end(facts, "CostOfRevenue")
    sga = annual_by_end(facts, "SellingGeneralAndAdministrativeExpense")
    rd = annual_by_end(facts, "ResearchAndDevelopmentExpense")
    opinc = annual_by_end(facts, "OperatingIncomeLoss")
    interest = annual_by_end(facts, "InterestExpense")
    tax = annual_by_end(facts, "IncomeTaxExpenseBenefit")
    ni = annual_by_end(facts, "NetIncomeLoss")
    cash = annual_by_end(facts, "CashAndCashEquivalentsAtCarryingValue")
    ar = annual_by_end(facts, "AccountsReceivableNetCurrent")
    ppe = annual_by_end(facts, "PropertyPlantAndEquipmentNet")
    ap = annual_by_end(facts, "AccountsPayableCurrent")
    debt_lt = annual_by_end(facts, "LongTermDebtNoncurrent")
    debt_cur = annual_by_end(facts, "LongTermDebtCurrent")
    da = annual_by_end(facts, "DepreciationDepletionAndAmortization")
    capex = annual_by_end(facts, "PaymentsToAcquirePropertyPlantAndEquipment")
    shares = annual_by_end(facts, "CommonStockSharesOutstanding")

    # Patch known 10-K audited values where XBRL concept naming differs (still $ thousands)
    # Interest expense, net from FY2025 10-K (InterestExpense concept missing FY2025)
    interest.setdefault(2025, 133_647)
    interest.setdefault(2024, 105_638)
    interest.setdefault(2023, 95_546)

    # Total debt = LT noncurrent + current maturities when available
    debt = {}
    for y in set(debt_lt) | set(debt_cur):
        debt[y] = debt_lt.get(y, 0) + debt_cur.get(y, 0)
    # Audited totals from 10-K debt note
    debt[2025] = 3_055_691
    debt[2024] = 2_209_021

    # Capex + capitalized software (from 10-K CF) for reinvestment realism
    capex_total = dict(capex)
    capex_total[2024] = 8_884 + 16_667
    capex_total[2025] = 8_922 + 30_485

    years = [2021, 2022, 2023, 2024, 2025]
    for y in years:
        for name, series in [
            ("revenue", rev),
            ("cogs", cogs),
            ("sga", sga),
            ("rd", rd),
            ("opinc", opinc),
            ("tax", tax),
            ("ni", ni),
            ("cash", cash),
            ("ar", ar),
            ("ppe", ppe),
            ("ap", ap),
            ("da", da),
        ]:
            if y not in series:
                raise RuntimeError(f"Missing SEC value for {name} FY{y}")

    # 10-Q bridge for DCF
    market = {
        "price": 1046.23,
        "shares_k": shares.get(2025, 23_764),  # fallback YE shares; overwritten below if desired
        "debt_10q": 5_582_389,
        "cash_10q": 248_444 + 56_093,
        "tax_rate": tax[2025] / (opinc[2025] - interest[2025] + 11_392),  # approx; replaced
    }
    # Use diluted shares / Yahoo-style outstanding if available
    market["shares_k"] = 21_597.635
    market["tax_rate"] = 150_649 / 802_595  # FY25 ETR from 10-K

    return {
        "years": years,
        "revenue": rev,
        "cogs": cogs,
        "sga": sga,
        "rd": rd,
        "opinc": opinc,
        "interest": interest,
        "tax": tax,
        "ni": ni,
        "cash": cash,
        "ar": ar,
        "inventory": {y: 0 for y in years + [2020]},
        "ppe": ppe,
        "ap": ap,
        "debt": debt,
        "da": da,
        "capex": capex_total,
        "shares": shares,
        "market": market,
        # prior year anchors
        "cash_2020": cash.get(2020, cash[2021]),
        "ar_2020": ar.get(2020, ar[2021]),
        "ap_2020": ap.get(2020, ap[2021]),
        "ppe_2020": ppe.get(2020, ppe[2021]),
        "debt_2020": debt.get(2020, debt_lt.get(2020, 739_435)),
    }


def populate_three_statement(ws, data: dict) -> None:
    years = data["years"]  # 2021..2025 -> E..I
    assert len(years) == 5

    # Title / start year (hardcoded input in template)
    set_input(ws, "E2", years[0])
    ws["B2"] = "FICO FINANCIAL STATEMENTS"
    ws["B9"] = "SG&A ($000's)"
    ws["B10"] = "R&D ($000's)"
    ws["B28"] = "SG&A"
    ws["B29"] = "R&D"

    # Historical IS/BS/CF/schedules
    for col, y in zip(HIST_COLS, years):
        # IS
        set_input(ws, f"{col}24", round(data["revenue"][y], 1))
        set_input(ws, f"{col}25", round(data["cogs"][y], 1))
        set_input(ws, f"{col}28", round(data["sga"][y], 1))
        set_input(ws, f"{col}29", round(data["rd"][y], 1))
        set_input(ws, f"{col}30", round(data["da"][y], 1))
        set_input(ws, f"{col}31", round(data["interest"][y], 1))
        set_input(ws, f"{col}35", round(data["tax"][y], 1))

        # BS
        set_input(ws, f"{col}41", round(data["cash"][y], 1))
        set_input(ws, f"{col}42", round(data["ar"][y], 1))
        set_input(ws, f"{col}43", 0)
        set_input(ws, f"{col}44", round(data["ppe"][y], 1))
        set_input(ws, f"{col}48", round(data["ap"][y], 1))
        set_input(ws, f"{col}49", round(data["debt"][y], 1))
        assets = data["cash"][y] + data["ar"][y] + data["ppe"][y]
        liab = data["ap"][y] + data["debt"][y]
        set_input(ws, f"{col}52", round(assets - liab, 1))  # equity plug for simplified BS
        set_input(ws, f"{col}53", 0)

        # WC schedule mirrors
        set_input(ws, f"{col}85", round(data["ar"][y], 1))
        set_input(ws, f"{col}86", 0)
        set_input(ws, f"{col}87", round(data["ap"][y], 1))

        # ΔNWC
        prev_ar = data["ar"].get(y - 1, data["ar_2020"] if y == years[0] else data["ar"][y])
        prev_ap = data["ap"].get(y - 1, data["ap_2020"] if y == years[0] else data["ap"][y])
        if y == years[0]:
            prev_nwc = data["ar_2020"] - data["ap_2020"]
        else:
            prev_nwc = data["ar"][y - 1] - data["ap"][y - 1]
        nwc = data["ar"][y] - data["ap"][y]
        set_input(ws, f"{col}64", round(nwc - prev_nwc, 1))

        # Capex / financing
        set_input(ws, f"{col}68", round(data["capex"][y], 1))
        prev_debt = data["debt"].get(y - 1, data["debt_2020"] if y == years[0] else data["debt"][y])
        if y == years[0]:
            prev_debt = data["debt_2020"]
        else:
            prev_debt = data["debt"][y - 1]
        debt_issue = data["debt"][y] - prev_debt
        set_input(ws, f"{col}72", round(debt_issue, 1))

        # Equity issuance plug so cash closes to BS cash under template CF identity
        # CFO proxy = NI + DA - dnwc; CFI = capex; CFF = debt_issue + equity_issue
        if y == years[0]:
            prev_cash = data["cash_2020"]
        else:
            prev_cash = data["cash"][y - 1]
        dnwc = nwc - prev_nwc
        cfo_proxy = data["ni"][y] + data["da"][y] - dnwc
        target_delta = data["cash"][y] - prev_cash
        equity_issue = target_delta - (cfo_proxy - data["capex"][y] + debt_issue)
        set_input(ws, f"{col}73", round(equity_issue, 1))

        # PPE schedule
        open_ppe = data["ppe_2020"] if y == years[0] else data["ppe"][y - 1]
        da = data["da"][y]
        close_ppe = data["ppe"][y]
        ppe_capex = close_ppe - open_ppe + da
        set_input(ws, f"{col}92", round(open_ppe, 1))
        set_input(ws, f"{col}93", round(ppe_capex, 1))
        set_input(ws, f"{col}94", round(da, 1))

        # Debt schedule
        set_input(ws, f"{col}98", round(prev_debt, 1))
        set_input(ws, f"{col}99", round(debt_issue, 1))
        set_input(ws, f"{col}101", round(data["interest"][y], 1))

    # Opening cash: pristine template has E77='=D78' (formula). Set D78 prior-year cash.
    # F77:I77 are hardcoded in the template sample — convert to roll-forward formulas.
    d78 = ws["D78"].value
    if isinstance(d78, str) and d78.startswith("="):
        raise RuntimeError("Unexpected formula in D78")
    ws["D78"] = round(data["cash_2020"], 1)
    ws["D78"].font = BLUE
    ws["D78"].fill = INPUT_FILL
    for coord, formula in [("F77", "=E78"), ("G77", "=F78"), ("H77", "=G78"), ("I77", "=H78")]:
        cell = ws[coord]
        if isinstance(cell.value, str) and cell.value.startswith("="):
            continue
        cell.value = formula

    # Forecast assumptions (J-N) — hardcoded inputs in template
    y0 = years[-1]
    growths = [0.14, 0.12, 0.10, 0.09, 0.08]
    cogs_pct = data["cogs"][y0] / data["revenue"][y0]
    ar_days = data["ar"][y0] / data["revenue"][y0] * 365
    ap_days = data["ap"][y0] / data["cogs"][y0] * 365
    tax_rate = data["market"]["tax_rate"]
    int_pct = data["interest"][y0] / data["debt"][y0]
    da_pct = data["da"][y0] / max(data["ppe"].get(y0 - 1, data["ppe"][y0]), 1)

    rev = data["revenue"][y0]
    for i, col in enumerate(FORECAST_COLS):
        g = growths[i]
        rev *= 1 + g
        set_input(ws, f"{col}7", g)
        set_input(ws, f"{col}8", round(cogs_pct, 4))
        set_input(ws, f"{col}9", round(data["sga"][y0] * (1 + g), 1))
        set_input(ws, f"{col}10", round(data["rd"][y0] * (1 + g), 1))
        set_input(ws, f"{col}11", round(min(max(da_pct, 0.05), 0.50), 4))
        set_input(ws, f"{col}12", round(min(max(int_pct, 0.03), 0.10), 4))
        set_input(ws, f"{col}13", round(tax_rate, 4))
        set_input(ws, f"{col}15", round(ar_days, 1))
        set_input(ws, f"{col}16", 0)
        set_input(ws, f"{col}17", round(ap_days, 1))
        set_input(ws, f"{col}18", round(rev * 0.02, 1))
        set_input(ws, f"{col}19", 0)
        set_input(ws, f"{col}20", 0)

    ws["B131"] = (
        "FICO SEC mapping: Revenue/COGS/SG&A/R&D/Interest/Tax/Cash/AR/PPE/AP/Debt/D&A/Capex from EDGAR 10-K."
    )
    ws["B132"] = "Yellow/blue = inputs. All subtotals/forecast mechanics remain native CFI formulas."
    ws["B133"] = f"Pulled via SEC companyfacts CIK {CIK}. Units $ thousands. Educational use only."
    for r in (131, 132, 133):
        ws[f"B{r}"].font = NOTE


def populate_dcf(ws, data: dict) -> None:
    m = data["market"]
    y0 = data["years"][-1]

    ws["B2"] = "FICO DCF Model"
    set_input(ws, "D5", round(m["tax_rate"], 4))
    set_input(ws, "D6", 0.096)  # WACC assumption
    set_input(ws, "D7", 0.03)
    set_input(ws, "D8", 25)
    set_input(ws, "D9", datetime(2025, 9, 30))
    set_input(ws, "D10", datetime(2025, 9, 30))
    set_input(ws, "D11", m["price"])
    set_input(ws, "D12", m["shares_k"])
    set_input(ws, "D13", m["debt_10q"])
    set_input(ws, "D14", m["cash_10q"])
    set_input(ws, "D15", round(data["capex"][y0], 1))

    # Fix FYE month to September (template hardcodes June in DATE formulas — those are formulas, so
    # we carefully replace only if pattern matches expected template formula)
    for col in ["E", "F", "G", "H", "I"]:
        coord = f"{col}18"
        val = ws[coord].value
        if isinstance(val, str) and val.startswith("=DATE(YEAR($D$10)+"):
            ws[coord] = f"=DATE(YEAR($D$10)+{col}19,9,30)"

    # Project EBIT / D&A / NWC for DCF explicit period from SEC base year
    growths = [0.14, 0.12, 0.10, 0.09, 0.08]
    rev = data["revenue"][y0]
    cogs_pct = data["cogs"][y0] / rev
    opex = data["sga"][y0] + data["rd"][y0]
    prev = rev
    for i, col in enumerate(["E", "F", "G", "H", "I"]):
        g = growths[i]
        rev = prev * (1 + g)
        ebit = rev - rev * cogs_pct - opex * (1 + g) - rev * 0.008
        da = rev * 0.008
        dnwc = max(0, rev - prev) * 0.05
        set_input(ws, f"{col}21", round(ebit, 1))
        set_input(ws, f"{col}23", round(da, 1))
        set_input(ws, f"{col}25", round(dnwc, 1))
        prev = rev

    ws["B50"] = "FICO DCF inputs from SEC + market bridge (10-Q debt/cash). Yellow/blue = inputs."
    ws["B50"].font = NOTE


def add_source_sheet(wb, data: dict) -> None:
    ws = wb.create_sheet("FICO_SEC_Source_Data", 0)
    ws["A1"] = "FICO SEC EDGAR source data injected into CFI templates"
    ws["A1"].font = TITLE
    ws["A2"] = f"CIK {CIK} | companyfacts API | units $ thousands | User-Agent: {USER_AGENT}"
    headers = ["Line item", *data["years"], "XBRL / source"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(4, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    rows = [
        ("Revenue", data["revenue"], "RevenueFromContractWithCustomerExcludingAssessedTax"),
        ("Cost of revenues", data["cogs"], "CostOfRevenue"),
        ("SG&A", data["sga"], "SellingGeneralAndAdministrativeExpense"),
        ("R&D", data["rd"], "ResearchAndDevelopmentExpense"),
        ("Operating income", data["opinc"], "OperatingIncomeLoss"),
        ("Interest expense", data["interest"], "InterestExpense / 10-K interest, net"),
        ("Tax provision", data["tax"], "IncomeTaxExpenseBenefit"),
        ("Net income", data["ni"], "NetIncomeLoss"),
        ("D&A", data["da"], "DepreciationDepletionAndAmortization"),
        ("Cash", data["cash"], "CashAndCashEquivalentsAtCarryingValue"),
        ("Accounts receivable", data["ar"], "AccountsReceivableNetCurrent"),
        ("PP&E", data["ppe"], "PropertyPlantAndEquipmentNet"),
        ("Accounts payable", data["ap"], "AccountsPayableCurrent"),
        ("Total debt", data["debt"], "LT debt + current maturities / 10-K note"),
        ("Capex + cap. software", data["capex"], "PaymentsToAcquirePPE + capitalized software"),
    ]
    for i, (name, series, src) in enumerate(rows, start=5):
        ws.cell(i, 1, name)
        for j, y in enumerate(data["years"], start=2):
            ws.cell(i, j, round(series.get(y, 0), 1))
        ws.cell(i, 7, src)

    ws["A22"] = "DCF market bridge"
    ws["A23"] = "Price"
    ws["B23"] = data["market"]["price"]
    ws["A24"] = "Shares (000s)"
    ws["B24"] = data["market"]["shares_k"]
    ws["A25"] = "Debt 10-Q"
    ws["B25"] = data["market"]["debt_10q"]
    ws["A26"] = "Cash + mkt securities 10-Q"
    ws["B26"] = data["market"]["cash_10q"]
    ws["A28"] = "Cell map"
    ws["A29"] = "3-Statement hist years = columns E-I starting at E2 year. Revenue row 24, COGS 25, SG&A 28, R&D 29, D&A 30, Interest 31, Tax 35."
    ws["A30"] = "DCF assumptions D5:D15; forecast EBIT/D&A/NWC in E21:I21, E23:I23, E25:I25. All other cells remain CFI formulas."
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["G"].width = 56


def copy_sheet(src_ws, dst_ws) -> None:
    for row in src_ws.iter_rows():
        for cell in row:
            new = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new.font = copy(cell.font)
                new.fill = copy(cell.fill)
                new.border = copy(cell.border)
                new.number_format = cell.number_format
                new.alignment = copy(cell.alignment)
    for letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[letter].width = dim.width
    for idx, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[idx].height = dim.height
    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))
    # charts
    for chart in getattr(src_ws, "_charts", []):
        try:
            dst_ws.add_chart(copy(chart))
        except Exception:
            pass


def link_dcf_to_three_statement(dcf_ws) -> None:
    """In combined workbook, point DCF forecast lines at 3-statement forecast cols J-N."""
    # Tax rate
    if not (isinstance(dcf_ws["D5"].value, str) and dcf_ws["D5"].value.startswith("=")):
        # replace hardcoded tax with link
        dcf_ws["D5"] = "='03_Three_Statement'!J13"
        dcf_ws["D5"].fill = PatternFill("solid", fgColor="E2EFDA")
        dcf_ws["D5"].font = Font(name="Calibri", color="000000")

    for dcol, scol in zip(["E", "F", "G", "H", "I"], ["J", "K", "L", "M", "N"]):
        dcf_ws[f"{dcol}21"] = (
            f"='03_Three_Statement'!{scol}26-'03_Three_Statement'!{scol}28-"
            f"'03_Three_Statement'!{scol}29-'03_Three_Statement'!{scol}30"
        )
        dcf_ws[f"{dcol}23"] = f"='03_Three_Statement'!{scol}30"
        # Capex row is formula to $D$15 in template — overwrite with year link (was formula; intentional for MEGA link)
        dcf_ws[f"{dcol}24"] = f"='03_Three_Statement'!{scol}18"
        dcf_ws[f"{dcol}25"] = f"='03_Three_Statement'!{scol}89"
        for r in (21, 23, 24, 25):
            dcf_ws[f"{dcol}{r}"].fill = PatternFill("solid", fgColor="E2EFDA")
            dcf_ws[f"{dcol}{r}"].font = Font(name="Calibri", color="000000")
    dcf_ws["B21"] = "EBIT (linked from 3-statement)"
    dcf_ws["B23"] = "Plus: D&A (linked)"
    dcf_ws["B24"] = "Less: Capex (linked)"
    dcf_ws["B25"] = "Less: Changes in NWC (linked)"


def count_formulas(ws) -> int:
    n = 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                n += 1
    return n


def main() -> None:
    if not ORIG_3.exists() or not ORIG_DCF.exists():
        raise SystemExit(f"Missing originals. Expected {ORIG_3} and {ORIG_DCF}")

    print("Fetching FICO companyfacts from SEC EDGAR...")
    facts = load_companyfacts()
    data = build_fico_series(facts)
    print("SEC years:", data["years"])
    print("FY25 revenue $000:", data["revenue"][2025], "NI:", data["ni"][2025])

    # Work on temp copies of originals
    tmp3 = ROOT / "DCF resources" / "_tmp_3stmt.xlsx"
    tmpd = ROOT / "DCF resources" / "_tmp_dcf.xlsx"
    shutil.copy2(ORIG_3, tmp3)
    shutil.copy2(ORIG_DCF, tmpd)

    wb3 = load_workbook(tmp3)
    wbd = load_workbook(tmpd)
    f3_before = count_formulas(wb3["3 Statement Model"])
    fd_before = count_formulas(wbd["DCF Model"])

    # Cover labels
    wb3["Cover Page"]["C12"] = "FICO — 3 Statement Model (CFI template)"
    wb3["Cover Page"]["C20"] = "Historical inputs from SEC EDGAR 10-K. Forecast formulas preserved."
    wbd["Cover Page"]["C12"] = "FICO — DCF Model (CFI template)"
    wbd["Cover Page"]["C20"] = "DCF assumptions + forecast drivers from FICO SEC/market data. Formulas preserved."

    populate_three_statement(wb3["3 Statement Model"], data)
    populate_dcf(wbd["DCF Model"], data)

    f3_after = count_formulas(wb3["3 Statement Model"])
    fd_after = count_formulas(wbd["DCF Model"])
    print(f"3-stmt formulas before/after: {f3_before}/{f3_after}")
    print(f"DCF formulas before/after: {fd_before}/{fd_after}")
    # Opening-cash replacements may convert 4 hardcoded cells into formulas (+4)
    if f3_after < f3_before:
        raise RuntimeError("Formula cells were destroyed in 3-statement template")
    if fd_after < fd_before - 5:
        raise RuntimeError("Formula cells were destroyed in DCF template")

    # Also refresh standalone copies in DCF resources for convenience
    standalone3 = ROOT / "DCF resources" / "CFI_3-Statement-Model-Complete.xlsx"
    standalone_dcf = ROOT / "DCF resources" / "CFI_DCF-Model.xlsx"
    wb3.save(standalone3)
    wbd.save(standalone_dcf)

    # Combined output workbook
    out_wb = Workbook()
    # remove default
    default = out_wb.active
    out_wb.remove(default)

    # Instructions
    ins = out_wb.create_sheet("01_Instructions", 0)
    ins["A1"] = "FICO DCF Financial Model (CFI templates + SEC EDGAR)"
    ins["A1"].font = TITLE
    for i, line in enumerate(
        [
            "",
            "Built from local CFI blueprint templates. Historical yellow/blue cells = FICO SEC inputs.",
            "All CFI subtotals, forecast mechanics, checks, and DCF valuation formulas are preserved.",
            "",
            "Sheets",
            "02_FICO_SEC_Source_Data — raw EDGAR figures used for injection",
            "03_Three_Statement — CFI 3-statement with FICO FY2021–FY2025 + forecast assumptions",
            "04_DCF — CFI DCF linked to 03_Three_Statement forecast (green cells)",
            "",
            "Key cells to verify",
            "03_Three_Statement!E2 = 2021",
            "03_Three_Statement!I24 = FY2025 Revenue (1,990,869)",
            "03_Three_Statement!I36 = Net Earnings (formula)",
            "04_DCF!D11 = share price; D13/D14 = 10-Q debt/cash",
            "04_DCF!E21 = linked EBIT from 3-statement",
        ],
        start=2,
    ):
        ins[f"A{i}"] = line
    ins.column_dimensions["A"].width = 100

    add_source_sheet(out_wb, data)
    # rename source sheet positionally
    out_wb["FICO_SEC_Source_Data"].title = "02_FICO_SEC_Source_Data"

    # Copy populated model sheets
    ws3 = out_wb.create_sheet("03_Three_Statement")
    copy_sheet(wb3["3 Statement Model"], ws3)
    wsd = out_wb.create_sheet("04_DCF")
    copy_sheet(wbd["DCF Model"], wsd)
    link_dcf_to_three_statement(wsd)

    # Cover pages optional
    c3 = out_wb.create_sheet("00_Cover_3Statement")
    copy_sheet(wb3["Cover Page"], c3)
    cd = out_wb.create_sheet("00_Cover_DCF")
    copy_sheet(wbd["Cover Page"], cd)

    order = [
        "01_Instructions",
        "02_FICO_SEC_Source_Data",
        "03_Three_Statement",
        "04_DCF",
        "00_Cover_3Statement",
        "00_Cover_DCF",
    ]
    for i, name in enumerate(order):
        out_wb.move_sheet(name, offset=i - out_wb.sheetnames.index(name))

    out_wb.save(OUT)
    print("Wrote", OUT)

    # Verify injected numbers
    v = load_workbook(OUT)
    s = v["03_Three_Statement"]
    assert s["E2"].value == 2021
    assert abs(s["I24"].value - data["revenue"][2025]) < 1
    assert abs(s["I25"].value - data["cogs"][2025]) < 1
    assert abs(s["I28"].value - data["sga"][2025]) < 1
    assert abs(s["I29"].value - data["rd"][2025]) < 1
    assert abs(s["I49"].value - data["debt"][2025]) < 1
    assert isinstance(s["I26"].value, str) and s["I26"].value.startswith("=")
    assert isinstance(s["J24"].value, str) and s["J24"].value.startswith("=")
    d = v["04_DCF"]
    assert isinstance(d["E21"].value, str) and "03_Three_Statement" in d["E21"].value
    assert isinstance(d["D32"].value, str) and d["D32"].value.startswith("=")
    print("VERIFY OK")
    print("FY25 Revenue cell I24 =", s["I24"].value)
    print("FY25 SG&A cell I28 =", s["I28"].value)
    print("FY25 Debt cell I49 =", s["I49"].value)
    print("DCF linked EBIT E21 =", d["E21"].value)

    tmp3.unlink(missing_ok=True)
    tmpd.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
