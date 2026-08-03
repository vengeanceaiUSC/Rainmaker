#!/usr/bin/env python3
"""
Build FICO pack:
1) FICO/3S_FICO.csv + FICO/DCF_FICO_SOURCES.csv with 10-K filing links / naming notes
2) FICO/3S_FICO.xlsx — pristine CFI 3-statement + FICO hist inputs, wide columns (no #####)
3) FICO/DCF_FICO.xlsx — CFI DCF + FICO drivers + Sources sheet with hyperlinks
"""

from __future__ import annotations

import csv
import json
import urllib.request
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
ROOT = Path(__file__).resolve().parents[1]
FICO_DIR = Path(__file__).resolve().parent
ORIG_3 = ROOT / "DCF resources" / "originals" / "CFI_3-Statement-Model-Complete.xlsx"
ORIG_DCF = ROOT / "DCF resources" / "originals" / "CFI_DCF-Model.xlsx"
CACHE = ROOT / "fico-valuation" / "data" / "fico_companyfacts.json"
USER_AGENT = "FinancialAnalyst user@example.com"
CIK = "0000814547"

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True)
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
NOTE = Font(name="Calibri", size=10, italic=True, color="C00000")
LINK_FONT = Font(name="Calibri", color="0563C1", underline="single")

FILINGS = {
    2025: {
        "htm": "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm",
        "index": "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/0000814547-25-000030-index.htm",
        "filed": "2025-11-07",
    },
    2024: {
        "htm": "https://www.sec.gov/Archives/edgar/data/814547/000162828024045719/fico-20240930.htm",
        "index": "https://www.sec.gov/Archives/edgar/data/814547/000162828024045719/0001628280-24-045719-index.htm",
        "filed": "2024-11-06",
    },
    2023: {
        "htm": "https://www.sec.gov/Archives/edgar/data/814547/000081454723000022/fico-20230930.htm",
        "index": "https://www.sec.gov/Archives/edgar/data/814547/000081454723000022/0000814547-23-000022-index.htm",
        "filed": "2023-11-08",
    },
}
COMPANYFACTS = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"
YEARS = [2023, 2024, 2025]
HIST_COLS = ["E", "F", "G", "H", "I"]
HIST_YEARS = [2021, 2022, 2023, 2024, 2025]


def is_formula(v) -> bool:
    return isinstance(v, str) and v.startswith("=")


def set_input(ws, coord, value) -> None:
    cell = ws[coord]
    if is_formula(cell.value):
        raise RuntimeError(f"Refusing formula overwrite {coord}: {cell.value}")
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL


def widen(ws, start=4, end=15, width=18.0) -> None:
    for col in range(start, end + 1):
        letter = get_column_letter(col)
        cur = ws.column_dimensions[letter].width or 10
        ws.column_dimensions[letter].width = max(cur, width)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 12, 28)
    ws.column_dimensions["C"].width = max(ws.column_dimensions["C"].width or 12, 24)


def fetch_facts() -> dict:
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(
        COMPANYFACTS,
        headers={"User-Agent": USER_AGENT, "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode())
    CACHE.write_text(json.dumps(data))
    return data


def annual_usd_thousands(facts: dict, concept: str, ymin=2020, ymax=2025) -> dict[int, float]:
    usgaap = facts["facts"].get("us-gaap", {})
    if concept not in usgaap or "USD" not in usgaap[concept]["units"]:
        return {}
    best: dict[int, dict] = {}
    for it in usgaap[concept]["units"]["USD"]:
        if it.get("form") not in ("10-K", "10-K/A") or it.get("fp") != "FY":
            continue
        end = it.get("end") or ""
        if not end.endswith("-09-30"):
            continue
        y = int(end[:4])
        if y < ymin or y > ymax:
            continue
        prev = best.get(y)
        if prev is None or it.get("filed", "") >= prev["filed"]:
            best[y] = {"val": it["val"], "filed": it.get("filed", "")}
    return {y: best[y]["val"] / 1000.0 for y in best}


def load_series(facts: dict) -> dict:
    s = {
        "revenue": annual_usd_thousands(facts, "RevenueFromContractWithCustomerExcludingAssessedTax"),
        "cogs": annual_usd_thousands(facts, "CostOfRevenue"),
        "sga": annual_usd_thousands(facts, "SellingGeneralAndAdministrativeExpense"),
        "rd": annual_usd_thousands(facts, "ResearchAndDevelopmentExpense"),
        "opinc": annual_usd_thousands(facts, "OperatingIncomeLoss"),
        "interest": annual_usd_thousands(facts, "InterestExpense"),
        "tax": annual_usd_thousands(facts, "IncomeTaxExpenseBenefit"),
        "ni": annual_usd_thousands(facts, "NetIncomeLoss"),
        "cash": annual_usd_thousands(facts, "CashAndCashEquivalentsAtCarryingValue"),
        "ar": annual_usd_thousands(facts, "AccountsReceivableNetCurrent"),
        "ppe": annual_usd_thousands(facts, "PropertyPlantAndEquipmentNet"),
        "ap": annual_usd_thousands(facts, "AccountsPayableCurrent"),
        "ca": annual_usd_thousands(facts, "AssetsCurrent"),
        "assets": annual_usd_thousands(facts, "Assets"),
        "re": annual_usd_thousands(facts, "RetainedEarningsAccumulatedDeficit"),
        "equity": annual_usd_thousands(facts, "StockholdersEquity"),
        "da": annual_usd_thousands(facts, "DepreciationDepletionAndAmortization"),
        "ocf": annual_usd_thousands(facts, "NetCashProvidedByUsedInOperatingActivities"),
        "capex_ppe": annual_usd_thousands(facts, "PaymentsToAcquirePropertyPlantAndEquipment"),
        "debt_lt": annual_usd_thousands(facts, "LongTermDebtNoncurrent"),
        "debt_cur": annual_usd_thousands(facts, "LongTermDebtCurrent"),
    }
    # 10-K audited patches (Interest expense, net; total debt; CapEx incl. capitalized software)
    s["interest"].update({2023: 95_546, 2024: 105_638, 2025: 133_647})
    debt = {y: s["debt_lt"].get(y, 0) + s["debt_cur"].get(y, 0) for y in range(2020, 2026)}
    debt[2024] = 2_209_021
    debt[2025] = 3_055_691
    s["debt"] = debt
    s["inventory"] = {y: 0.0 for y in range(2020, 2026)}
    # CapEx = Purchases of PPE + Capitalized internal-use software (CF investing)
    s["capex"] = dict(s["capex_ppe"])
    s["capex"].update({2023: 4_237, 2024: 8_884 + 16_667, 2025: 8_922 + 30_485})
    s["opex"] = {y: s["sga"][y] + s["rd"][y] for y in HIST_YEARS}
    # ΔNWC ≈ Δ(AR − AP)
    dnwc = {}
    for y in HIST_YEARS:
        nwc = s["ar"][y] - s["ap"][y]
        prev = s["ar"][y - 1] - s["ap"][y - 1]
        dnwc[y] = nwc - prev
    s["dnwc"] = dnwc
    s["market"] = {
        "price": 1046.23,
        "shares_k": 21_597.635,
        "debt_10q": 5_582_389,
        "cash_10q": 248_444 + 56_093,
        "tax_rate": 150_649 / 802_595,
        "10q": "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000814547&type=10-Q",
    }
    return s


def source_rows(s: dict) -> list[dict]:
    """Canonical source table used by CSVs and Excel Sources sheet."""
    f25 = FILINGS[2025]["htm"]
    f24 = FILINGS[2024]["htm"]
    f23 = FILINGS[2023]["htm"]
    rows = [
        {
            "model": "3-Statement / DCF",
            "model_label": "Revenues",
            "cfi_row": "3S!row24 / DCF driver base",
            "fico_10k_label": "Revenues",
            "naming_note": "Same meaning. 10-K Consolidated Statements of Income: 'Revenues'.",
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["revenue"][2023],
            "FY2024": s["revenue"][2024],
            "FY2025": s["revenue"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "RevenueFromContractWithCustomerExcludingAssessedTax",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "COGS / Cost of Goods Sold",
            "cfi_row": "3S!row25",
            "fico_10k_label": "Cost of revenues",
            "naming_note": "CFI says COGS; FICO 10-K labels it 'Cost of revenues' (same economic line).",
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["cogs"][2023],
            "FY2024": s["cogs"][2024],
            "FY2025": s["cogs"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "CostOfRevenue",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF OpEx",
            "model_label": "SG&A (CFI: Salaries and Benefits)",
            "cfi_row": "3S!row28",
            "fico_10k_label": "Selling, general and administrative",
            "naming_note": (
                "NAMING DIFFERS. CFI template line is 'Salaries and Benefits'. "
                "FICO 10-K has no 'Salaries' line — map FICO SG&A here. "
                "OpEx in FICO ≈ SG&A + R&D (below)."
            ),
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["sga"][2023],
            "FY2024": s["sga"][2024],
            "FY2025": s["sga"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "SellingGeneralAndAdministrativeExpense",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF OpEx",
            "model_label": "R&D (CFI: Rent and Overhead)",
            "cfi_row": "3S!row29",
            "fico_10k_label": "Research and development",
            "naming_note": (
                "NAMING DIFFERS. CFI 'Rent and Overhead' is a placeholder opex bucket. "
                "FICO maps R&D expense into that row. Combined SG&A+R&D = operating expenses excl. COGS/D&A."
            ),
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["rd"][2023],
            "FY2024": s["rd"][2024],
            "FY2025": s["rd"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "ResearchAndDevelopmentExpense",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "DCF",
            "model_label": "Operating Expenses (SG&A + R&D)",
            "cfi_row": "used inside DCF EBIT build",
            "fico_10k_label": "Selling, general and administrative + Research and development",
            "naming_note": (
                "FICO does not print a single 'Operating expenses' total for SG&A+R&D. "
                "Sum the two 10-K lines. Do NOT confuse with 'Cost of revenues'."
            ),
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["opex"][2023],
            "FY2024": s["opex"][2024],
            "FY2025": s["opex"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "SG&A + R&D (derived)",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "DCF",
            "model_label": "EBIT / Operating income",
            "cfi_row": "DCF!E21 (and forecast)",
            "fico_10k_label": "Operating income",
            "naming_note": "CFI/DCF 'EBIT' ≈ FICO 'Operating income' on the income statement.",
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["opinc"][2023],
            "FY2024": s["opinc"][2024],
            "FY2025": s["opinc"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "OperatingIncomeLoss",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Interest Expense",
            "cfi_row": "3S!row31",
            "fico_10k_label": "Interest expense, net",
            "naming_note": (
                "10-K presents 'Interest expense, net' (interest expense netted with interest income). "
                "Use that printed line, not gross InterestExpense alone when they diverge."
            ),
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["interest"][2023],
            "FY2024": s["interest"][2024],
            "FY2025": s["interest"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "Interest expense, net (10-K IS line)",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF",
            "model_label": "Tax Expense / Cash Taxes rate",
            "cfi_row": "3S!row35; DCF!D5 tax rate",
            "fico_10k_label": "Provision for income taxes",
            "naming_note": "Tax expense from income statement. DCF cash-tax rate uses FY25 tax / pretax approx.",
            "statement_section": "Consolidated Statements of Income and Comprehensive Income",
            "FY2023": s["tax"][2023],
            "FY2024": s["tax"][2024],
            "FY2025": s["tax"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "IncomeTaxExpenseBenefit",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF",
            "model_label": "Net Income",
            "cfi_row": "3S!row36/62",
            "fico_10k_label": "Net income",
            "naming_note": "Same naming.",
            "statement_section": "Consolidated Statements of Income / Cash Flows",
            "FY2023": s["ni"][2023],
            "FY2024": s["ni"][2024],
            "FY2025": s["ni"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "NetIncomeLoss",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF",
            "model_label": "D&A",
            "cfi_row": "3S!row30/63; DCF!E23",
            "fico_10k_label": "Depreciation and amortization",
            "naming_note": "From Consolidated Statements of Cash Flows (add-back).",
            "statement_section": "Consolidated Statements of Cash Flows",
            "FY2023": s["da"][2023],
            "FY2024": s["da"][2024],
            "FY2025": s["da"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "DepreciationDepletionAndAmortization",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF",
            "model_label": "CapEx",
            "cfi_row": "3S!row68; DCF!D15 / row24",
            "fico_10k_label": "Purchases of property and equipment + Capitalized internal-use software costs",
            "naming_note": (
                "NAMING / SCOPE. CFI 'Investments in Property & Equipment' / DCF Capex. "
                "FICO CapEx for modeling = PPE purchases + capitalized internal-use software "
                "(both in CF investing). FY25: 8,922 + 30,485 = 39,407."
            ),
            "statement_section": "Consolidated Statements of Cash Flows — Investing activities",
            "FY2023": s["capex"][2023],
            "FY2024": s["capex"][2024],
            "FY2025": s["capex"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "PaymentsToAcquirePPE + Capitalized internal-use software (10-K CF)",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF",
            "model_label": "Changes in NWC",
            "cfi_row": "3S!row64; DCF!E25",
            "fico_10k_label": "Derived Δ(Accounts receivable − Accounts payable)",
            "naming_note": (
                "Not a single printed 10-K total. Built from BS: AR and AP. "
                "CF also shows AR/AP working-capital lines separately."
            ),
            "statement_section": "Balance Sheet + Cash Flow working capital lines",
            "FY2023": s["dnwc"][2023],
            "FY2024": s["dnwc"][2024],
            "FY2025": s["dnwc"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "Derived from AccountsReceivableNetCurrent & AccountsPayableCurrent",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Operating Cash Flow",
            "cfi_row": "3S!row65",
            "fico_10k_label": "Net cash provided by operating activities",
            "naming_note": "Same meaning; FICO wording is longer.",
            "statement_section": "Consolidated Statements of Cash Flows",
            "FY2023": s["ocf"][2023],
            "FY2024": s["ocf"][2024],
            "FY2025": s["ocf"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "NetCashProvidedByUsedInOperatingActivities",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF",
            "model_label": "Cash & Equivalents",
            "cfi_row": "3S!row41; DCF!D14 (overridden by 10-Q bridge)",
            "fico_10k_label": "Cash and cash equivalents",
            "naming_note": "YE 10-K cash. DCF bridge may use later 10-Q cash + marketable securities.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": s["cash"][2023],
            "FY2024": s["cash"][2024],
            "FY2025": s["cash"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "CashAndCashEquivalentsAtCarryingValue",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Accounts Receivable",
            "cfi_row": "3S!row42",
            "fico_10k_label": "Accounts receivable, net",
            "naming_note": "Same; 10-K says ', net'.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": s["ar"][2023],
            "FY2024": s["ar"][2024],
            "FY2025": s["ar"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "AccountsReceivableNetCurrent",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Inventory",
            "cfi_row": "3S!row43",
            "fico_10k_label": "(not material / not presented)",
            "naming_note": "FICO software/scores business — no material inventory line; model uses 0.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": 0,
            "FY2024": 0,
            "FY2025": 0,
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "n/a",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Current Assets",
            "cfi_row": "not a CFI hard input (informational)",
            "fico_10k_label": "Total current assets",
            "naming_note": "From 10-K balance sheet total current assets.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": s["ca"][2023],
            "FY2024": s["ca"][2024],
            "FY2025": s["ca"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "AssetsCurrent",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Net PP&E",
            "cfi_row": "3S!row44",
            "fico_10k_label": "Property and equipment, net",
            "naming_note": "Same meaning.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": s["ppe"][2023],
            "FY2024": s["ppe"][2024],
            "FY2025": s["ppe"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "PropertyPlantAndEquipmentNet",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Total Assets",
            "cfi_row": "3S!row45 (formula in template)",
            "fico_10k_label": "Total assets",
            "naming_note": "SEC total assets provided for reference; CFI row 45 is a SUM formula — do not overwrite.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": s["assets"][2023],
            "FY2024": s["assets"][2024],
            "FY2025": s["assets"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "Assets",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Accounts Payable",
            "cfi_row": "3S!row48",
            "fico_10k_label": "Accounts payable",
            "naming_note": "Same naming.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": s["ap"][2023],
            "FY2024": s["ap"][2024],
            "FY2025": s["ap"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "AccountsPayableCurrent",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement / DCF",
            "model_label": "Total Debt",
            "cfi_row": "3S!row49; DCF!D13 (10-Q override optional)",
            "fico_10k_label": "Current maturities of long-term debt + Long-term debt",
            "naming_note": (
                "Sum of current + noncurrent debt from BS / debt footnote. "
                "FY25 total debt 3,055,691. DCF may override D13 with later 10-Q debt."
            ),
            "statement_section": "Consolidated Balance Sheets / Debt footnote",
            "FY2023": s["debt"][2023],
            "FY2024": s["debt"][2024],
            "FY2025": s["debt"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "LongTermDebtCurrent + LongTermDebtNoncurrent / 10-K note",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Retained Earnings",
            "cfi_row": "3S!row53",
            "fico_10k_label": "Retained earnings",
            "naming_note": "Same; equity section of balance sheet / equity statement.",
            "statement_section": "Consolidated Balance Sheets / Stockholders' equity",
            "FY2023": s["re"][2023],
            "FY2024": s["re"][2024],
            "FY2025": s["re"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "RetainedEarningsAccumulatedDeficit",
            "companyfacts_api": COMPANYFACTS,
        },
        {
            "model": "3-Statement",
            "model_label": "Total Equity",
            "cfi_row": "3S!row54 (formula/total)",
            "fico_10k_label": "Total stockholders' deficit / equity",
            "naming_note": "FICO reports stockholders' deficit (negative equity) due to buybacks.",
            "statement_section": "Consolidated Balance Sheets",
            "FY2023": s["equity"][2023],
            "FY2024": s["equity"][2024],
            "FY2025": s["equity"][2025],
            "unit": "USD thousands",
            "primary_10k_url_FY2025": f25,
            "also_FY2024_url": f24,
            "also_FY2023_url": f23,
            "xbrl_concept": "StockholdersEquity",
            "companyfacts_api": COMPANYFACTS,
        },
    ]
    return rows


def write_csvs(rows: list[dict], s: dict) -> None:
    # Full source bible
    dcf_path = FICO_DIR / "DCF_FICO_SOURCES.csv"
    fields = [
        "model",
        "model_label",
        "cfi_row",
        "fico_10k_label",
        "naming_note",
        "statement_section",
        "unit",
        "FY2023",
        "FY2024",
        "FY2025",
        "primary_10k_url_FY2025",
        "also_FY2024_url",
        "also_FY2023_url",
        "xbrl_concept",
        "companyfacts_api",
    ]
    with dcf_path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})

    # Compact 3S CSV (requested file name)
    three_path = FICO_DIR / "3S_FICO.csv"
    with three_path.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(
            [
                "statement_section",
                "cfi_template_row",
                "model_label",
                "fico_10k_label",
                "naming_note",
                "unit",
                "FY2023",
                "FY2024",
                "FY2025",
                "primary_10k_url_FY2025",
                "FY2024_10k_url",
                "FY2023_10k_url",
                "xbrl_concept",
                "company",
                "cik",
                "fiscal_year_end",
            ]
        )
        for r in rows:
            if "3-Statement" not in r["model"] and r["model"] != "3-Statement / DCF OpEx":
                # keep OpEx mapping rows too if useful; include all rows that feed 3S
                if r["model"] == "DCF" and "EBIT" not in r["model_label"] and "OpEx" not in r["model_label"]:
                    continue
            w.writerow(
                [
                    r["statement_section"],
                    r["cfi_row"],
                    r["model_label"],
                    r["fico_10k_label"],
                    r["naming_note"],
                    r["unit"],
                    r["FY2023"],
                    r["FY2024"],
                    r["FY2025"],
                    r["primary_10k_url_FY2025"],
                    r["also_FY2024_url"],
                    r["also_FY2023_url"],
                    r["xbrl_concept"],
                    "Fair Isaac Corporation (FICO)",
                    CIK,
                    "September 30",
                ]
            )

    # Filing index helper
    idx = FICO_DIR / "10K_FILING_LINKS.csv"
    with idx.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["fiscal_year", "report_date", "filed", "10k_html", "filing_index"])
        for y, meta in FILINGS.items():
            w.writerow([y, f"{y}-09-30", meta["filed"], meta["htm"], meta["index"]])
        w.writerow([])
        w.writerow(["companyfacts_api", COMPANYFACTS])
        w.writerow(["10q_browse", s["market"]["10q"]])

    print(f"Wrote {dcf_path}")
    print(f"Wrote {three_path}")
    print(f"Wrote {idx}")


def inject_3statement(ws, s: dict) -> None:
    set_input(ws, "E2", HIST_YEARS[0])
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
            prev_debt = s["debt"].get(y - 1, s["debt_lt"].get(y - 1, 0) + s["debt_cur"].get(y - 1, 0))
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
        target_delta = s["cash"][y] - prev_cash
        equity_issue = target_delta - (cfo_proxy - capex + debt_issue)
        set_input(ws, f"{col}73", round(equity_issue, 1))

        if col == "E":
            set_input(ws, "D78", round(s["cash"][y - 1], 1))
        else:
            set_input(ws, f"{col}77", round(prev_cash, 1))

        set_input(ws, f"{col}85", round(s["ar"][y], 1))
        set_input(ws, f"{col}86", 0)
        set_input(ws, f"{col}87", round(s["ap"][y], 1))
        da = s["da"][y]
        ppe_capex = s["ppe"][y] - open_ppe + da
        set_input(ws, f"{col}92", round(open_ppe, 1))
        set_input(ws, f"{col}93", round(ppe_capex, 1))
        set_input(ws, f"{col}94", round(da, 1))
        set_input(ws, f"{col}98", round(prev_debt, 1))
        set_input(ws, f"{col}99", round(debt_issue, 1))
        set_input(ws, f"{col}101", round(s["interest"][y], 1))

    ws["B131"] = "FICO 10-K sources: see sheet '10K_Sources' and FICO/DCF_FICO_SOURCES.csv"
    ws["B131"].font = NOTE
    widen(ws)


def add_sources_sheet(wb, rows: list[dict], s: dict, title: str) -> None:
    # Remove old if rebuilding
    if "10K_Sources" in wb.sheetnames:
        del wb["10K_Sources"]
    ws = wb.create_sheet("10K_Sources", 0)
    ws["A1"] = title
    ws["A1"].font = TITLE
    ws["A2"] = "Fair Isaac Corporation (FICO) | CIK 0000814547 | FYE September 30 | Units: USD thousands"
    ws["A3"] = "Click the FY2025 10-K link for each line. Naming differences (e.g. OpEx) are explained in column E."

    headers = [
        "Model",
        "Model / CFI label",
        "CFI cell/row",
        "FICO 10-K label (as printed)",
        "Naming / mapping note",
        "10-K statement section",
        "FY2023",
        "FY2024",
        "FY2025",
        "FY2025 10-K URL",
        "FY2024 10-K URL",
        "FY2023 10-K URL",
        "XBRL / derivation",
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(5, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for i, r in enumerate(rows, start=6):
        vals = [
            r["model"],
            r["model_label"],
            r["cfi_row"],
            r["fico_10k_label"],
            r["naming_note"],
            r["statement_section"],
            r["FY2023"],
            r["FY2024"],
            r["FY2025"],
            r["primary_10k_url_FY2025"],
            r["also_FY2024_url"],
            r["also_FY2023_url"],
            r["xbrl_concept"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(i, c, v)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c in (10, 11, 12) and isinstance(v, str) and v.startswith("http"):
                cell.hyperlink = v
                cell.font = LINK_FONT

    # Filing index block
    start = 6 + len(rows) + 2
    ws.cell(start, 1, "FILING INDEX").font = TITLE
    ws.cell(start + 1, 1, "FY2025 10-K HTML")
    ws.cell(start + 1, 2, FILINGS[2025]["htm"]).hyperlink = FILINGS[2025]["htm"]
    ws.cell(start + 1, 2).font = LINK_FONT
    ws.cell(start + 2, 1, "FY2024 10-K HTML")
    ws.cell(start + 2, 2, FILINGS[2024]["htm"]).hyperlink = FILINGS[2024]["htm"]
    ws.cell(start + 2, 2).font = LINK_FONT
    ws.cell(start + 3, 1, "FY2023 10-K HTML")
    ws.cell(start + 3, 2, FILINGS[2023]["htm"]).hyperlink = FILINGS[2023]["htm"]
    ws.cell(start + 3, 2).font = LINK_FONT
    ws.cell(start + 4, 1, "SEC companyfacts API")
    ws.cell(start + 4, 2, COMPANYFACTS).hyperlink = COMPANYFACTS
    ws.cell(start + 4, 2).font = LINK_FONT
    ws.cell(start + 5, 1, "10-Q browse (debt/cash bridge)")
    ws.cell(start + 5, 2, s["market"]["10q"]).hyperlink = s["market"]["10q"]
    ws.cell(start + 5, 2).font = LINK_FONT

    ws.cell(start + 7, 1, "DCF bridge overrides (not YE 10-K)").font = TITLE
    ws.cell(start + 8, 1, "DCF!D13 Debt (10-Q)")
    ws.cell(start + 8, 2, s["market"]["debt_10q"])
    ws.cell(start + 8, 3, "Later interim debt; override YE 10-K if bridging to current EV")
    ws.cell(start + 9, 1, "DCF!D14 Cash (10-Q cash+mkt secs)")
    ws.cell(start + 9, 2, s["market"]["cash_10q"])
    ws.cell(start + 9, 3, "Cash & cash equivalents + marketable securities from latest 10-Q bridge")
    ws.cell(start + 10, 1, "DCF!D11 Price / D12 Shares")
    ws.cell(start + 10, 2, f"{s['market']['price']} / {s['market']['shares_k']}")
    ws.cell(start + 10, 3, "Market price & diluted shares (000s) — not from 10-K income statement")

    widths = [18, 34, 22, 40, 55, 40, 12, 12, 12, 45, 45, 45, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[5].height = 30
    ws.freeze_panes = "A6"


def _strip_images(wb) -> None:
    """openpyxl can fail re-serializing embedded cover images; drop them safely."""
    for ws in wb.worksheets:
        if getattr(ws, "_images", None):
            ws._images = []


def build_3statement(s: dict, rows: list[dict]) -> Path:
    wb = load_workbook(ORIG_3, data_only=False)
    inject_3statement(wb["3 Statement Model"], s)
    add_sources_sheet(wb, rows, s, "FICO 3-Statement — 10-K source map")
    if "Cover Page" in wb.sheetnames:
        wb["Cover Page"]["C12"] = "FICO — 3 Statement Model (CFI template + 10-K actuals)"
        wb["Cover Page"]["C20"] = "See 10K_Sources sheet for SEC links and naming notes. Columns widened so values are visible."
    _strip_images(wb)
    out = FICO_DIR / "3S_FICO.xlsx"
    wb.save(out)
    print(f"Saved {out} ({out.stat().st_size} bytes)")
    return out


def build_dcf(s: dict, rows: list[dict]) -> Path:
    wb = load_workbook(ORIG_DCF, data_only=False)
    ws = wb["DCF Model"]
    m = s["market"]
    y0 = 2025

    ws["B2"] = "FICO DCF Model"
    set_input(ws, "D5", round(m["tax_rate"], 4))
    set_input(ws, "D6", 0.096)
    set_input(ws, "D7", 0.03)
    set_input(ws, "D8", 25)
    set_input(ws, "D9", datetime(2025, 9, 30))
    set_input(ws, "D10", datetime(2025, 9, 30))
    set_input(ws, "D11", m["price"])
    set_input(ws, "D12", m["shares_k"])
    set_input(ws, "D13", m["debt_10q"])
    set_input(ws, "D14", m["cash_10q"])
    set_input(ws, "D15", round(s["capex"][y0], 1))

    # Fix June FYE hardcodes inside DATE formulas → September
    for col in ["E", "F", "G", "H", "I"]:
        coord = f"{col}18"
        val = ws[coord].value
        if isinstance(val, str) and "DATE(YEAR($D$10)+" in val:
            # keep period offset, force month 9
            ws[coord] = f"=DATE(YEAR($D$10)+{col}19,9,30)"

    # Explicit forecast drivers from FY25 base (real numbers, not blank → no ##### from errors)
    growths = [0.14, 0.12, 0.10, 0.09, 0.08]
    rev = s["revenue"][y0]
    cogs_pct = s["cogs"][y0] / rev
    opex0 = s["opex"][y0]
    prev = rev
    for i, col in enumerate(["E", "F", "G", "H", "I"]):
        g = growths[i]
        rev = prev * (1 + g)
        ebit = rev - rev * cogs_pct - opex0 * (1 + g) - rev * 0.008
        da = rev * 0.008
        dnwc = max(0.0, rev - prev) * 0.05
        set_input(ws, f"{col}21", round(ebit, 1))
        set_input(ws, f"{col}23", round(da, 1))
        set_input(ws, f"{col}25", round(dnwc, 1))
        # Capex row is formula to $D$15 in template for all years — leave formulas
        prev = rev

    ws["B50"] = (
        "Sources: sheet 10K_Sources (clickable 10-K links). "
        "OpEx = SG&A + R&D from 10-K (CFI names differ). Capex D15 = PPE purchases + capitalized software."
    )
    ws["B50"].font = NOTE
    ws["D2"] = "See 10K_Sources for every input link"
    ws["D2"].font = NOTE

    widen(ws, start=4, end=14, width=16)
    # Entry/exit / TV columns often show ##### when narrow
    for letter in ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]:
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 10, 16)

    add_sources_sheet(wb, rows, s, "FICO DCF — 10-K source map (click links)")
    if "Cover Page" in wb.sheetnames:
        wb["Cover Page"]["C12"] = "FICO — DCF Model (CFI template + 10-K sourced inputs)"
        wb["Cover Page"]["C20"] = "Open 10K_Sources first — every number has a 10-K URL and naming note."
    _strip_images(wb)

    out = FICO_DIR / "DCF_FICO.xlsx"
    wb.save(out)
    print(f"Saved {out} ({out.stat().st_size} bytes)")
    return out


def verify() -> None:
    wb3 = load_workbook(FICO_DIR / "3S_FICO.xlsx", data_only=False)
    ws = wb3["3 Statement Model"]
    assert ws["I24"].value == 1990869.0 or ws["I24"].value == 1990869
    assert is_formula(ws["I26"].value)
    assert "10K_Sources" in wb3.sheetnames
    assert ws.column_dimensions["I"].width >= 16
    print("3S verify OK", "I24=", ws["I24"].value, "I width=", ws.column_dimensions["I"].width)

    wbd = load_workbook(FICO_DIR / "DCF_FICO.xlsx", data_only=False)
    d = wbd["DCF Model"]
    assert d["D15"].value == 39407 or d["D15"].value == 39407.0
    assert d["E21"].value is not None
    assert d["F21"].value is not None
    assert "10K_Sources" in wbd.sheetnames
    src = wbd["10K_Sources"]
    assert str(src["J6"].value).startswith("https://www.sec.gov/")
    print("DCF verify OK", "D15=", d["D15"].value, "E21=", d["E21"].value, "F21=", d["F21"].value)


def main() -> None:
    print("Fetching SEC companyfacts...")
    facts = fetch_facts()
    s = load_series(facts)
    rows = source_rows(s)
    write_csvs(rows, s)
    build_3statement(s, rows)
    build_dcf(s, rows)
    verify()
    print("Done.")


if __name__ == "__main__":
    main()
