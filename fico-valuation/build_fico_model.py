#!/usr/bin/env python3
"""Build FICO DCF / investor valuation Excel workbook from EDGAR fundamentals."""

from __future__ import annotations

import csv
import json
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"

END_YEARS = [2021, 2022, 2023, 2024, 2025]

# Market snapshot captured at model build (Yahoo Finance)
PRICE = 1046.23
SHARES_OUT_THOUSANDS = 21597.635  # shares outstanding / 1000
BETA = 1.321
MKT_CAP_M = 22596.09


def annual_by_end(usgaap: dict, dei: dict, concept: str) -> dict:
    src = usgaap.get(concept) or dei.get(concept)
    if not src:
        return {}
    units = src["units"]
    unit = next((u for u in ["USD", "shares", "USD/shares", "pure"] if u in units), next(iter(units)))
    best: dict = {}
    for it in units[unit]:
        if it.get("form") not in ("10-K", "10-K/A") or it.get("fp") != "FY":
            continue
        end = it.get("end") or ""
        if not end.endswith("-09-30"):
            continue
        y = int(end[:4])
        if y not in END_YEARS:
            continue
        prev = best.get(y)
        if prev is None or it.get("filed", "") >= prev["filed"]:
            best[y] = {
                "val": it["val"],
                "end": end,
                "filed": it.get("filed"),
                "frame": it.get("frame"),
                "unit": unit,
                "accn": it.get("accn"),
            }
    return best


def extract_xbrl() -> list[dict]:
    facts = json.loads((DATA / "fico_companyfacts.json").read_text())
    usgaap = facts["facts"]["us-gaap"]
    dei = facts["facts"].get("dei", {})
    concepts = [
        "Revenues",
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "CostOfRevenue",
        "ResearchAndDevelopmentExpense",
        "SellingGeneralAndAdministrativeExpense",
        "AmortizationOfIntangibleAssets",
        "RestructuringCharges",
        "OperatingIncomeLoss",
        "InterestExpense",
        "OtherNonoperatingIncomeExpense",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeTaxExpenseBenefit",
        "NetIncomeLoss",
        "EarningsPerShareBasic",
        "EarningsPerShareDiluted",
        "WeightedAverageNumberOfSharesOutstandingBasic",
        "WeightedAverageNumberOfDilutedSharesOutstanding",
        "CashAndCashEquivalentsAtCarryingValue",
        "AccountsReceivableNetCurrent",
        "AssetsCurrent",
        "PropertyPlantAndEquipmentNet",
        "Goodwill",
        "DeferredIncomeTaxAssetsNet",
        "OtherAssetsNoncurrent",
        "Assets",
        "AccountsPayableCurrent",
        "AccruedLiabilitiesCurrent",
        "ContractWithCustomerLiabilityCurrent",
        "LongTermDebtCurrent",
        "LiabilitiesCurrent",
        "LongTermDebtNoncurrent",
        "OperatingLeaseLiabilityNoncurrent",
        "OtherLiabilitiesNoncurrent",
        "Liabilities",
        "RetainedEarningsAccumulatedDeficit",
        "TreasuryStockValue",
        "AccumulatedOtherComprehensiveIncomeLossNetOfTax",
        "StockholdersEquity",
        "CommonStockSharesOutstanding",
        "NetCashProvidedByUsedInOperatingActivities",
        "NetCashProvidedByUsedInInvestingActivities",
        "NetCashProvidedByUsedInFinancingActivities",
        "DepreciationDepletionAndAmortization",
        "ShareBasedCompensation",
        "DeferredIncomeTaxExpenseBenefit",
        "IncreaseDecreaseInAccountsReceivable",
        "IncreaseDecreaseInAccountsPayable",
        "IncreaseDecreaseInDeferredRevenue",
        "PaymentsToAcquirePropertyPlantAndEquipment",
        "PaymentsForRepurchaseOfCommonStock",
        "ProceedsFromStockOptionsExercised",
        "OperatingLeaseRightOfUseAsset",
        "OperatingLeaseLiabilityCurrent",
    ]
    rows = []
    for c in concepts:
        vals = annual_by_end(usgaap, dei, c)
        if not vals:
            continue
        row = {"concept": c, "unit": next(iter(vals.values()))["unit"]}
        for y in END_YEARS:
            row[str(y)] = vals[y]["val"] if y in vals else None
        rows.append(row)

    out_csv = DATA / "fico_xbrl_by_period_end.csv"
    with out_csv.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["concept", "unit"] + [str(y) for y in END_YEARS])
        w.writeheader()
        w.writerows(rows)
    return rows


def audited_statements():
    """Audited FY figures from FY2025 10-K ($ thousands) + latest 10-Q BS."""
    income = OrderedDict(
        [
            ("On-premises and SaaS software", {2025: 740145, 2024: 711340, 2023: 640182}),
            ("Professional services", {2025: 82149, 2024: 86536, 2023: 99547}),
            ("Scores", {2025: 1168575, 2024: 919650, 2023: 773828}),
            ("Total revenues", {2025: 1990869, 2024: 1717526, 2023: 1513557}),
            ("Cost of revenues", {2025: 353722, 2024: 348206, 2023: 311053}),
            ("Research and development", {2025: 188347, 2024: 171940, 2023: 159950}),
            ("Selling, general and administrative", {2025: 513028, 2024: 462834, 2023: 400565}),
            ("Amortization of intangible assets", {2025: 0, 2024: 917, 2023: 1100}),
            ("Restructuring charges", {2025: 10922, 2024: 0, 2023: 0}),
            ("Gain on product line asset sale", {2025: 0, 2024: 0, 2023: -1941}),
            ("Total operating expenses", {2025: 1066019, 2024: 983897, 2023: 870727}),
            ("Operating income", {2025: 924850, 2024: 733629, 2023: 642830}),
            ("Interest expense, net", {2025: -133647, 2024: -105638, 2023: -95546}),
            ("Other income, net", {2025: 11392, 2024: 14034, 2023: 6340}),
            ("Income before income taxes", {2025: 802595, 2024: 642025, 2023: 553624}),
            ("Provision for income taxes", {2025: 150649, 2024: 129214, 2023: 124249}),
            ("Net income", {2025: 651946, 2024: 512811, 2023: 429375}),
            ("EPS basic", {2025: 26.90, 2024: 20.78, 2023: 17.18}),
            ("EPS diluted", {2025: 26.54, 2024: 20.45, 2023: 16.93}),
            ("Shares basic (000s)", {2025: 24239, 2024: 24676, 2023: 24986}),
            ("Shares diluted (000s)", {2025: 24561, 2024: 25079, 2023: 25367}),
        ]
    )

    balance = OrderedDict(
        [
            ("Cash and cash equivalents", {2025: 134136, 2024: 150667, "2026Q3": 248444}),
            ("Accounts receivable, net", {2025: 529148, 2024: 426642, "2026Q3": 592530}),
            ("Prepaid expenses and other current assets", {2025: 41881, 2024: 40104, "2026Q3": 40589}),
            ("Total current assets", {2025: 705165, 2024: 617413, "2026Q3": 881563}),
            ("Marketable securities", {2025: 54625, 2024: 45289, "2026Q3": 56093}),
            ("Property and equipment, net", {2025: 67713, 2024: 38465}),
            ("Operating lease right-of-use assets", {2025: 26213, 2024: 29580}),
            ("Goodwill", {2025: 783340, 2024: 782752}),
            ("Deferred income taxes", {2025: 118553, 2024: 86513}),
            ("Other assets", {2025: 112524, 2024: 117872}),
            ("Total assets", {2025: 1868133, 2024: 1717884, "2026Q3": 2037373}),
            ("Accounts payable", {2025: 32315, 2024: 22473, "2026Q3": 26878}),
            ("Accrued compensation and employee benefits", {2025: 115369, 2024: 106103}),
            ("Other accrued liabilities", {2025: 114618, 2024: 79812}),
            ("Deferred revenue", {2025: 187372, 2024: 156897, "2026Q3": 205424}),
            ("Current maturities on debt", {2025: 399541, 2024: 15000, "2026Q3": 300000}),
            ("Total current liabilities", {2025: 849215, 2024: 380285, "2026Q3": 745191}),
            ("Long-term debt", {2025: 2656150, 2024: 2194021, "2026Q3": 5282389}),
            ("Operating lease liabilities (noncurrent)", {2025: 19187, 2024: 21963, "2026Q3": 15621}),
            ("Other liabilities", {2025: 89365, 2024: 84294, "2026Q3": 91307}),
            ("Total liabilities", {2025: 3613917, 2024: 2680563, "2026Q3": 6134508}),
            ("Common stock (par)", {2025: 238, 2024: 244}),
            ("Additional paid-in-capital", {2025: 1331120, 2024: 1366572}),
            ("Treasury stock", {2025: -7537908, 2024: -6138736}),
            ("Retained earnings", {2025: 4552816, 2024: 3900870}),
            ("AOCI", {2025: -92050, 2024: -91629}),
            ("Total stockholders deficit", {2025: -1745784, 2024: -962679}),
            ("Shares outstanding (000s)", {2025: 23764, 2024: 24392}),
            ("Total debt (current + LT)", {2025: 3055691, 2024: 2209021, "2026Q3": 5582389}),
        ]
    )

    cashflow = OrderedDict(
        [
            ("Net income", {2025: 651946, 2024: 512811, 2023: 429375}),
            ("Depreciation and amortization", {2025: 14952, 2024: 13827, 2023: 14638}),
            ("Share-based compensation", {2025: 156667, 2024: 149439, 2023: 123847}),
            ("Deferred income taxes", {2025: -32486, 2024: -27330, 2023: -47378}),
            ("Net gain on marketable securities", {2025: -5024, 2024: -9834, 2023: -2908}),
            ("Non-cash operating lease costs", {2025: 9604, 2024: 12423, 2023: 14708}),
            ("Provision of doubtful accounts", {2025: 1485, 2024: 1675, 2023: 1475}),
            ("Gain on product line asset sale", {2025: 0, 2024: 0, 2023: -1941}),
            ("Net loss on sales/abandonment of PP&E", {2025: 210, 2024: 438, 2023: 547}),
            ("Change in accounts receivable", {2025: -100377, 2024: -34144, 2023: -70117}),
            ("Change in prepaid expenses and other assets", {2025: 610, 2024: -14034, 2023: -11904}),
            ("Change in accounts payable", {2025: 9849, 2024: 3316, 2023: 2236}),
            ("Change in accrued compensation", {2025: 10065, 2024: 3195, 2023: 4631}),
            ("Change in other liabilities", {2025: 31566, 2024: 7216, 2023: -7057}),
            ("Change in deferred revenue", {2025: 29740, 2024: 13966, 2023: 18763}),
            ("Net cash from operating activities", {2025: 778807, 2024: 632964, 2023: 468915}),
            ("Purchases of property and equipment", {2025: -8922, 2024: -8884, 2023: -4237}),
            ("Capitalized internal-use software costs", {2025: -30485, 2024: -16667, 2023: 0}),
            ("Proceeds from sales of marketable securities", {2025: 2184, 2024: 15930, 2023: 5032}),
            ("Purchases of marketable securities", {2025: -6496, 2024: -18372, 2023: -10623}),
            ("Cash transferred from product line asset sale", {2025: 0, 2024: 0, 2023: -6126}),
            ("Net cash used in investing activities", {2025: -43719, 2024: -27993, 2023: -15954}),
            ("Proceeds from revolving line / term loans", {2025: 725000, 2024: 947000, 2023: 407000}),
            ("Payments on revolving line / term loans", {2025: -1368750, 2024: -602000, 2023: -402000}),
            ("Proceeds from issuance of senior notes", {2025: 1500000, 2024: 0, 2023: 0}),
            ("Payments on debt issuance costs", {2025: -17163, 2024: -706, 2023: 0}),
            ("Payments on finance leases", {2025: -3144, 2024: -1333, 2023: 0}),
            ("Proceeds from issuance of treasury stock (ESP)", {2025: 32823, 2024: 25006, 2023: 22198}),
            ("Taxes paid related to net share settlement", {2025: -204593, 2024: -139188, 2023: -76673}),
            ("Repurchases of common stock", {2025: -1414502, 2024: -821702, 2023: -405526}),
            ("Net cash used in financing activities", {2025: -750329, 2024: -592923, 2023: -455001}),
            ("Effect of exchange rate changes on cash", {2025: -1290, 2024: 1841, 2023: 5616}),
            ("Increase (decrease) in cash", {2025: -16531, 2024: 13889, 2023: 3576}),
            ("Cash beginning of year", {2025: 150667, 2024: 136778, 2023: 133202}),
            ("Cash end of year", {2025: 134136, 2024: 150667, 2023: 136778}),
            ("Cash paid for income taxes", {2025: 162089, 2024: 133716, 2023: 152775}),
            ("Cash paid for interest", {2025: 103564, 2024: 106388, 2023: 96877}),
        ]
    )
    return income, balance, cashflow


def style_header_row(ws, row, start_col, end_col, header_font, header_fill):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def autosize(ws, min_width=12, max_width=48):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        length = 0
        for cell in col:
            if cell.value is not None:
                length = max(length, min(max_width, len(str(cell.value))))
        ws.column_dimensions[letter].width = max(min_width, length + 2)


def build() -> Path:
    xbrl_rows = extract_xbrl()
    income, balance, cashflow = audited_statements()

    title_font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    section_font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
    input_font = Font(name="Calibri", size=11, color="0000FF")
    formula_font = Font(name="Calibri", size=11, color="000000")
    link_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    input_fill = PatternFill("solid", fgColor="FFF2CC")
    section_fill = PatternFill("solid", fgColor="D6EAF8")
    good_fill = PatternFill("solid", fgColor="C6EFCE")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    pct = "0.0%"
    num = "#,##0"
    num1 = "#,##0.0"
    num2 = "#,##0.00"

    wb = Workbook()

    # ----- 00_ReadMe -----
    ws = wb.active
    ws.title = "00_ReadMe"
    ws["A1"] = "FICO (Fair Isaac Corporation) — Investor DCF / Valuation Workbook"
    ws["A1"].font = title_font
    lines = [
        "",
        "What this file is",
        "A full investor-style valuation pack for FICO built from SEC EDGAR 10-K / 10-Q fundamentals, with a working 5-year FCFF DCF, WACC, and sensitivity tables.",
        "",
        "Color key",
        "Yellow cells + blue font = INPUTS you should edit",
        "Black font = formulas (do not overwrite)",
        "",
        "How to use (guided path)",
        "1) Read 01_Template_Links for downloadable external DCF / LBO / comps templates (Damodaran, CFI, Macabacus, etc.).",
        "2) Review 02_EDGAR_Sources for filing links and CIK.",
        "3) Confirm historical line items on 03_IS / 04_BS / 05_CF (pre-filled from FY2025 10-K; override in yellow if needed).",
        "4) Set forecast drivers on 06_Assumptions (growth, margins, tax, reinvestment, WACC components).",
        "5) Read implied value on 07_DCF and 10_Output; stress with 09_Sensitivity.",
        "6) Use 11_Input_Checklist as a blank checklist of every line item to paste your own case.",
        "",
        "Units",
        "Historical financial statements are in $ thousands (as reported in the 10-K).",
        "DCF / WACC / Output tabs work in $ millions for readability.",
        "",
        "As-of market snapshot (Yahoo Finance pull during model build)",
        f"Price: ${PRICE:,.2f} | Shares out: {SHARES_OUT_THOUSANDS:,.1f}k | Mkt cap: ${MKT_CAP_M:,.1f}m | Beta: {BETA}",
        "Net debt bridge uses latest 10-Q (period ended 2026-06-30): Total debt $5,582.4m − Cash $248.4m − Marketable securities $56.1m.",
        "",
        "Disclaimer",
        "Educational / research model only. Not investment advice. Verify figures against primary SEC filings before relying on outputs.",
    ]
    for i, text in enumerate(lines, start=2):
        ws[f"A{i}"] = text
        if text in {
            "What this file is",
            "Color key",
            "How to use (guided path)",
            "Units",
            "As-of market snapshot (Yahoo Finance pull during model build)",
            "Disclaimer",
        }:
            ws[f"A{i}"].font = section_font
    ws.column_dimensions["A"].width = 120

    # ----- 01_Template_Links -----
    ws = wb.create_sheet("01_Template_Links")
    ws["A1"] = "Downloadable valuation / DCF templates (external)"
    ws["A1"].font = title_font
    headers = ["Source", "Template", "Direct / landing URL", "Notes"]
    for i, h in enumerate(headers, 1):
        ws.cell(3, i, h)
    style_header_row(ws, 3, 1, 4, header_font, header_fill)
    templates = [
        ("Aswath Damodaran (NYU)", "All valuation spreadsheets hub", "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/spreadsh.htm", "Best free academic suite; FCFF/FCFE/DDM/ginzu models"),
        ("Aswath Damodaran", "fcffsimpleginzu.xlsx (recommended FCFF)", "https://pages.stern.nyu.edu/~adamodar/pc/fcffsimpleginzu.xlsx", "Direct XLSX download — all-in-one firm DCF"),
        ("Aswath Damodaran", "fcffginzu.xlsx (full kitchen-sink FCFF)", "https://pages.stern.nyu.edu/~adamodar/pc/fcffginzu.xlsx", "Ratings, R&D capitalize, leases, bottom-up beta"),
        ("Aswath Damodaran", "Equity valuation spreadsheets page", "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/eqspread.htm", "Stable/2-stage/3-stage FCFF & FCFE models"),
        ("Aswath Damodaran", "fcff2st.xls (two-stage FCFF)", "https://pages.stern.nyu.edu/~adamodar/pc/fcff2st.xls", "Direct download"),
        ("Aswath Damodaran", "fcff3st.xls (three-stage FCFF)", "https://pages.stern.nyu.edu/~adamodar/pc/fcff3st.xls", "Direct download"),
        ("Aswath Damodaran", "firmmult.xls (implied multiples from DCF)", "https://pages.stern.nyu.edu/~adamodar/pc/firmmult.xls", "Direct download"),
        ("Babson / CFI mirror", "Bloomberg Practice DCF xlsx", "https://www.babson.edu/media/babson/assets/cutler-center/Bloomberg-Practice-Template.xlsx", "Direct XLSX download"),
        ("Smartsheet", "DCF Valuation xlsx", "https://www.smartsheet.com/sites/default/files/2020-07/IC-Discounted-Cash-Flow-Valuation-10840.xlsx", "Direct XLSX download"),
        ("Corporate Finance Institute (CFI)", "Free DCF Model Template", "https://corporatefinanceinstitute.com/resources/financial-modeling/dcf-model-template/", "3-statement linked + 5yr FCFF"),
        ("CFI (direct asset)", "DCF-Valuation-Compact-Complete.xlsx", "https://corporatefinanceinstitute.com/assets/DCF-Valuation-Compact-Complete.xlsx", "Direct XLSX linked from CFI DCF training"),
        ("CFI", "Comparable company analysis template", "https://corporatefinanceinstitute.com/resources/financial-modeling/comparable-company-analysis-template/", "Trading comps"),
        ("CFI", "3-statement case study template", "https://corporatefinanceinstitute.com/resources/financial-modeling/case-study-3-statement-model-template/", "Integrated model"),
        ("CT Acquisitions", "2026 Free DCF Excel Model", "https://ctacquisitions.com/dcf-template-and-excel/", "6-tab PE/IB style DCF walkthrough"),
        ("Wall Street Oasis", "DCF model template", "https://www.wallstreetoasis.com/resources/templates/excel-financial-modeling/discounted-cash-flow-model-template", "Community template"),
        ("Wall Street Oasis", "LBO model template", "https://www.wallstreetoasis.com/resources/templates/excel-financial-modeling/leveraged-buyout-model-template", "Community template"),
        ("Wall Street Prep", "LBO model lesson + template", "https://www.wallstreetprep.com/knowledge/lbo-model/", "Free lesson + Excel"),
        ("Wall Street Prep sample", "LBO + DCF sample xlsx", "https://s3.amazonaws.com/wsp_sample_file/excel-templates/lbo-w-dcf-model-sample.xlsx", "Direct sample download"),
        ("Exinfm", "LBO DCF Model xls", "https://exinfm.com/excel%20files/LBO_DCF_Model.xls", "Direct XLS download"),
        ("FE Training", "LBO model template", "https://www.fe.training/free-resources/financial-modeling/lbo-model-template/", "Free resource"),
        ("FE Training", "Trading comps template", "https://www.fe.training/free-resources/investment-banking/trading-comparables-model-template/", "Free resource"),
        ("Macabacus", "DCF Model Excel Template", "https://macabacus.com/excel/templates/discounted-cash-flow", "IB-style; requires free trial for download"),
        ("Macabacus", "Long-form LBO template", "https://macabacus.com/excel/templates/lbo-model-long", "Trial"),
        ("Limelight", "Free DCF Calculator Template", "https://www.golimelight.com/templates/dcf-calculator-template", "WACC + dual TV + scenario/sensitivity"),
        ("Keene Advisors", "Free DCF Model Template", "https://www.keeneadvisors.com/free-dcf-model-template", "3yr history + 5yr forecast DCF"),
        ("Cube Software", "Free 3-statement model", "https://www.cubesoftware.com/finance-templates/3-statement-model", "Free template"),
        ("SEC EDGAR", "FICO company filings (raw source)", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000814547&type=10-K&dateb=&owner=include&count=40", "Primary filings for FICO"),
        ("SEC data API", "FICO companyfacts XBRL JSON", "https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json", "Machine-readable fundamentals used in this repo"),
    ]
    for r, row in enumerate(templates, start=4):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(r, c, val)
            if c == 3:
                cell.font = link_font
                cell.hyperlink = val
    autosize(ws)
    ws.column_dimensions["C"].width = 70
    ws.column_dimensions["D"].width = 55

    # ----- 02_EDGAR_Sources -----
    ws = wb.create_sheet("02_EDGAR_Sources")
    ws["A1"] = "EDGAR source map — Fair Isaac Corp (FICO)"
    ws["A1"].font = title_font
    meta = [
        ("CIK", "0000814547"),
        ("Ticker", "FICO"),
        ("Fiscal year end", "September 30"),
        ("Latest 10-K", "FY2025 ended 2025-09-30, filed 2025-11-07"),
        ("10-K HTML", "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"),
        ("10-K index", "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/0000814547-25-000030-index.html"),
        ("Investor PDF copy", "https://investors.fico.com/static-files/663245db-39a2-4db7-8f89-284034bbccbf"),
        ("Latest 10-Q used", "Q3 FY2026 ended 2026-06-30, filed 2026-07-29"),
        ("10-Q HTML", "https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm"),
        ("Companyfacts API", "https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json"),
        ("Submissions API", "https://data.sec.gov/submissions/CIK0000814547.json"),
    ]
    ws["A3"] = "Field"
    ws["B3"] = "Value"
    style_header_row(ws, 3, 1, 2, header_font, header_fill)
    for i, (k, v) in enumerate(meta, start=4):
        ws.cell(i, 1, k)
        cell = ws.cell(i, 2, v)
        if v.startswith("http"):
            cell.font = link_font
            cell.hyperlink = v
    autosize(ws)
    ws.column_dimensions["B"].width = 100

    def write_statement(name, title, data, year_keys, note=None):
        ws = wb.create_sheet(name)
        ws["A1"] = title
        ws["A1"].font = title_font
        if note:
            ws["A2"] = note
        header_row = 4
        ws.cell(header_row, 1, "Line item")
        for j, y in enumerate(year_keys, start=2):
            ws.cell(header_row, j, str(y))
        style_header_row(ws, header_row, 1, 1 + len(year_keys), header_font, header_fill)
        for i, (item, vals) in enumerate(data.items(), start=header_row + 1):
            ws.cell(i, 1, item).border = thin
            for j, y in enumerate(year_keys, start=2):
                cell = ws.cell(i, j, vals.get(y))
                cell.number_format = num2 if str(item).startswith("EPS") else num
                cell.font = input_font
                cell.fill = input_fill
                cell.border = thin
        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, min_col=1, max_col=1):
            v = str(row[0].value or "")
            if v.startswith("Total") or v in {
                "Operating income",
                "Net income",
                "Net cash from operating activities",
                "Net cash used in investing activities",
                "Net cash used in financing activities",
            }:
                for c in range(1, 2 + len(year_keys)):
                    ws.cell(row[0].row, c).font = Font(name="Calibri", size=11, bold=True, color="0000FF")
        autosize(ws)
        ws.column_dimensions["A"].width = 52
        return ws

    write_statement(
        "03_IS_Historical",
        "Consolidated Statements of Income — FICO ($ thousands)",
        income,
        [2025, 2024, 2023],
        "Source: FY2025 Form 10-K. Yellow/blue = editable inputs. Values in thousands USD except EPS and share counts.",
    )
    write_statement(
        "04_BS_Historical",
        "Consolidated Balance Sheets — FICO ($ thousands)",
        balance,
        ["2026Q3", 2025, 2024],
        "FY columns from FY2025 10-K; 2026Q3 from 10-Q ended 2026-06-30 (latest public BS for net debt bridge).",
    )
    write_statement(
        "05_CF_Historical",
        "Consolidated Statements of Cash Flows — FICO ($ thousands)",
        cashflow,
        [2025, 2024, 2023],
        "Source: FY2025 Form 10-K. Capex for DCF = Purchases of PP&E + Capitalized internal-use software.",
    )

    # ----- 06_Assumptions -----
    ws = wb.create_sheet("06_Assumptions")
    ws["A1"] = "Forecast & valuation assumptions (edit yellow/blue cells)"
    ws["A1"].font = title_font
    ws["A2"] = "All $ figures on this tab and DCF are in $ millions unless noted."

    ws["A4"] = "A. Market & capital structure"
    ws["A4"].font = section_font
    ws["A4"].fill = section_fill
    for r, (label, val) in enumerate(
        [
            ("Share price ($)", PRICE),
            ("Diluted shares outstanding (millions)", SHARES_OUT_THOUSANDS / 1000),
            ("Cash & equivalents ($m)", 248.444),
            ("Marketable securities ($m)", 56.093),
            ("Total debt ($m)", 5582.389),
            ("Minority interest ($m)", 0.0),
            ("Preferred equity ($m)", 0.0),
        ],
        start=6,
    ):
        ws[f"A{r}"] = label
        ws[f"B{r}"] = val
        ws[f"B{r}"].font = input_font
        ws[f"B{r}"].fill = input_fill
        ws[f"B{r}"].border = thin
        ws[f"B{r}"].number_format = num2

    ws["A14"] = "B. WACC components"
    ws["A14"].font = section_font
    ws["A14"].fill = section_fill
    for r, (label, val) in enumerate(
        [
            ("Risk-free rate", 0.043),
            ("Equity risk premium", 0.050),
            ("Beta", BETA),
            ("Size / specific risk premium", 0.000),
            ("Pre-tax cost of debt", 0.055),
            ("Marginal tax rate", 0.188),
            ("Target equity weight", 0.80),
            ("Target debt weight", 0.20),
        ],
        start=16,
    ):
        ws[f"A{r}"] = label
        ws[f"B{r}"] = val
        ws[f"B{r}"].font = input_font
        ws[f"B{r}"].fill = input_fill
        ws[f"B{r}"].border = thin
        ws[f"B{r}"].number_format = pct if val < 2 else num2

    ws["A25"] = "Cost of equity (CAPM)"
    ws["B25"] = "=B16+B18*B17+B19"
    ws["B25"].number_format = pct
    ws["B25"].font = formula_font
    ws["A26"] = "After-tax cost of debt"
    ws["B26"] = "=B20*(1-B21)"
    ws["B26"].number_format = pct
    ws["B26"].font = formula_font
    ws["A27"] = "WACC"
    ws["B27"] = "=B22*B25+B23*B26"
    ws["B27"].number_format = pct
    ws["B27"].fill = good_fill
    ws["B27"].font = Font(bold=True)

    ws["A29"] = "C. Operating forecast drivers"
    ws["A29"].font = section_font
    ws["A29"].fill = section_fill
    for r, (label, val) in enumerate(
        [
            ("Base year revenue FY2025 ($m)", 1990.869),
            ("Base year EBIT FY2025 ($m)", 924.850),
            ("Base year D&A FY2025 ($m)", 14.952),
            ("Base year SBC FY2025 ($m) — optional add-back toggle below", 156.667),
        ],
        start=30,
    ):
        ws[f"A{r}"] = label
        ws[f"B{r}"] = val
        ws[f"B{r}"].font = input_font
        ws[f"B{r}"].fill = input_fill
        ws[f"B{r}"].number_format = num2
        ws[f"B{r}"].border = thin

    years = [2026, 2027, 2028, 2029, 2030]
    ws["A35"] = "Driver"
    for j, y in enumerate(years, start=2):
        ws.cell(35, j, f"FY{y}")
    style_header_row(ws, 35, 1, 6, header_font, header_fill)

    ws["A36"] = "Revenue growth"
    for j, g in enumerate([0.14, 0.12, 0.10, 0.09, 0.08], start=2):
        cell = ws.cell(36, j, g)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = pct
        cell.border = thin

    ws["A37"] = "EBIT margin"
    for j, m in enumerate([0.455, 0.450, 0.445, 0.440, 0.440], start=2):
        cell = ws.cell(37, j, m)
        cell.font = input_font
        cell.fill = input_fill
        cell.number_format = pct
        cell.border = thin

    for r, label, default in [
        (38, "Tax rate on EBIT", 0.188),
        (39, "D&A % of revenue", 0.008),
        (40, "Capex + capitalized software % of revenue", 0.020),
        (41, "ΔNWC % of ΔRevenue (use of cash if +)", 0.05),
    ]:
        ws[f"A{r}"] = label
        for j in range(2, 7):
            cell = ws.cell(r, j, default)
            cell.font = input_font
            cell.fill = input_fill
            cell.number_format = pct
            cell.border = thin

    ws["A42"] = "Add back SBC in FCFF? (1=yes, 0=no)"
    ws["B42"] = 0
    ws["B42"].font = input_font
    ws["B42"].fill = input_fill
    ws["B42"].border = thin
    ws["C42"] = "Conservative default = 0 (treat SBC as real economic cost)"

    ws["A44"] = "D. Terminal value"
    ws["A44"].font = section_font
    ws["A44"].fill = section_fill
    for r, label, val, fmt in [
        (45, "Perpetuity growth rate g", 0.03, pct),
        (46, "Exit EV/EBIT multiple", 22.0, num1),
        (47, "TV method weight on perpetuity (1=100% g, 0=100% exit multiple)", 0.5, pct),
        (48, "Mid-year discounting? (1=yes)", 1, num),
    ]:
        ws[f"A{r}"] = label
        ws[f"B{r}"] = val
        ws[f"B{r}"].font = input_font
        ws[f"B{r}"].fill = input_fill
        ws[f"B{r}"].number_format = fmt
        ws[f"B{r}"].border = thin

    ws["A50"] = "E. Historical reference (auto from 03_IS, $m)"
    ws["A50"].font = section_font
    ws["A50"].fill = section_fill
    ws["A51"] = "FY25 revenue ($m)"
    ws["B51"] = "='03_IS_Historical'!B5/1000"
    ws["B51"].number_format = num2
    ws["A52"] = "FY25 EBIT ($m)"
    ws["B52"] = "='03_IS_Historical'!B13/1000"
    ws["B52"].number_format = num2
    ws["A53"] = "FY25 net income ($m)"
    ws["B53"] = "='03_IS_Historical'!B18/1000"
    ws["B53"].number_format = num2
    ws["A54"] = "FY25 EBIT margin"
    ws["B54"] = "=B52/B51"
    ws["B54"].number_format = pct
    ws["A55"] = "FY25 effective tax rate"
    ws["B55"] = "='03_IS_Historical'!B17/'03_IS_Historical'!B16"
    ws["B55"].number_format = pct
    autosize(ws)
    ws.column_dimensions["A"].width = 58

    # ----- 08_WACC -----
    ws = wb.create_sheet("08_WACC")
    ws["A1"] = "WACC build"
    ws["A1"].font = title_font
    ws["A3"] = "Linked from 06_Assumptions"
    ws["A5"] = "Component"
    ws["B5"] = "Value"
    style_header_row(ws, 5, 1, 2, header_font, header_fill)
    items = [
        ("Risk-free rate", "='06_Assumptions'!B16"),
        ("ERP", "='06_Assumptions'!B17"),
        ("Beta", "='06_Assumptions'!B18"),
        ("Cost of equity", "='06_Assumptions'!B25"),
        ("Pre-tax cost of debt", "='06_Assumptions'!B20"),
        ("Tax rate", "='06_Assumptions'!B21"),
        ("After-tax cost of debt", "='06_Assumptions'!B26"),
        ("Equity weight", "='06_Assumptions'!B22"),
        ("Debt weight", "='06_Assumptions'!B23"),
        ("WACC", "='06_Assumptions'!B27"),
    ]
    for i, (lab, f) in enumerate(items, start=6):
        ws.cell(i, 1, lab)
        cell = ws.cell(i, 2, f)
        cell.number_format = num2 if lab == "Beta" else pct
        if lab == "WACC":
            cell.fill = good_fill
            cell.font = Font(bold=True)
    ws["A18"] = "Notes"
    ws["A19"] = "Beta sourced from Yahoo Finance at model build (~1.32). Cost of debt ~blended coupon / interest (~5.5% placeholder — replace with YTM if available)."
    ws["A20"] = "Target weights are judgmental (FICO runs with structural net leverage due to buybacks). Current market gearing is higher after 2026 debt raise — adjust B22/B23 accordingly."
    autosize(ws)

    # ----- 07_DCF -----
    ws = wb.create_sheet("07_DCF")
    ws["A1"] = "FCFF DCF valuation — FICO"
    ws["A1"].font = title_font
    ws["A2"] = "$ millions. Explicit forecast FY2026–FY2030. Edit drivers on 06_Assumptions."
    ws["A4"] = "($ millions)"
    for j, c in enumerate(["FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"], start=2):
        ws.cell(4, j, c)
    style_header_row(ws, 4, 1, 7, header_font, header_fill)

    ws["A5"] = "Revenue"
    ws["B5"] = "='06_Assumptions'!B30"
    ws["C5"] = "=B5*(1+C6)"
    ws["D5"] = "=C5*(1+D6)"
    ws["E5"] = "=D5*(1+E6)"
    ws["F5"] = "=E5*(1+F6)"
    ws["G5"] = "=F5*(1+G6)"

    ws["A6"] = "Revenue growth"
    ws["B6"] = "—"
    for j, col in enumerate(["C", "D", "E", "F", "G"]):
        ws[f"{col}6"] = f"='06_Assumptions'!{get_column_letter(j + 2)}36"

    ws["A7"] = "EBIT"
    ws["B7"] = "='06_Assumptions'!B31"
    for col in ["C", "D", "E", "F", "G"]:
        ws[f"{col}7"] = f"={col}5*{col}8"

    ws["A8"] = "EBIT margin"
    ws["B8"] = "=B7/B5"
    for j, col in enumerate(["C", "D", "E", "F", "G"]):
        ws[f"{col}8"] = f"='06_Assumptions'!{get_column_letter(j + 2)}37"

    ws["A9"] = "Cash tax on EBIT"
    ws["B9"] = "=B7*'06_Assumptions'!B21"
    for j, col in enumerate(["C", "D", "E", "F", "G"]):
        ws[f"{col}9"] = f"={col}7*'06_Assumptions'!{get_column_letter(j + 2)}38"

    ws["A10"] = "NOPAT"
    for col in ["B", "C", "D", "E", "F", "G"]:
        ws[f"{col}10"] = f"={col}7-{col}9"

    ws["A11"] = "(+) D&A"
    ws["B11"] = "='06_Assumptions'!B32"
    for j, col in enumerate(["C", "D", "E", "F", "G"]):
        ws[f"{col}11"] = f"={col}5*'06_Assumptions'!{get_column_letter(j + 2)}39"

    ws["A12"] = "(+) SBC add-back (optional)"
    ws["B12"] = "='06_Assumptions'!B33*'06_Assumptions'!B42"
    for col in ["C", "D", "E", "F", "G"]:
        ws[f"{col}12"] = f"='06_Assumptions'!B33*'06_Assumptions'!B42*({col}5/'06_Assumptions'!B30)"

    ws["A13"] = "Gross cash flow"
    for col in ["B", "C", "D", "E", "F", "G"]:
        ws[f"{col}13"] = f"={col}10+{col}11+{col}12"

    ws["A14"] = "(−) Capex + capitalized software"
    ws["B14"] = 39.407
    ws["B14"].font = input_font
    ws["B14"].fill = input_fill
    for j, col in enumerate(["C", "D", "E", "F", "G"]):
        ws[f"{col}14"] = f"={col}5*'06_Assumptions'!{get_column_letter(j + 2)}40"

    ws["A15"] = "(−) Δ Net working capital"
    ws["B15"] = 0
    for j, col in enumerate(["C", "D", "E", "F", "G"]):
        prev = "B" if col == "C" else chr(ord(col) - 1)
        ws[f"{col}15"] = f"=MAX(0,{col}5-{prev}5)*'06_Assumptions'!{get_column_letter(j + 2)}41"

    ws["A16"] = "Unlevered FCFF"
    ws["B16"] = "=B13-B14-B15"
    for col in ["C", "D", "E", "F", "G"]:
        ws[f"{col}16"] = f"={col}13-{col}14-{col}15"
        ws[f"{col}16"].fill = good_fill

    ws["A17"] = "Discount period (years)"
    ws["B17"] = "—"
    ws["C17"] = "=1-'06_Assumptions'!B48*0.5"
    ws["D17"] = "=C17+1"
    ws["E17"] = "=D17+1"
    ws["F17"] = "=E17+1"
    ws["G17"] = "=F17+1"

    ws["A18"] = "Discount factor"
    ws["B18"] = "—"
    for col in ["C", "D", "E", "F", "G"]:
        ws[f"{col}18"] = f"=1/(1+'06_Assumptions'!B27)^{col}17"

    ws["A19"] = "PV of FCFF"
    ws["B19"] = "—"
    for col in ["C", "D", "E", "F", "G"]:
        ws[f"{col}19"] = f"={col}16*{col}18"
        ws[f"{col}19"].fill = good_fill

    ws["A21"] = "Terminal value"
    ws["A21"].font = section_font
    ws["A22"] = "TV — Gordon growth (perpetuity)"
    ws["G22"] = "=G16*(1+'06_Assumptions'!B45)/('06_Assumptions'!B27-'06_Assumptions'!B45)"
    ws["A23"] = "TV — Exit multiple on EBIT"
    ws["G23"] = "=G7*'06_Assumptions'!B46"
    ws["A24"] = "Blended TV"
    ws["G24"] = "=G22*'06_Assumptions'!B47+G23*(1-'06_Assumptions'!B47)"
    ws["G24"].fill = good_fill
    ws["A25"] = "PV of terminal value"
    ws["G25"] = "=G24*G18"
    ws["G25"].fill = good_fill

    ws["A27"] = "Enterprise value bridge"
    ws["A27"].font = section_font
    bridge = [
        (28, "PV of explicit FCFF", "=SUM(C19:G19)"),
        (29, "PV of terminal value", "=G25"),
        (30, "Enterprise value", "=B28+B29"),
        (31, "(−) Total debt", "='06_Assumptions'!B10"),
        (32, "(+) Cash", "='06_Assumptions'!B8"),
        (33, "(+) Marketable securities", "='06_Assumptions'!B9"),
        (34, "(−) Minorities", "='06_Assumptions'!B11"),
        (35, "(−) Preferred", "='06_Assumptions'!B12"),
        (36, "Equity value", "=B30-B31+B32+B33-B34-B35"),
        (37, "Diluted shares (m)", "='06_Assumptions'!B7"),
        (38, "Intrinsic value / share", "=B36/B37"),
        (39, "Current price", "='06_Assumptions'!B6"),
        (40, "Upside / (downside)", "=B38/B39-1"),
    ]
    for r, lab, f in bridge:
        ws[f"A{r}"] = lab
        ws[f"B{r}"] = f
        ws[f"B{r}"].number_format = pct if r == 40 else num2
    ws["B30"].fill = good_fill
    ws["B30"].font = Font(bold=True)
    ws["B36"].fill = good_fill
    ws["B36"].font = Font(bold=True)
    ws["B38"].fill = good_fill
    ws["B38"].font = Font(bold=True, size=14, color="006100")

    for r in range(5, 20):
        for c in range(2, 8):
            cell = ws.cell(r, c)
            if cell.value not in (None, "—"):
                cell.number_format = pct if r in (6, 8) else num2
    for r in (22, 23, 24, 25):
        ws.cell(r, 7).number_format = num2

    ws["A42"] = "Implied metrics"
    ws["A42"].font = section_font
    ws["A43"] = "TV as % of EV"
    ws["B43"] = "=B29/B30"
    ws["B43"].number_format = pct
    ws["A44"] = "Implied FY25A EV/EBIT"
    ws["B44"] = "=B30/'06_Assumptions'!B31"
    ws["B44"].number_format = num1
    ws["A45"] = "Implied FY25A EV/Revenue"
    ws["B45"] = "=B30/'06_Assumptions'!B30"
    ws["B45"].number_format = num2
    autosize(ws)
    ws.column_dimensions["A"].width = 42

    # ----- 09_Sensitivity -----
    ws = wb.create_sheet("09_Sensitivity")
    ws["A1"] = "Sensitivity: intrinsic value / share vs WACC and terminal growth"
    ws["A1"].font = title_font
    ws["A2"] = "Grid uses Gordon TV only for isolation. Base case blended value remains on 07_DCF."
    ws["A4"] = "Value/share"
    for col, g in zip(["B", "C", "D", "E"], [0.025, 0.030, 0.035, 0.040]):
        ws[f"{col}4"] = g
        ws[f"{col}4"].number_format = pct
        ws[f"{col}4"].fill = input_fill
        ws[f"{col}4"].font = input_font
    for r, w in enumerate([0.08, 0.09, 0.10, 0.11, 0.12], start=5):
        ws.cell(r, 1, w).number_format = pct
        ws.cell(r, 1).fill = input_fill
        ws.cell(r, 1).font = input_font
        for c, col in enumerate(["B", "C", "D", "E"], start=2):
            formula = (
                f"((('07_DCF'!C16)/(1+$A{r})^'07_DCF'!C17)"
                f"+ (('07_DCF'!D16)/(1+$A{r})^'07_DCF'!D17)"
                f"+ (('07_DCF'!E16)/(1+$A{r})^'07_DCF'!E17)"
                f"+ (('07_DCF'!F16)/(1+$A{r})^'07_DCF'!F17)"
                f"+ (('07_DCF'!G16)/(1+$A{r})^'07_DCF'!G17)"
                f"+ ((('07_DCF'!G16)*(1+{col}$4)/($A{r}-{col}$4))/(1+$A{r})^'07_DCF'!G17)"
                f"-'06_Assumptions'!B10+'06_Assumptions'!B8+'06_Assumptions'!B9"
                f")/'06_Assumptions'!B7"
            )
            cell = ws.cell(r, c, formula)
            cell.number_format = "$#,##0.00"
            cell.border = thin

    ws["A11"] = "Exit multiple sensitivity (EV/EBIT) — value/share using exit-multiple TV only"
    ws["A11"].font = section_font
    ws["A13"] = "Value/share"
    for j, m in enumerate([18, 20, 22, 24, 26], start=2):
        cell = ws.cell(13, j, m)
        cell.fill = input_fill
        cell.font = input_font
        cell.number_format = num1
    for i, w in enumerate([0.08, 0.09, 0.10, 0.11, 0.12], start=14):
        cell = ws.cell(i, 1, w)
        cell.fill = input_fill
        cell.font = input_font
        cell.number_format = pct
        for j in range(2, 7):
            col = get_column_letter(j)
            formula = (
                f"((('07_DCF'!C16)/(1+$A{i})^'07_DCF'!C17)"
                f"+ (('07_DCF'!D16)/(1+$A{i})^'07_DCF'!D17)"
                f"+ (('07_DCF'!E16)/(1+$A{i})^'07_DCF'!E17)"
                f"+ (('07_DCF'!F16)/(1+$A{i})^'07_DCF'!F17)"
                f"+ (('07_DCF'!G16)/(1+$A{i})^'07_DCF'!G17)"
                f"+ ((('07_DCF'!G7*{col}$13))/(1+$A{i})^'07_DCF'!G17)"
                f"-'06_Assumptions'!B10+'06_Assumptions'!B8+'06_Assumptions'!B9"
                f")/'06_Assumptions'!B7"
            )
            cell = ws.cell(i, j, formula)
            cell.number_format = "$#,##0.00"
            cell.border = thin
    ws["A21"] = "Row headers = WACC; column headers = g or exit multiple."
    autosize(ws)

    # ----- 10_Output -----
    ws = wb.create_sheet("10_Output")
    ws["A1"] = "Valuation output summary — FICO"
    ws["A1"].font = title_font
    ws["A3"] = "Intrinsic value / share"
    ws["B3"] = "='07_DCF'!B38"
    ws["B3"].number_format = "$#,##0.00"
    ws["B3"].font = Font(bold=True, size=16, color="006100")
    ws["B3"].fill = good_fill
    ws["A4"] = "Current price"
    ws["B4"] = "='06_Assumptions'!B6"
    ws["B4"].number_format = "$#,##0.00"
    ws["A5"] = "Upside/(downside)"
    ws["B5"] = "='07_DCF'!B40"
    ws["B5"].number_format = pct
    for r, lab, f in [
        (7, "Enterprise value ($m)", "='07_DCF'!B30"),
        (8, "Equity value ($m)", "='07_DCF'!B36"),
        (9, "Net debt ($m)", "='06_Assumptions'!B10-'06_Assumptions'!B8-'06_Assumptions'!B9"),
        (10, "WACC", "='06_Assumptions'!B27"),
        (11, "Terminal growth", "='06_Assumptions'!B45"),
        (12, "Exit EV/EBIT", "='06_Assumptions'!B46"),
    ]:
        ws[f"A{r}"] = lab
        ws[f"B{r}"] = f
        ws[f"B{r}"].number_format = pct if r in (10, 11) else (num1 if r == 12 else num2)

    ws["A14"] = "Key FY2025 fundamentals ($m)"
    ws["A14"].font = section_font
    ws["A15"] = "Metric"
    ws["B15"] = "Value"
    style_header_row(ws, 15, 1, 2, header_font, header_fill)
    fund = [
        ("Revenue", 1990.869),
        ("Operating income (EBIT)", 924.850),
        ("Net income", 651.946),
        ("Operating cash flow", 778.807),
        ("Capex + capitalized software", 39.407),
        ("Approx FCFF proxy (OCF − capex − cap. software)", 739.400),
        ("EBIT margin", 0.4645),
        ("Effective tax rate", 0.1877),
    ]
    for i, (k, v) in enumerate(fund, start=16):
        ws.cell(i, 1, k)
        cell = ws.cell(i, 2, v)
        cell.number_format = pct if v < 1 else num2

    ws["A26"] = "Latest BS bridge (10-Q 2026-06-30, $m)"
    ws["A26"].font = section_font
    for i, (k, v) in enumerate(
        [
            ("Cash", 248.444),
            ("Marketable securities", 56.093),
            ("Total debt", 5582.389),
            ("Net debt", 5582.389 - 248.444 - 56.093),
            ("Shares diluted outstanding (Yahoo, m)", SHARES_OUT_THOUSANDS / 1000),
            ("Market cap ($m)", MKT_CAP_M),
            ("EV market ($m)", 29603.06),
        ],
        start=27,
    ):
        ws.cell(i, 1, k)
        cell = ws.cell(i, 2, v)
        cell.number_format = num2

    ws["A36"] = "What to change first if the output looks off"
    ws["A36"].font = section_font
    ws["A37"] = "1) Growth path (06_Assumptions row 36) — Scores growth has been the engine; fading too slowly inflates value."
    ws["A38"] = "2) EBIT margins (row 37) — FY25 margin ~46.5% is elevated vs history; mean-reversion matters."
    ws["A39"] = "3) WACC / leverage weights — buyback-funded leverage makes current gearing atypical."
    ws["A40"] = "4) Terminal multiple / g — with high ROIC software economics, TV dominates; keep g < WACC."
    ws["A41"] = "5) SBC treatment (B42) — turning on SBC add-back materially increases FCFF."
    autosize(ws)
    ws.column_dimensions["A"].width = 70

    # ----- 11_Input_Checklist -----
    ws = wb.create_sheet("11_Input_Checklist")
    ws["A1"] = "Blank input checklist — paste or override every line item"
    ws["A1"].font = title_font
    ws["A2"] = "Use this sheet if you want to hard-enter your own case from scratch. Units: $ thousands for statements; decimals for rates."
    for j, h in enumerate(["Section", "Line item", "Your input", "Model currently uses", "Unit", "Source suggestion"], 1):
        ws.cell(4, j, h)
    style_header_row(ws, 4, 1, 6, header_font, header_fill)

    checklist = []
    for item, vals in income.items():
        checklist.append(("Income Statement", item, None, vals.get(2025), "$000", "10-K IS"))
    for item, vals in balance.items():
        cur = vals.get(2025) if 2025 in vals else vals.get("2026Q3")
        checklist.append(("Balance Sheet", item, None, cur, "$000", "10-K / 10-Q BS"))
    for item, vals in cashflow.items():
        checklist.append(("Cash Flow", item, None, vals.get(2025), "$000", "10-K CF"))
    checklist.extend(
        [
            ("Valuation inputs", "Share price", None, PRICE, "$", "Market"),
            ("Valuation inputs", "Diluted shares (m)", None, round(SHARES_OUT_THOUSANDS / 1000, 3), "millions", "10-Q / Yahoo"),
            ("Valuation inputs", "Risk-free rate", None, 0.043, "decimal", "Treasury"),
            ("Valuation inputs", "ERP", None, 0.05, "decimal", "Damodaran"),
            ("Valuation inputs", "Beta", None, BETA, "number", "Yahoo"),
            ("Valuation inputs", "Pre-tax cost of debt", None, 0.055, "decimal", "Debt YTM"),
            ("Valuation inputs", "Tax rate", None, 0.188, "decimal", "FY25 ETR"),
            ("Valuation inputs", "Terminal growth g", None, 0.03, "decimal", "Judgment"),
            ("Valuation inputs", "Exit EV/EBIT", None, 22.0, "x", "Comps"),
            ("Forecast", "FY26 revenue growth", None, 0.14, "decimal", "Judgment"),
            ("Forecast", "FY26 EBIT margin", None, 0.455, "decimal", "Judgment"),
            ("Forecast", "Capex % sales", None, 0.02, "decimal", "FY25 run-rate"),
            ("Forecast", "ΔNWC % ΔSales", None, 0.05, "decimal", "Judgment"),
        ]
    )
    for i, row in enumerate(checklist, start=5):
        for j, v in enumerate(row, start=1):
            cell = ws.cell(i, j, v)
            cell.border = thin
            if j == 3:
                cell.fill = input_fill
                cell.font = input_font
            if j == 4 and isinstance(v, (int, float)):
                cell.number_format = pct if isinstance(v, float) and abs(v) <= 1 else num2
    autosize(ws)
    ws.column_dimensions["B"].width = 48

    # ----- 12_XBRL_Extract -----
    ws = wb.create_sheet("12_XBRL_Extract")
    ws["A1"] = "SEC companyfacts extract (keyed by period-end year, USD/shares)"
    ws["A1"].font = title_font
    ws["A2"] = "Raw XBRL pull from data.sec.gov — cross-check against audited statements on 03/04/05."
    ws["A4"] = "Concept"
    for j, y in enumerate(END_YEARS, start=2):
        ws.cell(4, j, y)
    ws.cell(4, 7, "Unit")
    style_header_row(ws, 4, 1, 7, header_font, header_fill)
    for i, r in enumerate(xbrl_rows, start=5):
        ws.cell(i, 1, r["concept"])
        for j, y in enumerate(END_YEARS, start=2):
            cell = ws.cell(i, j, r.get(str(y)))
            if isinstance(cell.value, (int, float)):
                cell.number_format = num2 if abs(cell.value) < 1000 else num
        ws.cell(i, 7, r["unit"])
    autosize(ws)
    ws.column_dimensions["A"].width = 70

    order = [
        "00_ReadMe",
        "01_Template_Links",
        "02_EDGAR_Sources",
        "03_IS_Historical",
        "04_BS_Historical",
        "05_CF_Historical",
        "06_Assumptions",
        "07_DCF",
        "08_WACC",
        "09_Sensitivity",
        "10_Output",
        "11_Input_Checklist",
        "12_XBRL_Extract",
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    out = ROOT / "FICO_DCF_Valuation_Model.xlsx"
    wb.save(out)

    (ROOT / "TEMPLATE_LINKS.md").write_text(
        """# Downloadable valuation templates

Direct and landing-page links for DCF / investor valuation Excel models.

| Source | Template | URL | Notes |
|---|---|---|---|
| Damodaran | Spreadsheet hub | https://pages.stern.nyu.edu/~adamodar/New_Home_Page/spreadsh.htm | Full free suite |
| Damodaran | fcffsimpleginzu.xlsx | https://pages.stern.nyu.edu/~adamodar/pc/fcffsimpleginzu.xlsx | **Best starting FCFF** (direct xlsx) |
| Damodaran | fcffginzu.xls | https://pages.stern.nyu.edu/~adamodar/pc/fcffginzu.xls | Full kitchen-sink FCFF |
| Damodaran | Equity spreadsheets | https://pages.stern.nyu.edu/~adamodar/New_Home_Page/eqspread.htm | 1/2/3-stage models |
| Damodaran | fcff2st.xls | https://pages.stern.nyu.edu/~adamodar/pc/fcff2st.xls | Two-stage FCFF |
| Damodaran | fcff3st.xls | https://pages.stern.nyu.edu/~adamodar/pc/fcff3st.xls | Three-stage FCFF |
| Damodaran | firmmult.xls | https://pages.stern.nyu.edu/~adamodar/pc/firmmult.xls | Implied multiples |
| CFI | DCF model template | https://corporatefinanceinstitute.com/resources/financial-modeling/dcf-model-template/ | 3-statement + DCF |
| CFI | Compact DCF xlsx | https://corporatefinanceinstitute.com/assets/DCF-Valuation-Compact-Complete.xlsx | Direct download |
| CT Acquisitions | 2026 free DCF | https://ctacquisitions.com/dcf-template-and-excel/ | 6-tab IB/PE style |
| Limelight | DCF calculator | https://www.golimelight.com/templates/dcf-calculator-template | WACC + scenarios |
| Excel Business Resource | Free DCF | https://excelbusinessresource.com/product/free-dcf-model-template-xls/ | FCFF dual TV |
| Keene Advisors | Free DCF | https://www.keeneadvisors.com/free-dcf-model-template | 3yr hist + 5yr proj |
| Macabacus | DCF | https://macabacus.com/excel/templates/discounted-cash-flow | Trial to download |
| Macabacus | Operating model | https://macabacus.com/excel/templates/operating-model | Trial |
| Macabacus | LBO templates | https://macabacus.com/lbo-model/templates | Trial |
| Vertex42 | Business valuation | https://www.vertex42.com/ExcelTemplates/business-valuation.html | Simple SME model |
| SEC EDGAR | FICO 10-K filings | https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000814547&type=10-K | Primary source |
| SEC API | FICO companyfacts | https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json | XBRL JSON |

This repo also includes a ready-to-edit FICO model: `fico-valuation/FICO_DCF_Valuation_Model.xlsx`.
""",
        encoding="utf-8",
    )

    # Numbers cheat sheet for README / users
    (ROOT / "FICO_NUMBERS_CHEATSHEET.md").write_text(
        f"""# FICO numbers you need (from SEC EDGAR)

All statement figures below are **$ thousands** unless noted. Source: FY2025 Form 10-K (filed 2025-11-07) and Q3 FY2026 10-Q (ended 2026-06-30).

## Income statement (FY)

| Line item | FY2025 | FY2024 | FY2023 |
|---|---:|---:|---:|
| On-premises and SaaS software | 740,145 | 711,340 | 640,182 |
| Professional services | 82,149 | 86,536 | 99,547 |
| Scores | 1,168,575 | 919,650 | 773,828 |
| **Total revenues** | **1,990,869** | **1,717,526** | **1,513,557** |
| Cost of revenues | 353,722 | 348,206 | 311,053 |
| Research and development | 188,347 | 171,940 | 159,950 |
| SG&A | 513,028 | 462,834 | 400,565 |
| Amortization of intangibles | — | 917 | 1,100 |
| Restructuring charges | 10,922 | — | — |
| **Operating income** | **924,850** | **733,629** | **642,830** |
| Interest expense, net | (133,647) | (105,638) | (95,546) |
| Other income, net | 11,392 | 14,034 | 6,340 |
| Income before taxes | 802,595 | 642,025 | 553,624 |
| Tax provision | 150,649 | 129,214 | 124,249 |
| **Net income** | **651,946** | **512,811** | **429,375** |
| Diluted EPS ($) | 26.54 | 20.45 | 16.93 |
| Diluted shares (000s) | 24,561 | 25,079 | 25,367 |

## Cash flow highlights (FY)

| Line item | FY2025 | FY2024 | FY2023 |
|---|---:|---:|---:|
| Operating cash flow | 778,807 | 632,964 | 468,915 |
| D&A | 14,952 | 13,827 | 14,638 |
| Share-based compensation | 156,667 | 149,439 | 123,847 |
| PP&E purchases | (8,922) | (8,884) | (4,237) |
| Capitalized internal-use software | (30,485) | (16,667) | — |
| Stock repurchases | (1,414,502) | (821,702) | (405,526) |

## Balance sheet / net debt bridge

| Line item | 2026-06-30 (10-Q) | 2025-09-30 (10-K) | 2024-09-30 |
|---|---:|---:|---:|
| Cash | 248,444 | 134,136 | 150,667 |
| Marketable securities | 56,093 | 54,625 | 45,289 |
| Accounts receivable | 592,530 | 529,148 | 426,642 |
| Total assets | 2,037,373 | 1,868,133 | 1,717,884 |
| Current maturities of debt | 300,000 | 399,541 | 15,000 |
| Long-term debt | 5,282,389 | 2,656,150 | 2,194,021 |
| **Total debt** | **5,582,389** | **3,055,691** | **2,209,021** |
| Deferred revenue (current) | 205,424 | 187,372 | 156,897 |
| Shares outstanding (000s, YE) | n/a in extract | 23,764 | 24,392 |

**Net debt for DCF (10-Q):** 5,582.4 − 248.4 − 56.1 = **$5,277.9 million**

## Market snapshot used in model

| Item | Value |
|---|---:|
| Price | ${PRICE:,.2f} |
| Shares outstanding | {SHARES_OUT_THOUSANDS/1000:.3f} million |
| Market cap | ${MKT_CAP_M:,.1f} million |
| Beta | {BETA} |

## Filing links

- 10-K FY2025: https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm
- 10-Q Q3 FY2026: https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm
- Companyfacts: https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json
""",
        encoding="utf-8",
    )
    return out


if __name__ == "__main__":
    path = build()
    print("Wrote", path)
