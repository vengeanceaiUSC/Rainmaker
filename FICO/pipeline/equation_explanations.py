"""~20-word explanations for every math equation on the 3-statement + key DCF links."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl.comments import Comment

from .assumption_explanations import URL_10K, URL_FACTS, URL_GUIDANCE, URL_10Q
from .named_range_map import SHEET_3S, SHEET_DCF

_AUTHOR = "vengeanceaiUSCMODEL8"
_FORECAST = ("J", "K", "L", "M", "N")

# Optional source URL by 3S row (shown in equation comments as LINK:)
_EQ_SOURCE_URL = {
    8: URL_10K,
    9: URL_10K,
    10: URL_10K,
    11: URL_10K,
    12: URL_10K,
    13: URL_10K,
    15: URL_FACTS,
    17: URL_10K,
    18: URL_10K,
    24: URL_GUIDANCE,
    25: URL_10K,
    42: URL_FACTS,
    48: URL_10K,
    68: URL_10K,
    98: URL_10Q,
}

# row -> (name, formula_pattern, explain ~20 words)
# Patterns use {c}=this col, {p}=prior col for display.
THREE_STATEMENT_EQS: List[Tuple[int, str, str, str]] = [
    (2, "Forecast year label", "={p}2+1",
     "Adds one year to the prior column’s fiscal-year label so forecast headers stay sequential."),
    (3, "Balance sheet check", '=IFERROR(IF(ABS({c}57)>1,"ERROR","OK"),"OK")',
     "Flags ERROR if Assets ≠ L+E by more than $1k; otherwise OK. Protects the plug."),
    (8, "COGS %", "=$I$25/$I$24",
     "FY25 COGS divided by FY25 revenue, held flat so gross margin matches latest reported mix."),
    (9, "SG&A %", "=MAX(15%,(($I$28-10922)/$I$24)-75bps×t)",
     "Normalized FY25 SG&A% after stripping restructuring, minus 75bps×year, floored at 15%."),
    (10, "R&D %", "=$I$29/$I$24",
     "FY25 R&D over FY25 revenue, held flat so product investment scales with sales."),
    (11, "D&A % of Avg PP&E", "=IF($I$44=0,0.25,$I$30/$I$44)",
     "FY25 DA/PPE rate; dollars use average PP&E so CapEx additions are depreciated."),
    (12, "Interest % of debt", "=IF($I$98=0,0.05,$I$101/$I$98)",
     "FY25 interest over opening debt approximates the book coupon; 5% fallback if no debt."),
    (13, "Tax % of EBT", "=IF($I$33=0,0.21,$I$35/$I$33)",
     "FY25 tax over FY25 EBT is the effective book rate; 21% statutory fallback if EBT zero."),
    (15, "Operating NWC % of Sales", "=2.5% flat (policy)",
     "Flat NWC/Sales for asset-light software; kills gross-AR/deferred double-count drain."),
    (17, "AP days", "=IF($I$25=0,0,ROUND($I$48/$I$25*365,0))",
     "Accounts payable over COGS × 365 from FY25; payable timing for operating NWC."),
    (18, "CapEx % fade", "=($I$68/$I$24)*w+1%*(1−w)",
     "Blends FY25 CapEx/Sales toward 1% steady so peak capitalized software does not persist."),
    (24, "Revenue", "={p}24*(1+{c}7)",
     "Prior-year revenue grown by this year’s policy growth rate; top-line driver of the model."),
    (25, "COGS $", "={c}24*{c}8",
     "Revenue times COGS%; converts the held margin assumption into dollar cost of sales."),
    (26, "Gross profit", "={c}24-{c}25",
     "Revenue minus COGS; contribution after direct cost of revenues before operating expenses."),
    (28, "SG&A $", "={c}24*{c}9",
     "Revenue times SG&A%; operating opex after the −75bps efficiency grind on the rate."),
    (29, "R&D $", "={c}24*{c}10",
     "Revenue times R&D%; research spend scales with sales at the FY25 reinvestment rate."),
    (30, "D&A $", "=((Open+Open+CapEx)/2)×DA%",
     "Average PP&E times DA% so new CapEx is depreciated in the year added."),
    (31, "Interest $", "={c}98*{c}12",
     "Opening debt times interest%; coupon on the debt stock while issuance policy is zero."),
    (32, "Total expenses", "=SUM({c}28:{c}31)",
     "Sum of SGA, R&D, D&A, and interest; all operating and financing costs above EBT."),
    (33, "EBT", "={c}26-{c}32",
     "Gross profit minus total expenses; earnings before tax in this simplified template."),
    (35, "Taxes $", "={c}33*{c}13",
     "EBT times effective tax rate; book tax expense for the income statement."),
    (36, "Net earnings", "={c}33*(1-{c}13)",
     "EBT kept after tax; full equation, not a pointer, so NI tracks the tax rate driver."),
    (41, "Cash (BS)", "={c}78",
     "Balance-sheet cash equals cash-flow closing cash so the three statements stay linked."),
    (42, "AR $", "={c}24*{c}15/365",
     "Revenue × AR days / 365; receivables from the days-sales-outstanding assumption."),
    (43, "Inventory $", "={c}25*{c}16/365",
     "COGS × inventory days / 365; stays zero because inventory days are policy zero."),
    (44, "PP&E (BS)", "={c}95",
     "Net PP&E equals the PPE schedule closing balance after CapEx and depreciation."),
    (45, "Total assets", "=SUM({c}41:{c}44)",
     "Cash + AR + inventory + PP&E; simplified asset base used for the balance-sheet check."),
    (48, "AP $", "={c}25*{c}17/365",
     "COGS × AP days / 365; payables funded by the payable-days assumption."),
    (49, "Debt (BS)", "={c}100",
     "Debt equals the debt-schedule closing balance (open + issuance, issuance usually 0)."),
    (50, "Total liabilities", "=SUM({c}48:{c}49)",
     "Accounts payable plus debt; simplified liability total in this teaching template."),
    (52, "Equity capital", "={p}52+{c}20",
     "Prior equity capital plus equity issued/(repurchased); policy keeps issuance at zero."),
    (53, "Retained earnings", "={p}53+{c}33*(1-{c}13)",
     "Prior RE plus this year’s net earnings; accumulates NI into equity without dividends."),
    (54, "Shareholders’ equity", "=SUM({c}52:{c}53)",
     "Equity capital plus retained earnings; book equity for the L+E side of the check."),
    (55, "Total L+E", "={c}50+{c}54",
     "Liabilities plus equity; must equal total assets when the model is in balance."),
    (57, "BS imbalance", "={c}55-{c}45",
     "L+E minus assets; should be ~0. Nonzero means a link or plug is broken."),
    (62, "CF net earnings", "={c}33*(1-{c}13)",
     "Same NI equation as the IS, written out so CFO does not hide behind a bare pointer."),
    (63, "CF + D&A", "=AvgPPE×DA%",
     "Adds back non-cash D&A on average PP&E in the cash-from-operations bridge."),
    (64, "CF − ΔNWC", "={c}90",
     "ΔNWC from GrossAR+Inv−AP−Deferred; increase uses cash in the CFO build."),
    (65, "Cash from operations", "=NI+DA−ΔNWC",
     "Full CFO: NI + avg-PPE D&A − ΔNWC written as one equation."),
    (68, "CapEx $", "={c}24*{c}18",
     "Revenue times CapEx%; investing outflow used in CF and the PPE roll-forward."),
    (69, "Cash from investing", "={c}68",
     "In this template CFI is CapEx only; equals the CapEx dollar line above."),
    (72, "Debt issuance CF", "={c}19",
     "Pulls the debt issuance/(repayment) policy into financing cash flow."),
    (73, "Equity issuance CF", "={c}20",
     "Pulls the equity issuance/(buyback) policy into financing cash flow."),
    (74, "Cash from financing", "={c}19+{c}20",
     "Debt plus equity financing cash; both policies are zero in the base case."),
    (76, "Δ Cash", "={c}65-{c}69+{c}74",
     "CFO minus CapEx plus financing; net change that rolls opening cash to closing."),
    (77, "Opening cash", "={p}41",
     "Starts at prior balance-sheet cash so the cash roll-forward matches history."),
    (78, "Closing cash", "={c}77+{c}76",
     "Opening cash plus Δ cash; feeds BS cash and the cash-flow check."),
    (80, "Cash check", "={c}78-{c}41",
     "Closing CF cash minus BS cash; must be zero when statements are linked."),
    (85, "WC: Gross AR", "={c}42",
     "WC gross AR mirrors BS gross AR used in standard DSO."),
    (86, "WC: Inventory", "={c}43",
     "Working-capital inventory mirrors BS inventory (zero for this software business)."),
    (87, "WC: AP", "={c}48",
     "Working-capital AP mirrors BS AP so NWC uses the same operating balances."),
    (88, "WC: Deferred Revenue", "={c}24×(I88/I24)",
     "Contract liability projected at FY25 deferred/sales; separate from AR."),
    (89, "Net working capital", "={c}85+{c}86-{c}87-{c}88",
     "GrossAR + inventory − AP − Deferred; standard operating NWC."),
    (90, "Change in NWC", "={c}89-{p}89",
     "Year-over-year NWC change; the cash investment/(release) used in CFO and FCFF."),
    (92, "PPE opening", "={p}95 or I44",
     "Prior closing PP&E (FY1 opens at last historical I44); base for CapEx and D&A."),
    (93, "PPE + CapEx", "={c}24*{c}18",
     "Same CapEx dollars as CF; additions that grow the gross PP&E stock."),
    (94, "PPE − D&A", "=((Open+Open+CapEx)/2)×DA%",
     "Depreciation on average PP&E; reduces net PP&E and adds back in FCFF."),
    (95, "PPE closing", "={c}92+{c}93-{c}94",
     "Open + CapEx − D&A; closing net PP&E posted to the balance sheet."),
    (98, "Debt opening", "={p}100 or I49",
     "Prior closing debt (FY1 opens at historical I49); base for interest and issuance."),
    (99, "Debt issuance", "={c}19",
     "Links the debt policy row into the schedule; base case keeps this at zero."),
    (100, "Debt closing", "={c}98+{c}99",
     "Open plus issuance/(repayment); closing debt posted to the balance sheet."),
    (101, "Interest (schedule)", "={c}98*{c}12",
     "Opening debt × interest%; same economics as IS interest, from the debt schedule."),
]

# DCF: (sheet_coord_template with {y} for year col E-I, name, formula note, explain)
DCF_EQS: List[Tuple[str, str, str, str]] = [
    ("D6", "WACC", "=R15",
     "Pulls the live CAPM WACC from the baked CAPM block so discounting uses researched Ke/Rd."),
    ("D5", "Tax rate", "=3S!J13",
     "Effective tax rate linked from the 3-statement so FCFF and WACC use the same t."),
    ("E21", "EBIT", "=3S EBT + Interest",
     "Operating profit before interest/tax: 3S earnings before tax plus interest (≡ GP−SGA−R&D−DA)."),
    ("E22", "Unlevered tax", "=EBIT×D5",
     "Taxes as if all-equity financed; D5 is the 3S tax rate for capital-structure-neutral FCFF."),
    ("E23", "D&A add-back", "=3S IS D&A",
     "Adds non-cash depreciation from the 3-statement into unlevered free cash flow."),
    ("E24", "CapEx", "=3S CF CapEx",
     "Subtracts growth CapEx linked from the 3-statement investing line."),
    ("E25", "ΔNWC", "=3S WC ΔNWC",
     "Subtracts WC-schedule ΔNWC (row 90) when operating working capital rises with revenue."),
    ("E26", "Unlevered FCF", "=EBIT−tax+DA−CapEx−ΔNWC",
     "FCFF available to debt and equity; the cash flow discounted in the DCF."),
    ("J27", "Exit TV", "=EBITDA_n×exit multiple",
     "Terminal enterprise value at year 5 using exit EV/EBITDA (primary exit method)."),
    ("E28", "Transaction CF", "=FCFF+TV (TV only in exit year)",
     "Periodic cash for XNPV: explicit FCFF each year; exit TV added in the terminal year."),
    ("D32", "Enterprise value", "=XNPV(WACC, CFs, mid-year dates)",
     "Present value of FCFF + TV using mid-year dates and the live WACC."),
    ("D35", "Equity value", "=EV+Cash−Debt",
     "Equity bridge: enterprise value plus cash minus gross debt (net debt adjustment)."),
    ("D37", "Value / share", "=Equity/Shares",
     "Intrinsic equity value per share using diluted shares outstanding."),
    ("R15", "WACC CAPM", "=We×Ke+Wd×Rd×(1−t)",
     "Weighted average cost of capital from CAPM Ke and after-tax cost of debt."),
]


def _comment(text: str) -> Comment:
    c = Comment(text, _AUTHOR)
    c.width = 300
    c.height = 110
    return c


def write_equation_comments(wb) -> int:
    """Attach ~20-word commentary to every forecast math cell + key DCF cells."""
    n = 0
    if SHEET_3S in wb.sheetnames:
        ws = wb[SHEET_3S]
        prev = {"J": "I", "K": "J", "L": "K", "M": "L", "N": "M"}
        for row, name, pattern, explain in THREE_STATEMENT_EQS:
            link = _EQ_SOURCE_URL.get(row, URL_10K)
            for col in _FORECAST:
                cell = ws[f"{col}{row}"]
                if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                    if row in (7, 16, 19, 20) and isinstance(cell.value, (int, float)):
                        pass
                    else:
                        continue
                text = (
                    f"{name}\nFORMULA: {pattern.replace('{c}', col).replace('{p}', prev[col])}\n"
                    f"WHY: {explain}\nLINK: {link}"
                )
                cell.comment = _comment(text)
                n += 1
            b = ws[f"B{row}"]
            b.comment = _comment(
                f"{name}\nWHY: {explain}\nPATTERN: {pattern}\nLINK: {link}"
            )
            n += 1

        # Policy inputs (not formulas) still get commentary + source link
        policy = {
            7: (
                "Revenue growth",
                "Yellow policy path 27/16/13/10/7%; Y1≈FY26 guidance, then fade to terminal.",
                URL_GUIDANCE,
            ),
            16: (
                "Inventory days",
                "Hard zero — FICO is software/scores; no inventory cycle to fund.",
                URL_10K,
            ),
            19: (
                "Debt issuance",
                "Policy zero — hold debt stock flat; financing excluded from FCFF.",
                URL_10Q,
            ),
            20: (
                "Equity issuance",
                "Policy zero — buybacks/issuance are financing, not in FCFF.",
                URL_10Q,
            ),
        }
        for row, (name, explain, link) in policy.items():
            for col in _FORECAST:
                ws[f"{col}{row}"].comment = _comment(
                    f"{name}\nWHY: {explain}\nLINK: {link}"
                )
                n += 1

    if SHEET_DCF in wb.sheetnames:
        ws = wb[SHEET_DCF]
        for coord, name, pattern, explain in DCF_EQS:
            cell = ws[coord]
            cell.comment = _comment(f"{name}\nFORMULA: {pattern}\nWHY: {explain}")
            n += 1
        # Replicate year-column FCFF comments E-I for rows 21-26,28
        for col in "EFGHI":
            for row, name, pattern, explain in [
                (21, "EBIT", "GP−SGA−R&D−DA", "Operating profit from 3-statement links."),
                (22, "Unlevered tax", "EBIT×t", "All-equity tax on EBIT for FCFF."),
                (23, "D&A", "3S D&A", "Non-cash add-back from 3-statement."),
                (24, "CapEx", "3S CapEx", "Reinvestment outflow from 3-statement CF."),
                (25, "ΔNWC", "3S ΔNWC", "Working-capital investment from 3-statement."),
                (26, "UFCF", "EBIT−tax+DA−CapEx−ΔNWC", "Unlevered free cash flow discounted in DCF."),
                (28, "Transaction CF", "FCFF (+ TV in exit yr)", "Cash flow vector for XNPV."),
            ]:
                ws[f"{col}{row}"].comment = _comment(f"{name}\nFORMULA: {pattern}\nWHY: {explain}")
                n += 1
    return n


def export_all_equations_csv(out_dir: Path, *, ticker: str = "FICO") -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"MODEL8_{ticker}_ALL_EQUATIONS_EXPLAINED.csv"
    rows: List[Dict[str, str]] = []
    for row, name, pattern, explain in THREE_STATEMENT_EQS:
        rows.append(
            {
                "sheet": "3 Statement Model",
                "row": str(row),
                "name": name,
                "formula_pattern": pattern,
                "explanation_20_words": explain,
                "source_url": _EQ_SOURCE_URL.get(row, URL_10K),
            }
        )
    for coord, name, pattern, explain in DCF_EQS:
        rows.append(
            {
                "sheet": "DCF Model",
                "row": coord,
                "name": name,
                "formula_pattern": pattern,
                "explanation_20_words": explain,
                "source_url": URL_10K,
            }
        )
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(
            f,
            fieldnames=[
                "sheet",
                "row",
                "name",
                "formula_pattern",
                "explanation_20_words",
                "source_url",
            ],
        )
        w.writeheader()
        w.writerows(rows)
    return path
