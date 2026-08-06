"""
Fix 3-statement supporting schedules (vengeanceaiUSCMODEL6).

MODEL6 WC schedule:
  85 Gross AR, 86 Inv, 87 AP, 88 Deferred Revenue, 89 NWC, 90 ΔNWC
NWC = GrossAR + Inv − AP − Deferred (standard operating WC).
BS AR (42) is GROSS; Total Liabilities includes Deferred via row 88.
"""

from __future__ import annotations

from typing import Any, Dict, List

from openpyxl.styles import Font, PatternFill

from .named_range_map import HIST_YEARS, SHEET_3S

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
BLACK = Font(name="Calibri", color="000000")

FY2020_PPE = 46_419.0
FY2020_DEFERRED = 105_400.0
FY2020_GROSS_AR = 334_180.0
FY2020_AP = 23_033.0
# MODEL16 Op NWC excludes Deferred (Deferred is an explicit CFO cash source)
FY2020_NWC = FY2020_GROSS_AR - FY2020_AP  # AR+Inv−AP (Inv≈0)
FY2020_CASH = 157_394.0
FY2020_DEBT = 739_435.0


def _input(cell, value) -> None:
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL


def _link(cell, formula: str) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL


def _debt_issuance(debts: List[float]) -> List[float]:
    out = [debts[0] - FY2020_DEBT]
    for i in range(1, len(debts)):
        out.append(debts[i] - debts[i - 1])
    return out


def fix_three_statement_schedules(wb, bundle: Dict[str, Any]) -> None:
    """Sync WC / PPE / Debt / Cash schedules to injected FICO history."""
    ws = wb[SHEET_3S]
    bs = bundle["balance_sheet"]
    is_ = bundle["income_statement"]
    cf = bundle["cash_flow"]

    hist_cols = ["E", "F", "G", "H", "I"]
    debts = [float(bs.loc["total_debt", y]) for y in HIST_YEARS]
    issuances = _debt_issuance(debts)

    ws["B85"] = "Accounts Receivable (Gross)"
    ws["B86"] = "Inventory"
    ws["B87"] = "Accounts Payable"
    ws["B88"] = "Deferred Revenue (contract liability)"
    ws["B89"] = "Operating NWC (AR+Inv−AP; excludes Deferred)"
    ws["B90"] = "Change in Operating NWC"

    # --- Working capital + BS gross AR ---
    for col, y in zip(hist_cols, HIST_YEARS):
        ar = float(bs.loc["accounts_receivable", y])
        deferred = (
            float(bs.loc["deferred_revenue", y])
            if "deferred_revenue" in bs.index
            else 0.0
        )
        inv = float(bs.loc["inventory", y])
        ap = float(bs.loc["accounts_payable", y])
        _input(ws[f"{col}42"], ar)  # BS GROSS AR
        _input(ws[f"{col}85"], ar)
        _input(ws[f"{col}86"], inv)
        _input(ws[f"{col}87"], ap)
        _input(ws[f"{col}88"], deferred)
        # MODEL16: Operating NWC excludes Deferred (Deferred is explicit CFO source)
        op_nwc = ar + inv - ap
        _input(ws[f"{col}89"], op_nwc)
        # Total Liabilities = AP + Debt + Deferred
        _link(ws[f"{col}50"], f"={col}48+{col}49+{col}88")

    # FY2020 opening Op NWC approx (strip deferred credit from legacy constant if needed)
    _input(ws["D89"], FY2020_NWC)
    # Hist Δ Op NWC
    prev_nwc = FY2020_NWC
    for col, y in zip(hist_cols, HIST_YEARS):
        nwc = float(ws[f"{col}89"].value)
        _input(ws[f"{col}90"], nwc - prev_nwc)
        prev_nwc = nwc

    # --- PPE schedule ---
    _input(ws["E92"], FY2020_PPE)
    for i, (col, y) in enumerate(zip(hist_cols, HIST_YEARS)):
        if i > 0:
            _link(ws[f"{col}92"], f"={hist_cols[i - 1]}95")
        _input(ws[f"{col}93"], float(cf.loc["capex", y]))
        _input(ws[f"{col}94"], float(is_.loc["da", y]))
        _input(ws[f"{col}95"], float(bs.loc["ppe_net", y]))
    _link(ws["J92"], "=I44")
    for prev, col in zip("JKLM", "KLMN"):
        _link(ws[f"{col}92"], f"={prev}95")

    # --- Debt schedule ---
    _input(ws["E98"], FY2020_DEBT)
    for i, (col, y) in enumerate(zip(hist_cols, HIST_YEARS)):
        if i > 0:
            _link(ws[f"{col}98"], f"={hist_cols[i - 1]}100")
        _input(ws[f"{col}99"], issuances[i])
        _input(ws[f"{col}100"], debts[i])
        _input(ws[f"{col}101"], float(is_.loc["interest_expense", y]))
        _input(ws[f"{col}72"], issuances[i])

    _link(ws["J98"], "=I49")
    for prev, col in zip("JKLM", "KLMN"):
        _link(ws[f"{col}98"], f"={prev}100")

    # --- Cash opening chain ---
    # Forecast cash opens from BS cash (I41), not drifted CF close.
    _input(ws["E77"], FY2020_CASH)
    for prev, col in zip("EFGHI", "FGHI"):
        _link(ws[f"{col}77"], f"={prev}41")
    _link(ws["J77"], "=I41")
    for prev, col in zip("JKLM", "KLMN"):
        _link(ws[f"{col}77"], f"={prev}78")

    for col in hist_cols:
        _link(ws[f"{col}76"], f"={col}41-{col}77")

    for col in hist_cols + ["J", "K", "L", "M", "N"]:
        ws[f"{col}2"].number_format = "0"
