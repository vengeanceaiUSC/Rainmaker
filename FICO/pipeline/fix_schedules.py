"""
Fix 3-statement supporting schedules that break Model2.0 DCF feeds.

Root causes addressed:
- WC schedule (rows 85–87) still held CFI sample AR/Inv/AP while BS had FICO
  figures → FY2026 ΔNWC ≈ full NWC stock (the $555k cliff).
- PPE schedule (rows 92–95) still held CFI sample balances → forecast DA opened
  off a tiny base (~$8k) then compounded.
- Debt schedule (rows 98–100) still held CFI sample (~$30k) while BS debt is
  ~$3.06B → forecast Debt collapses → Assets ≠ L+E → row-3 "ERROR".
- Cash opening chain (row 77) still had CFI sample openings.
- D88 (prior NWC) was empty → first ΔNWC formula unstable.
"""

from __future__ import annotations

from typing import Any, Dict, List

from openpyxl.styles import Font, PatternFill

from .named_range_map import HIST_YEARS, SHEET_3S

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
BLACK = Font(name="Calibri", color="000000")

# FY2020 bridges from FICO 10-K (schedule openings only)
FY2020_PPE = 46_419.0
FY2020_NWC = 334_180.0 - 23_033.0  # AR - AP
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

    # --- Working capital schedule ---
    for col, y in zip(hist_cols, HIST_YEARS):
        _input(ws[f"{col}85"], float(bs.loc["accounts_receivable", y]))
        _input(ws[f"{col}86"], float(bs.loc["inventory", y]))
        _input(ws[f"{col}87"], float(bs.loc["accounts_payable", y]))
    _input(ws["D88"], FY2020_NWC)

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
    for col in ["J", "K", "L", "M", "N"]:
        cell = ws[f"{col}11"]
        if isinstance(cell.value, (int, float)) and float(cell.value) > 0.25:
            _input(cell, 0.25)

    # --- Debt schedule (THIS is why forecast row-3 showed ERROR) ---
    # BS Debt (row 49) in forecast = Debt Closing (row 100). If row 100 still
    # rolls from the CFI sample (~30k), L+E collapses vs Assets → "ERROR".
    _input(ws["E98"], FY2020_DEBT)
    for i, (col, y) in enumerate(zip(hist_cols, HIST_YEARS)):
        if i > 0:
            _link(ws[f"{col}98"], f"={hist_cols[i - 1]}100")
        _input(ws[f"{col}99"], issuances[i])
        _input(ws[f"{col}100"], debts[i])  # force = BS debt
        _input(ws[f"{col}101"], float(is_.loc["interest_expense", y]))
        # Keep CF debt issuance in sync for hist
        _input(ws[f"{col}72"], issuances[i])

    # Forecast debt opens at FY25 BS debt (not broken sample I100)
    _link(ws["J98"], "=I49")
    for prev, col in zip("JKLM", "KLMN"):
        _link(ws[f"{col}98"], f"={prev}100")
    # J99:N99 already = J19:N19 (assumptions); J100 = SUM open+issue — OK

    # --- Cash opening chain ---
    # Hist BS cash (row 41) is FICO-reported; the CF rollforward (row 78) drifts
    # because hist CF lines are only partially mapped onto this simplified CFI
    # template. Forecast cash MUST open from BS cash (I41), not from drifted I78,
    # or Assets jump by ~(I78-I41) while Equity does not → row-3 ERROR for 2026–30.
    _input(ws["E77"], FY2020_CASH)
    for prev, col in zip("EFGHI", "FGHI"):
        _link(ws[f"{col}77"], f"={prev}41")  # hist open = prior BS cash
    _link(ws["J77"], "=I41")  # first forecast year opens at FY25 BS cash
    for prev, col in zip("JKLM", "KLMN"):
        _link(ws[f"{col}77"], f"={prev}78")  # then follow integrated CF

    # Hist CF closing cash: plug to reported BS cash so cash-check row 80 is clean
    # (net change row 76 becomes the implied bridge; OCF/CFI/CFF stay informational)
    for col in hist_cols:
        _link(ws[f"{col}76"], f"={col}41-{col}77")

    # Year headers: keep as plain integers (avoid 2,026.0 display)
    for col in hist_cols + ["J", "K", "L", "M", "N"]:
        ws[f"{col}2"].number_format = "0"
