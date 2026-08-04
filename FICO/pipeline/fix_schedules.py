"""
Fix 3-statement supporting schedules that break Model2.0 DCF feeds.

Root causes addressed:
- WC schedule (rows 85–87) still held CFI sample AR/Inv/AP while BS had FICO
  figures → FY2026 ΔNWC ≈ full NWC stock (the $555k cliff).
- PPE schedule (rows 92–95) still held CFI sample balances → forecast DA opened
  off a tiny base (~$8k) then compounded.
- D88 (prior NWC) was empty → first ΔNWC formula unstable.
"""

from __future__ import annotations

from typing import Any, Dict

from openpyxl.styles import Font, PatternFill

from .named_range_map import HIST_YEARS, SHEET_3S

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
BLACK = Font(name="Calibri", color="000000")

# FY2020 bridges from FICO 10-K (used only as schedule openings)
FY2020_PPE = 46_419.0
FY2020_NWC = 334_180.0 - 23_033.0  # AR - AP


def _input(cell, value) -> None:
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL


def _link(cell, formula: str) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL


def fix_three_statement_schedules(wb, bundle: Dict[str, Any]) -> None:
    """Sync WC + PPE supporting schedules to injected BS/IS/CF history."""
    ws = wb[SHEET_3S]
    bs = bundle["balance_sheet"]
    is_ = bundle["income_statement"]
    cf = bundle["cash_flow"]

    hist_cols = ["E", "F", "G", "H", "I"]

    # --- Working capital schedule: hist must match BS line items ---
    for col, y in zip(hist_cols, HIST_YEARS):
        _input(ws[f"{col}85"], float(bs.loc["accounts_receivable", y]))
        _input(ws[f"{col}86"], float(bs.loc["inventory", y]))
        _input(ws[f"{col}87"], float(bs.loc["accounts_payable", y]))

    # Prior-year NWC for FY2021 ΔNWC (=E88-D88)
    _input(ws["D88"], FY2020_NWC)

    # Forecast WC rows already link to BS (J85=J42 etc.) — leave them.

    # --- PPE schedule: hist close = BS PPE; hist DA/Capex = IS/CF ---
    _input(ws["E92"], FY2020_PPE)
    for i, (col, y) in enumerate(zip(hist_cols, HIST_YEARS)):
        if i > 0:
            prev = hist_cols[i - 1]
            _link(ws[f"{col}92"], f"={prev}95")
        _input(ws[f"{col}93"], float(cf.loc["capex", y]))
        _input(ws[f"{col}94"], float(is_.loc["da", y]))
        # Force closing PPE to equal BS (removes sample-template drift)
        _input(ws[f"{col}95"], float(bs.loc["ppe_net", y]))

    # Forecast PPE opening must use BS FY25 PPE, not a drifted I95 from samples
    _link(ws["J92"], "=I44")
    for prev, col in zip("JKLM", "KLMN"):
        _link(ws[f"{col}92"], f"={prev}95")

    # Keep forecast capex/DA formulas (J93=J18, J94=J92*J11) — intact.
    # Soften DA% of PPE if absurd vs revenue (cap at 25%)
    for col in ["J", "K", "L", "M", "N"]:
        cell = ws[f"{col}11"]
        if isinstance(cell.value, (int, float)) and float(cell.value) > 0.25:
            _input(cell, 0.25)
