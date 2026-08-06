"""
Wire the DCF sheet to the 3-statement sheet with live Excel formulas.

Model2.0 fixes:
1. EBIT / D&A / ΔNWC / CapEx ← 3-statement forecast (not CSV hardcodes)
2. CapEx year-specific (not $D$15)
3. Transaction CF not × Year Fraction (XNPV already dates cash flows)
4. Terminal Value = Gordon primary (aligned with g=3%); exit multiple as cross-check
5. Cash taxes = EBIT × tax rate (unlevered / capital-structure neutral)
6. Mid-year discounting via March 31 cash-flow dates (FICO FYE Sep 30)
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill

from .named_range_map import SHEET_3S, SHEET_DCF

_DCF_COLS = ["E", "F", "G", "H", "I"]
_S3_FORECAST_COLS = ["J", "K", "L", "M", "N"]

LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
BLACK = Font(name="Calibri", color="000000")


def _link(cell, formula: str) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL


def wire_dcf_to_three_statement(wb) -> None:
    if SHEET_DCF not in wb.sheetnames or SHEET_3S not in wb.sheetnames:
        raise RuntimeError("Template must contain both 3 Statement Model and DCF Model sheets")

    dcf = wb[SHEET_DCF]
    s3 = SHEET_3S

    # --- UFCF drivers linked to 3-statement ---
    for dcol, scol in zip(_DCF_COLS, _S3_FORECAST_COLS):
        _link(
            dcf[f"{dcol}21"],
            f"='{s3}'!{scol}26-'{s3}'!{scol}28-'{s3}'!{scol}29-'{s3}'!{scol}30",
        )
        # Unlevered cash tax on EBIT (NOT levered IS tax dollars)
        _link(dcf[f"{dcol}22"], f"={dcol}21*$D$5")
        _link(dcf[f"{dcol}23"], f"='{s3}'!{scol}30")
        _link(dcf[f"{dcol}24"], f"='{s3}'!{scol}68")
        _link(dcf[f"{dcol}25"], f"='{s3}'!{scol}64")

    _link(dcf["D15"], f"='{s3}'!J68")
    dcf["B15"] = "Capex (FY1 forecast, linked)"
    dcf["B5"] = "Tax Rate (unlevered / on EBIT)"
    dcf["B22"] = "Less: Unlevered Cash Taxes (EBIT×t)"

    # --- Transaction CF: full CFs; no year-fraction multiply ---
    for col in _DCF_COLS:
        _link(dcf[f"{col}28"], f"={col}27+{col}26")
        _link(dcf[f"{col}29"], f"={col}27+{col}26")
    _link(dcf["J28"], "=J27+J26")
    _link(dcf["J29"], "=J27")
    dcf["B20"] = "Year Fraction (display only; not applied to CF)"

    # --- Mid-year dates (FICO FYE 9/30 → cash flows at 3/31) ---
    # E18 was DATE(YEAR($D$10)+E19,6,30). Use month 3 for mid-year convention.
    for col in _DCF_COLS:
        _link(dcf[f"{col}18"], f"=DATE(YEAR($D$10)+{col}19,3,31)")
    dcf["B18"] = "Date (mid-year convention)"

    # --- Terminal Value: Gordon primary (consistent with g in D7) ---
    _link(
        dcf["J27"],
        '=IF(($D$6-$D$7)<=0,"WACC must exceed g",$I$26*(1+$D$7)/($D$6-$D$7))',
    )
    dcf["B27"] = "(Entry)/Exit TV — Gordon"

    dcf["L17"] = "Terminal value methods"
    dcf["L17"].font = Font(name="Calibri", bold=True, color="1F4E79")
    dcf["L18"] = "Gordon TV (in J27) — primary"
    _link(dcf["M18"], "=J27")
    dcf["L19"] = "Exit EV/EBITDA cross-check"
    _link(dcf["M19"], "=($I$21+$I$23)*$D$8")
    dcf["L20"] = "Implied exit multiple vs Gordon"
    _link(dcf["M20"], "=IF(($I$21+$I$23)=0,0,J27/($I$21+$I$23))")
    dcf.column_dimensions["L"].width = 36
    dcf.column_dimensions["M"].width = 18

    _link(dcf["D32"], "=XNPV(D6,D28:J28,D18:J18)")
    dcf["B32"] = "Enterprise Value (XNPV, mid-year dates)"

    dcf["B21"] = "EBIT (linked to 3-stmt)"
    dcf["B23"] = "Plus: D&A (linked to 3-stmt)"
    dcf["B24"] = "Less: Capex (linked to 3-stmt CF)"
    dcf["B25"] = "Less: ΔNWC (linked to 3-stmt CF)"
