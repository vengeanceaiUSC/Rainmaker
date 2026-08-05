"""
Wire the DCF sheet to the 3-statement sheet with live Excel formulas.

vengeanceaiUSCMODEL7:
1. Drivers linked to 3-statement
2. Primary exit = 16x (blended); 25x shown as bull sensitivity
3. Mid-year dates via EDATE from transaction date
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill

from .model7_assumptions import EXIT_EV_EBITDA, EXIT_EV_EBITDA_BULL
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

    for dcol, scol in zip(_DCF_COLS, _S3_FORECAST_COLS):
        _link(
            dcf[f"{dcol}21"],
            f"='{s3}'!{scol}26-'{s3}'!{scol}28-'{s3}'!{scol}29-'{s3}'!{scol}30",
        )
        _link(dcf[f"{dcol}22"], f"={dcol}21*$D$5")
        _link(dcf[f"{dcol}23"], f"='{s3}'!{scol}30")
        _link(dcf[f"{dcol}24"], f"='{s3}'!{scol}68")
        _link(dcf[f"{dcol}25"], f"='{s3}'!{scol}64")

    _link(dcf["D15"], f"='{s3}'!J68")
    dcf["B15"] = "Capex (FY1 forecast, linked)"
    dcf["B5"] = "Tax Rate (unlevered / on EBIT)"
    dcf["B22"] = "Less: Unlevered Cash Taxes (EBIT×t)"

    for col in _DCF_COLS:
        _link(dcf[f"{col}28"], f"={col}27+{col}26")
        _link(dcf[f"{col}29"], f"={col}27+{col}26")
    _link(dcf["J28"], "=J27+J26")
    _link(dcf["J29"], "=J27")
    dcf["B20"] = "Year Fraction (display only; XNPV uses exact dates below)"

    # Mid-year dates: EDATE(transaction, (t-0.5)*12)
    # E19..I19 are period indices 1..5
    for col in _DCF_COLS:
        _link(dcf[f"{col}18"], f"=EDATE($D$9,({col}19-0.5)*12)")
    dcf["B18"] = "Date (mid-year via EDATE from transaction)"
    # Exit column date = last explicit + 0 (TV at same mid-year as Y5) — keep J18 linked
    _link(dcf["J18"], "=I18")

    # Primary TV = blended exit multiple in D8 (MODEL7 default 16x)
    _link(dcf["J27"], "=($I$21+$I$23)*$D$8")
    dcf["B27"] = f"(Entry)/Exit TV — Exit EV/EBITDA (primary {EXIT_EV_EBITDA:.0f}x)"

    dcf["L17"] = "Terminal value methods (MODEL7)"
    dcf["L17"].font = Font(name="Calibri", bold=True, color="1F4E79")
    dcf["L18"] = f"Exit EV/EBITDA @ D8 (primary {EXIT_EV_EBITDA:.0f}x)"
    _link(dcf["M18"], "=J27")
    dcf["L19"] = "Gordon cross-check"
    _link(
        dcf["M19"],
        '=IF(($D$6-$D$7)<=0,"WACC must exceed g",$I$26*(1+$D$7)/($D$6-$D$7))',
    )
    dcf["L20"] = "Implied Gordon exit multiple"
    _link(dcf["M20"], "=IF(($I$21+$I$23)=0,0,M19/($I$21+$I$23))")
    dcf["L21"] = f"Bull case TV @ {EXIT_EV_EBITDA_BULL:.0f}x (sensitivity only)"
    _link(dcf["M21"], f"=($I$21+$I$23)*{EXIT_EV_EBITDA_BULL}")
    dcf["N18"] = "← used in XNPV"
    dcf["N19"] = "← sanity check"
    dcf["N21"] = "← not in base case"
    dcf.column_dimensions["L"].width = 42
    dcf.column_dimensions["M"].width = 18

    _link(dcf["D32"], "=XNPV(D6,D28:J28,D18:J18)")
    dcf["B32"] = "Enterprise Value (XNPV, mid-year EDATE dates)"

    dcf["B21"] = "EBIT (linked to 3-stmt)"
    dcf["B23"] = "Plus: D&A (linked to 3-stmt)"
    dcf["B24"] = "Less: Capex (linked to 3-stmt CF)"
    dcf["B25"] = "Less: ΔNWC (linked to 3-stmt CF)"
