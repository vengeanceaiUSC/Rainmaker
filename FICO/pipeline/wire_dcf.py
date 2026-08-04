"""
Wire the DCF sheet to the 3-statement sheet with live Excel formulas.

Fixes:
1. EBIT / D&A / ΔNWC / CapEx pull from projected 3-statement columns (not CSV hardcodes)
2. CapEx is year-specific (not locked to $D$15)
3. Transaction CF no longer multiplies by Year Fraction when EV uses XNPV
   (XNPV already time-weights by calendar dates — avoid double stub pro-rating)
4. Exit / Terminal Value uses Exit EV/EBITDA (D8) on final-year EBITDA,
   with Gordon growth shown as a cross-check in a notes cell
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .named_range_map import SHEET_3S, SHEET_DCF

# DCF projection cols E..I map to 3-statement forecast cols J..N (FY1..FY5 forecast)
_DCF_COLS = ["E", "F", "G", "H", "I"]
_S3_FORECAST_COLS = ["J", "K", "L", "M", "N"]

LINK_FILL = PatternFill("solid", fgColor="E2EFDA")  # green = linked formula
BLACK = Font(name="Calibri", color="000000")


def _link(cell, formula: str) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL


def wire_dcf_to_three_statement(wb) -> None:
    """Overwrite DCF driver/structure cells with correct linking formulas."""
    if SHEET_DCF not in wb.sheetnames or SHEET_3S not in wb.sheetnames:
        raise RuntimeError("Template must contain both 3 Statement Model and DCF Model sheets")

    dcf = wb[SHEET_DCF]
    s3 = SHEET_3S  # quoted in formulas below

    # --- 1 & 2: UFCF drivers + CapEx linked to 3-statement forecast years ---
    for dcol, scol in zip(_DCF_COLS, _S3_FORECAST_COLS):
        # EBIT ≈ Gross Profit - Salaries - Rent/Overhead - D&A  (excludes Interest)
        _link(
            dcf[f"{dcol}21"],
            f"='{s3}'!{scol}26-'{s3}'!{scol}28-'{s3}'!{scol}29-'{s3}'!{scol}30",
        )
        # D&A from income statement / depreciation schedule feed
        _link(dcf[f"{dcol}23"], f"='{s3}'!{scol}30")
        # CapEx from cash flow investing line (varies by year; NOT $D$15)
        _link(dcf[f"{dcol}24"], f"='{s3}'!{scol}68")
        # Change in NWC from CF working-capital line
        _link(dcf[f"{dcol}25"], f"='{s3}'!{scol}64")

    # D15 Capex assumption cell: point at first forecast year CapEx (documentation only)
    _link(dcf["D15"], f"='{s3}'!J68")
    dcf["B15"] = "Capex (FY1 forecast, linked)"

    # --- 3: Transaction CF — full cash flows; XNPV handles timing via dates ---
    # Entry column D stays as-is structurally; projection years drop * yearfrac
    for col in _DCF_COLS:
        _link(dcf[f"{col}28"], f"={col}27+{col}26")
        _link(dcf[f"{col}29"], f"={col}27+{col}26")
    # Exit column: TV + final UFCF (no yearfrac)
    _link(dcf["J28"], "=J27+J26")
    _link(dcf["J29"], "=J27")  # keep secondary row as TV-only (matches prior layout intent)

    # Clear misleading year-fraction dependency note by labeling row 20 as unused for EV
    dcf["B20"] = "Year Fraction (display only; XNPV uses dates — not applied to CF)"

    # --- 4: Terminal Value at exit (col J) ---
    # Primary: Exit EV/EBITDA on final projected EBITDA = EBIT + D&A
    # J27 was =N20 (orphan AVERAGE). Replace with real TV.
    _link(dcf["J27"], "=($I$21+$I$23)*$D$8")
    dcf["B27"] = "(Entry)/Exit TV"

    # Document Gordon cross-check (does not feed EV unless user swaps J27)
    dcf["L17"] = "Terminal value cross-check"
    dcf["L17"].font = Font(name="Calibri", bold=True, color="1F4E79")
    dcf["L18"] = "Exit EV/EBITDA TV (in J27)"
    dcf["M18"] = "=($I$21+$I$23)*$D$8"
    dcf["M18"].fill = LINK_FILL
    dcf["L19"] = "Gordon TV: UFCF_n*(1+g)/(WACC-g)"
    dcf["M19"] = "=IF(($D$6-$D$7)<=0,\"WACC must exceed g\",$I$26*(1+$D$7)/($D$6-$D$7))"
    dcf["M19"].fill = LINK_FILL
    dcf["L20"] = "Switch J27 to Gordon: copy M19 → J27"
    dcf.column_dimensions["L"].width = 36
    dcf.column_dimensions["M"].width = 18

    # Ensure EV still uses XNPV on the corrected Transaction CF row
    # D32 already =XNPV(D6,D28:J28,D18:J18) — leave it
    if not (isinstance(dcf["D32"].value, str) and "XNPV" in str(dcf["D32"].value).upper()):
        _link(dcf["D32"], "=XNPV(D6,D28:J28,D18:J18)")

    # Labels clarifying linkage
    dcf["B21"] = "EBIT (linked to 3-stmt)"
    dcf["B23"] = "Plus: D&A (linked to 3-stmt)"
    dcf["B24"] = "Less: Capex (linked to 3-stmt CF)"
    dcf["B25"] = "Less: ΔNWC (linked to 3-stmt CF)"
