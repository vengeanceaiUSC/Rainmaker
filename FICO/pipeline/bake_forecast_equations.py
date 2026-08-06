"""Bake FULL calculation equations into 3-statement forecast rows (J–N).

vengeanceaiUSCMODEL8:
  • Operating NWC = 2.5% × Revenue (flat); AR plugs to identity
  • SG&A floor 15% / −75 bps grind
  • CapEx fast fade → 1% (≈ D&A by Y5)
  • D&A on average PP&E; labels SG&A / R&D
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill, Border, Side

from .assumption_explanations import write_assumption_explanations
from .equation_explanations import write_equation_comments
from .model8_assumptions import (
    CAPEX_FADE_WEIGHTS,
    CAPEX_STEADY_PCT,
    NWC_STEADY_PCT,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
    SGA_FLOOR_PCT,
    SGA_IMPROVEMENT_BPS,
)
from .named_range_map import SHEET_3S

BLUE = Font(name="Calibri", color="0000FF")
BLACK = Font(name="Calibri", color="000000")
EQ_FONT = Font(name="Consolas", size=9, color="000000")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

_FORECAST_COLS = ["J", "K", "L", "M", "N"]
_PREV = {"J": "I", "K": "J", "L": "K", "M": "L", "N": "M"}


def _input(cell, value, fmt: str | None = None) -> None:
    cell.value = value
    cell.font = BLUE
    cell.fill = INPUT_FILL
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def _formula(cell, formula: str, fmt: str | None = None) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL
    cell.border = THIN
    if fmt:
        cell.number_format = fmt


def _da_avg_ppe(col: str) -> str:
    return f"=(({col}92+({col}92+{col}93))/2)*{col}11"


def bake_forecast_equations(wb) -> None:
    if SHEET_3S not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_3S}")
    ws = wb[SHEET_3S]

    base_sga = f"(($I$28-{RESTRUCTURING_NORMALIZE_000s})/$I$24)"
    bps = SGA_IMPROVEMENT_BPS / 10_000.0
    floor = SGA_FLOOR_PCT

    for col, g in zip(_FORECAST_COLS, REVENUE_GROWTH_PATH):
        _input(ws[f"{col}7"], float(g), "0.0%")

    for col in _FORECAST_COLS:
        _formula(ws[f"{col}8"], "=$I$25/$I$24", "0.00%")

    for i, col in enumerate(_FORECAST_COLS):
        grind = bps * (i + 1)
        _formula(ws[f"{col}9"], f"=MAX({floor},{base_sga}-{grind})", "0.00%")
        _formula(ws[f"{col}10"], "=$I$29/$I$24", "0.00%")
    ws["B9"] = f"SG&A % of Revenue (floor {floor:.0%}, −{SGA_IMPROVEMENT_BPS:.0f}bps×t)"
    ws["B10"] = "R&D % of Revenue (equation)"

    for col in _FORECAST_COLS:
        _formula(ws[f"{col}11"], '=IF($I$44=0,0.25,$I$30/$I$44)', "0.00%")
        _formula(ws[f"{col}12"], '=IF($I$98=0,0.05,$I$101/$I$98)', "0.00%")
        _formula(ws[f"{col}13"], '=IF($I$33=0,0.21,$I$35/$I$33)', "0.00%")
        _input(ws[f"{col}16"], 0, "0")
        _formula(ws[f"{col}17"], '=IF($I$25=0,0,ROUND($I$48/$I$25*365,0))', "0")

    # Row 15: flat Operating NWC % of Revenue = 2.5%
    for col in _FORECAST_COLS:
        _input(ws[f"{col}15"], float(NWC_STEADY_PCT), "0.00%")
    ws["B15"] = f"Operating NWC % of Revenue (flat {NWC_STEADY_PCT:.1%} — asset-light)"
    ws["B11"] = "D&A % of Avg PP&E (rate=FY25 DA/PPE; $ uses avg base)"
    ws["B16"] = "Inventory (Days)"
    ws["B17"] = "Accounts Payable (Days)"

    for col, w in zip(_FORECAST_COLS, CAPEX_FADE_WEIGHTS):
        _formula(
            ws[f"{col}18"],
            f"=($I$68/$I$24)*{w}+{CAPEX_STEADY_PCT}*(1-{w})",
            "0.00%",
        )
    ws["B18"] = "CapEx % of Revenue (fast fade → 1%)"

    for col in _FORECAST_COLS:
        _input(ws[f"{col}19"], 0.0, "#,##0.0")
        _input(ws[f"{col}20"], 0.0, "#,##0.0")

    # ===== INCOME STATEMENT =====
    ws["B28"] = "SG&A Expense"
    ws["B29"] = "Research & Development (R&D)"

    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}24"], f"={prev}24*(1+{col}7)", "#,##0.0")
        _formula(ws[f"{col}25"], f"={col}24*{col}8", "#,##0.0")
        _formula(ws[f"{col}26"], f"={col}24-{col}25", "#,##0.0")
        _formula(ws[f"{col}28"], f"={col}24*{col}9", "#,##0.0")
        _formula(ws[f"{col}29"], f"={col}24*{col}10", "#,##0.0")
        _formula(ws[f"{col}30"], _da_avg_ppe(col), "#,##0.0")
        _formula(ws[f"{col}31"], f"={col}98*{col}12", "#,##0.0")
        _formula(ws[f"{col}32"], f"=SUM({col}28:{col}31)", "#,##0.0")
        _formula(ws[f"{col}33"], f"={col}26-{col}32", "#,##0.0")
        _formula(ws[f"{col}35"], f"={col}33*{col}13", "#,##0.0")
        _formula(ws[f"{col}36"], f"={col}33*(1-{col}13)", "#,##0.0")

    for r, note in {
        24: "eqn: PriorRev × (1 + growth)",
        25: "eqn: Rev × COGS%",
        26: "eqn: Rev − COGS",
        28: "eqn: Rev × SG&A%",
        29: "eqn: Rev × R&D%",
        30: "eqn: AvgPPE × DA%",
        31: "eqn: Debt_open × Int%",
        33: "eqn: GP − Expenses",
        35: "eqn: EBT × tax%",
        36: "eqn: EBT × (1 − tax%)",
    }.items():
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== BALANCE SHEET =====
    # NWC = Rev×nwc%; Deferred & AP from hist ratios; AR plugs so identity holds
    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}41"], f"={col}78", "#,##0.0")
        # Deferred first (needed for AR plug) — stored on WC 88; BS shows via TotLiab
        _formula(ws[f"{col}43"], f"={col}25*{col}16/365", "#,##0.0")
        _formula(ws[f"{col}44"], f"={col}95", "#,##0.0")
        _formula(ws[f"{col}48"], f"={col}25*{col}17/365", "#,##0.0")
        # AR = NWC + AP + Deferred − Inv  (NWC on row 89)
        _formula(
            ws[f"{col}42"],
            f"={col}89+{col}48+{col}88-{col}43",
            "#,##0.0",
        )
        _formula(ws[f"{col}45"], f"=SUM({col}41:{col}44)", "#,##0.0")
        _formula(ws[f"{col}49"], f"={col}100", "#,##0.0")
        _formula(ws[f"{col}50"], f"={col}48+{col}49+{col}88", "#,##0.0")
        _formula(ws[f"{col}52"], f"={prev}52+{col}20", "#,##0.0")
        _formula(ws[f"{col}53"], f"={prev}53+{col}33*(1-{col}13)", "#,##0.0")
        _formula(ws[f"{col}54"], f"=SUM({col}52:{col}53)", "#,##0.0")
        _formula(ws[f"{col}55"], f"={col}50+{col}54", "#,##0.0")
        _formula(ws[f"{col}57"], f"={col}55-{col}45", "#,##0.0")

    for r, note in (
        (42, "eqn: NWC + AP + Deferred − Inv  (plug to NWC%)"),
        (48, "eqn: COGS × AP_days / 365"),
        (50, "eqn: AP + Debt + DeferredRev"),
        (53, "eqn: PriorRE + EBT×(1−t)"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== CASH FLOW =====
    for col in _FORECAST_COLS:
        prev = _PREV[col]
        da = _da_avg_ppe(col).lstrip("=")
        _formula(ws[f"{col}62"], f"={col}33*(1-{col}13)", "#,##0.0")
        _formula(ws[f"{col}63"], f"={da}", "#,##0.0")
        _formula(ws[f"{col}64"], f"={col}90", "#,##0.0")
        _formula(
            ws[f"{col}65"],
            f"={col}33*(1-{col}13)+({da})-{col}90",
            "#,##0.0",
        )
        _formula(ws[f"{col}68"], f"={col}24*{col}18", "#,##0.0")
        _formula(ws[f"{col}69"], f"={col}68", "#,##0.0")
        _formula(ws[f"{col}72"], f"={col}19", "#,##0.0")
        _formula(ws[f"{col}73"], f"={col}20", "#,##0.0")
        _formula(ws[f"{col}74"], f"={col}19+{col}20", "#,##0.0")
        _formula(ws[f"{col}76"], f"={col}65-{col}69+{col}74", "#,##0.0")
        _formula(ws[f"{col}77"], f"={prev}41", "#,##0.0")
        _formula(ws[f"{col}78"], f"={col}77+{col}76", "#,##0.0")
        _formula(ws[f"{col}80"], f"={col}78-{col}41", "#,##0.0")

    for r, note in (
        (62, "eqn: EBT × (1 − tax%)"),
        (63, "eqn: AvgPPE × DA%"),
        (64, "eqn: ΔNWC from consolidated NWC% of sales"),
        (65, "eqn: NI + DA − ΔNWC"),
        (68, "eqn: Rev × CapEx%"),
        (76, "eqn: CFO − CapEx + Financing"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== WC SCHEDULE =====
    ws["B85"] = "Accounts Receivable (plugged to NWC%)"
    ws["B86"] = "Inventory"
    ws["B87"] = "Accounts Payable"
    ws["B88"] = "Deferred Revenue (contract liability)"
    ws["B89"] = "Net Working Capital (NWC)"
    ws["B90"] = "Change in NWC"

    for col in _FORECAST_COLS:
        prev = _PREV[col]
        # NWC first (driver), then components
        _formula(ws[f"{col}89"], f"={col}24*{col}15", "#,##0.0")  # Rev × NWC%
        _formula(ws[f"{col}86"], f"={col}43", "#,##0.0")
        _formula(ws[f"{col}87"], f"={col}48", "#,##0.0")
        _formula(ws[f"{col}88"], f"={col}24*IF($I$24=0,0,$I$88/$I$24)", "#,##0.0")
        _formula(ws[f"{col}85"], f"={col}42", "#,##0.0")
        # ΔNWC must use the prior BS NWC (I89 hist, then prior forecast).
        # Seeding Y1 prior at 2.5%×FY25 Rev broke CF↔BS articulation: AR stepped
        # down to policy NWC% with no cash release → constant ~$260M BS ERROR.
        _formula(ws[f"{col}90"], f"={col}89-{prev}89", "#,##0.0")

        if col == "J":
            _formula(ws["J92"], "=I44", "#,##0.0")
        else:
            _formula(ws[f"{col}92"], f"={prev}95", "#,##0.0")
        _formula(ws[f"{col}93"], f"={col}24*{col}18", "#,##0.0")
        _formula(ws[f"{col}94"], _da_avg_ppe(col), "#,##0.0")
        _formula(ws[f"{col}95"], f"={col}92+{col}93-{col}94", "#,##0.0")

        if col == "J":
            _formula(ws["J98"], "=I49", "#,##0.0")
        else:
            _formula(ws[f"{col}98"], f"={prev}100", "#,##0.0")
        _formula(ws[f"{col}99"], f"={col}19", "#,##0.0")
        _formula(ws[f"{col}100"], f"={col}98+{col}99", "#,##0.0")
        _formula(ws[f"{col}101"], f"={col}98*{col}12", "#,##0.0")

    ws["C88"] = "eqn: Rev × (FY25 Deferred/FY25 Rev)"
    ws["C88"].font = EQ_FONT
    ws["C89"] = "eqn: Rev × NWC%   (consolidated operating WC)"
    ws["C89"].font = EQ_FONT
    ws["C90"] = "eqn: NWCt − NWCt−1 (hist I89 is Y1 prior — keeps BS balanced)"
    ws["C90"].font = EQ_FONT
    ws["C94"] = "eqn: ((Open+Open+CapEx)/2) × DA%"
    ws["C94"].font = EQ_FONT

    # Hist IS labels too
    for col in list("EFGHI") + _FORECAST_COLS:
        pass
    ws["B28"] = "SG&A Expense"
    ws["B29"] = "Research & Development (R&D)"

    write_equation_comments(wb)
    write_assumption_explanations(wb)
    ws.column_dimensions["C"].width = 52
