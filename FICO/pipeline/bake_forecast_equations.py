"""Bake FULL calculation equations into 3-statement forecast rows (J–N).

vengeanceaiUSCMODEL4: no pointer-only cells like =J36 for key lines.
Each forecast cell carries the economic equation (Rev×%, EBT×(1−t), etc.).
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill, Border, Side

from .assumption_explanations import write_assumption_explanations
from .equation_explanations import write_equation_comments
from .model4_assumptions import (
    CAPEX_FADE_WEIGHTS,
    CAPEX_STEADY_PCT,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
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


def bake_forecast_equations(wb) -> None:
    """Write hist-linked drivers + expanded IS/BS/CF forecast equations."""
    if SHEET_3S not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_3S}")
    ws = wb[SHEET_3S]

    # B4 banner is finalized in write_assumption_explanations (WHY/SOURCE guidance)
    base_sga = f"(($I$28-{RESTRUCTURING_NORMALIZE_000s})/$I$24)"
    bps = SGA_IMPROVEMENT_BPS / 10_000.0

    # ===== ASSUMPTION DRIVERS (hist-linked) =====
    for col, g in zip(_FORECAST_COLS, REVENUE_GROWTH_PATH):
        _input(ws[f"{col}7"], float(g), "0.0%")

    for col in _FORECAST_COLS:
        _formula(ws[f"{col}8"], "=$I$25/$I$24", "0.00%")

    for i, col in enumerate(_FORECAST_COLS):
        grind = bps * (i + 1)
        # Driver row shows the % equation (not dollars)
        _formula(ws[f"{col}9"], f"=MAX(0.05,{base_sga}-{grind})", "0.00%")
        _formula(ws[f"{col}10"], "=$I$29/$I$24", "0.00%")
    ws["B9"] = "SGA % of Revenue (equation)"
    ws["B10"] = "R&D % of Revenue (equation)"

    for col in _FORECAST_COLS:
        _formula(ws[f"{col}11"], '=IF($I$44=0,0.25,$I$30/$I$44)', "0.00%")
        _formula(ws[f"{col}12"], '=IF($I$98=0,0.05,$I$101/$I$98)', "0.00%")
        _formula(ws[f"{col}13"], '=IF($I$33=0,0.21,$I$35/$I$33)', "0.00%")
        _formula(ws[f"{col}15"], "=ROUND($I$42/$I$24*365,0)", "0")
        _input(ws[f"{col}16"], 0, "0")
        _formula(ws[f"{col}17"], '=IF($I$25=0,0,ROUND($I$48/$I$25*365,0))', "0")

    # CapEx % path (not dollars) then dollars = Rev × %
    for i, (col, w) in enumerate(zip(_FORECAST_COLS, CAPEX_FADE_WEIGHTS)):
        _formula(
            ws[f"{col}18"],
            f"=($I$68/$I$24)*{w}+{CAPEX_STEADY_PCT}*(1-{w})",
            "0.00%",
        )
    ws["B18"] = "CapEx % of Revenue (fade equation)"

    for col in _FORECAST_COLS:
        _input(ws[f"{col}19"], 0.0, "#,##0.0")
        _input(ws[f"{col}20"], 0.0, "#,##0.0")

    # ===== INCOME STATEMENT — full equations =====
    for i, col in enumerate(_FORECAST_COLS):
        prev = _PREV[col]

        # Revenue = prior × (1 + g)
        _formula(ws[f"{col}24"], f"={prev}24*(1+{col}7)", "#,##0.0")
        # COGS = Rev × COGS%
        _formula(ws[f"{col}25"], f"={col}24*{col}8", "#,##0.0")
        # Gross Profit = Rev − COGS
        _formula(ws[f"{col}26"], f"={col}24-{col}25", "#,##0.0")
        # SGA = Rev × SGA%
        _formula(ws[f"{col}28"], f"={col}24*{col}9", "#,##0.0")
        # R&D = Rev × R&D%
        _formula(ws[f"{col}29"], f"={col}24*{col}10", "#,##0.0")
        # D&A = Opening PPE × DA%
        _formula(ws[f"{col}30"], f"={col}92*{col}11", "#,##0.0")
        # Interest = Opening Debt × Int%
        _formula(ws[f"{col}31"], f"={col}98*{col}12", "#,##0.0")
        # Total expenses
        _formula(ws[f"{col}32"], f"=SUM({col}28:{col}31)", "#,##0.0")
        # EBT = GP − Expenses
        _formula(ws[f"{col}33"], f"={col}26-{col}32", "#,##0.0")
        # Tax = EBT × tax%
        _formula(ws[f"{col}35"], f"={col}33*{col}13", "#,##0.0")
        # Net Earnings = EBT × (1 − tax%)   ← FULL EQUATION, not a pointer
        _formula(ws[f"{col}36"], f"={col}33*(1-{col}13)", "#,##0.0")

    # Column C: text labels only — never start with "=" (avoids #NAME?)
    _c_notes = {
        24: "eqn: PriorRev × (1 + growth)",
        25: "eqn: Rev × COGS%",
        26: "eqn: Rev − COGS",
        28: "eqn: Rev × SGA%",
        29: "eqn: Rev × R&D%",
        30: "eqn: PPE_open × DA%",
        31: "eqn: Debt_open × Int%",
        33: "eqn: GP − Expenses",
        35: "eqn: EBT × tax%",
        36: "eqn: EBT × (1 − tax%)  — grows with Rev×(1+g)",
    }
    for r, note in _c_notes.items():
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== BALANCE SHEET — full equations =====
    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}41"], f"={col}78", "#,##0.0")  # cash = CF close
        _formula(ws[f"{col}42"], f"={col}24*{col}15/365", "#,##0.0")  # AR = Rev×days/365
        _formula(ws[f"{col}43"], f"={col}25*{col}16/365", "#,##0.0")
        _formula(ws[f"{col}44"], f"={col}95", "#,##0.0")  # PPE close from schedule
        _formula(ws[f"{col}45"], f"=SUM({col}41:{col}44)", "#,##0.0")
        _formula(ws[f"{col}48"], f"={col}25*{col}17/365", "#,##0.0")
        _formula(ws[f"{col}49"], f"={col}100", "#,##0.0")
        _formula(ws[f"{col}50"], f"=SUM({col}48:{col}49)", "#,##0.0")
        _formula(ws[f"{col}52"], f"={prev}52+{col}20", "#,##0.0")
        _formula(ws[f"{col}53"], f"={prev}53+{col}33*(1-{col}13)", "#,##0.0")  # RE += NI eqn
        _formula(ws[f"{col}54"], f"=SUM({col}52:{col}53)", "#,##0.0")
        _formula(ws[f"{col}55"], f"={col}50+{col}54", "#,##0.0")
        _formula(ws[f"{col}57"], f"={col}55-{col}45", "#,##0.0")

    for r, note in (
        (42, "eqn: Rev × AR_days / 365"),
        (48, "eqn: COGS × AP_days / 365"),
        (53, "eqn: PriorRE + EBT×(1−t)"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== CASH FLOW — full equations (NOT =J36 pointers) =====
    for col in _FORECAST_COLS:
        prev = _PREV[col]
        # Net Earnings = EBT × (1 − t)  [same equation as IS, written out]
        _formula(ws[f"{col}62"], f"={col}33*(1-{col}13)", "#,##0.0")
        # + D&A = PPE_open × DA%
        _formula(ws[f"{col}63"], f"={col}92*{col}11", "#,##0.0")
        # − ΔNWC = NWC_t − NWC_t−1
        _formula(ws[f"{col}64"], f"={col}88-{prev}88", "#,##0.0")
        # CFO = NI + DA − ΔNWC
        _formula(
            ws[f"{col}65"],
            f"={col}33*(1-{col}13)+{col}92*{col}11-({col}88-{prev}88)",
            "#,##0.0",
        )
        # CapEx = Rev × CapEx%  (positive outflow; subtracted in ΔCash)
        _formula(ws[f"{col}68"], f"={col}24*{col}18", "#,##0.0")
        _formula(ws[f"{col}69"], f"={col}68", "#,##0.0")
        # Financing
        _formula(ws[f"{col}72"], f"={col}19", "#,##0.0")
        _formula(ws[f"{col}73"], f"={col}20", "#,##0.0")
        _formula(ws[f"{col}74"], f"={col}19+{col}20", "#,##0.0")
        # ΔCash = CFO − CapEx + Financing
        _formula(ws[f"{col}76"], f"={col}65-{col}69+{col}74", "#,##0.0")
        _formula(ws[f"{col}77"], f"={prev}41", "#,##0.0")
        _formula(ws[f"{col}78"], f"={col}77+{col}76", "#,##0.0")
        _formula(ws[f"{col}80"], f"={col}78-{col}41", "#,##0.0")

    for r, note in (
        (62, "eqn: EBT × (1 − tax%)  — full eqn, not pointer to IS"),
        (63, "eqn: PPE_open × DA%"),
        (64, "eqn: NWC_t − NWC_t−1"),
        (65, "eqn: NI + DA − ΔNWC"),
        (68, "eqn: Rev × CapEx%"),
        (76, "eqn: CFO − CapEx + Financing"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== SCHEDULES — keep / reinforce equations =====
    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}85"], f"={col}42", "#,##0.0")
        _formula(ws[f"{col}86"], f"={col}43", "#,##0.0")
        _formula(ws[f"{col}87"], f"={col}48", "#,##0.0")
        _formula(ws[f"{col}88"], f"={col}85+{col}86-{col}87", "#,##0.0")
        _formula(ws[f"{col}89"], f"={col}88-{prev}88", "#,##0.0")
        # PPE: open, capex, da, close
        if col == "J":
            _formula(ws["J92"], "=I44", "#,##0.0")
        else:
            _formula(ws[f"{col}92"], f"={prev}95", "#,##0.0")
        _formula(ws[f"{col}93"], f"={col}24*{col}18", "#,##0.0")  # CapEx $ = Rev×%
        _formula(ws[f"{col}94"], f"={col}92*{col}11", "#,##0.0")
        _formula(ws[f"{col}95"], f"={col}92+{col}93-{col}94", "#,##0.0")
        # Debt
        if col == "J":
            _formula(ws["J98"], "=I49", "#,##0.0")
        else:
            _formula(ws[f"{col}98"], f"={prev}100", "#,##0.0")
        _formula(ws[f"{col}99"], f"={col}19", "#,##0.0")
        _formula(ws[f"{col}100"], f"={col}98+{col}99", "#,##0.0")
        _formula(ws[f"{col}101"], f"={col}98*{col}12", "#,##0.0")

    # ~20-word commentary on every forecast math equation cell first…
    write_equation_comments(wb)
    # …then assumption cells get richer HOW/WHY/SOURCE (overwrites rows 7–20)
    write_assumption_explanations(wb)

    ws.column_dimensions["C"].width = 50
