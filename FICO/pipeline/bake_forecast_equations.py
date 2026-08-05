"""Bake hist-linked Excel equations into 3-statement forecast assumption rows.

vengeanceaiUSCMODEL4: J8:N18 (and related) are FORMULAS off column I history,
not pre-baked dollar constants from Python.
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from .model4_assumptions import (
    CAPEX_FADE_WEIGHTS,
    CAPEX_STEADY_PCT,
    MODEL_NAME,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
    SGA_IMPROVEMENT_BPS,
)
from .named_range_map import SHEET_3S

BLUE = Font(name="Calibri", color="0000FF")
BLACK = Font(name="Calibri", color="000000")
BOLD = Font(name="Calibri", bold=True, color="1F4E79")
EQ_FONT = Font(name="Consolas", size=9, color="000000")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
NOTE_FILL = PatternFill("solid", fgColor="D6EAF8")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

_FORECAST_COLS = ["J", "K", "L", "M", "N"]


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
    """Overwrite 3-statement forecast assumption cells with hist-linked formulas."""
    if SHEET_3S not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_3S}")
    ws = wb[SHEET_3S]

    # ----- Header note above assumptions -----
    ws["B4"] = f"{MODEL_NAME}: green assumption cells = equations from FY25 (col I)"
    ws["B4"].font = BOLD
    ws["B4"].fill = NOTE_FILL

    # ----- Growth: explicit policy inputs (only forward judgment in this block) -----
    for col, g in zip(_FORECAST_COLS, REVENUE_GROWTH_PATH):
        _input(ws[f"{col}7"], float(g), "0.0%")
    ws["C7"] = "POLICY input (Y1≈FY26 guidance; fade→g)"
    ws["C7"].font = Font(name="Calibri", italic=True, size=8, color="666666")

    # ----- COGS% = FY25 COGS / Revenue -----
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}8"], "=$I$25/$I$24", "0.00%")
    ws["C8"] = "=I25/I24 (FY25 COGS%)"
    ws["C8"].font = EQ_FONT

    # ----- SGA $ = Rev × (normalized FY25 SGA% − 50bps×year) -----
    # I28 = Salaries/SGA hist; strip one-time restructuring
    base = f"(($I$28-{RESTRUCTURING_NORMALIZE_000s})/$I$24)"
    bps = SGA_IMPROVEMENT_BPS / 10_000.0
    for i, col in enumerate(_FORECAST_COLS):
        grind = bps * (i + 1)
        _formula(
            ws[f"{col}9"],
            f"={col}24*MAX(0.05,{base}-{grind})",
            "#,##0.0",
        )
    ws["C9"] = f"=Rev×((I28−{RESTRUCTURING_NORMALIZE_000s:g})/I24 − {SGA_IMPROVEMENT_BPS:.0f}bps×t)"
    ws["C9"].font = EQ_FONT

    # ----- R&D $ = Rev × (FY25 R&D / Revenue)  [template "Rent and Overhead" row] -----
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}10"], f"={col}24*($I$29/$I$24)", "#,##0.0")
    ws["C10"] = "=Rev×(I29/I24) FY25 R&D%"
    ws["C10"].font = EQ_FONT

    # ----- D&A % of PPE = FY25 DA / FY25 PPE (I44); forecast opens at I44 -----
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}11"], '=IF($I$44=0,0.25,$I$30/$I$44)', "0.00%")
    ws["C11"] = "=I30/I44 (FY25 DA / PPE); J92 opens at I44"
    ws["C11"].font = EQ_FONT

    # ----- Interest % of opening debt -----
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}12"], '=IF($I$98=0,0.05,$I$101/$I$98)', "0.00%")
    ws["C12"] = "=I101/I98 (FY25 Int / Debt open)"
    ws["C12"].font = EQ_FONT

    # ----- Tax % = FY25 tax / EBT -----
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}13"], '=IF($I$33=0,0.21,$I$35/$I$33)', "0.00%")
    ws["C13"] = "=I35/I33 (FY25 effective tax)"
    ws["C13"].font = EQ_FONT

    # ----- Working capital days from FY25 BS -----
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}15"], "=ROUND($I$42/$I$24*365,0)", "0")
        _input(ws[f"{col}16"], 0, "0")  # inventory days — FICO has none
        _formula(ws[f"{col}17"], '=IF($I$25=0,0,ROUND($I$48/$I$25*365,0))', "0")
    ws["C15"] = "=ROUND(I42/I24×365,0) net AR days"
    ws["C15"].font = EQ_FONT
    ws["C17"] = "=ROUND(I48/I25×365,0) AP days"
    ws["C17"].font = EQ_FONT

    # ----- CapEx $ = Rev × faded (FY25 CapEx/Sales → steady 1%) -----
    # peak = I68/I24; path_t = peak×w + steady×(1−w)
    for i, (col, w) in enumerate(zip(_FORECAST_COLS, CAPEX_FADE_WEIGHTS)):
        # Excel: J24*(($I$68/$I$24)*w + 0.01*(1-w))
        _formula(
            ws[f"{col}18"],
            f"={col}24*(($I$68/$I$24)*{w}+{CAPEX_STEADY_PCT}*(1-{w}))",
            "#,##0.0",
        )
    ws["C18"] = f"=Rev×fade(I68/I24 → {CAPEX_STEADY_PCT:.0%})"
    ws["C18"].font = EQ_FONT

    # ----- Financing: hold flat (explicit 0) -----
    for col in _FORECAST_COLS:
        _input(ws[f"{col}19"], 0.0, "#,##0.0")
        _input(ws[f"{col}20"], 0.0, "#,##0.0")
    ws["C19"] = "POLICY: no new debt issuance"
    ws["C20"] = "POLICY: no equity issuance"
    for r in (19, 20):
        ws[f"C{r}"].font = Font(name="Calibri", italic=True, size=8, color="666666")

    # ----- IS rows: bake growth into opex DIRECTLY (not via hardcoded $) -----
    # Override J28:N28 / J29:N29 so Net Earnings = f(Rev×(1+g), margins…)
    for i, col in enumerate(_FORECAST_COLS):
        grind = bps * (i + 1)
        _formula(
            ws[f"{col}28"],
            f"={col}24*MAX(0.05,{base}-{grind})",
            "#,##0.0",
        )
        _formula(ws[f"{col}29"], f"={col}24*($I$29/$I$24)", "#,##0.0")
        # Keep assumption rows J9/J10 as mirrors (same equations) for the drivers block
        _formula(
            ws[f"{col}9"],
            f"={col}28",
            "#,##0.0",
        )
        _formula(ws[f"{col}10"], f"={col}29", "#,##0.0")

    ws["C28"] = f"=Rev×((I28−{RESTRUCTURING_NORMALIZE_000s:g})/I24−{SGA_IMPROVEMENT_BPS:.0f}bps×t)"
    ws["C29"] = "=Rev×(I29/I24)  ← grows with revenue"
    ws["C36"] = "=EBT−Tax  ← GROWS via Rev×(1+g) in row 24"
    ws["C62"] = "=IS Net Earnings (row 36) — not hardcoded"
    for r in (28, 29, 36, 62):
        ws[f"C{r}"].font = EQ_FONT

    # CF Net Earnings / DA must stay linked (force formulas in case an old inject overwrote)
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}62"], f"={col}36", "#,##0.0")
        _formula(ws[f"{col}63"], f"={col}30", "#,##0.0")
        _formula(ws[f"{col}65"], f"={col}62+{col}63-{col}64", "#,##0.0")

    # Column C width for equation notes
    ws.column_dimensions["C"].width = 48

    # Flag on growth row that it drives NI
    ws["C7"] = "POLICY → drives Rev row24 → GP → EBT → Net Earnings / CF"
    ws["C7"].font = Font(name="Calibri", italic=True, size=8, color="C00000")
