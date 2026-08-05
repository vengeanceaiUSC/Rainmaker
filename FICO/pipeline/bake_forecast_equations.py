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

    # ----- D&A % of opening PPE = FY25 DA / FY25 opening PPE (I92) -----
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}11"], '=IF($I$92=0,0.25,$I$30/$I$92)', "0.00%")
    ws["C11"] = "=I30/I92 (FY25 DA / PPE open)"
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

    # ----- Ensure IS still pulls SGA/RD from assumption rows (template default) -----
    # J28=J9, J29=J10 already in CFI template — leave them.
    # Document on the IS side:
    ws["C28"] = "← J9 = Rev×SGA% equation"
    ws["C29"] = "← J10 = Rev×R&D% equation"
    ws["C28"].font = EQ_FONT
    ws["C29"].font = EQ_FONT

    # Column C width for equation notes
    ws.column_dimensions["C"].width = 42
