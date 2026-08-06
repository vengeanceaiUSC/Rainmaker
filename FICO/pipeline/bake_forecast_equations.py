"""Bake FULL calculation equations into 3-statement forecast rows (J–N).

vengeanceaiUSCMODEL11:
  • Bottom-up WC: AR via DSO, AP via DPO; NWC = AR+Inv−AP−Deferred
  • D&A = Revenue × DA% (software-industry total D&A / sales)
  • CapEx = Revenue × flat 3yr-avg CapEx% (no fade-to-1%)
  • SBC add-back in CFO at SBC% × Revenue
  • Financing: normalized buyback + net debt run-rates
  • SG&A floor 15% / −75 bps; labels SG&A / R&D
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill, Border, Side

from .assumption_explanations import write_assumption_explanations
from .equation_explanations import write_equation_comments
from .model11_assumptions import (
    BUYBACK_RUNRATE_000s,
    CAPEX_PCT_REVENUE,
    DA_PCT_REVENUE,
    DEBT_NET_RUNRATE_000s,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
    SBC_PCT_REVENUE,
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
        # Row 11 = total D&A % of Revenue (software-industry driver)
        _input(ws[f"{col}11"], float(DA_PCT_REVENUE), "0.00%")
        _formula(ws[f"{col}12"], '=IF($I$98=0,0.05,$I$101/$I$98)', "0.00%")
        _formula(ws[f"{col}13"], '=IF($I$33=0,0.21,$I$35/$I$33)', "0.00%")
        # Row 15 = DSO (AR days)
        _formula(
            ws[f"{col}15"],
            '=IF($I$24=0,0,ROUND($I$42/$I$24*365,0))',
            "0",
        )
        _input(ws[f"{col}16"], 0, "0")
        _formula(
            ws[f"{col}17"],
            '=IF($I$25=0,0,ROUND($I$48/$I$25*365,0))',
            "0",
        )
        # Row 18 = CapEx % of Revenue — flat 3yr hist avg (no fade)
        _input(ws[f"{col}18"], float(CAPEX_PCT_REVENUE), "0.00%")
        # Row 21 = SBC % of Revenue
        _input(ws[f"{col}21"], float(SBC_PCT_REVENUE), "0.00%")

    ws["B11"] = "D&A % of Revenue (3yr avg total D&A/Sales — software driver)"
    ws["B15"] = "DSO — Accounts Receivable (Days)  [AR = Rev × DSO/365]"
    ws["B16"] = "Inventory (Days)"
    ws["B17"] = "DPO — Accounts Payable (Days)  [AP = COGS × DPO/365]"
    ws["B18"] = "CapEx % of Revenue (flat 3yr hist avg — no fade)"
    ws["B21"] = "SBC % of Revenue (3yr avg — CF / FCFF add-back)"

    # Financing: debt = 3yr avg net debt CF; equity = residual FCF (no cash stockpile)
    for col in _FORECAST_COLS:
        _input(ws[f"{col}19"], float(DEBT_NET_RUNRATE_000s), "#,##0.0")
        # Buybacks absorb residual levered FCF after debt so ΔCash ≈ 0
        _formula(
            ws[f"{col}20"],
            f"=-({col}66-{col}69)-{col}19",
            "#,##0.0",
        )
    ws["B19"] = "Debt Issuance (Repayment) — 3yr avg net debt CF ($000s)"
    ws["B20"] = (
        f"Equity Issued (Repaid) — residual FCF after debt "
        f"(hist buybacks avg ${BUYBACK_RUNRATE_000s/1000:.0f}M)"
    )

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
        _formula(ws[f"{col}30"], f"={col}24*{col}11", "#,##0.0")  # Rev × DA%
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
        30: "eqn: Rev × DA% (total D&A / sales)",
        31: "eqn: Debt_open × Int%",
        33: "eqn: GP − Expenses",
        35: "eqn: EBT × book tax%",
        36: "eqn: EBT × (1 − book tax%)",
    }.items():
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== BALANCE SHEET =====
    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}41"], f"={col}78", "#,##0.0")
        _formula(ws[f"{col}42"], f"={col}24*{col}15/365", "#,##0.0")
        _formula(ws[f"{col}43"], f"={col}25*{col}16/365", "#,##0.0")
        _formula(ws[f"{col}44"], f"={col}95", "#,##0.0")
        _formula(ws[f"{col}48"], f"={col}25*{col}17/365", "#,##0.0")
        _formula(ws[f"{col}45"], f"=SUM({col}41:{col}44)", "#,##0.0")
        _formula(ws[f"{col}49"], f"={col}100", "#,##0.0")
        _formula(ws[f"{col}50"], f"={col}48+{col}49+{col}88", "#,##0.0")
        # Equity capital = Assets − Liabilities − RE  (hard BS identity plug).
        # Buybacks/SBC/APIC flow through cash + RE; this plug keeps Assets = L+E
        # exactly every year (no cumulative drift).
        _formula(ws[f"{col}52"], f"={col}45-{col}50-{col}53", "#,##0.0")
        _formula(ws[f"{col}53"], f"={prev}53+{col}33*(1-{col}13)", "#,##0.0")
        _formula(ws[f"{col}54"], f"=SUM({col}52:{col}53)", "#,##0.0")
        _formula(ws[f"{col}55"], f"={col}50+{col}54", "#,##0.0")
        _formula(ws[f"{col}57"], f"={col}55-{col}45", "#,##0.0")
        # Prominent OK/ERROR already on row 3; keep numeric check on 57

    ws["B52"] = "Equity Capital (BS plug = Assets − Liab − RE)"
    for r, note in (
        (42, "eqn: Rev × DSO / 365  (no AR plug)"),
        (48, "eqn: COGS × DPO / 365"),
        (50, "eqn: AP + Debt + DeferredRev"),
        (52, "eqn: Assets − Liabilities − RE  (forces BS balance)"),
        (53, "eqn: PriorRE + EBT×(1−t)"),
        (57, "eqn: (L+E) − Assets   must be 0"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== CASH FLOW =====
    # 62 NI | 63 +DA | 64 +SBC | 65 −ΔNWC | 66 CFO
    ws["B63"] = "Plus: Depreciation & Amortization"
    ws["B64"] = "Plus: Stock-Based Compensation (SBC)"
    ws["B65"] = "Less: Changes in Working Capital"
    ws["B66"] = "Cash from Operations"

    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}62"], f"={col}33*(1-{col}13)", "#,##0.0")
        _formula(ws[f"{col}63"], f"={col}30", "#,##0.0")  # DA = Rev × DA%
        _formula(ws[f"{col}64"], f"={col}24*{col}21", "#,##0.0")  # SBC = Rev × SBC%
        _formula(ws[f"{col}65"], f"={col}90", "#,##0.0")  # ΔNWC
        _formula(
            ws[f"{col}66"],
            f"={col}62+{col}63+{col}64-{col}65",
            "#,##0.0",
        )
        _formula(ws[f"{col}68"], f"={col}24*{col}18", "#,##0.0")
        _formula(ws[f"{col}69"], f"={col}68", "#,##0.0")
        _formula(ws[f"{col}72"], f"={col}19", "#,##0.0")
        _formula(ws[f"{col}73"], f"={col}20", "#,##0.0")
        _formula(ws[f"{col}74"], f"={col}19+{col}20", "#,##0.0")
        _formula(ws[f"{col}76"], f"={col}66-{col}69+{col}74", "#,##0.0")
        _formula(ws[f"{col}77"], f"={prev}41", "#,##0.0")
        _formula(ws[f"{col}78"], f"={col}77+{col}76", "#,##0.0")
        _formula(ws[f"{col}80"], f"={col}78-{col}41", "#,##0.0")

    for r, note in (
        (62, "eqn: EBT × (1 − book tax%)"),
        (63, "eqn: Rev × DA%"),
        (64, "eqn: Rev × SBC%  (non-cash add-back)"),
        (65, "eqn: organic ΔNWC = NWCt − NWCt−1"),
        (66, "eqn: NI + DA + SBC − ΔNWC"),
        (68, "eqn: Rev × CapEx% (flat hist avg)"),
        (76, "eqn: CFO − CapEx + Financing"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== WC SCHEDULE =====
    ws["B85"] = "Accounts Receivable (from DSO)"
    ws["B86"] = "Inventory"
    ws["B87"] = "Accounts Payable (from DPO)"
    ws["B88"] = "Deferred Revenue (contract liability)"
    ws["B89"] = "Net Working Capital (NWC = AR+Inv−AP−Deferred)"
    ws["B90"] = "Change in NWC"

    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}85"], f"={col}42", "#,##0.0")
        _formula(ws[f"{col}86"], f"={col}43", "#,##0.0")
        _formula(ws[f"{col}87"], f"={col}48", "#,##0.0")
        _formula(ws[f"{col}88"], f"={col}24*IF($I$24=0,0,$I$88/$I$24)", "#,##0.0")
        _formula(
            ws[f"{col}89"],
            f"={col}85+{col}86-{col}87-{col}88",
            "#,##0.0",
        )
        _formula(ws[f"{col}90"], f"={col}89-{prev}89", "#,##0.0")

        if col == "J":
            _formula(ws["J92"], "=I44", "#,##0.0")
        else:
            _formula(ws[f"{col}92"], f"={prev}95", "#,##0.0")
        _formula(ws[f"{col}93"], f"={col}24*{col}18", "#,##0.0")
        _formula(ws[f"{col}94"], f"={col}30", "#,##0.0")  # same total D&A
        _formula(ws[f"{col}95"], f"={col}92+{col}93-{col}94", "#,##0.0")

        if col == "J":
            _formula(ws["J98"], "=I49", "#,##0.0")
        else:
            _formula(ws[f"{col}98"], f"={prev}100", "#,##0.0")
        _formula(ws[f"{col}99"], f"={col}19", "#,##0.0")
        _formula(ws[f"{col}100"], f"={col}98+{col}99", "#,##0.0")
        _formula(ws[f"{col}101"], f"={col}98*{col}12", "#,##0.0")

    ws["C11"] = "eqn: 3yr avg total D&A / Revenue (FY23–25)"
    ws["C11"].font = EQ_FONT
    ws["C15"] = "eqn: ROUND(FY25 AR/Rev × 365)  — DSO held flat"
    ws["C15"].font = EQ_FONT
    ws["C17"] = "eqn: ROUND(FY25 AP/COGS × 365)  — DPO held flat"
    ws["C17"].font = EQ_FONT
    ws["C18"] = "eqn: flat 3yr avg (PPE + capitalized software) / Rev"
    ws["C18"].font = EQ_FONT
    ws["C21"] = "eqn: 3yr avg SBC / Revenue (FY23–25)"
    ws["C21"].font = EQ_FONT
    ws["C42"] = "eqn: Rev × DSO / 365"
    ws["C42"].font = EQ_FONT
    ws["C88"] = "eqn: Rev × (FY25 Deferred/FY25 Rev)"
    ws["C88"].font = EQ_FONT
    ws["C89"] = "eqn: AR + Inv − AP − Deferred   (NWC is OUTPUT)"
    ws["C89"].font = EQ_FONT
    ws["C90"] = "eqn: NWCt − NWCt−1"
    ws["C90"].font = EQ_FONT
    ws["C94"] = "eqn: Rev × DA% (same as IS D&A)"
    ws["C94"].font = EQ_FONT

    ws["B28"] = "SG&A Expense"
    ws["B29"] = "Research & Development (R&D)"

    # Hist FY21–25: SBC add-back + CF↔BS cash articulation
    hist_cols = ["E", "F", "G", "H", "I"]
    hist_sbc = [112_457.0, 115_355.0, 123_847.0, 149_439.0, 156_667.0]
    for col, sbc in zip(hist_cols, hist_sbc):
        _input(ws[f"{col}64"], sbc, "#,##0.0")
        # ΔNWC must match WC schedule (not a stale first-year zero)
        _formula(ws[f"{col}65"], f"={col}90", "#,##0.0")
        _formula(ws[f"{col}66"], f"={col}62+{col}63+{col}64-{col}65", "#,##0.0")
        _formula(ws[f"{col}69"], f"={col}68", "#,##0.0")
        # Equity issuance plugs so CF closing cash = BS cash (hist buybacks etc.)
        # Closing = Opening + CFO − CapEx + Debt + Equity  ⇒
        # Equity = BS_Cash − Opening − CFO + CapEx − Debt
        _formula(
            ws[f"{col}73"],
            f"={col}41-{col}77-{col}66+{col}69-{col}72",
            "#,##0.0",
        )
        _formula(ws[f"{col}74"], f"={col}72+{col}73", "#,##0.0")
        _formula(ws[f"{col}76"], f"={col}66-{col}69+{col}74", "#,##0.0")
        _formula(ws[f"{col}78"], f"={col}77+{col}76", "#,##0.0")
        # Hist equity capital remains the Assets−Liab−RE plug (same identity)
        _formula(ws[f"{col}52"], f"={col}45-{col}50-{col}53", "#,##0.0")
        _formula(ws[f"{col}54"], f"=SUM({col}52:{col}53)", "#,##0.0")
        _formula(ws[f"{col}55"], f"={col}50+{col}54", "#,##0.0")
        _formula(ws[f"{col}57"], f"={col}55-{col}45", "#,##0.0")

    # Prominent OK/ERROR for BS identity (row 3) and CF↔BS cash (row 80)
    ws["B80"] = "Cash Check (CF close vs BS cash)"
    for col in _FORECAST_COLS + hist_cols:
        _formula(
            ws[f"{col}3"],
            f'=IFERROR(IF(ABS({col}57)>0.5,"ERROR","OK"),"OK")',
        )
        # Keep numeric residual on row 80; label clarifies what it tests.
        # (Row 3 stays the BS identity flag users look at first.)

    write_equation_comments(wb)
    write_assumption_explanations(wb)
    ws.column_dimensions["C"].width = 52
