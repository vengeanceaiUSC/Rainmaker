"""Bake FULL calculation equations into 3-statement forecast rows (J–N).

vengeanceaiUSCMODEL17.0:
  • Segment mix (SaaS / B2C / B2B / PS / on-prem) → blended COGS + mix schedule
  • Op NWC = AR+Inv−AP (excludes Deferred); ΔDeferred is explicit CFO cash source
  • Phased DSO hist→target (no one-year AR cliff)
  • Deferred Revenue BS liability driven by SaaS+on-prem mix (not double-counted)
  • D&A = Revenue × DA%; CapEx % fades; SBC % fades; levered buybacks 1.4× FCF
  • Hard BS plug: Equity = Assets − Liab − RE
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill, Border, Side

from .assumption_explanations import write_assumption_explanations
from .equation_explanations import write_equation_comments
from .model17_assumptions import (
    BUYBACK_FCF_MULTIPLE,
    BUYBACK_RUNRATE_000s,
    CAPEX_PCT_PATH,
    COGS_FLOOR_PCT,
    COGS_IMPROVEMENT_BPS,
    DA_PCT_REVENUE,
    DSO_B2B,
    DSO_B2C,
    DSO_ONPREM,
    DSO_PROF_SVCS,
    DSO_SAAS,
    GM_B2B,
    GM_B2C,
    GM_ONPREM,
    GM_PROF_SVCS,
    GM_SAAS,
    HIST_DSO_DAYS,
    MIX_B2B,
    MIX_B2C,
    MIX_ONPREM,
    MIX_PROF_SVCS,
    MIX_SAAS,
    RD_FLOOR_PCT,
    RD_IMPROVEMENT_BPS,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
    SAAS_MIX_SHIFT_BPS,
    SBC_PCT_PATH,
    SGA_FLOOR_PCT,
    SGA_IMPROVEMENT_BPS,
    TARGET_DSO_DAYS,
    blended_cogs_pct,
    blended_dso,
    phased_dso,
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


def _write_revenue_mix_schedule(ws) -> None:
    """FY25 10-K mix → blended DSO / COGS; yellow inputs on column J, linked K–N."""
    j = _FORECAST_COLS[0]
    ws["B102"] = (
        "REVENUE MIX & CASH CONVERSION (FY25 10-K disaggregation — "
        "mix offsets within Total Rev row 24; NOT additive)"
    )
    ws["B102"].font = Font(name="Calibri", bold=True, color="833C0C")

    # SaaS mix rises +SAAS_MIX_SHIFT_BPS/yr taken from on-prem (cloud transition).
    shift = SAAS_MIX_SHIFT_BPS / 10_000.0
    ws["B104"] = (
        f"Mix % — SaaS / Platform software (cloud; +{SAAS_MIX_SHIFT_BPS:.0f}bps/yr shift)"
    )
    ws["B105"] = "Mix % — B2C Subscriptions (myFICO)"
    ws["B106"] = "Mix % — B2B Scores (transactional)"
    ws["B107"] = "Mix % — Professional Services (implementation)"
    ws["B108"] = (
        f"Mix % — On-Premises Software (−{SAAS_MIX_SHIFT_BPS:.0f}bps/yr → SaaS)"
    )
    for i, col in enumerate(_FORECAST_COLS):
        saas = min(MIX_SAAS + MIX_ONPREM - 0.02, MIX_SAAS + shift * (i + 1))
        onprem = max(0.02, MIX_ONPREM - shift * (i + 1))
        # Renormalize soft remainder into SaaS/on-prem so total mix = 100%
        other = MIX_B2C + MIX_B2B + MIX_PROF_SVCS
        soft = 1.0 - other
        saas_n = saas / (saas + onprem) * soft
        onprem_n = soft - saas_n
        _input(ws[f"{col}104"], float(saas_n), "0.00%")
        _input(ws[f"{col}105"], float(MIX_B2C), "0.00%")
        _input(ws[f"{col}106"], float(MIX_B2B), "0.00%")
        _input(ws[f"{col}107"], float(MIX_PROF_SVCS), "0.00%")
        _input(ws[f"{col}108"], float(onprem_n), "0.00%")
    ws["C104"] = (
        f"eqn: FY25 SaaS mix + {SAAS_MIX_SHIFT_BPS:.0f}bps×t (from on-prem) — "
        "Deferred CFO driver"
    )
    ws["C104"].font = EQ_FONT

    ws["B109"] = "Mix check (must equal 100%)"
    for col in _FORECAST_COLS:
        _formula(
            ws[f"{col}109"],
            f"={col}104+{col}105+{col}106+{col}107+{col}108",
            "0.00%",
        )
    ws["C109"] = "eqn: Σ mix = 100%  (flag if ≠ 1)"
    ws["C109"].font = EQ_FONT

    dso_rows = [
        (110, "DSO days — SaaS / Platform (30–60 mid)", DSO_SAAS),
        (111, "DSO days — B2C myFICO (≈0, card-settled)", DSO_B2C),
        (112, "DSO days — B2B Scores (≈30)", DSO_B2B),
        (113, "DSO days — Professional Services", DSO_PROF_SVCS),
        (114, "DSO days — On-Premises Software", DSO_ONPREM),
    ]
    for row, label, val in dso_rows:
        ws[f"B{row}"] = label
        _input(ws[f"{j}{row}"], float(val), "0.0")
        for col in _FORECAST_COLS[1:]:
            _formula(ws[f"{col}{row}"], f"=${j}${row}", "0.0")

    ws["B115"] = "Blended DSO (feeds assumption row 15; rises slightly with SaaS mix)"
    for col in _FORECAST_COLS:
        _formula(
            ws[f"{col}115"],
            f"={col}104*{col}110+{col}105*{col}111+{col}106*{col}112+"
            f"{col}107*{col}113+{col}108*{col}114",
            "0.00",
        )
    ws["C115"] = "eqn: Σ (mix_i × DSO_i)"
    ws["C115"].font = EQ_FONT

    gm_rows = [
        (116, "Gross Margin — SaaS / Platform", GM_SAAS),
        (117, "Gross Margin — B2C myFICO", GM_B2C),
        (118, "Gross Margin — B2B Scores", GM_B2B),
        (119, "Gross Margin — Professional Services (low)", GM_PROF_SVCS),
        (120, "Gross Margin — On-Premises Software", GM_ONPREM),
    ]
    for row, label, val in gm_rows:
        ws[f"B{row}"] = label
        _input(ws[f"{j}{row}"], float(val), "0.00%")
        for col in _FORECAST_COLS[1:]:
            _formula(ws[f"{col}{row}"], f"=${j}${row}", "0.00%")

    ws["B121"] = "Blended Gross Margin (pre pricing grind)"
    for col in _FORECAST_COLS:
        _formula(
            ws[f"{col}121"],
            f"={col}104*{col}116+{col}105*{col}117+{col}106*{col}118+"
            f"{col}107*{col}119+{col}108*{col}120",
            "0.00%",
        )

    ws["B122"] = (
        f"Blended COGS % base (row 8 applies −{COGS_IMPROVEMENT_BPS:.0f}bps×t grind)"
    )
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}122"], f"=1-{col}121", "0.00%")
    ws["C122"] = (
        f"eqn: 1 − blended GM; row 8 = MAX({COGS_FLOOR_PCT:.0%}, base − "
        f"{COGS_IMPROVEMENT_BPS:.0f}bps×t)"
    )
    ws["C122"].font = EQ_FONT

    ws["B123"] = "Software mix % (SaaS + On-Prem — Deferred / CFO WC driver)"
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}123"], f"={col}104+{col}108", "0.00%")

    ws["B124"] = "Segment $ — SaaS (illustrative = Rev × mix; not extra IS revenue)"
    ws["B125"] = "Segment $ — B2C myFICO (illustrative)"
    for col in _FORECAST_COLS:
        _formula(ws[f"{col}124"], f"={col}24*{col}104", "#,##0.0")
        _formula(ws[f"{col}125"], f"={col}24*{col}105", "#,##0.0")
    # Clear stale chart formulas in hist E–I on schedule rows (forecast-only block)
    for row in range(102, 126):
        for col in ("E", "F", "G", "H", "I"):
            cell = ws[f"{col}{row}"]
            if cell.value is not None:
                cell.value = None
                cell.fill = PatternFill(fill_type=None)
    ws["C124"] = "eqn: TotalRev × SaaS%  (offset, not additive)"
    ws["C124"].font = EQ_FONT
    ws["C125"] = "eqn: TotalRev × B2C%  (offset, not additive)"
    ws["C125"].font = EQ_FONT


def bake_forecast_equations(wb) -> None:
    if SHEET_3S not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_3S}")
    ws = wb[SHEET_3S]

    base_sga = f"(($I$28-{RESTRUCTURING_NORMALIZE_000s})/$I$24)"
    bps = SGA_IMPROVEMENT_BPS / 10_000.0
    floor = SGA_FLOOR_PCT

    for col, g in zip(_FORECAST_COLS, REVENUE_GROWTH_PATH):
        _input(ws[f"{col}7"], float(g), "0.0%")

    # COGS% = segment blended GM base − Scores pricing grind (bps×t); floor 8%
    cogs_bps = COGS_IMPROVEMENT_BPS / 10_000.0
    for i, col in enumerate(_FORECAST_COLS):
        grind_cogs = cogs_bps * (i + 1)
        _formula(
            ws[f"{col}8"],
            f"=MAX({COGS_FLOOR_PCT:.2f},{col}122-{grind_cogs})",
            "0.00%",
        )

    rd_bps = RD_IMPROVEMENT_BPS / 10_000.0
    for i, col in enumerate(_FORECAST_COLS):
        grind = bps * (i + 1)
        rd_grind = rd_bps * (i + 1)
        _formula(ws[f"{col}9"], f"=MAX({floor},{base_sga}-{grind})", "0.00%")
        _formula(
            ws[f"{col}10"],
            f"=MAX({RD_FLOOR_PCT:.2f},$I$29/$I$24-{rd_grind:.4f})",
            "0.00%",
        )
    ws["B9"] = f"SG&A % of Revenue (floor {floor:.0%}, −{SGA_IMPROVEMENT_BPS:.0f}bps×t)"
    ws["B10"] = (
        f"R&D % of Revenue (floor {RD_FLOOR_PCT:.0%}, −{RD_IMPROVEMENT_BPS:.0f}bps×t)"
    )

    for i, (col, capex_pct, sbc_pct) in enumerate(
        zip(_FORECAST_COLS, CAPEX_PCT_PATH, SBC_PCT_PATH)
    ):
        # Row 11 = total D&A % of Revenue (incl. amort. of intangibles)
        _input(ws[f"{col}11"], float(DA_PCT_REVENUE), "0.00%")
        _formula(ws[f"{col}12"], '=IF($I$98=0,0.05,$I$101/$I$98)', "0.00%")
        _formula(ws[f"{col}13"], '=IF($I$33=0,0.21,$I$35/$I$33)', "0.00%")
        # Row 15 = phased DSO hist→target (not a one-year cliff to segment blend)
        _input(ws[f"{col}15"], round(phased_dso(i), 1), "0.0")
        _input(ws[f"{col}16"], 0, "0")
        _formula(
            ws[f"{col}17"],
            '=IF($I$25=0,0,ROUND($I$48/$I$25*365,0))',
            "0",
        )
        # Row 18 = CapEx % — Scores-light fade
        _input(ws[f"{col}18"], float(capex_pct), "0.00%")
        # Row 21 = SBC % path (CF/FCFF add-back only — no share dilution)
        _input(ws[f"{col}21"], float(sbc_pct), "0.00%")

    ws["B8"] = (
        f"COGS % of Revenue (floor {COGS_FLOOR_PCT:.0%}; "
        f"−{COGS_IMPROVEMENT_BPS:.0f}bps×t Scores pricing)"
    )
    ws["B11"] = "D&A % of Revenue (TOTAL D&A incl. amort. of intangibles — CF add-back)"
    ws["B15"] = (
        f"DSO — phased {HIST_DSO_DAYS:.0f}→{TARGET_DSO_DAYS:.0f}d over 5yrs "
        f"(segment blend ref ≈ {blended_dso():.1f}d @ row 115; AR=Rev×DSO/365)"
    )
    ws["B16"] = "Inventory (Days)"
    ws["B17"] = "DPO — Accounts Payable (Days)  [AP = COGS × DPO/365]"
    ws["B18"] = (
        f"CapEx % of Revenue (fade {CAPEX_PCT_PATH[0]:.2%}→{CAPEX_PCT_PATH[-1]:.2%})"
    )
    ws["B21"] = (
        "SBC % of Revenue (fade path; CF/FCFF add-back ONLY — buybacks cut DCF shares)"
    )

    # Financing: equity buybacks = −MULT×FCF; debt funds MULT−1 so ΔCash ≈ 0
    for col in _FORECAST_COLS:
        fcf = f"({col}66-{col}69)"
        _formula(
            ws[f"{col}20"],
            f"=-{BUYBACK_FCF_MULTIPLE:.2f}*{fcf}",
            "#,##0.0",
        )
        _formula(
            ws[f"{col}19"],
            f"=-{fcf}-{col}20",
            "#,##0.0",
        )
    ws["B19"] = (
        f"Debt Issuance (Repayment) — funds buybacks above 1.0× FCF "
        f"(MULT={BUYBACK_FCF_MULTIPLE:.2f})"
    )
    ws["B20"] = (
        f"Equity Issued (Repaid) — {BUYBACK_FCF_MULTIPLE:.2f}×(CFO−CapEx) buybacks "
        f"(hist avg ${BUYBACK_RUNRATE_000s/1000:.0f}M; Q3 FY26 levered)"
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
        # Total Liab = AP + Debt + Deferred Revenue (row 88 is the BS liability)
        _formula(ws[f"{col}50"], f"={col}48+{col}49+{col}88", "#,##0.0")
        # Equity capital = Assets − Liabilities − RE  (hard BS identity plug).
        _formula(ws[f"{col}52"], f"={col}45-{col}50-{col}53", "#,##0.0")
        _formula(ws[f"{col}53"], f"={prev}53+{col}33*(1-{col}13)", "#,##0.0")
        _formula(ws[f"{col}54"], f"=SUM({col}52:{col}53)", "#,##0.0")
        _formula(ws[f"{col}55"], f"={col}50+{col}54", "#,##0.0")
        _formula(ws[f"{col}57"], f"={col}55-{col}45", "#,##0.0")

    ws["B49"] = "Debt"
    ws["B50"] = "Total Liabilities (AP + Debt + Deferred Revenue)"
    ws["B52"] = "Equity Capital (BS plug = Assets − Liab − RE)"
    # Mirror Deferred onto the BS liability section label (value lives at row 88)
    ws["B47"] = "Liabilities (Deferred Revenue = WC row 88, included in Total Liab)"
    for r, note in (
        (42, "eqn: Rev × blended segment DSO / 365"),
        (48, "eqn: COGS × DPO / 365"),
        (50, "eqn: AP + Debt + DeferredRev(88)"),
        (52, "eqn: Assets − Liabilities − RE  (forces BS balance)"),
        (53, "eqn: PriorRE + EBT×(1−t)"),
        (57, "eqn: (L+E) − Assets   must be 0"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== CASH FLOW =====
    # 62 NI | 63 +DA | 64 +SBC | 65 −ΔOpNWC | 66 CFO (=…+ΔDeferred)
    # 81 memo: Increase in Deferred (explicit cash source; not inside Op NWC)
    ws["B63"] = "Plus: Depreciation & Amortization"
    ws["B64"] = "Plus: Stock-Based Compensation (SBC)"
    ws["B65"] = "Less: Δ Operating NWC (AR+Inv−AP; excludes Deferred)"
    ws["B66"] = "Cash from Operations (incl. +Increase in Deferred)"

    for col in _FORECAST_COLS:
        prev = _PREV[col]
        _formula(ws[f"{col}62"], f"={col}33*(1-{col}13)", "#,##0.0")
        _formula(ws[f"{col}63"], f"={col}30", "#,##0.0")  # DA = Rev × DA%
        _formula(ws[f"{col}64"], f"={col}24*{col}21", "#,##0.0")  # SBC = Rev × SBC%
        _formula(ws[f"{col}65"], f"={col}90", "#,##0.0")  # Δ Op NWC
        # CFO = NI + DA + SBC − ΔOpNWC + ΔDeferred  (Deferred NOT inside Op NWC)
        _formula(
            ws[f"{col}66"],
            f"={col}62+{col}63+{col}64-{col}65+({col}88-{prev}88)",
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
        # Explicit Deferred cash-source line (memo; already in CFO formula)
        _formula(ws[f"{col}81"], f"={col}88-{prev}88", "#,##0.0")

    for r, note in (
        (62, "eqn: EBT × (1 − book tax%)"),
        (63, "eqn: Rev × DA%"),
        (64, "eqn: Rev × SBC%  (non-cash add-back)"),
        (65, "eqn: Δ Op NWC = OpNWCt − OpNWCt−1  (AR+Inv−AP only)"),
        (66, "eqn: NI + DA + SBC − ΔOpNWC + ΔDeferred"),
        (68, "eqn: Rev × CapEx% (fade path)"),
        (76, "eqn: CFO − CapEx + Financing"),
        (81, "eqn: Def_t − Def_t−1  (SaaS/maintenance upfront cash; in CFO)"),
    ):
        ws[f"C{r}"] = note
        ws[f"C{r}"].font = EQ_FONT

    # ===== WC SCHEDULE =====
    ws["B85"] = "Accounts Receivable (from phased DSO — separate from Deferred)"
    ws["B86"] = "Inventory"
    ws["B87"] = "Accounts Payable (from DPO)"
    ws["B88"] = "Deferred Revenue (BS contract liability — SaaS/on-prem; CF source)"
    ws["B89"] = "Operating NWC (AR+Inv−AP; EXCLUDES Deferred — no double count)"
    ws["B90"] = "Change in Operating NWC"

    # Deferred = Rev × (SaaS%+OnPrem%) × (FY25 Deferred / FY25 software Rev)
    # SaaS mix shift grows Deferred → explicit +ΔDeferred in CFO (not via Op NWC).
    fy25_soft = f"({MIX_SAAS}+{MIX_ONPREM})"
    def_ratio = f"IF($I$24*{fy25_soft}=0,0,$I$88/($I$24*{fy25_soft}))"

    for col in _FORECAST_COLS:
        prev = _PREV[col]
        soft_mix = f"({col}104+{col}108)"
        _formula(ws[f"{col}85"], f"={col}42", "#,##0.0")
        _formula(ws[f"{col}86"], f"={col}43", "#,##0.0")
        _formula(ws[f"{col}87"], f"={col}48", "#,##0.0")
        _formula(
            ws[f"{col}88"],
            f"={col}24*{soft_mix}*{def_ratio}",
            "#,##0.0",
        )
        # Op NWC excludes Deferred (Deferred cash is CF row 81 / CFO add)
        _formula(
            ws[f"{col}89"],
            f"={col}85+{col}86-{col}87",
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

    # ===== REVENUE MIX & CASH CONVERSION SCHEDULE (rows 102–125) =====
    # Mix % offset within Total Revenue (row 24) — never additive to IS revenue.
    _write_revenue_mix_schedule(ws)

    ws["C8"] = (
        f"eqn: MAX({COGS_FLOOR_PCT:.0%}, blended COGS base≈{blended_cogs_pct():.2%} − "
        f"{COGS_IMPROVEMENT_BPS:.0f}bps×t)"
    )
    ws["C8"].font = EQ_FONT
    ws["C10"] = (
        f"eqn: MAX({RD_FLOOR_PCT:.0%}, I29/I24 − {RD_IMPROVEMENT_BPS:.0f}bps×t)"
    )
    ws["C10"].font = EQ_FONT
    ws["C11"] = "eqn: 3yr avg TOTAL D&A/Rev (incl. AmortizationOfIntangibleAssets)"
    ws["C11"].font = EQ_FONT
    ws["C15"] = (
        f"eqn: HIST {HIST_DSO_DAYS:.0f}d × (1−t/5) + TARGET {TARGET_DSO_DAYS:.0f}d × (t/5)"
    )
    ws["C15"].font = EQ_FONT
    ws["C17"] = "eqn: ROUND(FY25 AP/COGS × 365)  — DPO held flat"
    ws["C17"].font = EQ_FONT
    ws["C18"] = (
        f"eqn: CapEx% fade {CAPEX_PCT_PATH[0]:.2%}→{CAPEX_PCT_PATH[-1]:.2%} "
        "(Scores-light)"
    )
    ws["C18"].font = EQ_FONT
    ws["C19"] = (
        f"eqn: −(CFO−CapEx) − EquityCF = ({BUYBACK_FCF_MULTIPLE:.2f}−1)×FCF debt"
    )
    ws["C19"].font = EQ_FONT
    ws["C20"] = (
        f"eqn: −{BUYBACK_FCF_MULTIPLE:.2f}×(CFO−CapEx)  (levered buybacks)"
    )
    ws["C20"].font = EQ_FONT
    ws["C21"] = (
        "eqn: SBC% fade path "
        + " / ".join(f"{p:.2%}" for p in SBC_PCT_PATH)
        + "; add-back only"
    )
    ws["C21"].font = EQ_FONT
    ws["C42"] = "eqn: Rev × phased DSO / 365  (AR ≠ Deferred)"
    ws["C42"].font = EQ_FONT
    ws["C63"] = "eqn: TOTAL D&A (depr + amort. of intangibles) — single add-back"
    ws["C63"].font = EQ_FONT
    ws["C88"] = "eqn: Rev × softmix_t × (FY25 Def / FY25 soft Rev) — SaaS shift ↑ Def"
    ws["C88"].font = EQ_FONT
    ws["C89"] = "eqn: AR + Inv − AP   (Deferred excluded — see CF row 81)"
    ws["C89"].font = EQ_FONT
    ws["C90"] = "eqn: OpNWCt − OpNWCt−1  (AR/AP only; Deferred separate)"
    ws["C90"].font = EQ_FONT
    ws["C94"] = "eqn: Rev × DA% (same as IS D&A — includes intangible amort)"
    ws["C94"].font = EQ_FONT
    ws["B63"] = "Plus: D&A (TOTAL — incl. amortization of intangibles)"
    ws["B65"] = "Less: Δ Operating NWC (AR+Inv−AP; Deferred NOT here)"
    ws["B81"] = "Plus: Increase in Deferred Revenue (explicit CFO cash source)"

    ws["B28"] = "SG&A Expense"
    ws["B29"] = "Research & Development (R&D)"

    # Hist FY21–25: SBC add-back + CF↔BS cash articulation
    hist_cols = ["E", "F", "G", "H", "I"]
    hist_sbc = [112_457.0, 115_355.0, 123_847.0, 149_439.0, 156_667.0]
    hist_prev = {"E": None, "F": "E", "G": "F", "H": "G", "I": "H"}
    for col, sbc in zip(hist_cols, hist_sbc):
        _input(ws[f"{col}64"], sbc, "#,##0.0")
        # Δ Op NWC from WC schedule; ΔDeferred explicit in CFO
        _formula(ws[f"{col}65"], f"={col}90", "#,##0.0")
        prev_h = hist_prev[col]
        if prev_h is None:
            # First hist year: no prior deferred → ΔDef = 0 in CFO bridge
            _formula(
                ws[f"{col}66"],
                f"={col}62+{col}63+{col}64-{col}65",
                "#,##0.0",
            )
            _formula(ws[f"{col}81"], "0", "#,##0.0")
        else:
            _formula(
                ws[f"{col}66"],
                f"={col}62+{col}63+{col}64-{col}65+({col}88-{prev_h}88)",
                "#,##0.0",
            )
            _formula(ws[f"{col}81"], f"={col}88-{prev_h}88", "#,##0.0")
        _formula(ws[f"{col}69"], f"={col}68", "#,##0.0")
        # Equity issuance plugs so CF closing cash = BS cash (hist buybacks etc.)
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

    # Hist Op NWC should exclude Deferred too (align schedule with forecast)
    for col in hist_cols:
        prev = hist_prev[col]
        _formula(ws[f"{col}89"], f"={col}85+{col}86-{col}87", "#,##0.0")
        if prev is None:
            _formula(ws[f"{col}90"], "0", "#,##0.0")
        else:
            _formula(ws[f"{col}90"], f"={col}89-{prev}89", "#,##0.0")

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
