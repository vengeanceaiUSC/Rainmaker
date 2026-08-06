"""
Wire the DCF sheet to the 3-statement sheet with live Excel formulas.

vengeanceaiUSCMODEL8:
1. Drivers linked to 3-statement
2. Base exit = 17.5x (public peer median ~18.1x); Bull 25x / Bear 12.8x
3. Mid-year dates via EDATE from transaction date (0-based periods)
4. 5×WACC × Exit-multiple sensitivity matrix (EV and $/share)
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .model8_assumptions import (
    EXIT_EV_EBITDA,
    EXIT_EV_EBITDA_BEAR,
    EXIT_EV_EBITDA_BULL,
    MODEL_NAME,
    PEER_EV_EBITDA,
    PEER_MEDIAN_EV_EBITDA,
    URL_FICO_EV_EBITDA,
    URL_PEER_COMPS,
)
from .named_range_map import SHEET_3S, SHEET_DCF

_DCF_COLS = ["E", "F", "G", "H", "I"]
_S3_FORECAST_COLS = ["J", "K", "L", "M", "N"]

LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
SUB_FILL = PatternFill("solid", fgColor="D6EAF8")
BLACK = Font(name="Calibri", color="000000")
WHITE_BOLD = Font(name="Calibri", bold=True, color="FFFFFF")
BOLD = Font(name="Calibri", bold=True, color="1F4E79")
LINK_FONT = Font(name="Calibri", size=8, color="0563C1", underline="single")
NOTE_FONT = Font(name="Calibri", italic=True, size=8, color="595959")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

# Sensitivity axes (checklist)
_WACC_AXIS = (0.0824, 0.0874, 0.0924, 0.0974, 0.1024)
_EXIT_AXIS = (12.5, 15.0, 17.5, 20.0, 22.5, 25.0)  # 12.5→25.0 step 2.5x


def _link(cell, formula: str) -> None:
    cell.value = formula
    cell.font = BLACK
    cell.fill = LINK_FILL


def _write_peer_comps(dcf) -> None:
    """On-sheet public comps table used to justify the 17.5x base exit."""
    dcf["L1"] = f"{MODEL_NAME} — Public Peer EV/EBITDA Comps"
    dcf["L1"].font = WHITE_BOLD
    dcf["L1"].fill = HDR_FILL
    dcf.merge_cells("L1:O1")

    dcf["L2"] = "Ticker"
    dcf["M2"] = "Company"
    dcf["N2"] = "EV/EBITDA"
    dcf["O2"] = "Source"
    for col in "LMNO":
        dcf[f"{col}2"].font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
        dcf[f"{col}2"].fill = PatternFill("solid", fgColor="833C0C")
        dcf[f"{col}2"].border = THIN

    for i, (ticker, name, mult) in enumerate(PEER_EV_EBITDA):
        r = 3 + i
        dcf[f"L{r}"] = ticker
        dcf[f"M{r}"] = name
        dcf[f"N{r}"] = mult
        dcf[f"N{r}"].number_format = '0.00"x"'
        dcf[f"O{r}"] = "VCP Scanner peer set"
        for col in "LMNO":
            dcf[f"{col}{r}"].border = THIN
            dcf[f"{col}{r}"].font = Font(name="Calibri", size=9)

    r = 3 + len(PEER_EV_EBITDA)
    dcf[f"L{r}"] = "Median"
    dcf[f"M{r}"] = "Peer set"
    dcf[f"N{r}"] = PEER_MEDIAN_EV_EBITDA
    dcf[f"N{r}"].number_format = '0.00"x"'
    dcf[f"O{r}"] = "Sorted median (SPGI)"
    for col in "LMNO":
        dcf[f"{col}{r}"].font = Font(name="Calibri", bold=True, size=9)
        dcf[f"{col}{r}"].fill = SUB_FILL
        dcf[f"{col}{r}"].border = THIN

    r += 1
    dcf[f"L{r}"] = "Base exit (D8)"
    dcf[f"N{r}"] = EXIT_EV_EBITDA
    dcf[f"N{r}"].number_format = '0.0"x"'
    dcf[f"O{r}"] = "Near peer median; audit 17.5x"
    for col in "LMNO":
        dcf[f"{col}{r}"].fill = INPUT_FILL
        dcf[f"{col}{r}"].border = THIN

    r += 1
    dcf[f"L{r}"] = "Peer comps (click)"
    dcf[f"L{r}"].hyperlink = URL_PEER_COMPS
    dcf[f"L{r}"].font = LINK_FONT
    dcf[f"M{r}"] = "FICO spot EV/EBITDA (click)"
    dcf[f"M{r}"].hyperlink = URL_FICO_EV_EBITDA
    dcf[f"M{r}"].font = LINK_FONT


def _write_scenarios(dcf) -> None:
    dcf["L17"] = f"Terminal value & scenarios ({MODEL_NAME})"
    dcf["L17"].font = BOLD
    dcf.merge_cells("L17:N17")

    dcf["L18"] = f"Exit EV/EBITDA @ D8 (BASE {EXIT_EV_EBITDA:.1f}x)"
    _link(dcf["M18"], "=J27")
    dcf["N18"] = "← used in XNPV"

    dcf["L19"] = "Gordon cross-check (g=D7)"
    _link(
        dcf["M19"],
        '=IF(($D$6-$D$7)<=0,"WACC must exceed g",$I$26*(1+$D$7)/($D$6-$D$7))',
    )
    dcf["N19"] = "← sanity check"

    dcf["L20"] = "Implied Gordon exit multiple"
    _link(dcf["M20"], "=IF(($I$21+$I$23)=0,0,M19/($I$21+$I$23))")
    dcf["N20"] = "← ~bear check"

    dcf["L21"] = f"BASE TV @ {EXIT_EV_EBITDA:.1f}x (primary)"
    _link(dcf["M21"], f"=($I$21+$I$23)*{EXIT_EV_EBITDA}")
    dcf["N21"] = "← = D8 policy"

    dcf["L22"] = f"BULL TV @ {EXIT_EV_EBITDA_BULL:.1f}x (spot/policy)"
    _link(dcf["M22"], f"=($I$21+$I$23)*{EXIT_EV_EBITDA_BULL}")
    dcf["N22"] = "← sensitivity only"

    dcf["L23"] = f"BEAR TV @ {EXIT_EV_EBITDA_BEAR:.1f}x (Gordon-implied)"
    _link(dcf["M23"], f"=($I$21+$I$23)*{EXIT_EV_EBITDA_BEAR}")
    dcf["N23"] = "← sensitivity only"

    dcf["L24"] = (
        f"Reconciliation: Bull {EXIT_EV_EBITDA_BULL:.0f}x → Base {EXIT_EV_EBITDA:.1f}x "
        f"(peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x) → Bear {EXIT_EV_EBITDA_BEAR:.1f}x"
    )
    dcf["L24"].font = NOTE_FONT
    dcf.merge_cells("L24:O24")

    dcf.column_dimensions["L"].width = 44
    dcf.column_dimensions["M"].width = 18
    dcf.column_dimensions["N"].width = 18
    dcf.column_dimensions["O"].width = 22


def _sens_ev_formula(wacc_addr: str, exit_addr: str) -> str:
    """EV at alternate WACC/exit without mutating D6/D8.

    Explicit FCFF (D28:I28) discounted at scenario WACC via XNPV; terminal
    cash (Y5 FCFF + Exit×EBITDA) discounted with XNPV's 365-day convention.
    """
    return (
        f"=XNPV({wacc_addr},$D$28:$I$28,$D$18:$I$18)"
        f"+($I$26+($I$21+$I$23)*{exit_addr})"
        f"/((1+{wacc_addr})^(($J$18-$D$18)/365))"
    )


def _write_sensitivity(dcf) -> None:
    """Two-way sensitivity: rows = WACC, cols = Exit EV/EBITDA → EV and $/share.

    Placed at row 55+ so it does not collide with DCF!Q:V CAPM/source block (rows 2–51).
    """
    # --- EV table ---
    start = 55
    dcf[f"L{start}"] = "SENSITIVITY — Enterprise Value ($000s)"
    dcf[f"L{start}"].font = WHITE_BOLD
    dcf[f"L{start}"].fill = HDR_FILL
    dcf.merge_cells(f"L{start}:R{start}")

    dcf[f"L{start + 1}"] = "WACC \\ Exit EV/EBITDA"
    dcf[f"L{start + 1}"].font = Font(name="Calibri", bold=True, size=9)
    dcf[f"L{start + 1}"].fill = SUB_FILL
    dcf[f"L{start + 1}"].border = THIN

    exit_row = start + 1
    exit_cols = ["M", "N", "O", "P", "Q", "R"]
    for col, mx in zip(exit_cols, _EXIT_AXIS):
        cell = dcf[f"{col}{exit_row}"]
        cell.value = mx
        cell.number_format = '0.0"x"'
        cell.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="833C0C")
        cell.border = THIN
        cell.alignment = Alignment(horizontal="center")

    for i, w in enumerate(_WACC_AXIS):
        r = start + 2 + i
        wcell = dcf[f"L{r}"]
        wcell.value = w
        wcell.number_format = "0.00%"
        wcell.font = Font(name="Calibri", bold=True, size=9)
        wcell.fill = SUB_FILL
        wcell.border = THIN
        for col, mx in zip(exit_cols, _EXIT_AXIS):
            # Absolute refs to header WACC (col L) and exit (exit_row)
            formula = _sens_ev_formula(f"$L{r}", f"{col}${exit_row}")
            cell = dcf[f"{col}{r}"]
            cell.value = formula
            cell.number_format = "#,##0.0"
            cell.font = BLACK
            cell.fill = LINK_FILL
            cell.border = THIN
            # Highlight base case cell (WACC≈9.24%, Exit=17.5x)
            if abs(w - 0.0924) < 1e-9 and abs(mx - EXIT_EV_EBITDA) < 1e-9:
                cell.fill = INPUT_FILL

    # --- $/share table ---
    sh = start + 9
    dcf[f"L{sh}"] = "SENSITIVITY — Equity Value / Share ($)"
    dcf[f"L{sh}"].font = WHITE_BOLD
    dcf[f"L{sh}"].fill = HDR_FILL
    dcf.merge_cells(f"L{sh}:R{sh}")

    dcf[f"L{sh + 1}"] = "WACC \\ Exit EV/EBITDA"
    dcf[f"L{sh + 1}"].font = Font(name="Calibri", bold=True, size=9)
    dcf[f"L{sh + 1}"].fill = SUB_FILL
    dcf[f"L{sh + 1}"].border = THIN

    for col, mx in zip(exit_cols, _EXIT_AXIS):
        cell = dcf[f"{col}{sh + 1}"]
        cell.value = mx
        cell.number_format = '0.0"x"'
        cell.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="833C0C")
        cell.border = THIN

    for i, w in enumerate(_WACC_AXIS):
        r = sh + 2 + i
        ev_r = start + 2 + i
        wcell = dcf[f"L{r}"]
        wcell.value = w
        wcell.number_format = "0.00%"
        wcell.font = Font(name="Calibri", bold=True, size=9)
        wcell.fill = SUB_FILL
        wcell.border = THIN
        for col, mx in zip(exit_cols, _EXIT_AXIS):
            # Equity/share = (EV + Cash − Debt) / Shares
            cell = dcf[f"{col}{r}"]
            cell.value = f"=({col}{ev_r}+$D$33-$D$34)/$D$12"
            cell.number_format = "$#,##0.00"
            cell.font = BLACK
            cell.fill = LINK_FILL
            cell.border = THIN
            if abs(w - 0.0924) < 1e-9 and abs(mx - EXIT_EV_EBITDA) < 1e-9:
                cell.fill = INPUT_FILL

    note_r = sh + 8
    dcf[f"L{note_r}"] = (
        "Yellow = base (WACC≈9.24%, Exit 17.5x). "
        "EV formula: XNPV(explicit FCFF) + (Y5 FCFF + Exit×EBITDA) / (1+WACC)^daycount. "
        f"Peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x → base {EXIT_EV_EBITDA:.1f}x; "
        f"bull {EXIT_EV_EBITDA_BULL:.0f}x; bear {EXIT_EV_EBITDA_BEAR:.1f}x."
    )
    dcf[f"L{note_r}"].font = NOTE_FONT
    dcf.merge_cells(f"L{note_r}:R{note_r}")


def _write_three_statement_bridge(dcf) -> None:
    """On-sheet audit: every FCFF/assumption driver and its 3-statement source."""
    s3 = SHEET_3S
    dcf["L26"] = f"{MODEL_NAME} — DCF ← 3-Statement linkage audit"
    dcf["L26"].font = WHITE_BOLD
    dcf["L26"].fill = HDR_FILL
    dcf.merge_cells("L26:O26")

    headers = ("DCF item", "Value / link", "3-Statement source", "Notes")
    for col, h in zip("LMNO", headers):
        cell = dcf[f"{col}27"]
        cell.value = h
        cell.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="833C0C")
        cell.border = THIN

    rows = [
        (
            28,
            "Tax rate (D5)",
            "=$D$5",
            f"='{s3}'!J13",
            "Same FY25 effective rate as 3S forecast (I35/I33)",
        ),
        (
            29,
            "NWC % of sales",
            f"='{s3}'!J15",
            f"='{s3}'!J15",
            "Flat policy NWC%; drives ΔNWC in FCFF",
        ),
        (
            30,
            "CapEx % (FY1)",
            f"='{s3}'!J18",
            f"='{s3}'!J18",
            "3S CapEx fade path → CF CapEx $",
        ),
        (
            31,
            "Revenue growth FY1",
            f"='{s3}'!J7",
            f"='{s3}'!J7",
            "Policy path on 3S assumption row 7",
        ),
        (
            32,
            "EBIT FY1 (DCF E21)",
            "=$E$21",
            f"='{s3}'!J33+'{s3}'!J31",
            "EBT + Interest = GP − SG&A − R&D − D&A",
        ),
        (
            33,
            "EBIT cross-check",
            f"='{s3}'!J26-'{s3}'!J28-'{s3}'!J29-'{s3}'!J30",
            "GP−SGA−R&D−DA",
            "Must equal E21 (live check)",
        ),
        (
            34,
            "D&A FY1",
            "=$E$23",
            f"='{s3}'!J30",
            "Avg-PP&E D&A from 3S IS",
        ),
        (
            35,
            "CapEx FY1",
            "=$E$24",
            f"='{s3}'!J68",
            "3S CF CapEx (also D15)",
        ),
        (
            36,
            "ΔNWC FY1",
            "=$E$25",
            f"='{s3}'!J90",
            "WC schedule change (CF row 64 = J90)",
        ),
        (
            37,
            "EBITDA FY5 (TV base)",
            "=$I$21+$I$23",
            f"='{s3}'!N33+'{s3}'!N31+'{s3}'!N30",
            "EBIT+D&A; Exit TV = this × D8",
        ),
        (
            38,
            "3S FY25 Cash (ref)",
            f"='{s3}'!I41",
            f"='{s3}'!I41",
            "FYE reference only — bridge uses 10-Q D14",
        ),
        (
            39,
            "3S FY25 Debt (ref)",
            f"='{s3}'!I49",
            f"='{s3}'!I49",
            "FYE reference only — bridge uses 10-Q D13",
        ),
        (
            40,
            "EBIT link OK?",
            '=IF(ABS(M32-M33)<1,"OK","MISMATCH")',
            "",
            "Flags if EBIT link ≠ GP−opex build",
        ),
    ]
    for r, item, val, src, note in rows:
        dcf[f"L{r}"] = item
        dcf[f"L{r}"].font = Font(name="Calibri", size=9)
        dcf[f"L{r}"].border = THIN
        if isinstance(val, str) and val.startswith("="):
            _link(dcf[f"M{r}"], val)
        else:
            dcf[f"M{r}"] = val
            dcf[f"M{r}"].border = THIN
        if isinstance(src, str) and src.startswith("="):
            _link(dcf[f"N{r}"], src)
        else:
            dcf[f"N{r}"] = src
            dcf[f"N{r}"].border = THIN
            dcf[f"N{r}"].font = Font(name="Calibri", size=8, color="595959")
        dcf[f"O{r}"] = note
        dcf[f"O{r}"].font = NOTE_FONT
        dcf[f"O{r}"].border = THIN
        for col in "LMNO":
            dcf[f"{col}{r}"].border = THIN

    for addr in ("M28", "M29", "M30", "M31"):
        dcf[addr].number_format = "0.00%"
    for addr in ("M32", "M33", "M34", "M35", "M36", "M37", "M38", "M39", "N32", "N34", "N35", "N36", "N37", "N38", "N39"):
        dcf[addr].number_format = "#,##0.0"

    dcf["L41"] = (
        "Green cells are live links. FCFF uses 3S EBIT/D&A/CapEx/ΔNWC. "
        "D5 tax = 3S J13. D13/D14 stay 10-Q for a current equity bridge "
        "(3S I41/I49 shown above as FY25 reference)."
    )
    dcf["L41"].font = NOTE_FONT
    dcf.merge_cells("L41:O41")


def wire_dcf_to_three_statement(wb) -> None:
    if SHEET_DCF not in wb.sheetnames or SHEET_3S not in wb.sheetnames:
        raise RuntimeError("Template must contain both 3 Statement Model and DCF Model sheets")

    dcf = wb[SHEET_DCF]
    s3 = SHEET_3S

    # --- Assumptions that must match the 3-statement ---
    # Tax: use the same effective rate the 3S forecast applies (row 13 = I35/I33).
    # Set as a formula BEFORE CSV inject so formula-protection keeps it.
    _link(dcf["D5"], f"='{s3}'!J13")
    dcf["D5"].number_format = "0.00%"
    dcf["B5"] = "Tax Rate (from 3-Statement J13 = FY25 tax/EBT)"
    dcf["C5"] = "Linked to 3S — do not hardcode"
    dcf["C5"].font = NOTE_FONT
    dcf["D5"].comment = Comment(
        "Tax rate for unlevered FCFF and WACC after-tax Rd.\n"
        "HOW: D5 = '3 Statement Model'!J13 (I35/I33 on hist, held in forecast).\n"
        "WHY: DCF must use the same t as the 3-statement so NOPAT/FCFF articulate.\n"
        "SOURCE: SEC 10-K FY2025 tax / EBT (via 3S hist).",
        MODEL_NAME,
    )

    # FCFF drivers — all live from 3-statement forecast columns J–N
    for dcol, scol in zip(_DCF_COLS, _S3_FORECAST_COLS):
        # EBIT = EBT + Interest (≡ GP − SG&A − R&D − D&A)
        _link(dcf[f"{dcol}21"], f"='{s3}'!{scol}33+'{s3}'!{scol}31")
        _link(dcf[f"{dcol}22"], f"={dcol}21*$D$5")
        _link(dcf[f"{dcol}23"], f"='{s3}'!{scol}30")  # D&A IS
        _link(dcf[f"{dcol}24"], f"='{s3}'!{scol}68")  # CapEx CF
        _link(dcf[f"{dcol}25"], f"='{s3}'!{scol}90")  # ΔNWC from WC schedule

    _link(dcf["D15"], f"='{s3}'!J68")
    dcf["B15"] = "Capex FY1 (linked to 3S J68)"
    dcf["B22"] = "Less: Unlevered Cash Taxes (EBIT × 3S tax rate)"

    for col in _DCF_COLS:
        _link(dcf[f"{col}28"], f"={col}27+{col}26")
        _link(dcf[f"{col}29"], f"={col}27+{col}26")
    _link(dcf["J28"], "=J27+J26")
    _link(dcf["J29"], "=J27")
    dcf["B20"] = "Year Fraction (display only; XNPV uses exact dates below)"

    # Mid-year dates via EDATE from transaction date D9.
    # Template periods E19:I19 are 0-based (0,1,2,3,4) — NOT 1-based.
    # months = (period + 0.5) * 12 → Y1 mid = +6m after D9.
    for col in _DCF_COLS:
        _link(dcf[f"{col}18"], f"=EDATE($D$9,({col}19+0.5)*12)")
    dcf["B18"] = "Date (mid-year via EDATE; periods are 0-based)"
    _link(dcf["J18"], "=I18")
    _link(dcf["D18"], "=$D$9")

    # Transaction / FYE dates align to 3S fiscal calendar (FICO FYE = 9/30)
    dcf["B9"] = "Transaction Date (= last hist FYE on 3S)"
    dcf["B10"] = "Fiscal Year End (first forecast FYE)"
    dcf["C9"] = "Matches 3S hist FYE (Sep 30)"
    dcf["C9"].font = NOTE_FONT

    # Equity bridge: keep 10-Q cash/debt (more current than FY25 3S) but label clearly
    dcf["B13"] = "Debt (10-Q bridge — not 3S FY25 I49)"
    dcf["B14"] = "Cash+mkt secs (10-Q bridge — not 3S FY25 I41)"
    dcf["C13"] = "See L39 for 3S FY25 debt ref"
    dcf["C14"] = "See L38 for 3S FY25 cash ref"
    dcf["C13"].font = NOTE_FONT
    dcf["C14"].font = NOTE_FONT

    # Base-case exit multiple (comps-sourced policy — not from 3S)
    dcf["D8"] = float(EXIT_EV_EBITDA)
    dcf["D8"].fill = INPUT_FILL
    dcf["D8"].font = Font(name="Calibri", color="0000FF")
    dcf["D8"].number_format = "0.0"
    dcf["B8"] = f"Exit EV/EBITDA (BASE {EXIT_EV_EBITDA:.1f}x — peer comps)"

    _link(dcf["J27"], "=($I$21+$I$23)*$D$8")
    dcf["B27"] = f"(Entry)/Exit TV — Exit EV/EBITDA (BASE {EXIT_EV_EBITDA:.1f}x)"

    exit_comment = (
        f"Exit EV/EBITDA multiple (BASE)\n"
        f"HOW: Yellow POLICY input D8 = {EXIT_EV_EBITDA:.1f}x; "
        f"TV = Year-5 EBITDA × D8.\n"
        f"WHY: Sourced from public peer EV/EBITDA (EFX/VRSK/SPGI/MCO/MSCI). "
        f"Peer median ≈ {PEER_MEDIAN_EV_EBITDA:.1f}x → base {EXIT_EV_EBITDA:.1f}x "
        f"(also ≈ 50/50 of bull {EXIT_EV_EBITDA_BULL:.0f}x and bear/Gordon "
        f"{EXIT_EV_EBITDA_BEAR:.1f}x). NOT FICO spot (~25–27x).\n"
        f"SOURCE (peer comps): {URL_PEER_COMPS}\n"
        f"SOURCE (FICO spot): {URL_FICO_EV_EBITDA}"
    )
    dcf["D8"].comment = Comment(exit_comment, MODEL_NAME)
    dcf["D8"].comment.width = 380
    dcf["D8"].comment.height = 200

    dcf["C8"] = (
        f"BASE {EXIT_EV_EBITDA:.1f}x ≈ peer median {PEER_MEDIAN_EV_EBITDA:.1f}x "
        f"(VCP Scanner). Bull {EXIT_EV_EBITDA_BULL:.0f}x / Bear {EXIT_EV_EBITDA_BEAR:.1f}x."
    )
    dcf["C8"].font = NOTE_FONT
    dcf["F8"] = "Peer EV/EBITDA comps (click)"
    dcf["F8"].hyperlink = URL_PEER_COMPS
    dcf["F8"].font = LINK_FONT
    dcf["G8"] = "FICO spot EV/EBITDA (click)"
    dcf["G8"].hyperlink = URL_FICO_EV_EBITDA
    dcf["G8"].font = LINK_FONT

    _write_peer_comps(dcf)
    _write_scenarios(dcf)
    _write_three_statement_bridge(dcf)
    _write_sensitivity(dcf)

    _link(dcf["D32"], "=XNPV(D6,D28:J28,D18:J18)")
    dcf["B32"] = "Enterprise Value (XNPV, mid-year EDATE dates)"

    dcf["B21"] = "EBIT (3S: EBT + Interest)"
    dcf["B23"] = "Plus: D&A (3S IS row 30)"
    dcf["B24"] = "Less: Capex (3S CF row 68)"
    dcf["B25"] = "Less: ΔNWC (3S WC row 90)"
