#!/usr/bin/env python3
"""
vengeanceaiUSC-LBOMODEL1 — ZoomInfo AI-SDR value-creation LBO
Implements: FY2024 baseline, 15% Y1 S&M cut + $15m AI cost, 2% S&M growth thereafter,
≥500 bps EBITDA margin expansion by Y2, debt sweep, exit 12–14x, IRR/MOIC uplift vs base.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from copy import copy

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "vengeanceaiUSC_LBOMODEL1.xlsx"
OUT2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL1.xlsx"
PRISTINE = REPO / "LBO" / "templates" / "WSP_LBO_ORIGINAL_IRR_MoM.xlsx"

BLUE = Font(name="Calibri", color="0000FF")
BLACK = Font(name="Calibri", color="000000")
BOLD = Font(name="Calibri", bold=True)
TITLE = Font(name="Calibri", size=16, bold=True, color="1F4E79")
HDR_F = Font(name="Calibri", color="FFFFFF", bold=True)
HDR = PatternFill("solid", fgColor="1F4E79")
IN = PatternFill("solid", fgColor="FFF2CC")
CALC = PatternFill("solid", fgColor="E2EFDA")
SEC = PatternFill("solid", fgColor="D6DCE4")
GOOD = PatternFill("solid", fgColor="C6EFCE")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
LINK = Font(name="Calibri", color="0563C1", underline="single")


def lab(ws, r, c, t, bold=False):
    cell = ws.cell(r, c, t)
    cell.font = BOLD if bold else BLACK
    return cell


def inp(ws, r, c, v, fmt=None):
    cell = ws.cell(r, c, v)
    cell.font = BLUE
    cell.fill = IN
    cell.border = THIN
    if fmt:
        cell.number_format = fmt
    return cell


def calc(ws, r, c, v, fmt=None, fill=CALC):
    cell = ws.cell(r, c, v)
    cell.font = BLACK
    cell.fill = fill
    cell.border = THIN
    if fmt:
        cell.number_format = fmt
    return cell


def hdr_row(ws, r, values, start=2):
    for i, v in enumerate(values):
        cell = ws.cell(r, start + i, v)
        cell.fill = HDR
        cell.font = HDR_F


# ---- FY2024 baseline (user prompt + SEC cross-check) ----
REV0 = 1214.3
GP0 = 1024.5
COGS0 = REV0 - GP0  # 189.8
OPINC0 = 97.4
DA0 = 85.7
EBITDA0 = OPINC0 + DA0  # 183.1
SM0 = round(REV0 * 0.35, 1)  # ~$425.0 per prompt
RD0 = 196.1
GA0 = 295.3  # actual FY24 G&A (includes one-time-ish items)
# Cash interest etc. simplified
TAX = 0.21
PRICE = 4.05
SHARES = 307.294  # YE2025 diluted/outstanding proxy (m)
PREMIUM = 0.25
OFFER = PRICE * (1 + PREMIUM)
DEBT0 = 1329.9  # YE2025 for take-private timing
CASH0 = 179.9
MIN_CASH = 50.0
EXIT_MULT = 13.0  # mid of 12–14x
HOLD = 5
# New debt 5.0x entry EBITDA
TLA_X, TLB_X, NOTE_X = 2.5, 1.5, 1.0
RD_RATE = 0.075  # blended interest on new debt

# Operating forecast
GROWTHS = [0.03, 0.04, 0.04, 0.03, 0.03]  # Y1..Y5
GM = GP0 / REV0
RD_PCT = RD0 / REV0
GA_PCT = 0.18  # normalize G&A to leaner run-rate % of rev (strip FY24 spike)
DA_PCT = DA0 / REV0
CAPEX_PCT = 0.02
NWC_PCT = 0.02  # ΔNWC use of cash as % of Δrev approx via NWC level


def project_case(ai: bool):
    """Return year-by-year dicts for base (ai=False) or AI strategy (ai=True)."""
    years = []
    rev = REV0
    sm = SM0
    for i, g in enumerate(GROWTHS):
        y = 2025 + i  # Y1=2025 projection from FY24 baseline entry
        rev = rev * (1 + g)
        if ai:
            if i == 0:
                sm = SM0 * 0.85  # 15% cut
                ai_cost = 15.0
            else:
                sm = sm * 1.02  # 2% growth cap
                ai_cost = 15.0
        else:
            # Base: S&M stays ~35% of revenue (bloated)
            sm = rev * 0.35
            ai_cost = 0.0
        cogs = rev * (1 - GM)
        rd = rev * RD_PCT
        ga = rev * GA_PCT
        da = rev * DA_PCT
        ebit = rev - cogs - sm - ai_cost - rd - ga  # D&A embedded in opex below ebitda bridge
        # Treat SM/RD/GA as cash opex; D&A separate for EBITDA
        ebitda = rev - cogs - sm - ai_cost - rd - ga  # = EBIT + DA if DA not in opex
        # Above ebitda already excludes DA; add clarity:
        ebit = ebitda - da
        years.append(
            {
                "year": y,
                "rev": rev,
                "cogs": cogs,
                "sm": sm,
                "ai": ai_cost,
                "rd": rd,
                "ga": ga,
                "da": da,
                "ebitda": ebitda,
                "ebit": ebit,
                "ebitda_m": ebitda / rev,
            }
        )
    return years


def debt_and_returns(op_years, new_debt, sponsor_eq):
    """Interest, FCF, cash sweep, exit MoIC/IRR."""
    debt = new_debt
    rows = []
    for i, y in enumerate(op_years):
        interest = debt * RD_RATE
        ebt = y["ebit"] - interest
        tax = max(0, ebt) * TAX
        ni = ebt - tax
        # UFCF-ish to equity after interest (cash available for debt paydown)
        # NI + DA - capex - ΔNWC
        capex = y["rev"] * CAPEX_PCT
        dnwc = (y["rev"] - (REV0 if i == 0 else op_years[i - 1]["rev"])) * NWC_PCT
        fcf = ni + y["da"] - capex - dnwc
        paydown = min(debt, max(0, fcf))
        debt_end = debt - paydown
        rows.append(
            {
                **y,
                "interest": interest,
                "ni": ni,
                "fcf": fcf,
                "paydown": paydown,
                "debt_beg": debt,
                "debt_end": debt_end,
            }
        )
        debt = debt_end
    exit_ebitda = rows[-1]["ebitda"]
    exit_ev = exit_ebitda * EXIT_MULT
    exit_equity = exit_ev - rows[-1]["debt_end"]
    moic = exit_equity / sponsor_eq if sponsor_eq else 0
    irr = moic ** (1 / HOLD) - 1 if moic > 0 else 0
    return rows, exit_ebitda, exit_ev, exit_equity, moic, irr


def build():
    ai_ops = project_case(True)
    base_ops = project_case(False)

    # Entry
    equity_buy = OFFER * SHARES
    fees = 0.02 * equity_buy
    uses = equity_buy + DEBT0 + fees
    excess = max(0, CASH0 - MIN_CASH)
    new_debt = EBITDA0 * (TLA_X + TLB_X + NOTE_X)
    sponsor = uses - excess - new_debt

    ai_rows, ai_xebitda, ai_xev, ai_xeq, ai_moic, ai_irr = debt_and_returns(ai_ops, new_debt, sponsor)
    base_rows, b_xebitda, b_xev, b_xeq, b_moic, b_irr = debt_and_returns(base_ops, new_debt, sponsor)
    irr_uplift = (ai_irr - b_irr) * 100  # percentage points

    base_m0 = EBITDA0 / REV0
    ai_y2_m = ai_ops[1]["ebitda_m"]
    bps_y2 = (ai_y2_m - base_m0) * 10000

    # Start from pristine WSP to keep LBO/DCF shells, then overwrite + add strategy sheets
    if PRISTINE.exists():
        wb = load_workbook(PRISTINE)
    else:
        wb = Workbook()

    # --- Strategy_Summary (front) ---
    if "Strategy_Summary" in wb.sheetnames:
        del wb["Strategy_Summary"]
    s = wb.create_sheet("Strategy_Summary", 0)
    s["B2"] = "vengeanceaiUSC-LBOMODEL1 — ZoomInfo AI-SDR LBO Strategy"
    s["B2"].font = TITLE
    s["B3"] = "Take-private LBO: slash outbound SDR S&M, replace with AI agents, ≥500 bps EBITDA margin → higher FCF → faster debt paydown → higher IRR"
    lab(s, 5, 2, "THESIS", True)
    s["B5"].fill = SEC
    s["B6"] = "Replacing human SDRs with AI agents to drive 500+ bps EBITDA margin expansion."
    s["B7"] = "S&M is the largest SaaS opex lever → lower CAC → higher FCF → sweep debt → juice LBO IRR."

    lab(s, 9, 2, "ENTRY (ZoomInfo FY2024 10-K baseline + latest BS)", True)
    s["B9"].fill = SEC
    entry = [
        (10, "Revenue (FY2024)", REV0, "10-K"),
        (11, "Gross Profit / GM", f"{GP0} / {GM:.1%}", "10-K"),
        (12, "GAAP Operating Income", OPINC0, "10-K"),
        (13, "D&A", DA0, "XBRL OtherDepreciationAndAmortization"),
        (14, "Baseline EBITDA (OpInc+D&A)", EBITDA0, "Calculated"),
        (15, "Baseline EBITDA margin", base_m0, "Calculated"),
        (16, "S&M baseline (35% of rev)", SM0, "Prompt: ~35% / ~$425m"),
        (17, "Share price / shares (m)", f"${PRICE} / {SHARES}", "Market + 10-K shares"),
        (18, "Offer (25% premium)", f"${OFFER:.4f}", "Assumption"),
        (19, "Gross debt / cash", f"{DEBT0} / {CASH0}", "YE2025 10-K/Q"),
        (20, "Sponsor equity check", sponsor, "Uses − excess cash − new debt"),
        (21, "New debt (5.0x EBITDA)", new_debt, "TLA 2.5x + TLB 1.5x + Notes 1.0x"),
        (22, "Exit multiple", f"{EXIT_MULT:.1f}x", "Mid of 12–14x software band"),
    ]
    for r, name, val, src in entry:
        lab(s, r, 2, name)
        if isinstance(val, float):
            calc(s, r, 3, round(val, 4) if val < 2 else round(val, 1), "0.0%" if val < 2 and name.endswith("margin") else "#,##0.0")
        else:
            calc(s, r, 3, val)
        lab(s, r, 4, src)

    lab(s, 24, 2, "AI STRATEGY RULES (as specified)", True)
    s["B24"].fill = SEC
    s["B25"] = "Y1: cut S&M 15% vs baseline $425m; add $15m AI software/compute cost"
    s["B26"] = "Y2–Y5: S&M grows only 2%/yr (AI scales without headcount)"
    s["B27"] = "Target: ≥500 bps EBITDA margin expansion by Year 2 vs baseline year"

    lab(s, 29, 2, "RESULTS — AI case vs Base (no AI) case", True)
    s["B29"].fill = SEC
    hdr_row(s, 30, ["Metric", "Base (no AI)", "AI-SDR case", "Delta"])
    rows_out = [
        (31, "Y2 EBITDA margin", base_ops[1]["ebitda_m"], ai_ops[1]["ebitda_m"]),
        (32, "Y2 margin vs baseline (bps)", (base_ops[1]["ebitda_m"] - base_m0) * 10000, bps_y2),
        (33, "Y5 EBITDA ($m)", base_ops[4]["ebitda"], ai_ops[4]["ebitda"]),
        (34, "Exit EV @ 13x ($m)", b_xev, ai_xev),
        (35, "Exit net debt ($m)", base_rows[-1]["debt_end"], ai_rows[-1]["debt_end"]),
        (36, "Exit equity value ($m)", b_xeq, ai_xeq),
        (37, "MOIC", b_moic, ai_moic),
        (38, "IRR", b_irr, ai_irr),
    ]
    for r, name, b, a in rows_out:
        lab(s, r, 2, name)
        if "margin" in name and "bps" not in name:
            calc(s, r, 3, b, "0.0%")
            calc(s, r, 4, a, "0.0%", GOOD)
            calc(s, r, 5, a - b, "0.0%")
        elif "bps" in name:
            calc(s, r, 3, b, "#,##0.0")
            calc(s, r, 4, a, "#,##0.0", GOOD)
            calc(s, r, 5, a - b, "#,##0.0")
        elif name in ("MOIC",):
            calc(s, r, 3, b, "0.00x")
            calc(s, r, 4, a, "0.00x", GOOD)
            calc(s, r, 5, a - b, "0.00x")
        elif name == "IRR":
            calc(s, r, 3, b, "0.0%")
            calc(s, r, 4, a, "0.0%", GOOD)
            calc(s, r, 5, a - b, "0.0%")
        else:
            calc(s, r, 3, round(b, 1), "#,##0.0")
            calc(s, r, 4, round(a, 1), "#,##0.0", GOOD)
            calc(s, r, 5, round(a - b, 1), "#,##0.0")

    lab(s, 40, 2, "IRR POINTS FROM AI-DRIVEN S&M REDUCTION", True)
    s["B40"].fill = GOOD
    calc(s, 41, 2, f"AI case IRR {ai_irr:.1%} − Base IRR {b_irr:.1%} = {irr_uplift:.1f} percentage points of IRR")
    s["B41"].fill = GOOD
    calc(s, 42, 2, f"Y2 AI EBITDA margin {ai_y2_m:.1%} vs baseline {base_m0:.1%} = {bps_y2:.0f} bps (target ≥500: {'YES' if bps_y2 >= 500 else 'NO'})")
    s["B42"].fill = GOOD if bps_y2 >= 500 else IN

    s["B44"] = "See AI_Operating, Base_Operating, Debt_Sweep_AI, LBO (WSP shell inputs), 00_Assumptions_List"
    s["B45"] = "10-K: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm"
    s["B45"].font = LINK
    s["B45"].hyperlink = "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm"
    s.column_dimensions["B"].width = 48
    s.column_dimensions["C"].width = 16
    s.column_dimensions["D"].width = 16
    s.column_dimensions["E"].width = 14
    s.column_dimensions["F"].width = 40

    def write_op_sheet(name, rows, title):
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(name, 1)
        ws["B2"] = title
        ws["B2"].font = TITLE
        hdr_row(ws, 4, ["", "FY2024A"] + [f"Y{i} {2024+i}" for i in range(1, 6)])
        # baseline col C, then D-H forecast
        metrics = [
            ("Revenue", "rev"),
            ("COGS", "cogs"),
            ("S&M (human)", "sm"),
            ("AI agent cost", "ai"),
            ("R&D", "rd"),
            ("G&A", "ga"),
            ("EBITDA", "ebitda"),
            ("EBITDA margin", "ebitda_m"),
            ("D&A", "da"),
            ("EBIT", "ebit"),
            ("Interest", "interest"),
            ("Net income", "ni"),
            ("FCF (after interest)", "fcf"),
            ("Debt paydown", "paydown"),
            ("Debt end", "debt_end"),
        ]
        # baseline row values
        base_map = {
            "rev": REV0,
            "cogs": COGS0,
            "sm": SM0,
            "ai": 0.0,
            "rd": RD0,
            "ga": GA0,
            "ebitda": EBITDA0,
            "ebitda_m": base_m0,
            "da": DA0,
            "ebit": OPINC0,
            "interest": None,
            "ni": None,
            "fcf": None,
            "paydown": None,
            "debt_end": DEBT0,
        }
        for i, (label, key) in enumerate(metrics):
            r = 5 + i
            lab(ws, r, 2, label)
            bv = base_map.get(key)
            if bv is not None:
                calc(ws, r, 3, bv, "0.0%" if key.endswith("_m") else "#,##0.0")
            for j, row in enumerate(rows):
                v = row.get(key)
                if v is None:
                    continue
                calc(ws, r, 4 + j, v, "0.0%" if key.endswith("_m") else "#,##0.0", GOOD if key in ("ebitda", "ebitda_m", "fcf") else CALC)
        ws.column_dimensions["B"].width = 22
        for c in range(3, 10):
            ws.column_dimensions[get_column_letter(c)].width = 12
        return ws

    write_op_sheet("AI_Operating", ai_rows, "AI-SDR case — S&M cut 15% Y1 + $15m AI; S&M +2%/yr thereafter")
    write_op_sheet("Base_Operating", base_rows, "Base case — S&M stays ~35% of revenue (no AI transformation)")

    # Debt sweep compare
    if "Debt_Sweep_AI" in wb.sheetnames:
        del wb["Debt_Sweep_AI"]
    d = wb.create_sheet("Debt_Sweep_AI", 1)
    d["B2"] = "Debt schedule — AI case cash sweep (vs base ending debt)"
    d["B2"].font = TITLE
    hdr_row(d, 4, ["Year", "AI FCF", "AI paydown", "AI debt end", "Base debt end", "Extra paydown vs base"])
    for i, (ar, br) in enumerate(zip(ai_rows, base_rows)):
        r = 5 + i
        calc(d, r, 2, ar["year"], "0")
        calc(d, r, 3, ar["fcf"], "#,##0.0")
        calc(d, r, 4, ar["paydown"], "#,##0.0")
        calc(d, r, 5, ar["debt_end"], "#,##0.0", GOOD)
        calc(d, r, 6, br["debt_end"], "#,##0.0")
        calc(d, r, 7, br["debt_end"] - ar["debt_end"], "#,##0.0", GOOD)
    lab(d, 11, 2, "New debt at close")
    calc(d, 11, 3, new_debt, "#,##0.0")
    lab(d, 12, 2, "Blended interest rate")
    inp(d, 12, 3, RD_RATE, "0.0%")
    for c, w in enumerate([12, 12, 12, 14, 14, 18], start=2):
        d.column_dimensions[get_column_letter(c)].width = w

    # Patch WSP LBO sheet inputs if present
    if "LBO" in wb.sheetnames:
        lbo = wb["LBO"]
        lbo["B2"] = "Leveraged buyout model for ZoomInfo (GTM) — vengeanceaiUSC-LBOMODEL1 AI-SDR"
        lbo["D6"] = "ZoomInfo Technologies Inc."
        lbo["D7"] = "GTM"
        lbo["D8"] = PRICE
        lbo["D9"] = datetime(2026, 8, 15)
        lbo["D13"] = round(EBITDA0, 1)
        lbo["D14"] = -round(DEBT0, 1)
        lbo["D15"] = round(CASH0, 1)
        lbo["D16"] = MIN_CASH
        lbo["D17"] = EXIT_MULT
        lbo["H7"] = 1
        entry_mult = round((equity_buy + DEBT0 - CASH0) / EBITDA0, 2)
        for col in ["H", "I", "J"]:
            lbo[f"{col}10"] = round(EBITDA0, 1)
            lbo[f"{col}11"] = entry_mult
            lbo[f"{col}18"] = round(SHARES, 3)
            lbo[f"{col}20"] = round(OFFER, 4)
            lbo[f"{col}21"] = PREMIUM
        lbo["D20"] = round(equity_buy, 1)
        lbo["D21"] = round(DEBT0, 1)
        lbo["D22"] = round(fees, 1)
        lbo["D23"] = round(uses, 1)
        lbo["D27"] = round(excess, 1)
        lbo["C29"] = TLA_X
        lbo["D29"] = round(EBITDA0 * TLA_X, 1)
        lbo["C30"] = TLB_X
        lbo["D30"] = round(EBITDA0 * TLB_X, 1)
        lbo["C31"] = NOTE_X
        lbo["D31"] = round(EBITDA0 * NOTE_X, 1)
        lbo["D35"] = round(sponsor, 1)
        lbo["D36"] = round(uses, 1)
        # Returns summary
        lbo["B266"] = f"SUMMARY AT {EXIT_MULT:.1f}x EXIT — AI-SDR case vs Base"
        lbo["B275"] = "Sponsor equity (AI case)"
        lbo["F275"] = round(sponsor, 1)
        lbo["I275"] = round(ai_moic, 3)
        lbo["J275"] = round(ai_irr, 4)
        lbo["B276"] = "Sponsor equity (Base case)"
        lbo["I276"] = round(b_moic, 3)
        lbo["J276"] = round(b_irr, 4)
        lbo["B277"] = "IRR uplift from AI S&M (percentage points)"
        lbo["J277"] = round(irr_uplift / 100, 4)
        # Forecast years on LBO IS from AI case
        for i, row in enumerate(ai_rows):
            col = 6 + i  # F onward; align Y1 with col F
            if col > 10:
                break
            lbo.cell(39, col, row["year"])
            lbo.cell(42, col, round(row["rev"], 1))
            lbo.cell(43, col, round(-row["cogs"], 1))
            lbo.cell(45, col, round(-row["rd"], 1))
            # SG&A = S&M + AI + G&A in combined template line
            lbo.cell(46, col, round(-(row["sm"] + row["ai"] + row["ga"]), 1))
            lbo.cell(48, col, round(row["ebit"], 1))
            lbo.cell(50, col, round(-row["interest"], 1))
            lbo.cell(54, col, round(row["ni"], 1))
            lbo.cell(58, col, round(row["da"], 1))
            lbo.cell(62, col, round(row["ebitda"], 1))
            lbo.cell(65, col, GROWTHS[i])
            lbo.cell(68, col, round((row["sm"] + row["ai"] + row["ga"]) / row["rev"], 4))

    # Assumptions list sheet
    if "00_Assumptions_List" in wb.sheetnames:
        del wb["00_Assumptions_List"]
    a = wb.create_sheet("00_Assumptions_List", 1)
    a["B2"] = "Assumptions List — every key driver (What / How / Why / Source)"
    a["B2"].font = TITLE
    hdr_row(a, 4, ["#", "Tab", "Assumption", "Value", "What", "How", "Why", "Source"])
    assumptions = [
        ("Strategy_Summary", "Baseline year", "FY2024", "Entry operating year", "Prompt + 10-K", "User-specified baseline", "FY2024 10-K"),
        ("Strategy_Summary", "Revenue", f"${REV0}m", "FY2024 revenue", "10-K XBRL", "Historical anchor", "SEC 10-K"),
        ("Strategy_Summary", "Gross profit / GM", f"${GP0}m / {GM:.0%}", "FY2024 GP", "10-K", "Software GM", "SEC 10-K"),
        ("Strategy_Summary", "OpInc / D&A / EBITDA", f"{OPINC0} / {DA0} / {EBITDA0}", "Baseline earnings", "OpInc+D&A", "Standard add-back", "SEC XBRL"),
        ("Strategy_Summary", "S&M baseline", f"${SM0}m (35% rev)", "Bloated S&M", "35% × revenue", "Prompt ~$425m", "User prompt"),
        ("AI_Operating", "Y1 S&M cut", "−15%", "Terminate outbound SDRs", "SM1=SM0×0.85", "Remove SDR comp/seats", "User prompt"),
        ("AI_Operating", "AI replacement cost", "$15m / yr", "AI compute/software", "Fixed +$15m from Y1", "Agent infrastructure", "User prompt"),
        ("AI_Operating", "S&M growth Y2+", "2% / yr", "Cap S&M after cut", "SM_t=SM_t-1×1.02", "AI scales without HC", "User prompt"),
        ("AI_Operating", "Margin target", "≥500 bps by Y2", "EBITDA margin expansion", "Y2 margin − baseline", "Value-creation hurdle", "User prompt"),
        ("Base_Operating", "Base S&M policy", "35% of revenue", "No AI transformation", "SM=0.35×Rev", "Counterfactual", "User prompt"),
        ("Strategy_Summary", "Exit multiple", f"{EXIT_MULT:.1f}x", "Software exit EV/EBITDA", "Mid 12–14x band", "PE exit assumption", "User prompt"),
        ("Strategy_Summary", "New leverage", "5.0x EBITDA", "TLA/TLB/Notes", "2.5+1.5+1.0 turns", "Financing structure", "Model"),
        ("Debt_Sweep_AI", "Cash sweep", "Max FCF to debt", "Accelerated paydown", "min(debt, max(0,FCF))", "IRR via lower exit debt", "LBO mechanics"),
        ("Strategy_Summary", "Offer premium", "25%", "Take-private premium", "Offer=Px1.25", "Deal assumption", "Model"),
        ("Strategy_Summary", "IRR uplift metric", f"{irr_uplift:.1f} ppt", "AI vs Base IRR", "IRR_AI − IRR_Base", "Isolates AI S&M value", "Calculated"),
    ]
    for i, (tab, name, val, what, how, why, src) in enumerate(assumptions, 1):
        r = 4 + i
        for c, v in enumerate([i, tab, name, val, what, how, why, src], start=2):
            cell = a.cell(r, c, v)
            cell.border = THIN
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        a.row_dimensions[r].height = 30
    for i, w in enumerate([4, 18, 22, 22, 28, 28, 28, 22], start=2):
        a.column_dimensions[get_column_letter(i)].width = w
    a.freeze_panes = "B5"

    if "Coversheet" in wb.sheetnames:
        cs = wb["Coversheet"]
        cs["B2"] = "vengeanceaiUSC-LBOMODEL1 — ZoomInfo AI-SDR take-private"
        cs["B40"] = "Open Strategy_Summary first — AI vs Base IRR and 500+ bps check"
        cs["B41"] = f"AI IRR {ai_irr:.1%} vs Base {b_irr:.1%} → +{irr_uplift:.1f} IRR points from AI S&M"
        cs["B42"] = f"Y2 margin expansion {bps_y2:.0f} bps (target ≥500)"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT2.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    wb.save(OUT2)

    print("=== LBOMODEL1 AI-SDR RESULTS ===")
    print(f"Baseline EBITDA margin: {base_m0:.1%}")
    print(f"AI Y2 EBITDA margin: {ai_y2_m:.1%} ({bps_y2:.0f} bps vs baseline) target>=500: {bps_y2>=500}")
    print(f"Base IRR {b_irr:.1%} / MoIC {b_moic:.2f}x")
    print(f"AI   IRR {ai_irr:.1%} / MoIC {ai_moic:.2f}x")
    print(f"IRR uplift from AI S&M: {irr_uplift:.1f} percentage points")
    print(f"Wrote {OUT}")
    return {
        "bps_y2": bps_y2,
        "ai_irr": ai_irr,
        "base_irr": b_irr,
        "irr_uplift_ppt": irr_uplift,
        "ai_moic": ai_moic,
        "base_moic": b_moic,
    }


if __name__ == "__main__":
    build()
