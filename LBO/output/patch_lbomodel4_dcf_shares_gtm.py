#!/usr/bin/env python3
"""Wire LBOMODEL4 DCF + Shares tabs to ZoomInfo (GTM) data from other LBO tabs.

DCF previously still showed BMC Software template hardcodes (price $46.25, years
2011–2018, ~146m shares). This patch:
  - Relabels DCF/Shares for ZoomInfo (GTM)
  - Links price, date, shares, offer, debt, cash, exit multiple, tax, rates
    to LBO / Assumptions_Drivers / Debt_Sweep_AI / AI_Operating / Strategy_Summary
  - Rebuilds a 5-year unlevered FCF DCF from AI_Operating (thesis case)
  - Rebuilds Shares diluted-share bridge from LBO diluted share count
  - Updates 52wkHL summary high/low to GTM (for football-field inputs)
"""

from __future__ import annotations

from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "vengeanceaiUSC_LBOMODEL2 (29).xlsx"
OUT = ROOT / "vengeanceaiUSC_LBOMODEL4.xlsx"
OUT_COPY = ROOT / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL4.xlsx"

YELLOW = PatternFill("solid", fgColor="FFFF99")
HEADER = Font(bold=True, size=14)
BOLD = Font(bold=True)
THIN = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def set_val(ws, addr, value, *, number_format=None, fill=None, font=None):
    cell = ws[addr]
    cell.value = value
    if number_format:
        cell.number_format = number_format
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font


def clear_block(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).value = None


def patch_shares(ws):
    """GTM diluted shares — use LBO filing diluted count (no BMC option lattice)."""
    clear_block(ws, 2, 28, 2, 5)

    set_val(ws, "B2", "Diluted shares for ZoomInfo (GTM)", font=HEADER)
    set_val(ws, "B3", "$ mm except per share / share counts in millions")

    set_val(ws, "B5", "Offer price (LBO)")
    set_val(ws, "E5", "=LBO!H20", number_format="0.0000", fill=YELLOW)

    set_val(ws, "B6", "Latest close (LBO)")
    set_val(ws, "E6", "=LBO!D8", number_format="0.00", fill=YELLOW)

    set_val(ws, "B7", "Diluted shares outstanding (mm) — from LBO tab")
    set_val(ws, "E7", "=LBO!H18", number_format="0.000", fill=YELLOW)

    set_val(ws, "B8", "In-the-money exercisable options (mm)")
    set_val(ws, "E8", "=E28", number_format="0.000")
    set_val(ws, "B9", "Total proceeds ($mm)")
    set_val(ws, "E9", "=SUMPRODUCT((D19:D27<$E$5)*(C19:C27)*(D19:D27))", number_format="0.0")
    set_val(ws, "B10", "Total shares repurchased (mm) — treasury method")
    set_val(ws, "E10", "=IF(E5=0,0,E9/E5)", number_format="0.000")
    set_val(ws, "B11", "Net dilutive options")
    set_val(ws, "E11", "=E8-E10", number_format="0.000")
    set_val(ws, "B12", "Dilutive impact of shares from other securities")
    set_val(ws, "E12", 0, number_format="0.000", fill=YELLOW)

    set_val(ws, "B14", "Net diluted shares outstanding (mm)", font=BOLD)
    set_val(ws, "E14", "=E7+E11+E12", number_format="0.000", font=BOLD)

    set_val(ws, "B16", "Options outstanding — GTM note", font=BOLD)
    set_val(
        ws,
        "B17",
        "LBO!H18 already stores diluted shares used for offer value / share. "
        "No ZoomInfo option-tranche lattice was provided in this workbook; "
        "leave option rows at 0 unless you add filing detail. Source map: Strategy_Summary!C17.",
    )
    set_val(ws, "C18", "Out. shares")
    set_val(ws, "D18", "Exercise price")
    set_val(ws, "E18", "In-the-$-shares")
    for i, r in enumerate(range(19, 28), start=1):
        set_val(ws, f"B{r}", i)
        set_val(ws, f"C{r}", 0, number_format="0.0", fill=YELLOW)
        set_val(ws, f"D{r}", 0, number_format="0.00", fill=YELLOW)
        set_val(ws, f"E{r}", f'=IF(AND(C{r}>0,D{r}<$E$5),C{r},0)', number_format="0.0")
    set_val(ws, "B28", "ITM options total")
    set_val(ws, "E28", "=SUM(E19:E27)", number_format="0.0")


def patch_dcf(ws):
    """Rebuild DCF header + FCF + WACC + valuation using live sheet links."""
    # Keep football-field chart; clear numeric body we rewrite.
    clear_block(ws, 2, 91, 2, 11)

    # --- Title / general ---
    set_val(ws, "B2", "Discounted Cash Flow Model for ZoomInfo (GTM)", font=HEADER)
    set_val(ws, "B3", "$ mm except per share — forecast from AI_Operating (AI-SDR thesis case)")

    set_val(ws, "B5", "General assumptions", font=BOLD)
    set_val(ws, "B6", "Share price as of last close")
    set_val(ws, "C6", "=LBO!D8", number_format="0.00", fill=YELLOW)
    set_val(ws, "B7", "Latest closing share price date")
    set_val(ws, "C7", "=LBO!D9", number_format="YYYY-MM-DD", fill=YELLOW)
    set_val(ws, "B8", "Diluted shares outstanding (mm)")
    set_val(ws, "C8", "=Shares!E14", number_format="0.000", fill=YELLOW)
    set_val(ws, "B9", "Weighted average cost of capital")
    set_val(ws, "C9", "=C59", number_format="0.00%")

    set_val(ws, "B11", "Free cash flow buildup", font=BOLD)

    # Years: C=FY2024A hist, D–H = Y1–Y5
    years = [
        ("C", "AI_Operating!C", 2024, "FY2024A"),
        ("D", "AI_Operating!D", 2025, "Y1 2025"),
        ("E", "AI_Operating!E", 2026, "Y2 2026"),
        ("F", "AI_Operating!F", 2027, "Y3 2027"),
        ("G", "AI_Operating!G", 2028, "Y4 2028"),
        ("H", "AI_Operating!H", 2029, "Y5 2029"),
    ]
    set_val(ws, "B12", "Fiscal year")
    set_val(ws, "B13", "Fiscal year end date")
    for col, _, yr, _label in years:
        set_val(ws, f"{col}12", yr, number_format="0")
        set_val(ws, f"{col}13", datetime(yr, 12, 31), number_format="YYYY-MM-DD")

    # Line items
    set_val(ws, "B15", "EBITDA")
    set_val(ws, "B16", "EBIT")
    set_val(ws, "B17", "Tax rate")
    set_val(ws, "B18", "EBIAT (NOPAT)")
    set_val(ws, "B20", "Depreciation and amortization")
    set_val(ws, "B21", "Stock based compensation")
    set_val(ws, "B22", "Changes in net working capital")
    set_val(ws, "B23", "Other assets & liabilities")
    set_val(ws, "B24", "Unlevered CFO")
    set_val(ws, "B25", "Less: Capital expenditures")
    set_val(ws, "B26", "Less: Purchases of intangible assets")
    set_val(ws, "B27", "Unlevered FCF", font=BOLD)
    set_val(ws, "B28", "% growth")

    for col, pref, yr, _ in years:
        set_val(ws, f"{col}15", f"={pref}13", number_format="#,##0.0")
        set_val(ws, f"{col}16", f"={pref}16", number_format="#,##0.0")
        set_val(ws, f"{col}17", "=Assumptions_Drivers!$C$19", number_format="0.0%")
        set_val(ws, f"{col}18", f"={col}16*(1-{col}17)", number_format="#,##0.0")
        if yr >= 2025:
            set_val(ws, f"{col}20", f"={pref}15", number_format="#,##0.0")
            set_val(ws, f"{col}21", 0, number_format="#,##0.0")
            set_val(ws, f"{col}22", f"={pref}20", number_format="#,##0.0")
            set_val(ws, f"{col}23", 0, number_format="#,##0.0")
            # Unlevered CFO = EBIAT + D&A + SBC − ΔNWC + Other
            # AI_Operating ΔNWC is signed (negative = cash source); subtract as-is.
            set_val(
                ws,
                f"{col}24",
                f"={col}18+{col}20+{col}21-{col}22+{col}23",
                number_format="#,##0.0",
            )
            set_val(ws, f"{col}25", f"=-{pref}19", number_format="#,##0.0")
            set_val(ws, f"{col}26", 0, number_format="#,##0.0")
            set_val(ws, f"{col}27", f"={col}24+{col}25+{col}26", number_format="#,##0.0", font=BOLD)

    # Growth rates D→H on forecast
    for prev, col in zip("DEFG", "EFGH"):
        set_val(ws, f"{col}28", f"=IF({prev}27=0,0,{col}27/{prev}27-1)", number_format="0.0%")

    # Discount timing — forecast years D–H only (t=1..5)
    set_val(ws, "B30", "Discount period (years, mid-year)")
    for i, col in enumerate("DEFGH", start=1):
        set_val(ws, f"{col}30", i - 0.5, number_format="0.0")
    set_val(ws, "B31", "Assume cash flows are generated at:")
    set_val(ws, "C31", "Middle of period")
    set_val(ws, "B32", "Midperiod adjustment factor")
    set_val(ws, "C32", "=(1+$C$9)^0.5", number_format="0.0000")
    set_val(ws, "B33", "Present value of Unlevered FCF")
    for col in "DEFGH":
        set_val(
            ws,
            f"{col}33",
            f"={col}27/(1+$C$9)^{col}30",
            number_format="#,##0.0",
        )

    # Terminal value — perpetuity (left) and exit multiple (right)
    set_val(ws, "B35", "Perpetuity approach", font=BOLD)
    set_val(ws, "E35", "Exit EBITDA multiple approach", font=BOLD)

    set_val(ws, "B36", "Normalized FCFt+1")
    set_val(ws, "C36", "=H27*(1+C37)", number_format="#,##0.0")
    set_val(ws, "B37", "Long term growth rate (g)")
    set_val(ws, "C37", "=Assumptions_Drivers!C10", number_format="0.0%", fill=YELLOW)
    set_val(ws, "B38", "Terminal value")
    set_val(ws, "C38", "=IF(($C$9-C37)<=0,\"NM\",C36/($C$9-C37))", number_format="#,##0.0")
    set_val(ws, "B39", "Present value of terminal value")
    set_val(ws, "C39", "=IF(ISNUMBER(C38),C38/(1+$C$9)^H30,0)", number_format="#,##0.0")
    set_val(ws, "B40", "Present value of stage 1 cash flows")
    set_val(ws, "C40", "=SUM(D33:H33)", number_format="#,##0.0")
    set_val(ws, "B41", "Enterprise value", font=BOLD)
    set_val(ws, "C41", "=C39+C40", number_format="#,##0.0", font=BOLD)
    set_val(ws, "B42", "Implied TV exit EBITDA multiple")
    set_val(ws, "C42", "=IF(H15=0,0,C38/H15)", number_format="0.00x")

    set_val(ws, "E36", "Terminal year EBITDA")
    set_val(ws, "H36", "=H15", number_format="#,##0.0")
    set_val(ws, "E37", "Terminal value EBITDA multiple")
    set_val(ws, "H37", "=Assumptions_Drivers!C20", number_format="0.0x", fill=YELLOW)
    set_val(ws, "E38", "Terminal value")
    set_val(ws, "H38", "=H36*H37", number_format="#,##0.0")
    set_val(ws, "E39", "Present value of terminal value")
    set_val(ws, "H39", "=H38/(1+$C$9)^H30", number_format="#,##0.0")
    set_val(ws, "E40", "Present value of stage 1 cash flows")
    set_val(ws, "H40", "=C40", number_format="#,##0.0")
    set_val(ws, "E41", "Enterprise value", font=BOLD)
    set_val(ws, "H41", "=H39+H40", number_format="#,##0.0", font=BOLD)
    set_val(ws, "E42", "Implied TV perpetual growth rate")
    # g = WACC − UFCF_T+1 / TV ; UFCF_T+1 ≈ H27*(1+g) → solve g = (TV*WACC − UFCF)/(TV+UFCF)
    set_val(
        ws,
        "H42",
        "=IF(H38=0,0,($C$9*H38-H27)/(H38+H27))",
        number_format="0.00%",
    )

    # Net debt + fair value per share
    set_val(ws, "B44", "Net Debt", font=BOLD)
    set_val(ws, "E44", "Fair value per share", font=BOLD)

    set_val(ws, "B45", "Cash & equivalents (LBO latest)")
    set_val(ws, "C45", "=LBO!D15", number_format="#,##0.0", fill=YELLOW)
    set_val(ws, "B46", "Gross debt (LBO latest filing)")
    set_val(ws, "C46", "=ABS(LBO!D14)", number_format="#,##0.0", fill=YELLOW)
    set_val(ws, "B47", "Net debt")
    set_val(ws, "C47", "=C46-C45", number_format="#,##0.0")

    set_val(ws, "H45", "LBO Offer")
    set_val(ws, "I45", "Perpetuity")
    set_val(ws, "J45", "EBITDA exit")

    set_val(ws, "E46", "Enterprise value")
    # Offer EV = offer equity value + net debt (LBO!H12/H17 still hold BMC leftovers)
    set_val(ws, "H46", "=H48+C47", number_format="#,##0.0")
    set_val(ws, "I46", "=C41", number_format="#,##0.0")
    set_val(ws, "J46", "=H41", number_format="#,##0.0")

    set_val(ws, "E47", "Less: Net debt")
    set_val(ws, "H47", "=-C47", number_format="#,##0.0")
    set_val(ws, "I47", "=-C47", number_format="#,##0.0")
    set_val(ws, "J47", "=-C47", number_format="#,##0.0")

    set_val(ws, "E48", "Equity value")
    set_val(ws, "H48", "=LBO!H20*$C$8", number_format="#,##0.0")
    set_val(ws, "I48", "=I46+I47", number_format="#,##0.0")
    set_val(ws, "J48", "=J46+J47", number_format="#,##0.0")

    set_val(ws, "E49", "Diluted shares")
    set_val(ws, "H49", "=$C$8", number_format="0.000")
    set_val(ws, "I49", "=$C$8", number_format="0.000")
    set_val(ws, "J49", "=$C$8", number_format="0.000")

    set_val(ws, "E50", "Equity value per share")
    set_val(ws, "H50", "=LBO!H20", number_format="0.00")
    set_val(ws, "I50", "=IF(I49=0,0,I48/I49)", number_format="0.00")
    set_val(ws, "J50", "=IF(J49=0,0,J48/J49)", number_format="0.00")

    set_val(ws, "E51", "Premium (discount) to LBO offer")
    set_val(ws, "H51", "NM")
    set_val(ws, "I51", "=IF(H50=0,0,I50/H50-1)", number_format="0.0%")
    set_val(ws, "J51", "=IF(H50=0,0,J50/H50-1)", number_format="0.0%")

    # Cost of capital — post-LBO target stack from Debt_Sweep + sponsor equity
    set_val(ws, "B49", "Cost of capital assumptions", font=BOLD)
    set_val(ws, "B50", "Cost of debt (wtd TLA/TLB/Notes)")
    set_val(
        ws,
        "C50",
        "=(Debt_Sweep_AI!G12*Assumptions_Drivers!C27"
        "+Debt_Sweep_AI!G13*Assumptions_Drivers!C28"
        "+Debt_Sweep_AI!G14*(Assumptions_Drivers!C14+0.035))"
        "/Debt_Sweep_AI!G15",
        number_format="0.00%",
    )
    set_val(ws, "B51", "Tax rate")
    set_val(ws, "C51", "=Assumptions_Drivers!C19", number_format="0.0%")
    set_val(ws, "B52", "Debt (new LBO stack $m)")
    set_val(ws, "C52", "=Debt_Sweep_AI!G15", number_format="#,##0.0")
    set_val(ws, "B53", "Debt as a % of total capital")
    set_val(ws, "C53", "=C52/(C52+C54)", number_format="0.0%")
    set_val(ws, "B54", "Sponsor equity check ($m)")
    set_val(ws, "C54", "=Strategy_Summary!C20", number_format="#,##0.0")

    set_val(ws, "B55", "Risk free rate (SOFR from Drivers)")
    set_val(ws, "C55", "=Assumptions_Drivers!C14", number_format="0.00%", fill=YELLOW)
    set_val(ws, "B56", "Beta (MODEL CONST — SaaS mid)")
    set_val(ws, "C56", 1.25, number_format="0.00", fill=YELLOW)
    set_val(ws, "B57", "Market risk premium")
    set_val(ws, "C57", 0.05, number_format="0.0%", fill=YELLOW)
    set_val(ws, "B58", "Equity as % of total capital structure")
    set_val(ws, "C58", "=1-C53", number_format="0.0%")
    set_val(ws, "B59", "Cost of capital (WACC)", font=BOLD)
    set_val(
        ws,
        "C59",
        "=C58*(C55+C56*C57)+C53*C50*(1-C51)",
        number_format="0.00%",
        font=BOLD,
    )

    # Sensitivity stubs (static structure; values formula-driven where possible)
    set_val(ws, "B61", "Sensitivity analysis", font=BOLD)
    set_val(ws, "C63", "Equity value per share (perpetuity)")
    set_val(ws, "D64", "Long term growth rate (g):")
    set_val(ws, "C65", "=I50", number_format="0.00")
    for i, g in enumerate([0.01, 0.015, 0.02, 0.025, 0.03]):
        col = get_column_letter(4 + i)
        set_val(ws, f"{col}65", g, number_format="0.0%")
    for ri, w in enumerate([0.11, 0.10, 0.09, 0.08, 0.07, 0.06]):
        r = 66 + ri
        set_val(ws, f"C{r}", w, number_format="0.0%")
        for i in range(5):
            col = get_column_letter(4 + i)
            # EV = PV stage1@w + TV; approximate TV with g from header row
            # Simplified: rescale from base I50 using (WACC-g) ratio on TV-heavy value
            set_val(
                ws,
                f"{col}{r}",
                f'=IF(($C{r}-{col}$65)<=0,"NM",'
                f"(($C$40+($H$27*(1+{col}$65))/($C{r}-{col}$65))/(1+$C{r})^$H$30-$C$47)/$C$8)",
                number_format="0.00",
            )
    set_val(ws, "B68", "WACC:")

    set_val(ws, "C73", "Equity value per share (exit multiple)")
    set_val(ws, "D74", "Exit EBITDA Multiple")
    set_val(ws, "C75", "=J50", number_format="0.00")
    for i, m in enumerate([10.9, 11.9, 12.9, 13.9, 14.9]):
        col = get_column_letter(4 + i)
        set_val(ws, f"{col}75", m, number_format="0.0x")
    for ri, w in enumerate([0.11, 0.10, 0.09, 0.08, 0.07, 0.06]):
        r = 76 + ri
        set_val(ws, f"C{r}", w, number_format="0.0%")
        for i in range(5):
            col = get_column_letter(4 + i)
            set_val(
                ws,
                f"{col}{r}",
                f"=(($C$40+($H$15*{col}$75)/(1+$C{r})^$H$30)-$C$47)/$C$8",
                number_format="0.00",
            )
    set_val(ws, "B78", "WACC:")

    # EBITDA % of plan x exit multiple
    set_val(ws, "C83", "Equity value per share")
    set_val(ws, "D84", "Exit EBITDA Multiple")
    set_val(ws, "C85", "EBITDA % of plan")
    set_val(ws, "D85", "=J50", number_format="0.00")
    for i, m in enumerate([10.9, 11.9, 12.9, 13.9]):
        col = get_column_letter(5 + i)
        set_val(ws, f"{col}85", m, number_format="0.0x")
    for ri, pct in enumerate([0.8, 0.9, 1.0, 1.1, 1.2, 1.3]):
        r = 86 + ri
        set_val(ws, f"C{r}", pct, number_format="0%")
        set_val(ws, f"D{r}", f"=$H$15*C{r}", number_format="#,##0.0")
        for i in range(4):
            col = get_column_letter(5 + i)
            set_val(
                ws,
                f"{col}{r}",
                f"=(($C$40+(D{r}*{col}$85)/(1+$C$9)^$H$30)-$C$47)/$C$8",
                number_format="0.00",
            )

    set_val(ws, "B93", "Football field", font=BOLD)
    set_val(ws, "B94", "Sources: AI_Operating FCF build; Assumptions_Drivers tax/g/exit; "
                     "LBO price/shares/debt/cash; Debt_Sweep_AI stack for kd; "
                     "Strategy_Summary equity check; 52wkHL for market range.")

    # Football field data (chart may reference these)
    set_val(ws, "C114", "Low")
    set_val(ws, "D114", "Diff.")
    set_val(ws, "E114", "High")
    set_val(ws, "B115", "DCF Value at exit-multiple band (10.9x–14.9x)")
    set_val(ws, "C115", "=D76", number_format="0.00")  # high WACC / low multiple
    set_val(ws, "E115", "=H81", number_format="0.00")  # low WACC / high multiple
    set_val(ws, "D115", "=E115-C115", number_format="0.00")
    set_val(ws, "B116", "DCF Value at 1.0%–3.0% Perpetuity Range")
    set_val(ws, "C116", "=D66", number_format="0.00")
    set_val(ws, "E116", "=H71", number_format="0.00")
    set_val(ws, "D116", "=E116-C116", number_format="0.00")
    set_val(ws, "B117", "LBO Offer price / share")
    set_val(ws, "C117", "=LBO!H20", number_format="0.00")
    set_val(ws, "E117", "=LBO!H20", number_format="0.00")
    set_val(ws, "D117", 0, number_format="0.00")
    set_val(ws, "B118", "52 Week Market High/Low (GTM)")
    set_val(ws, "C118", "='52wkHL'!E5", number_format="0.00")
    set_val(ws, "E118", "='52wkHL'!E4", number_format="0.00")
    set_val(ws, "D118", "=E118-C118", number_format="0.00")


def patch_52wkhl(ws):
    """Replace BMC 52w summary with GTM Yahoo 1y high/low (price matches LBO $4.05)."""
    set_val(ws, "A2", "52 week high low — ZoomInfo (GTM)")
    set_val(ws, "A4", "52 week high")
    set_val(ws, "E4", 12.51, number_format="0.00", fill=YELLOW)
    set_val(ws, "A5", "52 week low")
    set_val(ws, "E5", 2.54, number_format="0.00", fill=YELLOW)
    set_val(
        ws,
        "A6",
        "Summary high/low from Yahoo Finance GTM 1y range (regularMarketPrice 4.05 matches LBO!D8). "
        "Daily OHLC rows below are legacy BMC template history and are not used by DCF.",
    )


def patch_coversheet(ws):
    set_val(
        ws,
        "B2",
        "vengeanceaiUSC-LBOMODEL4 — ZoomInfo (GTM) AI-SDR take-private",
        font=HEADER,
    )
    set_val(
        ws,
        "B40",
        "Open Strategy_Summary first — AI vs Base IRR; DCF/Shares now wired to GTM tabs",
    )


def main():
    wb = load_workbook(SRC)
    patch_shares(wb["Shares"])
    patch_dcf(wb["DCF"])
    patch_52wkhl(wb["52wkHL"])
    patch_coversheet(wb["Coversheet"])

    # Light note on Assumptions_List engine column for DCF relevance
    als = wb["00_Assumptions_List"]
    als["B3"] = (
        "Each row states Where (Tab!Cell), the math location, commentary, and the formula meaning. "
        "Operating engine is LBO; DCF tab now consumes AI_Operating + Drivers + LBO + Debt_Sweep for GTM valuation."
    )

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT_COPY.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    wb.save(OUT_COPY)
    # Also overwrite the branch upload name for continuity
    wb.save(SRC)
    print(f"Wrote {OUT}")
    print(f"Wrote {OUT_COPY}")
    print(f"Updated {SRC}")


if __name__ == "__main__":
    main()
