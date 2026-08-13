#!/usr/bin/env python3
"""Build downloadable assumptions list (MD/TXT/CSV/PDF) for every tab in LBOMODEL1."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent.parent
XLSX = REPO / "vengeanceaiUSC_LBOMODEL1.xlsx"
OUT_MD = ROOT / "LBOMODEL1_ASSUMPTIONS_LIST.md"
OUT_TXT = ROOT / "LBOMODEL1_ASSUMPTIONS_LIST.txt"
OUT_CSV = ROOT / "LBOMODEL1_ASSUMPTIONS_LIST.csv"
OUT_PDF = ROOT / "LBOMODEL1_ASSUMPTIONS_LIST.pdf"

RAW_XLSX = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx"
RAW_PDF = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.pdf"
RAW_MD = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.md"
TENK = "https://www.sec.gov/Archives/edgar/data/1794515/000179451526000012/zi-20251231.htm"
FACTS = "https://data.sec.gov/api/xbrl/companyfacts/CIK0001794515.json"
Q2 = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001794515&type=10-Q"


def A(tab, cell, name, value, what, how, why, source):
    return {
        "tab": tab,
        "cell": cell,
        "name": name,
        "value": value,
        "what": what,
        "how": how,
        "why": why,
        "source": source,
    }


def sync_dcf_shares(wb) -> None:
    """Align DCF + Shares tabs with ZoomInfo LBO inputs (still BMC leftovers otherwise)."""
    lbo = wb["LBO"]
    price = float(lbo["D8"].value)
    offer = float(lbo["H20"].value)
    shares = float(lbo["H18"].value)
    ebitda = float(lbo["D13"].value)
    debt = abs(float(lbo["D14"].value))
    cash = float(lbo["D15"].value)
    min_cash = float(lbo["D16"].value)
    exit_mult = float(lbo["D17"].value)
    premium = float(lbo["H21"].value)
    sponsor = float(lbo["D35"].value)
    new_debt = float(lbo["D29"].value) + float(lbo["D30"].value) + float(lbo["D31"].value)

    dcf = wb["DCF"]
    dcf["B2"] = "Discounted Cash Flow Model for ZoomInfo (GTM) — LBOMODEL1"
    dcf["C6"] = offer  # aligned to LBO offer for football-field compare
    dcf["C7"] = lbo["D9"].value
    dcf["C8"] = shares
    # Keep a simple SaaS-ish WACC for DCF tab; document as assumption
    rf, beta, erp, rd, tax = 0.043, 1.20, 0.05, 0.07, 0.21
    we = 0.55
    wd = 0.45
    wacc = we * (rf + beta * erp) + wd * rd * (1 - tax)
    dcf["C9"] = round(wacc, 4)
    dcf["C50"] = rd
    dcf["C51"] = tax
    dcf["C55"] = rf
    dcf["C56"] = beta
    dcf["C57"] = erp
    dcf["C58"] = we
    dcf["C53"] = wd
    dcf["C59"] = round(wacc, 4)
    dcf["C45"] = round(max(min_cash, cash * 0.5), 1)  # surplus cash concept simplified
    dcf["C46"] = round(new_debt, 1)  # PF debt for LBO bridge compare
    dcf["C47"] = round(new_debt - max(min_cash, cash * 0.5), 1)
    dcf["C52"] = dcf["C47"].value
    dcf["J37"] = exit_mult
    dcf["H48"] = round(offer * shares, 1)
    dcf["H49"] = shares
    dcf["I49"] = shares
    dcf["J49"] = shares
    dcf["H50"] = offer

    # Copy LBO forecast EBITDA/EBIT into DCF hist+forecast cols where present
    # LBO: C-J years 2023-2030 on rows 42/48/62; DCF uses C-I (7 years) historically — map 2023-2029
    for i, col in enumerate(["C", "D", "E", "F", "G", "H", "I"]):
        ycol = 3 + i  # LBO cols
        dcf[f"{col}12"] = lbo.cell(39, ycol).value
        dcf[f"{col}13"] = lbo.cell(40, ycol).value
        dcf[f"{col}15"] = lbo.cell(62, ycol).value  # EBITDA
        dcf[f"{col}16"] = lbo.cell(48, ycol).value  # EBIT
        dcf[f"{col}17"] = 0.21
        if lbo.cell(48, ycol).value is not None:
            ebit = float(lbo.cell(48, ycol).value)
            dcf[f"{col}18"] = round(ebit * (1 - 0.21), 1)

    # Terminal / stage uses J in template for exit approach — set terminal year from LBO J
    dcf["J36"] = lbo["J62"].value
    dcf["J12"] = lbo["J39"].value
    dcf["J13"] = lbo["J40"].value
    dcf["J15"] = lbo["J62"].value
    dcf["J16"] = lbo["J48"].value
    dcf["J17"] = 0.21
    if lbo["J48"].value is not None:
        dcf["J18"] = round(float(lbo["J48"].value) * (1 - 0.21), 1)

    # Shares tab
    sh = wb["Shares"]
    sh["B2"] = "Diluted shares for ZoomInfo (GTM) — LBOMODEL1"
    sh["E5"] = offer
    sh["E6"] = offer
    sh["E7"] = shares  # using latest basic/outstanding as diluted proxy (no detailed options ladder filed for MODEL1)
    sh["E8"] = 0.0
    sh["E9"] = 0.0
    sh["E10"] = 0.0
    sh["E11"] = 0.0
    sh["E12"] = 0.0
    sh["E14"] = shares
    # Clear old BMC option ladder ITM flags
    for r in range(18, 28):
        sh.cell(r, 3, 0)
        sh.cell(r, 5, 0)
    sh["E28"] = 0.0

    # 52wkHL — leave market tape as-is but label; document as BMC legacy tape not yet replaced
    hl = wb["52wkHL"]
    hl["A2"] = "52 week high low — NOTE: tape still sample/legacy; LBOMODEL1 uses GTM $4.31 spot on LBO!D8"

    if "Coversheet" in wb.sheetnames:
        cs = wb["Coversheet"]
        cs["B44"] = "Assumptions list (all tabs): LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.md|.pdf"
        cs["B45"] = RAW_PDF


def assumptions() -> list[dict]:
    items: list[dict] = []

    # Coversheet / overview
    items += [
        A("Coversheet", "B2", "Model identity", "vengeanceaiUSC-LBOMODEL1 — ZoomInfo (GTM)",
          "Name of this LBO case.", "Set on Coversheet and LBO!B2.",
          "Separates ZoomInfo LBOMODEL1 from the blank WSP BMC sample.",
          RAW_XLSX),
        A("Coversheet", "—", "Units", "$ millions except per share / rates / multiples",
          "Unit convention for the workbook.", "Stated on LBO!B3 / DCF!B3.",
          "Matches WSP template convention and SEC $ presentation scaled to millions.",
          TENK),
    ]

    # LBO tab
    items += [
        A("LBO", "D6/D7", "Company / ticker", "ZoomInfo Technologies Inc. / GTM",
          "Target company identity.", "Hardcoded inputs on LBO general inputs.",
          "Public GTM (legacy ZI) is the LBOMODEL1 case company.",
          TENK),
        A("LBO", "D8", "Current share price", "$4.31",
          "Spot equity price used for premium math.", "LBO!D8 market input.",
          "Latest available GTM print used when MODEL1 was built (~2026-08-12).",
          "https://finance.yahoo.com/quote/GTM"),
        A("LBO", "D9", "Price date", "2026-08-12",
          "As-of date for the spot price.", "LBO!D9.",
          "Timestamps the market input for auditability.",
          "https://finance.yahoo.com/quote/GTM"),
        A("LBO", "H21", "Acquisition premium", "25%",
          "Premium of offer price over spot.", "Offer/sh = D8×(1+premium).",
          "Standard illustrative take-private premium; editable deal assumption.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "H20", "Offer price / share", "$5.3875",
          "Implied takeout price per share.", "D8 × 1.25.",
          "Drives equity purchase price in Uses.",
          "Model assumption + market price"),
        A("LBO", "H18", "Diluted shares (millions)", "292.312",
          "Share count for equity value.", "From latest shares outstanding.",
          "10-Q 2026-06-30 common shares outstanding used as diluted proxy (options ladder not fully modeled).",
          Q2),
        A("LBO", "D20", "Equity purchase price", "$1,574.8m",
          "Cash to buy equity.", "Offer/sh × shares.",
          "Primary Use of funds line.",
          "Derived"),
        A("LBO", "D13", "LTM EBITDA", "$314.5m",
          "Entry EBITDA for leverage / multiple math.", "FY2025 OpInc $225.7m + D&A $88.8m.",
          "GAAP EBITDA proxy from latest annual 10-K (not adjusted EBITDA).",
          f"{TENK} ; {FACTS}"),
        A("LBO", "D14", "Gross debt", "−$1,269.9m",
          "Debt refinanced at close (entered as negative in template).", "LT debt + current maturities.",
          "10-Q 2026-06-30: LongTermDebt + LongTermDebtCurrent.",
          Q2),
        A("LBO", "D15", "Cash", "$150.1m",
          "Cash available at entry.", "Cash & equivalents + short-term investments.",
          "10-Q 2026-06-30.",
          Q2),
        A("LBO", "D16", "Minimum cash", "$50.0m",
          "Cash retained for ops; excess used as a Source.", "Hardcoded policy input.",
          "Leaves liquidity post-close; excess cash = 150.1 − 50 = 100.1.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "I11", "Entry EV / LTM EBITDA", "8.57x",
          "Implied entry multiple.", "(Equity buy + debt − cash) / EBITDA.",
          "Output of offer + net debt vs LTM EBITDA.",
          "Derived from SEC + market inputs"),
        A("LBO", "D17", "Exit EV / EBITDA", "8.0x",
          "Exit multiple in hold-year 5.", "LBO!D17 policy input.",
          "Near entry; assumes modest multiple compression vs entry 8.57x.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "C29/D29", "Term Loan A", "2.5x / $786.2m",
          "New TLA quantum.", "EBITDA × turns.",
          "Part of 5.0x new-money debt stack.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "C30/D30", "Term Loan B", "1.5x / $471.8m",
          "New TLB quantum.", "EBITDA × turns.",
          "Part of 5.0x new-money debt stack.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "C31/D31", "Senior Notes", "1.0x / $314.5m",
          "New senior notes quantum.", "EBITDA × turns.",
          "Completes 5.0x new debt with TLA+TLB.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "D35", "Sponsor equity", "$1,203.6m",
          "Equity check (Sources plug).", "Total Uses − excess cash − new debt.",
          "Balancing figure so Sources = Uses ($2,876.2m).",
          "Derived"),
        A("LBO", "G36", "Transaction fee %", "2% of equity value",
          "M&A / financing fee assumption.", "Fees $ = 2% × equity purchase.",
          "Simple fee load for Uses; financing fee dollars also shown in fee block.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "F65:J65", "Revenue growth (FY26–30)", "4% / 5% / 5% / 4% / 3%",
          "YoY revenue growth in forecast.", "Hardcoded on growth row.",
          "Modest recovery vs FY24 dip; not a hyper-growth case.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "E66/F66", "Gross margin", "FY25 82.35% → Fcst 82.85%",
          "Gross profit % of sales.", "From FY25 COGS implied by 10-K; slight +50 bps in forecast.",
          "Software gross margin structure from SEC P&L.",
          TENK),
        A("LBO", "E67/F67", "R&D margin", "14.57% flat in forecast",
          "R&D % of sales.", "FY2025 R&D / Revenue held flat.",
          "Keeps product investment while S&M is the efficiency lever.",
          TENK),
        A("LBO", "F68:J68", "SG&A margin path", "47.72% → 46.22% → 44.72% → 44.22% → 44.22%",
          "SG&A (S&M+G&A) % of sales in forecast.", "Compresses ~500 bps of revenue by Y3 vs FY25 49.72%.",
          "Operationalizes ZI AI-SDR thesis into the LBO P&L.",
          "LBOMODEL1 thesis + ZI_* tabs; FY25 base from " + TENK),
        A("LBO", "C69:J69", "Tax rate", "21%",
          "Book tax rate on EBT.", "Flat statutory-like rate in MODEL1.",
          "Simplifies vs actual cash tax / NOL complexity.",
          "Model assumption (LBOMODEL1)"),
        A("LBO", "I275/J275", "Sponsor MoIC / IRR", "~2.97x / ~24.4%",
          "Base returns at 8.0x exit.", "Exit equity / sponsor equity; IRR = MoIC^(1/5)−1.",
          "Illustrative; sensitive to exit multiple, leverage, and SG&A path.",
          "Derived on LBO returns block"),
    ]

    # DCF tab
    items += [
        A("DCF", "C6", "DCF reference share price", "$5.3875 (LBO offer)",
          "Price shown in DCF header for football-field compare.", "Synced to LBO offer/sh.",
          "Makes DCF tab speak the same deal price as the LBO.",
          "LBO!H20"),
        A("DCF", "C8", "Basic / diluted share count", "292.312m",
          "Share count for per-share value.", "Synced to LBO shares.",
          "Same 10-Q share count as LBO.",
          Q2),
        A("DCF", "C55/C56/C57", "CAPM inputs", "Rf 4.3% / β 1.20 / ERP 5.0%",
          "Building blocks of cost of equity.", "Ke = Rf + β×ERP.",
          "Illustrative SaaS CAPM stack for DCF tab (not Yacktman-adjusted).",
          "Model assumption (LBOMODEL1)"),
        A("DCF", "C50/C51", "Cost of debt / tax", "7.0% / 21%",
          "After-tax debt cost inputs.", "Rd×(1−t) in WACC.",
          "Rounded financing cost vs new LBO stack.",
          "Model assumption (LBOMODEL1)"),
        A("DCF", "C53/C58", "Capital structure weights", "Wd 45% / We 55%",
          "Target weights in WACC.", "Hardcoded policy on DCF tab.",
          "Closer to PF leveraged structure than pre-deal equity-heavy weights.",
          "Model assumption (LBOMODEL1)"),
        A("DCF", "C9/C59", "WACC", "~8.53%",
          "Discount rate for unlevered FCF.", "We×Ke + Wd×Rd×(1−t).",
          "Used for DCF PV math on this tab.",
          "Derived from CAPM + Rd assumptions"),
        A("DCF", "J37", "Exit EBITDA multiple (DCF)", "8.0x",
          "Terminal multiple in exit-EBITDA approach.", "Synced to LBO!D17.",
          "Keeps LBO and DCF exit frameworks comparable.",
          "LBO!D17"),
        A("DCF", "C37", "Perpetuity g", "0% in template cell (sensitivity uses 2–4%)",
          "Long-term growth in Gordon approach.", "Template default; grid varies g.",
          "Football-field sensitivity still shows 2–4% g range.",
          "WSP template convention + MODEL1 note"),
        A("DCF", "C31", "Cash-flow timing", "Middle of period",
          "Mid-year discounting convention.", "Midperiod adjustment factor on DCF tab.",
          "Standard IB mid-year convention for annual FCFs.",
          "WSP template"),
    ]

    # Shares
    items += [
        A("Shares", "E5/E6", "Offer price for dilution", "$5.3875",
          "Price used in treasury-stock method.", "Synced to LBO offer.",
          "Aligns diluted-share math with deal price.",
          "LBO!H20"),
        A("Shares", "E7/E14", "Basic & net diluted shares", "292.312m / 292.312m",
          "Share count used for equity value.", "Outstanding shares; options set to 0 in MODEL1.",
          "10-Q share count; detailed options ladder not loaded (net dilutive options = 0).",
          Q2),
        A("Shares", "E8:E11", "Options / other dilutives", "0",
          "Incremental dilutive securities.", "Cleared BMC ladder; set to zero pending full award rollforward.",
          "Conservative vs overstating dilution without a current award table.",
          "Model assumption (LBOMODEL1) — update from proxy/10-K award note when refined"),
    ]

    # 52wkHL
    items += [
        A("52wkHL", "E4/E5", "52-week high / low", "Legacy sample tape (not GTM)",
          "Market range for football-field.", "Still WSP BMC sample series in MODEL1.",
          "Not yet replaced with GTM price history; LBO uses spot $4.31 instead.",
          "WSP sample residual — replace with GTM tape in MODEL2+"),
    ]

    # ZI tabs
    items += [
        A("ZI_01_10K_Source", "E5", "FY2025 Revenue", "$1,249.5m",
          "Annual revenue anchor.", "XBRL RevenueFromContractWithCustomerExcludingAssessedTax.",
          "Latest annual 10-K fact.",
          f"{TENK} ; {FACTS}"),
        A("ZI_01_10K_Source", "E6", "FY2025 Selling & Marketing", "$414.6m",
          "S&M expense (the bloat line).", "XBRL SellingAndMarketingExpense.",
          "33.2% of FY25 revenue — thesis starting point.",
          FACTS),
        A("ZI_01_10K_Source", "E10/E13/E20", "OpInc / D&A / EBITDA proxy", "$225.7 / $88.8 / $314.5m",
          "Operating income, D&A, EBITDA proxy.", "OpInc + OtherDepreciationAndAmortization.",
          "Bridges to LBO!D13 LTM EBITDA.",
          FACTS),
        A("ZI_01_10K_Source", "E28", "S&M employees", "1,370",
          "Headcount in sales & marketing.", "10-K Human Capital disclosure (3,180 total).",
          "Denominator for SDR % assumption.",
          TENK),
        A("ZI_02_SM_Isolation", "C16", "Target mature SaaS S&M % rev", "22%",
          "Benchmark efficient S&M intensity.", "Yellow input vs FY25 33.2%.",
          "Defines “bloat” dollars vs a mature SaaS target.",
          "Model assumption (LBOMODEL1)"),
        A("ZI_03_Headcount_AI", "C6", "% of S&M that are outbound SDR/BDR", "40%",
          "Share of S&M headcount treated as AI-replaceable SDRs.", "Yellow input × 1,370.",
          "10-K does not disclose SDR count — explicit model assumption.",
          "Model assumption (LBOMODEL1)"),
        A("ZI_03_Headcount_AI", "C8", "Fully loaded cost / SDR / yr", "$155k",
          "Cash cost per outbound SDR.", "Yellow input.",
          "Below blended residual S&M $/employee (commissions/ads stay).",
          "Model assumption (LBOMODEL1)"),
        A("ZI_03_Headcount_AI", "C13/C15", "AI seat / oversight cost", "$15k per seat; $150k per oversight FTE",
          "Replacement opex for AI agents + humans.", "Yellow inputs; 1 oversight FTE / 20 seats.",
          "Net savings must clear AI + oversight to hit margin target.",
          "Model assumption (LBOMODEL1)"),
        A("ZI_03_Headcount_AI", "C21:E21", "Cumulative % SDRs replaced", "45% / 80% / 95% (2026–28)",
          "Headcount reduction schedule.", "Yellow cumulative exit fractions.",
          "Phased cut to avoid one-year cliff risk.",
          "Model assumption (LBOMODEL1)"),
        A("ZI_03_Headcount_AI", "C16/C17", "Severance / AI build", "$25k per SDR; $8m Y1 build",
          "One-time costs.", "Yellow inputs in P&L impact.",
          "Hits in-year EBITDA but excluded from run-rate bps.",
          "Model assumption (LBOMODEL1)"),
        A("ZI_04_EBITDA_Bridge", "C7", "Target margin expansion", "500 bps",
          "Minimum success hurdle for the AI program.", "Yellow input; YES/NO vs run-rate bps.",
          " equates to ~$62.5m EBITDA on $1,249.5m revenue.",
          "User thesis requirement"),
        A("ZI_05_Sensitivity", "B12:H17", "Sensitivity grid", "Replace-% × AI $/seat → bps",
          "Y3 run-rate bps under alternate AI costs / cut depths.", "Algebra on SDR0 & unit costs.",
          "Shows what it takes to stay above 500 bps if AI is more expensive.",
          "Derived from ZI_03 inputs"),
    ]

    return items


def render_md(items: list[dict]) -> str:
    lines = [
        "# vengeanceaiUSC-LBOMODEL1 — Assumptions List (every tab)",
        "",
        "**Ticker / issuer:** GTM — ZoomInfo Technologies Inc.",
        "",
        "**Format:** Plain list. Every item uses **What** / **How** / **Why** / **Source**.",
        "",
        f"**Workbook:** [{RAW_XLSX}]({RAW_XLSX})",
        "",
        f"**PDF:** [{RAW_PDF}]({RAW_PDF})",
        "",
        f"**Generated:** {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}",
        "",
        "Educational / research use only — not investment advice.",
        "",
    ]
    current = None
    for i, a in enumerate(items, 1):
        if a["tab"] != current:
            current = a["tab"]
            lines += ["", f"# Tab: `{current}`", ""]
        lines += [
            f"## {i}. {a['name']} ({a['cell']})",
            f"- **Value:** {a['value']}",
            f"- **What:** {a['what']}",
            f"- **How:** {a['how']}",
            f"- **Why:** {a['why']}",
            f"- **Source:** {a['source']}",
            "",
        ]
    return "\n".join(lines)


def render_txt(items: list[dict]) -> str:
    return render_md(items).replace("**", "").replace("`", "")


def write_csv(items: list[dict]) -> None:
    import csv

    with OUT_CSV.open("w", newline="") as f:
        w = csv.DictWriter(
            f,
            fieldnames=["tab", "cell", "name", "value", "what", "how", "why", "source"],
        )
        w.writeheader()
        w.writerows(items)


def write_pdf(items: list[dict]) -> None:
    styles = getSampleStyleSheet()
    title = ParagraphStyle("T", parent=styles["Heading1"], fontSize=14, spaceAfter=8)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=12, spaceBefore=10, spaceAfter=4)
    h3 = ParagraphStyle("H3", parent=styles["Heading3"], fontSize=10, spaceBefore=6, spaceAfter=2)
    body = ParagraphStyle("B", parent=styles["Normal"], fontSize=9, leading=12)
    doc = SimpleDocTemplate(
        str(OUT_PDF),
        pagesize=letter,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
        topMargin=0.6 * inch,
        bottomMargin=0.6 * inch,
    )
    story = [
        Paragraph("vengeanceaiUSC-LBOMODEL1 — Assumptions List (every tab)", title),
        Paragraph("Ticker: GTM (ZoomInfo Technologies Inc.). Format: What / How / Why / Source.", body),
        Paragraph(f"Workbook: {RAW_XLSX}", body),
        Paragraph("Educational / research use only — not investment advice.", body),
        Spacer(1, 8),
    ]
    current = None
    for i, a in enumerate(items, 1):
        if a["tab"] != current:
            current = a["tab"]
            story.append(Paragraph(f"Tab: {current}", h2))
        story.append(Paragraph(f"{i}. {a['name']} ({a['cell']}) — {a['value']}", h3))
        story.append(Paragraph(f"<b>What:</b> {a['what']}", body))
        story.append(Paragraph(f"<b>How:</b> {a['how']}", body))
        story.append(Paragraph(f"<b>Why:</b> {a['why']}", body))
        src = a["source"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        story.append(Paragraph(f"<b>Source:</b> {src}", body))
    doc.build(story)


def write_excel_assumptions(wb, items: list[dict]) -> None:
    """Embed filterable assumptions list sheets inside the workbook."""
    from collections import OrderedDict

    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    for name in list(wb.sheetnames):
        if name.startswith("Assumptions") or name == "00_Assumptions_List":
            del wb[name]

    hdr = PatternFill("solid", fgColor="1F4E79")
    hdr_f = Font(name="Calibri", color="FFFFFF", bold=True, size=11)
    title = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    bold = Font(name="Calibri", bold=True)
    black = Font(name="Calibri", size=10)
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    link = Font(name="Calibri", color="0563C1", underline="single", size=10)

    ws = wb.create_sheet("00_Assumptions_List", 1)
    ws["B2"] = "vengeanceaiUSC-LBOMODEL1 — Assumptions List (every tab)"
    ws["B2"].font = title
    ws["B3"] = "In-workbook list. Every row: Tab · Cell · Name · Value · What · How · Why · Source"
    ws["B4"] = "Ticker: GTM (ZoomInfo). Educational / research use only — not investment advice."
    headers = ["#", "Tab", "Cell", "Assumption", "Value", "What", "How", "Why", "Source"]
    for c, h in enumerate(headers, start=2):
        cell = ws.cell(6, c, h)
        cell.fill = hdr
        cell.font = hdr_f
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, a in enumerate(items, start=1):
        r = 6 + i
        vals = [i, a["tab"], a["cell"], a["name"], a["value"], a["what"], a["how"], a["why"], a["source"]]
        for c, v in enumerate(vals, start=2):
            cell = ws.cell(r, c, v)
            cell.font = black
            cell.alignment = wrap
            cell.border = thin
            if c == 10 and isinstance(v, str) and v.startswith("http"):
                cell.hyperlink = v.split(" ; ")[0].strip() if " ; " in v else v
                cell.font = link
        ws.row_dimensions[r].height = 48
    for i, w in enumerate([4, 22, 14, 32, 28, 36, 36, 40, 55], start=2):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B7"
    ws.auto_filter.ref = f"B6:J{6 + len(items)}"

    by_tab: OrderedDict[str, list] = OrderedDict()
    for a in items:
        by_tab.setdefault(a["tab"], []).append(a)
    insert_at = 2
    for tab, rows in by_tab.items():
        name = f"Assumptions_{tab.replace(' ', '_')}"[:31]
        tw = wb.create_sheet(name, insert_at)
        insert_at += 1
        tw["B2"] = f"Assumptions — {tab}"
        tw["B2"].font = title
        tw["B3"] = f"Subset of 00_Assumptions_List for tab `{tab}` only."
        for c, h in enumerate(["#", "Cell", "Assumption", "Value", "What", "How", "Why", "Source"], start=2):
            cell = tw.cell(5, c, h)
            cell.fill = hdr
            cell.font = hdr_f
        for i, a in enumerate(rows, start=1):
            r = 5 + i
            vals = [i, a["cell"], a["name"], a["value"], a["what"], a["how"], a["why"], a["source"]]
            for c, v in enumerate(vals, start=2):
                cell = tw.cell(r, c, v)
                cell.font = black
                cell.alignment = wrap
                cell.border = thin
                if c == 9 and isinstance(v, str) and v.startswith("http"):
                    cell.hyperlink = v.split(" ; ")[0].strip() if " ; " in v else v
                    cell.font = link
            tw.row_dimensions[r].height = 48
        for i, w in enumerate([4, 14, 28, 26, 36, 36, 40, 55], start=2):
            tw.column_dimensions[get_column_letter(i)].width = w
        tw.freeze_panes = "B6"

    if "Coversheet" in wb.sheetnames:
        cs = wb["Coversheet"]
        cs["B44"] = "IN-WORKBOOK ASSUMPTIONS LIST"
        cs["B44"].font = bold
        cs["B45"] = "See sheet 00_Assumptions_List (all tabs) and Assumptions_* sheets (one per tab)"
        cs["B46"] = f"{len(items)} assumptions · What / How / Why / Source"


def main() -> None:
    wb = load_workbook(XLSX)
    sync_dcf_shares(wb)
    items = assumptions()
    write_excel_assumptions(wb, items)
    wb.save(XLSX)
    wb.save(ROOT / "vengeanceaiUSC_LBOMODEL1.xlsx")
    wb.save(REPO / "LBO_Complex_Template_IRR_MoM.xlsx")

    OUT_MD.write_text(render_md(items))
    OUT_TXT.write_text(render_txt(items))
    write_csv(items)
    write_pdf(items)
    print(f"Wrote Excel assumptions sheets into {XLSX}")
    print(f"Wrote {OUT_MD}")
    print(f"Wrote {OUT_TXT}")
    print(f"Wrote {OUT_CSV}")
    print(f"Wrote {OUT_PDF}")
    print(f"Assumptions: {len(items)} across tabs {sorted({a['tab'] for a in items})}")


if __name__ == "__main__":
    main()
