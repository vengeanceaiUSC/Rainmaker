"""Plain-English explanations for every 3-statement assumption row (MODEL9).

SOURCE fields always include a clickable URL (col U) so users can open the filing/data.
Every assumption also links to the downloadable Assumptions List PDF (no charts).
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import List, Tuple

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .model9_assumptions import (
    ASSUMPTIONS_PDF_FILENAME,
    ASSUMPTIONS_PDF_URL,
    ASSUMPTIONS_PDF_VIEW_URL,
    DA_METHODOLOGY_NOTE,
    EXIT_EV_EBITDA,
    EXIT_EV_EBITDA_BEAR,
    EXIT_EV_EBITDA_BULL,
    MODEL_NAME as _MODEL_NAME,
    PEER_EV_EBITDA,
    PEER_MEDIAN_EV_EBITDA,
    URL_PEER_COMPS,
    WC_METHODOLOGY_NOTE,
)
from .named_range_map import SHEET_3S, SHEET_DCF
from .wacc import MODEL3_WACC

_FORECAST_COLS = ("J", "K", "L", "M", "N")
_COMMENT_AUTHOR = _MODEL_NAME

# Canonical source URLs (clickable in Excel col U)
URL_10K = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"
URL_10Q = "https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm"
URL_GUIDANCE = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm"
)
URL_FACTS = "https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json"
URL_8K_NOTES = "https://www.sec.gov/Archives/edgar/data/814547/000119312526117936/d56220d8k.htm"
URL_DAMODARAN_FCFF = "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histfcff.html"
URL_IR = "https://www.fico.com/en/investors"

LINK_FONT = Font(name="Calibri", size=9, color="0563C1", underline="single")

# (row, short_name, what_it_is, how_set, why, source_label, source_url)
ASSUMPTION_EXPLANATIONS: List[Tuple[int, str, str, str, str, str, str]] = [
    (
        7,
        "Revenue Growth",
        "Year-over-year % increase in revenue for each forecast year.",
        "Yellow POLICY input: 27% / 16% / 13% / 10% / 7%.",
        "Y1 ≈ company FY2026 revenue guidance (~$2.53B / FY25 $1.991B − 1 ≈ 27%). "
        "Then fade toward terminal g=3% (not straight-lined). This is the main "
        "forward-looking judgment; it drives Rev → GP → EBT → Net Earnings → CF.",
        "SEC EX-99.1 Q3 FY2026 — updated FY2026 revenue guidance $2.53B",
        URL_GUIDANCE,
    ),
    (
        8,
        "COGS % of Revenue",
        "Cost of revenues as a % of sales (gross margin = 1 − this).",
        "Excel equation: I25/I24 (FY25 COGS ÷ FY25 Revenue) held flat.",
        "Locks in latest reported cost structure (~17.8%). Scores mix is high-margin; "
        "holding FY25 is conservative vs further mix shift.",
        "SEC 10-K FY2025 — Cost of revenues / Total revenues",
        URL_10K,
    ),
    (
        9,
        "SG&A % of Revenue",
        "Operating opex (SG&A) as % of sales.",
        "Equation: MAX(15%, (I28 − 10,922)/I24 − 75bps × year). "
        "Strips FY25 restructuring; then −0.75% of sales each year.",
        "MODEL9: restore software operating leverage. Floor 15%; grind −75 bps/yr "
        "so SG&A can scale toward mid-teens as revenue expands (not stuck ~24–25%).",
        "SEC 10-K FY2025 — SG&A + restructuring note",
        URL_10K,
    ),
    (
        10,
        "R&D % of Revenue",
        "Research & development expense as % of sales (IS label: R&D, not Rent).",
        "Excel equation: I29/I24 held flat (~9.46%).",
        "FICO reinvests steadily in analytics/software. Flat % grows dollars with revenue.",
        "SEC 10-K FY2025 — Research and development",
        URL_10K,
    ),
    (
        11,
        "D&A % of PPE",
        "Depreciation rate from FY25; dollars = (Opening PPE + CapEx/2) × DA%.",
        "Rate = I30/I44 (FY25 DA ÷ FY25 PPE). "
        "Forecast DA$ = (Open + CapEx/2) × rate.",
        f"MODEL9: {DA_METHODOLOGY_NOTE} "
        "Half of current CapEx is treated as in service for the year "
        "(mid-year convention) without a circular closing-PPE loop.",
        "SEC 10-K FY2025 — D&A and PP&E",
        URL_10K,
    ),
    (
        12,
        "Interest % of Debt Open",
        "Interest expense as % of opening debt balance.",
        "Excel equation: I101/I98 (FY25 interest ÷ FY25 opening debt in schedule).",
        "Approximates average coupon / cost of debt on the book. Debt held flat "
        "(issuance = 0), so interest stays linked to that stock.",
        "SEC 10-K FY2025 — Interest expense; see also 8-K 6.250% notes",
        URL_10K,
    ),
    (
        13,
        "Tax Rate (% of EBT)",
        "Effective book tax rate applied to forecast EBT.",
        "Excel equation: I35/I33 (FY25 tax ÷ simplified FY25 EBT in this template).",
        "Uses reported effective rate (~19% on template EBT; ~18.8% on 10-K EBT). "
        "UFCF in DCF also uses unlevered EBIT × t.",
        "SEC 10-K FY2025 — Provision for income taxes / Income before taxes",
        URL_10K,
    ),
    (
        15,
        "DSO — Accounts Receivable (Days)",
        "Days Sales Outstanding. AR = Revenue × DSO / 365.",
        "Excel equation: ROUND(I42/I24×365, 0) from FY25; held flat each forecast year.",
        f"MODEL9: {WC_METHODOLOGY_NOTE} "
        "Row 15 is DSO again (not NWC%). AR is organic from collections days — "
        "no top-down NWC% AR plug.",
        "SEC companyfacts XBRL — AccountsReceivable / Revenue",
        URL_FACTS,
    ),
    (
        16,
        "Inventory (Days)",
        "Inventory days (Inv = COGS × days/365).",
        "Hard zero — FICO is software / scores; inventory is immaterial.",
        "No inventory cycle to fund.",
        "SEC 10-K FY2025 — Inventory ≈ $0",
        URL_10K,
    ),
    (
        17,
        "DPO — Accounts Payable (Days)",
        "Days Payable Outstanding. AP = COGS × DPO / 365.",
        "Excel equation: ROUND(I48/I25×365, 0) from FY25; held flat.",
        "MODEL9: DPO is an explicit WC driver paired with DSO. "
        "NWC = AR + Inventory − AP − Deferred (output); ΔNWC = NWCt − NWCt−1.",
        "SEC 10-K FY2025 — Accounts payable",
        URL_10K,
    ),
    (
        18,
        "CapEx % of Revenue",
        "Capital investment (PPE + capitalized software) as % of sales.",
        "Equation: fade from I68/I24 toward 1.0% steady. "
        "MODEL9 weights 40%→20%→10%→0%→0% on the peak (fast fade).",
        "Prior path left CapEx ≫ D&A (~$80M cumulative drag). "
        "Fade to maintenance ~1% so CapEx ≈ D&A by Y5 (terminal cash conversion).",
        "SEC 10-K FY2025 — PP&E purchases + capitalized software",
        URL_10K,
    ),
    (
        19,
        "Debt Issuance (Repayment)",
        "New borrowing (+) or repayment (−) in the forecast.",
        "POLICY input = 0 every year.",
        "Hold debt stock flat at FY25 level for the explicit period; interest follows. "
        "Net debt for DCF equity bridge uses the later 10-Q amount separately.",
        "SEC 10-Q Q3 FY2026 — debt stock (financing excluded from FCFF)",
        URL_10Q,
    ),
    (
        20,
        "Equity Issued (Repaid)",
        "New equity issued (+) or buybacks (−) in the forecast BS/CF.",
        "POLICY input = 0 every year.",
        "Buybacks are real historically but are financing — excluded from FCFF. "
        "Share count for $/share is the spot shares outstanding assumption.",
        "Damodaran FCFF framework — financing (buybacks) not in FCFF",
        URL_DAMODARAN_FCFF,
    ),
]


def _source_line(label: str, url: str) -> str:
    return f"{label} | {url}"


def _cell_commentary(name: str, how: str, why: str, label: str, url: str) -> str:
    """Text embedded in Excel cell comments (hover on assumption cells)."""
    return (
        f"{name}\n"
        f"HOW: {how}\n"
        f"WHY: {why}\n"
        f"SOURCE: {label}\n"
        f"LINK: {url}"
    )


def _column_c_note(how: str, why: str, label: str, url: str) -> str:
    """Visible note beside the label — must NOT start with '=' (#NAME?)."""
    why_short = why if len(why) <= 120 else why[:117] + "…"
    return f"WHY: {why_short} | SOURCE: {label} | LINK: {url} | HOW: {how}"


def _set_hyperlink(cell, url: str, display: str) -> None:
    """Clickable link that works in Excel and Google Sheets.

    Prefer Excel HYPERLINK() formula (Sheets-friendly). Also set the
    openpyxl hyperlink property as a fallback for desktop Excel.
    """
    # Escape quotes in display/url for formula safety
    safe_url = url.replace('"', '""')
    safe_disp = display.replace('"', '""')
    cell.value = f'=HYPERLINK("{safe_url}","{safe_disp}")'
    cell.hyperlink = url
    cell.font = LINK_FONT
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.border = THIN


def _set_hyperlink_plain(cell, url: str, display: str) -> None:
    """Plain-text display + hyperlink (no formula) — for cells that must stay text."""
    cell.value = display
    cell.hyperlink = url
    cell.font = LINK_FONT
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.border = THIN


HEADER_FILL = PatternFill("solid", fgColor="C65911")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
SUB_FILL = PatternFill("solid", fgColor="FCE4D6")
SUB_FONT = Font(name="Calibri", bold=True, size=9, color="833C0C")
BODY_FONT = Font(name="Calibri", size=9)
TITLE_FONT = Font(name="Calibri", bold=True, size=12, color="C65911")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def explanation_rows() -> List[dict]:
    rows = []
    for row, name, what, how, why, label, url in ASSUMPTION_EXPLANATIONS:
        rows.append(
            {
                "row": row,
                "assumption": name,
                "what_it_is": what,
                "how_set_in_model": how,
                "why_this_choice": why,
                "source": label,
                "source_url": url,
            }
        )
    return rows


def dcf_assumption_rows() -> List[dict]:
    """DCF policy assumptions included in the PDF list (not 3S rows 7–20)."""
    peers = ", ".join(f"{t} {m:.2f}x" for t, _n, m in PEER_EV_EBITDA)
    return [
        {
            "row": "DCF-D5",
            "assumption": "Tax Rate (unlevered)",
            "what_it_is": "Effective tax rate used for FCFF NOPAT and after-tax cost of debt.",
            "how_set_in_model": "Linked live to 3-Statement!J13 (= FY25 tax / simplified EBT).",
            "why_this_choice": "DCF must use the same t as the 3-statement forecast.",
            "source": "SEC 10-K FY2025 — via 3S J13",
            "source_url": URL_10K,
        },
        {
            "row": "DCF-D6",
            "assumption": "WACC (discount rate)",
            "what_it_is": "Weighted average cost of capital for XNPV of FCFF + TV.",
            "how_set_in_model": f"CAPM live formula ≈ {MODEL3_WACC:.2%} (We×Ke + Wd×Rd(1−t)).",
            "why_this_choice": "Market-consistent discount rate from Rf, β, ERP, and note coupon.",
            "source": "FRED DGS10 / Damodaran ERP / Yahoo β / 8-K 6.250% notes",
            "source_url": "https://fred.stlouisfed.org/series/DGS10",
        },
        {
            "row": "DCF-D7",
            "assumption": "Perpetual Growth (g)",
            "what_it_is": "Long-run growth used only in the Gordon TV cross-check.",
            "how_set_in_model": "POLICY input = 3.0%.",
            "why_this_choice": "Long-run nominal GDP-like floor; primary TV uses exit multiple.",
            "source": "Damodaran long-run growth framing",
            "source_url": "https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/histimpl.html",
        },
        {
            "row": "DCF-D8",
            "assumption": "Exit EV/EBITDA (BASE)",
            "what_it_is": "Terminal enterprise value multiple on Year-5 EBITDA.",
            "how_set_in_model": (
                f"POLICY BASE {EXIT_EV_EBITDA:.1f}x; Bull {EXIT_EV_EBITDA_BULL:.1f}x; "
                f"Bear {EXIT_EV_EBITDA_BEAR:.1f}x. Peer set: {peers}. "
                f"Median ≈ {PEER_MEDIAN_EV_EBITDA:.2f}x."
            ),
            "why_this_choice": (
                f"Base {EXIT_EV_EBITDA:.1f}x is near the public peer median "
                f"({PEER_MEDIAN_EV_EBITDA:.1f}x), between Gordon-implied bear and spot/bull."
            ),
            "source": "VCP Scanner FICO peer EV/EBITDA comps",
            "source_url": URL_PEER_COMPS,
        },
        {
            "row": "DCF-D13/D14",
            "assumption": "Debt & Cash (equity bridge)",
            "what_it_is": "Gross debt and cash+marketable securities for EV → equity.",
            "how_set_in_model": "10-Q Q3 FY2026 totals (more current than FY25 3S balances).",
            "why_this_choice": "Bridge should reflect the latest capital structure; 3S FY25 shown as reference.",
            "source": "SEC 10-Q Q3 FY2026",
            "source_url": URL_10Q,
        },
    ]


def write_assumption_explanations(wb) -> None:
    """Put WHY + SOURCE+URL in each assumption cell, col C, and clickable col U/W."""
    if SHEET_3S not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_3S}")
    ws = wb[SHEET_3S]

    ws["B4"] = (
        f"{_MODEL_NAME}: hover J–N for WHY+SOURCE; col U = filing source; "
        f"col W = download Assumptions List PDF (no charts)"
    )
    ws["B4"].font = Font(name="Calibri", bold=True, color="1F4E79")
    ws["B4"].fill = PatternFill("solid", fgColor="D6EAF8")

    # Banner link to the full assumptions PDF list
    _set_hyperlink(
        ws["A4"],
        ASSUMPTIONS_PDF_URL,
        "⬇ Download Assumptions List PDF",
    )
    ws["A4"].font = Font(name="Calibri", bold=True, size=10, color="0563C1", underline="single")

    for row, name, what, how, why, label, url in ASSUMPTION_EXPLANATIONS:
        ws[f"C{row}"] = (
            _column_c_note(how, why, label, url)
            + f" | ASSUMPTIONS PDF: {ASSUMPTIONS_PDF_URL}"
        )
        ws[f"C{row}"].font = Font(name="Calibri", italic=True, size=8, color="595959")
        ws[f"C{row}"].alignment = WRAP

        text = (
            _cell_commentary(name, how, why, label, url)
            + f"\nASSUMPTIONS LIST PDF: {ASSUMPTIONS_PDF_URL}"
        )
        for col in _FORECAST_COLS:
            cell = ws[f"{col}{row}"]
            comment = Comment(text, _COMMENT_AUTHOR)
            comment.width = 340
            comment.height = 180
            cell.comment = comment

        label_cell = ws[f"B{row}"]
        label_comment = Comment(text, _COMMENT_AUTHOR)
        label_comment.width = 340
        label_comment.height = 180
        label_cell.comment = label_comment

        ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 15, 40)

    # Legend block P–V (the chart the user sees) — PDF link lives INSIDE this block
    # Unmerge any prior wide merges that hid the link
    for merge in list(ws.merged_cells.ranges):
        if str(merge).startswith("P3:") or str(merge).startswith("P4:") or str(merge).startswith("P5:"):
            ws.unmerge_cells(str(merge))

    # ===== Title line = the PDF hyperlink (LARGE ORANGE letters) =====
    # Exact chart title text the user expects, made into the clickable PDF link.
    TITLE_LINK_TEXT = (
        "ASSUMPTIONS EXPLAINED — click blue Source (col U) for the filing/data"
    )
    ORANGE_FILL = PatternFill("solid", fgColor="FF6B00")
    ORANGE_TITLE_FONT = Font(
        name="Calibri",
        bold=True,
        size=20,
        color="FFFFFF",
        underline="single",
    )

    for merge in list(ws.merged_cells.ranges):
        m = str(merge)
        if m.startswith("P2:") or m.startswith("P3:") or m.startswith("P4:") or m.startswith("Q6:"):
            try:
                ws.unmerge_cells(m)
            except Exception:
                pass

    # Clear old banner cells so only the title hyperlink shows
    for addr in ("P2", "R3", "P4", "Q6", "U3", "U4"):
        ws[addr].value = None
        ws[addr].hyperlink = None

    # P4 = the chart title — LARGE ORANGE hyperlink to the Assumptions PDF
    _set_hyperlink(ws["P4"], ASSUMPTIONS_PDF_URL, TITLE_LINK_TEXT)
    ws["P4"].font = ORANGE_TITLE_FONT
    ws["P4"].fill = ORANGE_FILL
    ws["P4"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.merge_cells("P4:V4")
    ws.row_dimensions[4].height = 44

    # Subtitle under the orange title (still in the chart)
    ws["P5"] = (
        "Yellow = policy. Green = FY25-linked equations. "
        "Orange title above = click to download the Assumptions List PDF (no charts). "
        f"PDF: {ASSUMPTIONS_PDF_URL}"
    )
    ws["P5"].font = Font(name="Calibri", italic=True, size=9, color="595959")
    ws.merge_cells("P5:V5")

    headers = [
        ("P6", "Row"),
        ("Q6", "Assumption"),
        ("R6", "What it is"),
        ("S6", "How set in model"),
        ("T6", "Why this choice"),
        ("U6", "Source (click)"),
        ("V6", "Source URL"),
    ]
    for coord, label in headers:
        cell = ws[coord]
        cell.value = label
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN
        cell.alignment = Alignment(vertical="center")

    for row, name, what, how, why, label, url in ASSUMPTION_EXPLANATIONS:
        for col_idx, val in enumerate([row, name, what, how, why], start=16):  # P–T
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.font = BODY_FONT if col_idx > 17 else SUB_FONT
            cell.alignment = WRAP
            cell.border = THIN
            if col_idx <= 17:
                cell.fill = SUB_FILL

        # U = clickable filing source
        u = ws.cell(row=row, column=21)
        _set_hyperlink(u, url, label)
        u.comment = Comment(
            _cell_commentary(name, how, why, label, url)
            + f"\nASSUMPTIONS LIST PDF: {ASSUMPTIONS_PDF_URL}",
            _COMMENT_AUTHOR,
        )

        # V = filing source URL (original chart column)
        v = ws.cell(row=row, column=22)
        _set_hyperlink(v, url, url)
        v.comment = Comment(
            f"Filing source for this row:\n{url}\n\n"
            f"Full Assumptions List PDF (click orange title P4):\n{ASSUMPTIONS_PDF_URL}",
            _COMMENT_AUTHOR,
        )

    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["P"].width = 8
    ws.column_dimensions["Q"].width = 22
    ws.column_dimensions["R"].width = 36
    ws.column_dimensions["S"].width = 42
    ws.column_dimensions["T"].width = 48
    ws.column_dimensions["U"].width = 42
    ws.column_dimensions["V"].width = 55

    # DCF sheet: link to the same PDF near assumptions
    if SHEET_DCF in wb.sheetnames:
        dcf = wb[SHEET_DCF]
        _set_hyperlink(
            dcf["A4"],
            ASSUMPTIONS_PDF_URL,
            "⬇ Download Assumptions List PDF",
        )
        dcf["A4"].font = Font(
            name="Calibri", bold=True, size=10, color="0563C1", underline="single"
        )
        dcf["B4"] = f"Assumptions — full list PDF (no charts): {ASSUMPTIONS_PDF_FILENAME}"
        dcf["C4"] = ASSUMPTIONS_PDF_VIEW_URL
        dcf["C4"].hyperlink = ASSUMPTIONS_PDF_VIEW_URL
        dcf["C4"].font = LINK_FONT


def _write_assumptions_pdf(out_dir: Path, *, ticker: str = "FICO") -> Path:
    """Clean printable PDF list of every assumption (no charts / no tables graphics)."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / ASSUMPTIONS_PDF_FILENAME
    doc = SimpleDocTemplate(
        str(path),
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch,
        title=f"{_MODEL_NAME} Assumptions List",
        author=_MODEL_NAME,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "Title2",
        parent=styles["Heading1"],
        fontSize=14,
        spaceAfter=6,
        textColor="#1F4E79",
    )
    h_style = ParagraphStyle(
        "AssumpHead",
        parent=styles["Heading2"],
        fontSize=11,
        spaceBefore=12,
        spaceAfter=4,
        textColor="#833C0C",
    )
    body = ParagraphStyle(
        "Body2",
        parent=styles["Normal"],
        fontSize=9,
        leading=12,
        spaceAfter=2,
    )
    small = ParagraphStyle(
        "Small2",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        textColor="#595959",
        spaceAfter=2,
    )

    story = [
        Paragraph(f"{_MODEL_NAME} — Assumptions List", title_style),
        Paragraph(
            f"Ticker: {ticker}. Plain list only (no charts). "
            f"Each item includes What / How / Why / Source. "
            f"Download URL: {ASSUMPTIONS_PDF_URL}",
            small,
        ),
        Paragraph(
            f"Exit context: BASE {EXIT_EV_EBITDA:.1f}x EV/EBITDA "
            f"(peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x; "
            f"bull {EXIT_EV_EBITDA_BULL:.1f}x; bear {EXIT_EV_EBITDA_BEAR:.1f}x). "
            f"WACC ≈ {MODEL3_WACC:.2%}.",
            body,
        ),
        Paragraph(WC_METHODOLOGY_NOTE, body),
        Paragraph(DA_METHODOLOGY_NOTE, body),
        Spacer(1, 0.1 * inch),
        Paragraph("A. Three-Statement Forecast Assumptions", h_style),
    ]

    for r in explanation_rows():
        story.append(Paragraph(f"Row {r['row']}: {r['assumption']}", h_style))
        story.append(Paragraph(f"<b>What it is:</b> {r['what_it_is']}", body))
        story.append(Paragraph(f"<b>How set in model:</b> {r['how_set_in_model']}", body))
        story.append(Paragraph(f"<b>Why this choice:</b> {r['why_this_choice']}", body))
        story.append(
            Paragraph(
                f"<b>Source:</b> {r['source']}<br/>"
                f"<b>Source URL:</b> <link href='{r['source_url']}' "
                f"color='blue'>{r['source_url']}</link>",
                small,
            )
        )

    story.append(Paragraph("B. DCF Assumptions", h_style))
    for r in dcf_assumption_rows():
        story.append(Paragraph(f"{r['row']}: {r['assumption']}", h_style))
        story.append(Paragraph(f"<b>What it is:</b> {r['what_it_is']}", body))
        story.append(Paragraph(f"<b>How set in model:</b> {r['how_set_in_model']}", body))
        story.append(Paragraph(f"<b>Why this choice:</b> {r['why_this_choice']}", body))
        story.append(
            Paragraph(
                f"<b>Source:</b> {r['source']}<br/>"
                f"<b>Source URL:</b> <link href='{r['source_url']}' "
                f"color='blue'>{r['source_url']}</link>",
                small,
            )
        )

    story.append(Spacer(1, 0.2 * inch))
    story.append(
        Paragraph(
            "This PDF is the canonical printable assumptions list for the workbook. "
            "Filings and data sources remain clickable in Excel columns U/V; "
            "column W on every assumption row links back to this file.",
            small,
        )
    )
    doc.build(story)
    return path


def export_assumption_explanations_csv(out_dir: Path, *, ticker: str = "FICO") -> Path:
    """Write MODEL9_*_ASSUMPTIONS_EXPLAINED.csv, TXT, and PDF list (no charts)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"MODEL9_{ticker}_ASSUMPTIONS_EXPLAINED.csv"
    rows = explanation_rows() + dcf_assumption_rows()
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "row",
                "assumption",
                "what_it_is",
                "how_set_in_model",
                "why_this_choice",
                "source",
                "source_url",
                "assumptions_pdf_url",
            ],
        )
        writer.writeheader()
        for r in rows:
            writer.writerow({**r, "assumptions_pdf_url": ASSUMPTIONS_PDF_URL})

    txt_path = out_dir / f"MODEL9_{ticker}_ASSUMPTIONS_EXPLAINED.txt"
    lines = [
        f"{_MODEL_NAME} — Assumptions Explained",
        "=" * 60,
        "",
        f"DOWNLOADABLE ASSUMPTIONS LIST PDF (no charts): {ASSUMPTIONS_PDF_URL}",
        f"PDF view page: {ASSUMPTIONS_PDF_VIEW_URL}",
        "",
        f"Exit multiple context (DCF D8): BASE {EXIT_EV_EBITDA:.1f}x from public peer "
        f"EV/EBITDA (median ~{PEER_MEDIAN_EV_EBITDA:.1f}x). Peer comps: {URL_PEER_COMPS}",
        "",
        WC_METHODOLOGY_NOTE,
        DA_METHODOLOGY_NOTE,
        "",
    ]
    for r in rows:
        lines.extend(
            [
                f"ROW {r['row']}: {r['assumption']}",
                f"What it is: {r['what_it_is']}",
                f"How set in model: {r['how_set_in_model']}",
                f"Why this choice: {r['why_this_choice']}",
                f"Source: {r['source']}",
                f"Source URL: {r['source_url']}",
                f"Assumptions PDF: {ASSUMPTIONS_PDF_URL}",
                "",
                "-" * 40,
                "",
            ]
        )
    txt_path.write_text("\n".join(lines), encoding="utf-8")

    pdf_path = _write_assumptions_pdf(out_dir, ticker=ticker)
    # Also write a Google-Docs-friendly plain markdown the user can File→Open
    md_path = out_dir / f"MODEL9_{ticker}_ASSUMPTIONS_LIST.md"
    md = [
        f"# {_MODEL_NAME} — Assumptions List",
        "",
        f"**PDF download:** {ASSUMPTIONS_PDF_URL}",
        "",
        "Plain list only (no charts). Paste into Google Docs via File → Open if needed.",
        "",
        f"> {WC_METHODOLOGY_NOTE}",
        "",
        f"> {DA_METHODOLOGY_NOTE}",
        "",
    ]
    for r in rows:
        md.extend(
            [
                f"## {r['row']}: {r['assumption']}",
                f"- **What it is:** {r['what_it_is']}",
                f"- **How set in model:** {r['how_set_in_model']}",
                f"- **Why this choice:** {r['why_this_choice']}",
                f"- **Source:** [{r['source']}]({r['source_url']})",
                "",
            ]
        )
    md_path.write_text("\n".join(md), encoding="utf-8")
    print(f"  Assumptions PDF → {pdf_path}")
    print(f"  Assumptions MD  → {md_path}")
    return path
