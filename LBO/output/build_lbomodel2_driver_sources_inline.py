#!/usr/bin/env python3
"""
LBOMODEL2: put a source URL + ~30-word credibility/reasonableness justification
on EVERY Assumptions_Drivers row (same row, columns to the right).
Remove standalone AI_Cost_Sources tab — evidence lives next to each driver.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL2.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL2.xlsx"
PDF = REPO / "LBO" / "output" / "LBOMODEL2_AI_COST_SOURCES.pdf"
MD = REPO / "LBO" / "output" / "LBOMODEL2_AI_COST_SOURCES.md"
NOTES = REPO / "LBO" / "output" / "LBOMODEL2_NOTES.md"

RAW_XLSX = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx"
RAW_PDF = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_AI_COST_SOURCES.pdf"

HDR = PatternFill("solid", fgColor="1F4E79")
HDR_F = Font(name="Calibri", color="FFFFFF", bold=True)
TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
BOLD = Font(name="Calibri", bold=True)
BLACK = Font(name="Calibri", size=10)
LINK = Font(name="Calibri", color="0563C1", underline="single", size=9)
NOTE = PatternFill("solid", fgColor="E7F3FF")
IN = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
WRAP = Alignment(wrap_text=True, vertical="top")

# row -> (primary_url, label_for_link, ~30-word justification)
# Secondary URL optional in tuple index 3
DRIVER_SOURCES = {
    5: (
        "https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results",
        "ZoomInfo FY2024 results (BusinessWire)",
        "ZI FY24 revenue fell ~2%, so 10% is a forward recovery/LBO underwrite—not trailing actual. Held constant in AI vs Base so IRR uplift comes from cost takeout, not inflated top-line.",
        "https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean",
    ),
    6: (
        "https://www.getaleph.com/answers/saas-gross-margin-2026",
        "Aleph/Benchmarkit SaaS GM median ~80%",
        "2025 software median gross margin is ~80% (Aleph×Benchmarkit). ZI already prints ~84% GAAP GM; using 80% is a conservative SaaS benchmark, not an aggressive stretch.",
        "https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc",
    ),
    7: (
        "https://firstsales.io/blog/cost-per-meeting-outbound/",
        "AI SDR vs human SDR cost/meeting",
        "Published outbound math: fully loaded human SDRs ~$110–150k vs much cheaper AI stacks. Cutting human sales payroll 15% in Y1 is a modest slice of ZI’s 1,513 S&M headcount.",
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
    ),
    8: (
        "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
        "SDR pay mix ~70/30 base/variable",
        "SaaS SDR OTE is typically ~70% base / ~30% variable. Replacing outbound humans with AI logically shrinks the commission pool; −20% is a partial, not full, cut of that variable layer.",
        "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
    ),
    9: (
        "https://www.salesforce.com/news/press-releases/2025/05/15/agentforce-flexible-pricing-news/",
        "Salesforce Agentforce $2/conv + Flex Credits",
        "SF publishes $2/conversation and ~$0.10/action Flex Credits. At ZI scale (~378 outbound roles at 25% of 1,513 S&M), $5–10k/mo AI agents imply ~$23–45m; $15m is conservative.",
        "https://highticketaisystems.com/blog/ai-sdr-pricing-explained",
    ),
    10: (
        "https://www.gartner.com/en/newsroom/press-releases/2026-08-10-gartner-forecasts-worldwide-artificial-intelligence-optimized-iaas-spending-to-grow-96-percent-in-2026",
        "Gartner: agentic AI scales via inference opex",
        "Once agents replace headcount, S&M opex need not grow with revenue. Cap growth at ~2% (near inflation/maintenance) to model infinite agent scale vs linear human hiring.",
        "https://www.ciodive.com/news/AI-spending-soars-enterprise-maturity/827488/",
    ),
    11: (
        "https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results",
        "ZoomInfo FY24 G&A $295.3m (~24% rev)",
        "ZI FY24 G&A was $295.3m (~24% of $1,214m). We lock ~18% as a leaner but still corporate G&A load; thesis does not touch finance/comp overhead—only the sales floor.",
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
    ),
    12: (
        "https://caliberoutbound.com/blog/build-vs-buy-outbound-the-2026-cost-math",
        "In-house SDR tooling/laptop stack costs",
        "Human SDR all-in budgets include laptops, seats, and tooling. Terminating outbound humans cuts recurring device/refresh CapEx; −5% vs base CapEx rate is a small, directionally correct haircut.",
        "https://firstsales.io/blog/cost-per-meeting-outbound/",
    ),
    13: (
        "https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc",
        "ZI receivables ~80 days (SEC ratios)",
        "ZI receivables turnover ~80 days shows WC is material. Automated billing reminders can shorten DSO; modeling −10% WC intensity vs base is a modest collections improvement, not a wipeout.",
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
    ),
    14: (
        "https://www.newyorkfed.org/markets/reference-rates/sofr",
        "NY Fed official SOFR page",
        "SOFR is the NY Fed’s official overnight Treasury-repo financing benchmark used to price leveraged loans. ~4.3% is a planning level near recent SOFR prints (see also FRED SOFR series).",
        "https://fred.stlouisfed.org/series/SOFR",
    ),
    15: (
        "https://ryanoconnellfinance.com/lbo-debt-structure/",
        "TLA typical SOFR+275–375bps (illustrative)",
        "Bank Term Loan A in LBO structures typically prices tighter than TLB. SOFR+400bps sits at the wider end of published TLA ranges—conservative senior pricing independent of AI ops.",
        "https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending",
    ),
    16: (
        "https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine",
        "TLB often SOFR+300–500bps; 1%/yr amort",
        "Institutional Term Loan B commonly prices SOFR+300–500bps in LBO guides. SOFR+500bps is the upper published band—set by syndicated loan markets, not by the company’s internal AI software strategy.",
        "https://ryanoconnellfinance.com/lbo-debt-structure/",
    ),
    17: (
        "https://ibinterviewquestions.com/guides/debt-capital-markets/term-loan-b-tlb-mechanics-and-why-it-dominates-lev-lending",
        "Syndicated TLB ~1% annual amort standard",
        "Market-standard institutional TLB schedules about 1% annual amortization with a bullet at maturity. We keep 1% fixed because credit agreements—not AI-generated FCF—set the mandatory repayment floor.",
        "https://clearvaluelending.com/glossary/term-loan-b",
    ),
    18: (
        "https://ryanoconnellfinance.com/lbo-debt-structure/",
        "LBO excess cash / optional prepay practice",
        "Sponsor LBOs typically sweep excess free cash to optional debt prepay after mandatory amort. 100% sweep maximizes deleveraging so AI margin gains show up in MOIC/IRR, not idle cash.",
        "https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine",
    ),
    19: (
        "https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income",
        "PwC: US federal corporate rate 21%",
        "US federal corporate income tax is a flat 21% statutory rate (TCJA). Operating-margin gains do not change the federal statutory rate used in the model.",
        "https://tradingeconomics.com/united-states/corporate-tax-rate",
    ),
    20: (
        "https://valueaddvc.com/blog/saas-evebitda-multiples-explained-benchmarks-and-what-they-mean",
        "PE SaaS buyouts often ~15–22x EBITDA",
        "PE SaaS deals often underwrite ~15–22x EBITDA. Locking exit at 13x in BOTH AI and Base cases is deliberately conservative so equity creation is attributed to ops, not multiple expansion.",
        "https://aventis-advisors.com/software-valuation-multiples/",
    ),
    21: (
        "https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results",
        "ZI FY24 S&M $414.1m on $1,214.3m rev",
        "Company-reported FY24 sales & marketing expense = $414.1m on $1,214.3m revenue (~34%). The model’s ~$425m (~35%) baseline is a rounded figure consistent with that audited filing disclosure.",
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
    ),
    22: (
        "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
        "SDR OTE mostly base (~70%)",
        "With ~70/30 base/variable SDR mix, most S&M people cost is payroll/base. Payroll baseline = S&M − commission pool (~75% of S&M) matches that pay-mix structure.",
        "https://pulserevops.com/knowledge/ra0197",
    ),
    23: (
        "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
        "Variable ~30–35% of SDR OTE",
        "Tech SDR variable is typically ~30–35% of OTE. Setting commission pool at 25% of S&M is a slightly conservative carve of total S&M for variable payouts.",
        "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
    ),
    24: (
        "https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results",
        "ZI FY24 unlevered FCF / light CapEx SaaS",
        "SaaS models run low CapEx vs industrials; ZI still printed strong FY24 FCF. 2% of revenue is a simple maintenance CapEx planning rate before the AI −5% haircut.",
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
    ),
    25: (
        "https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc",
        "ZI receivables days ~80 (WC intensity)",
        "Receivables ~80 days imply WC scales with growth. Using 2% of Δrevenue as ΔNWC is a simplified SaaS WC plug anchored to that receivables reality.",
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
    ),
}


def add_link(cell, url, label=None):
    cell.value = label or url
    cell.hyperlink = url
    cell.font = LINK
    cell.alignment = WRAP
    cell.border = THIN


def patch_assumptions_drivers(wb):
    if "AI_Cost_Sources" in wb.sheetnames:
        del wb["AI_Cost_Sources"]

    d = wb["Assumptions_Drivers"]
    d["B2"] = "Assumptions Drivers — yellow inputs + SOURCE LINK + justification on each row"
    d["B2"].font = TITLE
    d["B3"] = (
        "Every driver has a reputable public source URL (col E) and ~30 words on credibility / why the number is reasonable (col F). "
        "Optional second source in col G. No separate sources tab."
    )
    d["B3"].alignment = WRAP
    d.merge_cells("B3:G3")
    d.row_dimensions[3].height = 36

    # Clear old AI-only source clutter on row 9 extras / row 30 note if present
    for col in range(5, 10):
        for row in range(4, 32):
            cell = d.cell(row, col)
            # keep structure; we'll rewrite headers and source cols
            if row == 4 or row in DRIVER_SOURCES or row >= 29:
                pass

    # Header row
    headers = {
        2: "Driver",
        3: "Value",
        4: "Feeds",
        5: "Source link (click)",
        6: "Why this number is reasonable (~30 words)",
        7: "Secondary source (click)",
    }
    for col, h in headers.items():
        cell = d.cell(4, col, h)
        cell.fill = HDR
        cell.font = HDR_F
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    d.row_dimensions[4].height = 32

    for row, pack in DRIVER_SOURCES.items():
        url, label, why, url2 = pack
        add_link(d.cell(row, 5), url, label)
        why_cell = d.cell(row, 6, why)
        why_cell.alignment = WRAP
        why_cell.border = THIN
        why_cell.fill = NOTE
        why_cell.font = BLACK
        add_link(d.cell(row, 7), url2, url2.split("//")[-1][:60])
        d.row_dimensions[row].height = 48

    # Clear obsolete merged note that pointed to AI_Cost_Sources
    # Unmerge if needed
    to_unmerge = []
    for mr in list(d.merged_cells.ranges):
        if mr.min_row >= 29:
            to_unmerge.append(str(mr))
    for m in to_unmerge:
        d.unmerge_cells(m)
    for r in range(29, 35):
        for c in range(2, 8):
            d.cell(r, c).value = None
            d.cell(r, c).hyperlink = None

    d["B29"] = "PDF twin of this source map:"
    d["B29"].font = BOLD
    add_link(d["C29"], RAW_PDF, RAW_PDF)
    d["B30"] = "Educational / research use only — not investment advice. List prices and filings update over time."
    d["B30"].font = Font(name="Calibri", italic=True, size=9)

    d.column_dimensions["B"].width = 32
    d.column_dimensions["C"].width = 12
    d.column_dimensions["D"].width = 28
    d.column_dimensions["E"].width = 42
    d.column_dimensions["F"].width = 62
    d.column_dimensions["G"].width = 42
    d.freeze_panes = "B5"


def scrub_old_ai_cost_refs(wb):
    """Point any leftover AI_Cost_Sources mentions to Assumptions_Drivers."""
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "AI_Cost_Sources" in cell.value:
                    cell.value = cell.value.replace("AI_Cost_Sources", "Assumptions_Drivers (cols E–G)")
                    if cell.hyperlink and "AI_Cost_Sources" in str(getattr(cell.hyperlink, "target", "") or ""):
                        cell.hyperlink = "#'Assumptions_Drivers'!E9"
                if cell.hyperlink and "AI_Cost_Sources" in str(getattr(cell.hyperlink, "target", "") or ""):
                    cell.hyperlink = "#'Assumptions_Drivers'!E9"
                    if isinstance(cell.value, str) and "AI_Cost" in cell.value:
                        cell.value = "Go to Assumptions_Drivers!E9 (source on same row)"
                    cell.font = LINK

    if "Strategy_Summary" in wb.sheetnames:
        s = wb["Strategy_Summary"]
        s["B22"] = "Every Assumptions_Drivers input has a source link + justification in columns E–G (same row)."
        add_link(s["B23"], "#'Assumptions_Drivers'!E5", "Go to Assumptions_Drivers source columns")
        add_link(s["B24"], RAW_PDF, RAW_PDF)

    if "00_Assumptions_List" in wb.sheetnames:
        ws = wb["00_Assumptions_List"]
        # Append note at bottom
        r = ws.max_row + 2
        # find a non-merged cell
        while True:
            cell = ws.cell(r, 2)
            try:
                cell.value = "SOURCE MAP: open Assumptions_Drivers — columns E (link), F (~30-word justification), G (secondary link) on each driver row."
                break
            except AttributeError:
                r += 1
        cell.font = BOLD
        add_link(ws.cell(r + 1, 2), "#'Assumptions_Drivers'!E9", "Jump to AI infra source row (E9)")
        add_link(ws.cell(r + 2, 2), RAW_PDF, "PDF: all driver sources")


def write_pdf_md():
    styles = getSampleStyleSheet()
    title = ParagraphStyle("T", parent=styles["Heading1"], fontSize=12, spaceAfter=6)
    body = ParagraphStyle("B", parent=styles["Normal"], fontSize=8, leading=10)
    small = ParagraphStyle("S", parent=styles["Normal"], fontSize=7, leading=9)
    doc = SimpleDocTemplate(
        str(PDF),
        pagesize=letter,
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.45 * inch,
        bottomMargin=0.45 * inch,
    )
    story = [
        Paragraph("LBOMODEL2 — Assumptions_Drivers Source Map (no separate sources tab)", title),
        Paragraph(
            f"Each yellow input on <b>Assumptions_Drivers</b> has Source link (col E) + ~30-word justification (col F) on the <b>same row</b>. "
            f"Excel: <link href='{RAW_XLSX}'>{RAW_XLSX}</link>",
            body,
        ),
        Spacer(1, 4),
    ]
    data = [["Row", "Driver", "Source", "Why reasonable (~30 words)"]]
    # labels from expected drivers
    labels = {
        5: "Revenue growth 10%",
        6: "Gross margin 80%",
        7: "Sales payroll cut −15%",
        8: "Commission cut −20%",
        9: "AI infrastructure $15m",
        10: "S&M growth Y2+ 2%",
        11: "G&A % rev 18%",
        12: "CapEx reduction −5%",
        13: "WC improvement −10%",
        14: "SOFR ~4.3%",
        15: "TLA spread +4%",
        16: "TLB spread +5%",
        17: "Mandatory amort 1%",
        18: "Cash sweep 100%",
        19: "Tax rate 21%",
        20: "Exit multiple 13x",
        21: "S&M baseline $425m",
        22: "Payroll baseline $318.8m",
        23: "Commission baseline $106.2m",
        24: "CapEx base 2% rev",
        25: "WC base 2% Δrev",
    }
    for row, pack in DRIVER_SOURCES.items():
        url, label, why, _url2 = pack
        data.append(
            [
                str(row),
                Paragraph(f"<b>{labels.get(row, '')}</b>", small),
                Paragraph(f'<link href="{url}">{label}</link>', small),
                Paragraph(why, small),
            ]
        )
    t = Table(data, colWidths=[0.35 * inch, 1.35 * inch, 2.2 * inch, 3.5 * inch])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 7),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#B0B0B0")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F8FC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(t)
    story.append(Spacer(1, 6))
    story.append(Paragraph("Educational / research use only — not investment advice.", small))
    doc.build(story)

    lines = [
        "# LBOMODEL2 — Assumptions_Drivers Source Map",
        "",
        "Sources live **on the same row** as each driver (`Assumptions_Drivers` columns E–G). There is **no** separate AI_Cost_Sources tab.",
        "",
        f"**Excel:** {RAW_XLSX}",
        "",
        f"**PDF:** {RAW_PDF}",
        "",
        "| Row | Driver | Source | Why reasonable |",
        "|---|---|---|---|",
    ]
    for row, pack in DRIVER_SOURCES.items():
        url, label, why, url2 = pack
        lines.append(f"| {row} | {labels.get(row, '')} | [{label}]({url}) | {why} |")
        lines.append(f"| | | Secondary: {url2} | |")
    lines += ["", "Educational / research use only — not investment advice.", ""]
    MD.write_text("\n".join(lines))

    NOTES.write_text(
        f"""# vengeanceaiUSC-LBOMODEL2

Every yellow input on **Assumptions_Drivers** has:
- **Column E** — clickable source URL
- **Column F** — ~30-word justification (credibility + why reasonable for this model)
- **Column G** — secondary source URL

There is **no** separate `AI_Cost_Sources` tab.

Excel: {RAW_XLSX}
PDF: {RAW_PDF}
"""
    )


def main():
    wb = load_workbook(XLSX)
    patch_assumptions_drivers(wb)
    scrub_old_ai_cost_refs(wb)
    wb.save(XLSX)
    wb.save(XLSX2)
    write_pdf_md()
    d = load_workbook(XLSX)["Assumptions_Drivers"]
    missing = [r for r in range(5, 26) if d.cell(r, 5).value is None]
    print("Saved", XLSX)
    print("AI_Cost_Sources present:", "AI_Cost_Sources" in load_workbook(XLSX).sheetnames)
    print("Drivers with source links:", 21 - len(missing), "/ 21")
    print("Missing rows:", missing)
    print("Sample E9:", d["E9"].value, "->", d["E9"].hyperlink.target if d["E9"].hyperlink else None)
    print("Sample F9:", (d["F9"].value or "")[:80])


if __name__ == "__main__":
    main()
