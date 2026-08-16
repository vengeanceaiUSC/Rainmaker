#!/usr/bin/env python3
"""
Patch Assumptions_Drivers: only cite sources whose numbers enter the C-column EQ.
Replace prior MODEL CONST / non-EQ citations where a true EQ-equal source was found.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL2.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL2.xlsx"
MD = REPO / "LBO" / "output" / "LBOMODEL2_DRIVER_SOURCES.md"
PDF = REPO / "LBO" / "output" / "LBOMODEL2_DRIVER_SOURCES.pdf"
NOTES = REPO / "LBO" / "output" / "LBOMODEL2_NOTES.md"

RAW_XLSX = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/vengeanceaiUSC_LBOMODEL2.xlsx"
RAW_PDF = "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_DRIVER_SOURCES.pdf"

LINK = Font(name="Calibri", color="0563C1", underline="single", size=9)
BLACK = Font(name="Calibri", size=9)
GREEN = PatternFill("solid", fgColor="C6EFCE")  # sourced number in EQ
YELLOW = PatternFill("solid", fgColor="FFF2CC")  # MODEL CONST remains
WRAP = Alignment(wrap_text=True, vertical="top")

# row -> E, F, H_text, E_url, H_url, fill_E
UPDATES = {
    7: {
        "E": "IN EQ: C7=15% = midpoint of sourced SDR HC declines 12% (RevHeat/Bridge) and ~18% YoY (2026 SaaS) — click",
        "F": (
            "EQ: Payroll_Y1=318.8×(1−0.15)=$270.98m. SOURCED in EQ: 15% = midpoint of 12% "
            "(RevHeat: Bridge Group active SDR HC −12% in 18 months) and ~18% YoY US B2B SaaS SDR "
            "HC decline. SOURCED $: 318.8 from row22. NOT used as driver alone: either endpoint."
        ),
        "H": "Secondary: ~18% YoY SDR HC decline (US B2B SaaS 2026)",
        "E_url": "https://revheat.com/sdr-model-decline/",
        "H_url": "https://www.digitalapplied.com/blog/ai-sdr-statistics-2026-outbound-sales-data-points",
        "fill": "green",
    },
    9: {
        "E": "IN EQ: $5,000/mo (entry) and $10,000/mo (enterprise) AI SDR — White Space + Miniloop print both — click",
        "F": (
            "EQ: Roles=0.25×1,513=378; Low=378×$5,000×12=$22.7m; High=378×$10,000×12=$45.4m; "
            "C9=avg=$34.0m. SOURCED $/mo IN EQ: White Space lists 11x Alice $5,000–10,000+/mo; "
            "Miniloop entry $5,000 and enterprise $10,000–15,000+ (we use $10k high). "
            "SOURCED HC: 1,513 S&M (ZI 10-K, H). MODEL CONST: 0.25 outbound share. "
            "NOT in EQ: 11x official Growth $3,750/mo (below $5k low — cross-check only)."
        ),
        "H": "Secondary EQ $/mo: Miniloop 11x pricing table Entry $5k / Enterprise $10k+",
        "E_url": "https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026",
        "H_url": "https://www.miniloop.ai/blog/11x-pricing",
        "fill": "green",
        "extra_hc_note": True,
    },
    11: {
        "E": "IN EQ: C11=18% = Blossom Street 2025 public SaaS median G&A 18% of rev — click",
        "F": (
            "EQ: G&A=Rev×0.18 → Y1 1335.73×0.18=$240.43m. SOURCED in EQ: 18% = Blossom Street "
            "median G&A for SaaS IPOs-since-2017 in FY2025 (18%, 20%, 21%, 23% for 2025–2022). "
            "NOT in EQ: ZI filing G&A 295.3/1214.3=24.3% (would be Y1 $324.9m)."
        ),
        "H": None,
        "E_url": "https://blossomstreetventures.medium.com/what-should-saas-spend-on-cogs-s-m-r-d-and-g-a-the-data-from-57-public-cos-8e88eabe99ca",
        "H_url": None,
        "fill": "green",
    },
    14: {
        "E": "IN EQ: C14=4.30% = SOFR 2024 calendar-year LOW (Global-Rates / NY Fed series) — click",
        "F": (
            "EQ: rates use C14 directly (TLA=C14+C15). SOURCED in EQ: 4.30% = lowest SOFR print "
            "in 2024 (Global-Rates summary of NY Fed SOFR; Dec 2024 trough). Official series: FRED/NY Fed. "
            "NOT used: latest FRED ~3.62% (would lower interest vs this underwrite)."
        ),
        "H": "Official SOFR publisher: NY Fed",
        "E_url": "https://www.global-rates.com/en/interest-rates/sofr/historical/2024/",
        "H_url": "https://www.newyorkfed.org/markets/reference-rates/sofr",
        "fill": "green",
    },
    15: {
        "E": "IN EQ: C15=+400bps = lower bound of senior Term Loan A SOFR+400–500bps (CT Acquisitions) — click",
        "F": (
            "EQ: TLA_rate=4.30%+4.00%=8.30%. SOURCED in EQ: +400bps = floor of CT Acquisitions "
            "senior term loan A band SOFR+400–500bps (2026 LMM acquisition financing). "
            "Removed prior Ryan O’Connell 275–375bps band (did not contain 400)."
        ),
        "H": "Also: senior bank term loan SOFR+350–500 (400 in band)",
        "E_url": "https://ctacquisitions.com/acquisition-financing/",
        "H_url": "https://ctacquisitions.com/debt-financing/",
        "fill": "green",
    },
    18: {
        "E": "IN EQ: C18=100% = upper bound of LBO excess-cash-flow sweep 50–100%; models often use 100% — click",
        "F": (
            "EQ: Sweep=1.00×max(0,FCF−Mandatory). SOURCED in EQ: Ryan O’Connell LBO fundamentals — "
            "ECF sweep typically 50–100% of excess FCF; LBO models typically use 100% for max delever path. "
            "100% = upper bound of that published band. NOT claiming every credit agreement is 100%."
        ),
        "H": None,
        "E_url": "https://ryanoconnellfinance.com/lbo-model-fundamentals/",
        "H_url": None,
        "fill": "green",
    },
    20: {
        "E": "IN EQ: C20=13x ≈ Aventis private SaaS EV/EBITDA lower quartile 12.8x–12.9x (round) — click",
        "F": (
            "EQ: Exit_EV=Y5_EBITDA×13. SOURCED in EQ: Aventis Advisors private SaaS EV/EBITDA "
            "1st quartile 12.9x (table) / ~12.8x (text) → model rounds to 13.0x conservative exit. "
            "NOT in EQ: PE SaaS entry/exit bands 15–22x / 20–28x (would raise exit EV)."
        ),
        "H": None,
        "E_url": "https://aventis-advisors.com/saas-valuation-multiples-how-much-is-my-saas-business-worth/",
        "H_url": None,
        "fill": "green",
    },
}

# Still MODEL CONST — keep honesty labels (refresh if needed)
STILL_CONST = {
    5: {
        "E": "IN EQ: FY24 Rev=$1,214.3m (ZI). NOT in EQ: −2% growth. C5=10% = MODEL CONST (no source equals 10%).",
        "F": (
            "EQ: Y1_Rev=1214.3×(1+0.10)=$1,335.73m. SOURCED $ base: 1214.3 ZI FY24. "
            "MODEL CONST: 0.10. Blossom public SaaS medians ~15–18% and Aventis ‘under 10%’ guidance "
            "do not equal 10% — not cited as the driver of C5."
        ),
        "E_url": "https://www.businesswire.com/news/home/20250225694335/en/ZoomInfo-Announces-Fourth-Quarter-and-Full-Year-2024-Financial-Results",
        "fill": "yellow",
    },
    8: {
        "E": "IN EQ: ~30% variable OTE (explains commission pool); pool $106.2m; C8=20% cut = MODEL CONST",
        "F": (
            "EQ: Commission_AI=106.2×(1−0.20)=$84.96m. SOURCED: ~30% variable share (why pool exists). "
            "SOURCED $: 106.2=0.25×425. MODEL CONST: 0.20 cut — no publication prints a 20% commission cut."
        ),
        "E_url": "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
        "H_url": "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
        "H": "syncgtm.com variable OTE mix",
        "fill": "yellow",
    },
    12: {
        "E": "C12=5% CapEx cut = MODEL CONST (no sourced 5%). Device $ in SDR stacks is qualitative only.",
        "F": (
            "EQ: CapEx_AI=Rev×0.02×(1−0.05)=Rev×0.019 → Y1=$25.38m. MODEL CONST: 0.05. "
            "No source prints a 5% CapEx cut — not citing hardware articles as if they did."
        ),
        "E_url": None,
        "fill": "yellow",
    },
    13: {
        "E": "~80 DSO motivates WC but is NOT equal to C13=10%. 10% = MODEL CONST.",
        "F": (
            "EQ: ΔNWC_AI=ΔRev×0.02×0.90 → Y1 121.43×0.018=$2.19m. MODEL CONST: 0.10 improvement. "
            "80 DSO ≠ 10% improvement rate (80/365≈21.9% AR/sales stock)."
        ),
        "E_url": "https://www.readyratios.com/sec/GTM_zoominfo-technologies-inc",
        "fill": "yellow",
    },
    25: {
        "E": "C25=2% ΔNWC/ΔRev = MODEL CONST. ~80 DSO is not equal to 2%.",
        "F": (
            "EQ: ΔNWC=ΔRev×0.02 → Y1=$2.43m. MODEL CONST: 0.02. "
            "No source equals a 2% ΔNWC/ΔRev plug (stock NWC/rev metrics are different)."
        ),
        "E_url": None,
        "fill": "yellow",
    },
}


def set_link(cell, url: str | None, text: str):
    cell.value = text
    cell.font = LINK if url else BLACK
    cell.alignment = WRAP
    if url:
        cell.hyperlink = url
    elif cell.hyperlink:
        cell.hyperlink = None


def apply_row(ws, r: int, cfg: dict):
    fill = GREEN if cfg.get("fill") == "green" else YELLOW
    ws[f"E{r}"] = cfg["E"]
    ws[f"E{r}"].fill = fill
    ws[f"E{r}"].alignment = WRAP
    if cfg.get("E_url"):
        set_link(ws[f"E{r}"], cfg["E_url"], cfg["E"])
    else:
        ws[f"E{r}"].font = BLACK
        if ws[f"E{r}"].hyperlink:
            ws[f"E{r}"].hyperlink = None

    ws[f"F{r}"] = cfg["F"]
    ws[f"F{r}"].font = BLACK
    ws[f"F{r}"].alignment = WRAP

    if cfg.get("H"):
        set_link(ws[f"H{r}"], cfg.get("H_url"), cfg["H"])
    elif "H" in cfg and cfg["H"] is None:
        ws[f"H{r}"] = None
        if ws[f"H{r}"].hyperlink:
            ws[f"H{r}"].hyperlink = None


def patch_workbook(path: Path):
    wb = load_workbook(path)
    ws = wb["Assumptions_Drivers"]

    ws["B2"] = (
        "Assumptions Drivers — source listed ONLY if its number is an input to the C-column equation"
    )
    ws["B3"] = (
        "AI $34m: White Space + Miniloop print $5k and $10k/mo used in EQ + ZI 1,513 S&M HC. "
        "Newly EQ-sourced (were MODEL CONST): G&A 18%, exit ~13x, sweep 100%, TLA +400bps, "
        "SOFR 4.30% (2024 low), payroll cut 15% (mid of 12–18%). Still MODEL CONST: rev +10%, "
        "commission cut 20%, CapEx −5%, WC −10%/2%, outbound share 25%."
    )

    for r, cfg in {**STILL_CONST, **UPDATES}.items():
        apply_row(ws, r, cfg)

    # AI footer rows — dual primary sources + HC
    ws["B29"] = "AI $34m — sources that are IN the equation:"
    ws["C29"] = (
        "(1) https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026 → "
        "11x Alice $5,000–$10,000+/mo. "
        "(2) https://www.miniloop.ai/blog/11x-pricing → Entry $5,000 / Enterprise $10,000+. "
        "(3) https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm → "
        "1,513 S&M employees. "
        "(4) MODEL CONST 25% outbound share. "
        "NOT in EQ: 11x Growth $3,750/mo https://www.11x.ai/products/alice/pricing"
    )
    ws["C29"].alignment = WRAP

    # Add tertiary HC link on row 9 via note in H already Miniloop; put HC on row 30 area
    ws["B30"] = "ZI 10-K S&M headcount 1,513 (EQ input):"
    ws["C30"] = "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm"
    set_link(
        ws["C30"],
        "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
        ws["C30"].value,
    )

    ws["B31"] = "11x official Growth $3,750/mo (NOT in EQ — below $5k low):"
    ws["C31"] = "https://www.11x.ai/products/alice/pricing"
    set_link(ws["C31"], "https://www.11x.ai/products/alice/pricing", ws["C31"].value)

    ws["B32"] = "PDF map:"
    ws["C32"] = RAW_PDF
    set_link(ws["C32"], RAW_PDF, RAW_PDF)

    wb.save(path)
    print(f"Saved {path}")


def write_md():
    lines = [
        "# Sources only when IN the equation",
        "",
        f"**Excel:** {RAW_XLSX}",
        f"**PDF:** {RAW_PDF}",
        "",
        "## AI $34m — sources that feed the math",
        "1. **$5,000–$10,000+/mo** (White Space 11x Alice enterprise): https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026",
        "2. **$5,000 entry / $10,000+ enterprise** (Miniloop 11x pricing table): https://www.miniloop.ai/blog/11x-pricing",
        "3. **1,513 S&M employees**: https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
        "4. **MODEL CONST:** 25% outbound share",
        "5. NOT in EQ: 11x Growth **$3,750/mo** https://www.11x.ai/products/alice/pricing",
        "",
        "EQ: `0.25×1513=378` → `378×$5k×12=$22.7m` / `378×$10k×12=$45.4m` → avg **$34m**",
        "",
        "## Newly EQ-sourced (previously MODEL CONST / wrong band)",
        "| Driver | Value | Source number in EQ | Link |",
        "|---|---|---|---|",
        "| G&A % rev | 18% | Blossom 2025 SaaS median G&A **18%** | https://blossomstreetventures.medium.com/what-should-saas-spend-on-cogs-s-m-r-d-and-g-a-the-data-from-57-public-cos-8e88eabe99ca |",
        "| Exit EV/EBITDA | 13x | Aventis private SaaS EV/EBITDA Q1 **12.8–12.9x** → round 13 | https://aventis-advisors.com/saas-valuation-multiples-how-much-is-my-saas-business-worth/ |",
        "| Cash sweep | 100% | Ryan O’Connell ECF sweep **50–100%**; models often **100%** | https://ryanoconnellfinance.com/lbo-model-fundamentals/ |",
        "| TLA spread | +400bps | CT Acquisitions senior TLA **SOFR+400–500** (400 = floor) | https://ctacquisitions.com/acquisition-financing/ |",
        "| SOFR | 4.30% | SOFR **2024 low = 4.30%** (Global-Rates / NY Fed) | https://www.global-rates.com/en/interest-rates/sofr/historical/2024/ |",
        "| Payroll cut Y1 | 15% | Midpoint of **12%** (RevHeat/Bridge) and **~18%** YoY | https://revheat.com/sdr-model-decline/ |",
        "",
        "## Still MODEL CONST (no source equals the yellow number)",
        "- Revenue growth **10%** (ZI FY24 −2%; public SaaS medians ~15–18% ≠ 10%)",
        "- Commission cut **20%** (~30% variable OTE explains the pool, not the 20% cut)",
        "- CapEx cut **5%**",
        "- WC improvement **10%** and WC plug **2% of Δrev**",
        "- AI roles outbound share **25% of S&M HC**",
        "",
        "## Removed / not allowed as EQ sources",
        "- Salesforce $/conversation or Flex Credits (never in $34m EQ)",
        "- Ryan O’Connell TLA **275–375bps** for a **400bps** driver (band did not contain 400)",
        "- PE SaaS **15–22x** as the source of **13x** exit (replaced with Aventis lower quartile)",
        "- ZI filing G&A **24.3%** as the source of **18%** (replaced with Blossom 18%)",
        "- Latest FRED SOFR **~3.62%** as the source of **4.30%** (replaced with 2024 trough 4.30%)",
        "",
    ]
    MD.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {MD}")


def write_pdf():
    styles = getSampleStyleSheet()
    title = ParagraphStyle("t", parent=styles["Heading1"], fontSize=14, spaceAfter=8)
    body = ParagraphStyle("b", parent=styles["Normal"], fontSize=8, leading=10)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=11, spaceBefore=10, spaceAfter=4)

    story = [
        Paragraph("LBOMODEL2 — Sources only when IN the equation", title),
        Paragraph(f"Excel: <link href='{RAW_XLSX}'>{RAW_XLSX}</link>", body),
        Spacer(1, 6),
        Paragraph("AI $34m EQ sources", h2),
        Paragraph(
            "Roles=0.25×1,513=378; Low=378×$5,000×12=$22.7m; High=378×$10,000×12=$45.4m; avg=$34m. "
            "White Space prints $5,000–10,000+/mo; Miniloop prints Entry $5,000 and Enterprise $10,000+. "
            "ZI 10-K: 1,513 S&M HC. MODEL CONST: 25% outbound. NOT in EQ: 11x Growth $3,750/mo.",
            body,
        ),
        Spacer(1, 8),
        Paragraph("Newly EQ-sourced drivers", h2),
    ]

    data = [
        [
            Paragraph("<b>Driver</b>", body),
            Paragraph("<b>Value</b>", body),
            Paragraph("<b>EQ source number</b>", body),
        ],
        [
            Paragraph("G&A % rev", body),
            Paragraph("18%", body),
            Paragraph("Blossom 2025 SaaS median G&A 18%", body),
        ],
        [
            Paragraph("Exit multiple", body),
            Paragraph("13x", body),
            Paragraph("Aventis EV/EBITDA Q1 12.8–12.9x → 13", body),
        ],
        [
            Paragraph("Cash sweep", body),
            Paragraph("100%", body),
            Paragraph("Ryan O’Connell ECF 50–100%; models often 100%", body),
        ],
        [
            Paragraph("TLA spread", body),
            Paragraph("+400bps", body),
            Paragraph("CT Acquisitions TLA SOFR+400–500 (floor)", body),
        ],
        [
            Paragraph("SOFR", body),
            Paragraph("4.30%", body),
            Paragraph("SOFR 2024 calendar low 4.30%", body),
        ],
        [
            Paragraph("Payroll cut Y1", body),
            Paragraph("15%", body),
            Paragraph("Midpoint of 12% and ~18% SDR HC declines", body),
        ],
    ]
    t = Table(data, colWidths=[1.4 * inch, 0.8 * inch, 4.5 * inch])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(Paragraph("Still MODEL CONST (no EQ-equal source)", h2))
    story.append(
        Paragraph(
            "Rev growth 10%; commission cut 20%; CapEx cut 5%; WC improve 10%; WC plug 2% of Δrev; "
            "AI outbound share 25% of S&M HC.",
            body,
        )
    )
    story.append(Spacer(1, 8))
    story.append(Paragraph("Removed non-EQ citations", h2))
    story.append(
        Paragraph(
            "Salesforce unit prices; TLA 275–375bps for a 400bps driver; PE 15–22x for 13x exit; "
            "ZI filing 24.3% G&A for 18%; latest FRED ~3.62% for a 4.30% SOFR input.",
            body,
        )
    )

    doc = SimpleDocTemplate(str(PDF), pagesize=letter, leftMargin=0.6 * inch, rightMargin=0.6 * inch)
    doc.build(story)
    print(f"Wrote {PDF}")


def patch_notes():
    if not NOTES.exists():
        return
    text = NOTES.read_text(encoding="utf-8")
    block = (
        "\n\n## Source honesty pass (EQ-only citations)\n"
        "- AI $5–10k/mo dual-sourced: White Space + Miniloop (both print $5k and $10k); "
        "official 11x Growth $3,750 NOT in EQ.\n"
        "- Newly EQ-sourced: G&A 18% (Blossom 2025 median), exit 13x≈Aventis Q1 12.8–12.9x, "
        "sweep 100% (Ryan O’Connell 50–100% upper / model practice), TLA +400 (CT Acquisitions "
        "SOFR+400–500 floor), SOFR 4.30% (2024 low), payroll cut 15% (mid of 12–18%).\n"
        "- Still MODEL CONST: rev +10%, commission cut 20%, CapEx −5%, WC −10%/2%, outbound 25%.\n"
    )
    marker = "## Source honesty pass (EQ-only citations)"
    if marker in text:
        # replace from marker to end or next ## after? just append once uniquely
        start = text.index(marker)
        # find next ## after start+2
        rest = text[start + 3 :]
        nxt = rest.find("\n## ")
        if nxt >= 0:
            text = text[:start] + block.strip() + "\n" + rest[nxt:]
        else:
            text = text[:start] + block.strip() + "\n"
    else:
        text = text.rstrip() + block
    NOTES.write_text(text, encoding="utf-8")
    print(f"Updated {NOTES}")


def main():
    patch_workbook(XLSX)
    if XLSX2.exists() or True:
        XLSX2.parent.mkdir(parents=True, exist_ok=True)
        # copy by re-patching same logic on copy
        import shutil

        shutil.copy2(XLSX, XLSX2)
        print(f"Copied to {XLSX2}")
    write_md()
    write_pdf()
    patch_notes()


if __name__ == "__main__":
    main()
