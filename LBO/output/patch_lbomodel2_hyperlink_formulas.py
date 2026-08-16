#!/usr/bin/env python3
"""
Fix Verbatim FIND links: use Excel =HYPERLINK("url","label") formulas so links
actually appear clickable in Excel / Excel Online / Sheets.
"""

from __future__ import annotations

import shutil
from pathlib import Path
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.hyperlink import Hyperlink

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL2.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL2.xlsx"

LINK_F = Font(name="Calibri", color="0563C1", underline="single", size=9)
BLACK = Font(name="Calibri", size=9)
HDR = PatternFill("solid", fgColor="1F4E79")
HDR_F = Font(name="Calibri", color="FFFFFF", bold=True, size=9)
GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GRAY = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def frag(url: str, text: str) -> str:
    base = url.split("#")[0]
    return base + "#:~:text=" + quote(text, safe="")


def esc(s: str) -> str:
    """Escape for Excel formula string literal."""
    return s.replace('"', '""')


def hyperlink_formula(url: str, label: str) -> str:
    return f'=HYPERLINK("{esc(url)}","{esc(label)}")'


def set_hl(cell, url: str | None, label: str, fill):
    cell.alignment = WRAP
    cell.border = THIN
    # clear prior relationship hyperlink
    if cell.hyperlink:
        cell.hyperlink = None
    if url:
        cell.value = hyperlink_formula(url, label)
        cell.font = LINK_F
        cell.fill = fill
    else:
        cell.value = label
        cell.font = BLACK
        cell.fill = fill if fill else GRAY


# (display_label_with_verbatim, url)
FINDS = {
    5: [
        (
            'Verbatim: "GAAP Revenue of $1,214.3 million" → click',
            frag(
                "https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
                "GAAP Revenue of $1,214.3 million",
            ),
        ),
        (
            'Verbatim: "We generated revenue of $1,214.3 million" (10-K) → click',
            frag(
                "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
                "We generated revenue of $1,214.3 million",
            ),
        ),
        ("10% growth: N/A — MODEL CONST (no link)", None),
    ],
    6: [
        (
            'Verbatim: "The 2025 median software gross margin is 80%" → click',
            frag(
                "https://www.getaleph.com/answers/saas-gross-margin-2026",
                "The 2025 median software gross margin is 80%",
            ),
        ),
    ],
    7: [
        (
            'Verbatim: "down 18% YoY" → click',
            frag(
                "https://www.digitalapplied.com/blog/ai-sdr-statistics-2026-outbound-sales-data-points",
                "down 18% YoY",
            ),
        ),
    ],
    8: [
        (
            'Verbatim: "70/30 base/variable split" → click',
            frag(
                "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
                "70/30 base/variable split",
            ),
        ),
        (
            'Verbatim: "30–35% of total pay" → click',
            frag(
                "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
                "30–35% of total pay",
            ),
        ),
        ("20% cut: N/A — MODEL CONST (no link)", None),
    ],
    9: [
        (
            'Verbatim: "11x.ai (Alice) $5,000-10,000+" → click',
            frag(
                "https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026",
                "$5,000-10,000+",
            ),
        ),
        (
            'Verbatim: "Enterprise $10,000-$15,000+" → click',
            frag(
                "https://www.miniloop.ai/blog/11x-pricing",
                "$10,000-$15,000+",
            ),
        ),
        (
            'Verbatim: "1,513 in sales and marketing" → click',
            frag(
                "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
                "1,513 in sales and marketing",
            ),
        ),
    ],
    10: [
        (
            'Verbatim: "inflation of 2 percent over the longer run" → click',
            frag(
                "https://www.federalreserve.gov/faqs/economy_14400.htm",
                "inflation of 2 percent over the longer run",
            ),
        ),
    ],
    11: [
        (
            'Verbatim: "G&A was 18%" → click',
            frag(
                "https://blossomstreetventures.medium.com/what-should-saas-spend-on-cogs-s-m-r-d-and-g-a-the-data-from-57-public-cos-8e88eabe99ca",
                "G&A was 18%",
            ),
        ),
        (
            'Verbatim table: "MEDIAN" (G&A col 18%) → click',
            frag(
                "https://www.blossomstreetventures.com/saas-metric/2025-annual-percent-of-revenue",
                "MEDIAN",
            ),
        ),
    ],
    12: [("5% CapEx cut: N/A — MODEL CONST (no link)", None)],
    13: [("10% WC improve: N/A — MODEL CONST (no link)", None)],
    14: [
        (
            'Verbatim: "4.30" (SOFR 2024 Lowest) → click',
            frag(
                "https://www.global-rates.com/en/interest-rates/sofr/historical/2024/",
                "4.30",
            ),
        ),
    ],
    15: [
        (
            'Verbatim: "SOFR + 400-500 bps" → click',
            frag(
                "https://ctacquisitions.com/acquisition-financing/",
                "SOFR + 400-500 bps",
            ),
        ),
    ],
    16: [
        (
            'Verbatim: "SOFR + 300-500 basis points" → click',
            frag(
                "https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine",
                "SOFR + 300-500 basis points",
            ),
        ),
    ],
    17: [
        (
            'Verbatim: "minimal amortization (1% annually)" → click',
            frag(
                "https://clearvaluelending.com/glossary/term-loan-b",
                "minimal amortization (1% annually)",
            ),
        ),
    ],
    18: [
        (
            'Verbatim: "50-100% of annual excess free cash flow" → click',
            frag(
                "https://ryanoconnellfinance.com/lbo-model-fundamentals/",
                "50-100% of annual excess free cash flow",
            ),
        ),
        (
            'Verbatim: "typically modeled at 100%" → click',
            frag(
                "https://ryanoconnellfinance.com/lbo-model-fundamentals/",
                "typically modeled at 100%",
            ),
        ),
    ],
    19: [
        (
            'Verbatim: "flat rate of 21%" → click',
            frag(
                "https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income",
                "flat rate of 21%",
            ),
        ),
    ],
    20: [
        (
            'Verbatim: "12.9x" → click',
            frag(
                "https://aventis-advisors.com/saas-valuation-multiples-how-much-is-my-saas-business-worth/",
                "12.9x",
            ),
        ),
    ],
    21: [
        (
            'Verbatim: "$1,214.3" → click',
            frag(
                "https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
                "$1,214.3",
            ),
        ),
        (
            'Verbatim: "414.1" → click',
            frag(
                "https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
                "414.1",
            ),
        ),
        (
            'Verbatim: "$414.1" (10-K) → click',
            frag(
                "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
                "$414.1",
            ),
        ),
    ],
    22: [
        (
            'Verbatim: "70/30 base/variable split" → click',
            frag(
                "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
                "70/30 base/variable split",
            ),
        ),
        ("$318.8 / 75%: MODEL carve (no link)", None),
    ],
    23: [
        (
            'Verbatim: "30–35% of total pay" → click',
            frag(
                "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
                "30–35% of total pay",
            ),
        ),
        ("$106.2 / 25%: MODEL carve (no link)", None),
    ],
    24: [
        (
            'Verbatim: "Most pure-play SaaS companies report CapEx of 1–3% of revenue" → click',
            frag(
                "https://saasdb.app/learn/financials/fcf-margin/",
                "Most pure-play SaaS companies report CapEx of",
            ),
        ),
        ("2%: MODEL midpoint of 1–3% band (no link)", None),
    ],
    25: [("2% WC plug: N/A — MODEL CONST (no link)", None)],
    27: [("See FIND links on rows 14–15", None)],
    28: [("See FIND links on rows 14 & 16", None)],
}


def set_source_hl(cell, url: str):
    """Keep descriptive text, attach both HYPERLINK formula is awkward with long text —
    use relationship hyperlink AND also write a companion — better: keep value text,
    set hyperlink target via Hyperlink object with display."""
    # For E/H keep the explanatory text as display but force working hyperlink via formula
    # would wipe text. Instead use openpyxl Hyperlink properly:
    text = cell.value
    if not isinstance(text, str) or text.startswith("="):
        return
    cell.hyperlink = None
    # Excel-reliable: formula with truncated label from existing text
    short = text if len(text) <= 80 else text[:77] + "..."
    cell.value = hyperlink_formula(url, short)
    cell.font = LINK_F


def main():
    wb = load_workbook(XLSX)
    ws = wb["Assumptions_Drivers"]

    # Headers
    ws["I4"] = "Verbatim 1 (click blue link — opens source at quote)"
    ws["J4"] = "Verbatim 2 (click blue link — opens source at quote)"
    ws["K4"] = "Verbatim 3 (click blue link — opens source at quote)"
    for col in ("I", "J", "K"):
        ws[f"{col}4"].fill = HDR
        ws[f"{col}4"].font = HDR_F
        ws[f"{col}4"].alignment = WRAP
        ws.column_dimensions[col].width = 48

    ws["B2"] = (
        "Assumptions Drivers — columns I/J/K are =HYPERLINK() to the verbatim quote "
        "(click the blue text). Opens source with quote highlighted in Chrome/Edge."
    )
    ws["B3"] = (
        "If a cell says MODEL CONST there is no link. Scroll right to columns I, J, K "
        "for Verbatim hyperlinks on every sourced row."
    )

    for r, finds in FINDS.items():
        for i, col in enumerate(("I", "J", "K")):
            cell = ws[f"{col}{r}"]
            if i < len(finds):
                label, url = finds[i]
                fill = GREEN if url else YELLOW
                set_hl(cell, url, label, fill)
            else:
                set_hl(cell, None, "", GRAY)
                cell.value = None

        # Also make E (and H when present) real HYPERLINK formulas to primary finds
        if finds and finds[0][1]:
            e = ws[f"E{r}"]
            if e.value and not (isinstance(e.value, str) and e.value.startswith("=HYPERLINK")):
                # Keep longer E description as hyperlink label
                label = str(e.value)
                if len(label) > 120:
                    label = label[:117] + "..."
                set_hl(e, finds[0][1], label, GREEN if finds[0][1] else YELLOW)
            elif e.value and isinstance(e.value, str) and e.value.startswith("=HYPERLINK"):
                pass
            elif e.value:
                set_hl(e, finds[0][1], str(e.value)[:120], GREEN)

        if len(finds) > 1 and finds[1][1]:
            h = ws[f"H{r}"]
            hlab = h.value if isinstance(h.value, str) and h.value and not h.value.startswith("=") else finds[1][0]
            if isinstance(hlab, str) and hlab.startswith("=HYPERLINK"):
                pass
            else:
                set_hl(h, finds[1][1], str(hlab)[:120] if hlab else finds[1][0], GREEN)

    # Row 9: ensure K is SEC (3rd find) — already in FINDS
    # Footer
    ws["B33"] = "HOW TO USE VERBATIM LINKS:"
    ws["C33"] = (
        "Go to columns I, J, K. Blue cells are Excel HYPERLINK formulas. "
        "Click the blue verbatim text → browser opens the source and jumps to/highlights that quote. "
        "Example row 9: I9 White Space $5,000-10,000+ | J9 Miniloop $10,000-$15,000+ | K9 SEC 1,513."
    )
    ws["C33"].alignment = WRAP
    ws["C33"].font = BLACK

    # Unhide / ensure columns visible
    ws.column_dimensions["I"].hidden = False
    ws.column_dimensions["J"].hidden = False
    ws.column_dimensions["K"].hidden = False

    wb.save(XLSX)
    shutil.copy2(XLSX, XLSX2)

    # Verify
    wb2 = load_workbook(XLSX)
    ws2 = wb2["Assumptions_Drivers"]
    for col in "IJK":
        v = ws2[f"{col}9"].value
        print(f"{col}9:", v[:90] if isinstance(v, str) else v)


if __name__ == "__main__":
    main()
