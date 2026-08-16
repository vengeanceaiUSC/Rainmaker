#!/usr/bin/env python3
"""
Make Verbatim Ctrl+F tips clickable via Chrome Scroll-to-Text fragments (#:~:text=).
Clicking opens the source page with that quote already highlighted (browser find).
"""

from __future__ import annotations

import shutil
from pathlib import Path
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL2.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL2.xlsx"
MD = REPO / "LBO" / "output" / "LBOMODEL2_DRIVER_SOURCES.md"

LINK = Font(name="Calibri", color="0563C1", underline="single", size=9)
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


def text_frag(url: str, *parts: str) -> str:
    """Build Chrome Scroll-to-Text URL. parts are exact page substrings to highlight."""
    # Strip existing fragments
    base = url.split("#")[0]
    encoded = [quote(p, safe="") for p in parts if p]
    if not encoded:
        return base
    if len(encoded) == 1:
        frag = f"#:~:text={encoded[0]}"
    else:
        # start,end form
        frag = f"#:~:text={encoded[0]},{encoded[1]}"
    return base + frag


def set_find_cell(cell, label: str | None, url: str | None, fill):
    cell.alignment = WRAP
    cell.border = THIN
    if label and url:
        cell.value = label
        cell.font = LINK
        cell.hyperlink = url
        cell.fill = fill
    elif label:
        cell.value = label
        cell.font = BLACK
        cell.fill = GRAY
        if cell.hyperlink:
            cell.hyperlink = None
    else:
        cell.value = None
        cell.fill = GRAY
        if cell.hyperlink:
            cell.hyperlink = None


# row -> list of (label, url_with_fragment) up to 3 finds; also update E/H primary urls
# Labels shown in I/J/K; clicking opens doc with highlight
FINDS = {
    5: [
        (
            'CLICK → FIND "$1,214.3"',
            text_frag(
                "https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
                "GAAP Revenue of $1,214.3 million",
            ),
        ),
        (
            'CLICK → FIND "$1,214.3" (10-K)',
            text_frag(
                "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
                "We generated revenue of $1,214.3 million",
            ),
        ),
        ("10% growth: N/A MODEL CONST", None),
    ],
    6: [
        (
            'CLICK → FIND "median software gross margin is 80"',
            text_frag(
                "https://www.getaleph.com/answers/saas-gross-margin-2026",
                "The 2025 median software gross margin is 80%",
            ),
        ),
    ],
    7: [
        (
            'CLICK → FIND "down 18%"',
            text_frag(
                "https://www.digitalapplied.com/blog/ai-sdr-statistics-2026-outbound-sales-data-points",
                "down 18% YoY",
            ),
        ),
    ],
    8: [
        (
            'CLICK → FIND "70/30"',
            text_frag(
                "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
                "70/30 base/variable split",
            ),
        ),
        (
            'CLICK → FIND "30–35%"',
            text_frag(
                "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
                "30–35% of total pay",
            ),
        ),
        ("20% cut: N/A MODEL CONST", None),
    ],
    9: [
        (
            'CLICK → FIND "$5,000-10,000+" (White Space)',
            text_frag(
                "https://www.whitespacesolutions.ai/content/ai-sdr-pricing-guide-2026",
                "$5,000-10,000+",
            ),
        ),
        (
            'CLICK → FIND "$10,000-$15,000+" (Miniloop)',
            text_frag(
                "https://www.miniloop.ai/blog/11x-pricing",
                "$10,000-$15,000+",
            ),
        ),
        (
            'CLICK → FIND "1,513 in sales and marketing" (SEC)',
            text_frag(
                "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
                "1,513 in sales and marketing",
            ),
        ),
    ],
    10: [
        (
            'CLICK → FIND "2 percent"',
            text_frag(
                "https://www.federalreserve.gov/faqs/economy_14400.htm",
                "inflation of 2 percent over the longer run",
            ),
        ),
    ],
    11: [
        (
            'CLICK → FIND "G&A was 18%"',
            text_frag(
                "https://blossomstreetventures.medium.com/what-should-saas-spend-on-cogs-s-m-r-d-and-g-a-the-data-from-57-public-cos-8e88eabe99ca",
                "G&A was 18%",
            ),
        ),
        (
            'CLICK → FIND table MEDIAN 18% G&A',
            text_frag(
                "https://www.blossomstreetventures.com/saas-metric/2025-annual-percent-of-revenue",
                "MEDIAN",
            ),
        ),
    ],
    12: [("5% CapEx cut: N/A MODEL CONST", None)],
    13: [("10% WC improve: N/A MODEL CONST", None)],
    14: [
        (
            'CLICK → FIND "4.30"',
            text_frag(
                "https://www.global-rates.com/en/interest-rates/sofr/historical/2024/",
                "4.30",
            ),
        ),
    ],
    15: [
        (
            'CLICK → FIND "400-500"',
            text_frag(
                "https://ctacquisitions.com/acquisition-financing/",
                "SOFR + 400-500 bps",
            ),
        ),
    ],
    16: [
        (
            'CLICK → FIND "300-500"',
            text_frag(
                "https://ibinterviewquestions.com/guides/valuation-investment-banking/lbo-debt-structures-senior-subordinated-mezzanine",
                "SOFR + 300-500 basis points",
            ),
        ),
    ],
    17: [
        (
            'CLICK → FIND "1% annually"',
            text_frag(
                "https://clearvaluelending.com/glossary/term-loan-b",
                "minimal amortization (1% annually)",
            ),
        ),
    ],
    18: [
        (
            'CLICK → FIND "50-100%"',
            text_frag(
                "https://ryanoconnellfinance.com/lbo-model-fundamentals/",
                "50-100% of annual excess free cash flow",
            ),
        ),
        (
            'CLICK → FIND "modeled at 100%"',
            text_frag(
                "https://ryanoconnellfinance.com/lbo-model-fundamentals/",
                "typically modeled at 100%",
            ),
        ),
    ],
    19: [
        (
            'CLICK → FIND "flat rate of 21%"',
            text_frag(
                "https://taxsummaries.pwc.com/united-states/corporate/taxes-on-corporate-income",
                "flat rate of 21%",
            ),
        ),
    ],
    20: [
        (
            'CLICK → FIND "12.9x"',
            text_frag(
                "https://aventis-advisors.com/saas-valuation-multiples-how-much-is-my-saas-business-worth/",
                "12.9x",
            ),
        ),
    ],
    21: [
        (
            'CLICK → FIND "$1,214.3"',
            text_frag(
                "https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
                "$1,214.3",
            ),
        ),
        (
            'CLICK → FIND "414.1"',
            text_frag(
                "https://markets.financialcontent.com/stocks/article/bizwire-2025-2-25-zoominfo-announces-fourth-quarter-and-full-year-2024-financial-results",
                "414.1",
            ),
        ),
        (
            'CLICK → FIND "414.1" (10-K)',
            text_frag(
                "https://www.sec.gov/Archives/edgar/data/1794515/000179451525000045/zi-20241231.htm",
                "$414.1",
            ),
        ),
    ],
    22: [
        (
            'CLICK → FIND "70/30"',
            text_frag(
                "https://syncgtm.com/blog/how-to-pay-sales-development-rep",
                "70/30 base/variable split",
            ),
        ),
        ("$318.8 / 75%: MODEL carve", None),
    ],
    23: [
        (
            'CLICK → FIND "30–35%"',
            text_frag(
                "https://syncgtm.com/blog/how-much-do-sales-development-representatives-make-at-tech-companies",
                "30–35% of total pay",
            ),
        ),
        ("$106.2 / 25%: MODEL carve", None),
    ],
    24: [
        (
            'CLICK → FIND "CapEx of 1" (en-dash 1–3%)',
            text_frag(
                "https://saasdb.app/learn/financials/fcf-margin/",
                "Most pure-play SaaS companies report CapEx of",
            ),
        ),
        ("2%: MODEL midpoint of 1–3% band", None),
    ],
    25: [("2% WC plug: N/A MODEL CONST", None)],
    27: [("Uses rows 14–15 FIND links", None)],
    28: [("Uses rows 14 & 16 FIND links", None)],
}

# Also update E/H hyperlinks to primary find fragment
E_FRAG = {
    5: FINDS[5][0][1],
    6: FINDS[6][0][1],
    7: FINDS[7][0][1],
    8: FINDS[8][0][1],
    9: FINDS[9][0][1],
    10: FINDS[10][0][1],
    11: FINDS[11][0][1],
    14: FINDS[14][0][1],
    15: FINDS[15][0][1],
    16: FINDS[16][0][1],
    17: FINDS[17][0][1],
    18: FINDS[18][0][1],
    19: FINDS[19][0][1],
    20: FINDS[20][0][1],
    21: FINDS[21][0][1],
    22: FINDS[22][0][1],
    23: FINDS[23][0][1],
    24: FINDS[24][0][1],
}
H_FRAG = {
    5: FINDS[5][1][1],
    8: FINDS[8][1][1],
    9: FINDS[9][1][1],
    11: FINDS[11][1][1],
    18: FINDS[18][1][1],
    21: FINDS[21][1][1],
}


def main():
    wb = load_workbook(XLSX)
    ws = wb["Assumptions_Drivers"]

    ws["I4"] = "FIND 1 — click opens page with quote highlighted"
    ws["J4"] = "FIND 2 — click opens page with quote highlighted"
    ws["K4"] = "FIND 3 — click opens page with quote highlighted"
    for col in ("I", "J", "K"):
        ws[f"{col}4"].fill = HDR
        ws[f"{col}4"].font = HDR_F
        ws[f"{col}4"].alignment = WRAP
        ws.column_dimensions[col].width = 42

    ws["B3"] = (
        "FIND columns I/J/K: click the blue link — Chrome/Edge opens the source and highlights "
        "the quote (#:~:text= fragment = browser Find). Works best in Chrome or Edge."
    )

    for r, finds in FINDS.items():
        fill = YELLOW if any(u is None for _, u in finds) and r in (5, 8, 12, 13, 22, 23, 25) else GREEN
        # greener if at least one real find
        if any(u for _, u in finds):
            fill = GREEN if r not in (12, 13, 25) else YELLOW
        if r in (5, 8, 12, 13, 22, 23, 25):
            # mixed / const
            fill = YELLOW if r in (12, 13, 25) else GREEN

        for i, col in enumerate(("I", "J", "K")):
            if i < len(finds):
                label, url = finds[i]
                set_find_cell(ws[f"{col}{r}"], label, url, GREEN if url else YELLOW)
            else:
                set_find_cell(ws[f"{col}{r}"], None, None, GRAY)

        # Point Source E / Secondary H at fragment URLs too
        if r in E_FRAG and E_FRAG[r] and ws[f"E{r}"].value:
            ws[f"E{r}"].hyperlink = E_FRAG[r]
            ws[f"E{r}"].font = LINK
        if r in H_FRAG and H_FRAG[r]:
            if ws[f"H{r}"].value:
                ws[f"H{r}"].hyperlink = H_FRAG[r]
                ws[f"H{r}"].font = LINK
            else:
                # ensure H has something clickable for secondary find when needed
                pass

    # Row 9: also put SEC find on a clear H if H is Miniloop — keep H=Miniloop fragment, K=SEC
    # Footer note
    ws["B33"] = "How FIND links work:"
    ws["C33"] = (
        "Blue FIND cells use Chrome Scroll-to-Text (#:~:text=...). Click → browser opens the URL "
        "and highlights the quote (same effect as Ctrl+F landing on that string). Use Chrome or Edge."
    )
    ws["C33"].alignment = WRAP

    # Example row 9 clarification
    ws["B34"] = "AI $34m FIND example (row 9):"
    ws["C34"] = (
        "I → White Space $5,000-10,000+ highlighted | "
        "J → Miniloop $10,000-$15,000+ highlighted | "
        "K → SEC 1,513 in sales and marketing highlighted"
    )
    ws["C34"].alignment = WRAP

    wb.save(XLSX)
    shutil.copy2(XLSX, XLSX2)

    # MD update
    lines = [
        "# Clickable FIND links (Scroll-to-Text)",
        "",
        "Columns **I / J / K** on Assumptions_Drivers are blue hyperlinks.",
        "Each uses Chrome `#:~:text=` so the page opens with that quote **already highlighted** (browser Find).",
        "",
        "## Row 9 — AI $34m (your example)",
        f"1. {FINDS[9][0][0]}",
        f"   {FINDS[9][0][1]}",
        f"2. {FINDS[9][1][0]}",
        f"   {FINDS[9][1][1]}",
        f"3. {FINDS[9][2][0]}",
        f"   {FINDS[9][2][1]}",
        "",
        "Use **Chrome or Edge**. Safari/Firefox may ignore text fragments.",
        "",
    ]
    MD.write_text(MD.read_text(encoding="utf-8") + "\n\n" + "\n".join(lines), encoding="utf-8")
    print("Patched FIND fragment links")
    print("R9 I:", FINDS[9][0][1][:100])
    print("R9 J:", FINDS[9][1][1][:100])
    print("R9 K:", FINDS[9][2][1][:100])


if __name__ == "__main__":
    main()
