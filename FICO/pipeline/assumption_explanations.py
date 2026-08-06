"""Plain-English explanations for every 3-statement assumption row (MODEL11).

SOURCE fields always include a clickable URL (col U) so users can open the filing/data.
Every assumption also links to the downloadable Assumptions List PDF (no charts).
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import List, Tuple

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .model11_assumptions import (
    ASSUMPTIONS_PDF_FILENAME,
    ASSUMPTIONS_PDF_URL,
    ASSUMPTIONS_PDF_VIEW_URL,
    BUYBACK_RUNRATE_000s,
    CAPEX_METHODOLOGY_NOTE,
    CAPEX_PCT_REVENUE,
    CASH_TAX_METHODOLOGY_NOTE,
    CASH_TAX_RATE,
    DA_METHODOLOGY_NOTE,
    DA_PCT_REVENUE,
    DEBT_NET_RUNRATE_000s,
    DEFERRED_METHODOLOGY_NOTE,
    DILUTION_METHODOLOGY_NOTE,
    DSO_B2B,
    DSO_B2C,
    DSO_ONPREM,
    DSO_PROF_SVCS,
    DSO_SAAS,
    EXIT_EV_EBITDA,
    EXIT_EV_EBITDA_BEAR,
    EXIT_EV_EBITDA_BULL,
    BS_METHODOLOGY_NOTE,
    FINANCING_METHODOLOGY_NOTE,
    FY25_B2B_SCORES_000s,
    FY25_B2C_000s,
    FY25_ONPREM_000s,
    FY25_PLATFORM_ARR_000s,
    FY25_PROF_SERVICES_000s,
    FY25_REVENUE_000s,
    FY25_SAAS_000s,
    MIX_B2B,
    MIX_B2C,
    MIX_ONPREM,
    MIX_PROF_SVCS,
    MIX_SAAS,
    MODEL_NAME as _MODEL_NAME,
    PEER_EV_EBITDA,
    PEER_MEDIAN_EV_EBITDA,
    SBC_METHODOLOGY_NOTE,
    SBC_PCT_REVENUE,
    SEGMENT_METHODOLOGY_NOTE,
    URL_10K_HTM_IR,
    URL_PEER_COMPS,
    URL_SEC_FILINGS_IR,
    WC_METHODOLOGY_NOTE,
    blended_cogs_pct,
    blended_dso,
)
from .named_range_map import SHEET_3S, SHEET_DCF
from .wacc import MODEL3_WACC

_FORECAST_COLS = ("J", "K", "L", "M", "N")
_COMMENT_AUTHOR = _MODEL_NAME

# Canonical source URLs (clickable in Excel col U)
URL_10K = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"
URL_10Q = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000016/fico-20250331.htm"
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
        f"Excel equation: =$J$122 blended COGS from segment GMs "
        f"(≈ {blended_cogs_pct():.2%}; calibrated near FY25 ~17.8%).",
        "MODEL11: COGS reflects segment margin mix — B2B Scores high GM, "
        "SaaS/on-prem high GM, Professional Services much lower GM (~28%). "
        "Mix % offset within Total Revenue (not a second revenue stack).",
        "SEC 10-K FY2025 — Note 9 / MD&A disaggregated revenue + Cost of revenues",
        URL_10K,
    ),
    (
        9,
        "SG&A % of Revenue",
        "Operating opex (SG&A) as % of sales.",
        "Equation: MAX(15%, (I28 − 10,922)/I24 − 75bps × year). "
        "Strips FY25 restructuring; then −0.75% of sales each year.",
        "MODEL11: restore software operating leverage. Floor 15%; grind −75 bps/yr "
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
        "D&A % of Revenue",
        "Total depreciation & amortization as % of sales (software-industry driver).",
        f"Yellow POLICY input: flat {DA_PCT_REVENUE:.2%} (3yr FY23–25 avg). "
        "Forecast DA$ = Revenue × DA%.",
        f"MODEL11: {DA_METHODOLOGY_NOTE}",
        "SEC companyfacts — DepreciationDepletionAndAmortization / Revenue",
        URL_FACTS,
    ),
    (
        12,
        "Interest % of Debt Open",
        "Interest expense as % of opening debt balance.",
        "Excel equation: I101/I98 (FY25 interest ÷ FY25 opening debt in schedule).",
        "Approximates average coupon / cost of debt on the book. Debt stock follows "
        "the net debt issuance assumption (row 19).",
        "SEC 10-K FY2025 — Interest expense; see also 8-K 6.250% notes",
        URL_10K,
    ),
    (
        13,
        "Book Tax Rate (% of EBT)",
        "Effective book tax rate applied to forecast EBT → Net Income.",
        "Excel equation: I35/I33 (FY25 tax ÷ simplified FY25 EBT in this template).",
        f"Book rate (~19%) remains on the income statement. "
        f"DCF unlevered cash taxes use a separate cash tax rate "
        f"({CASH_TAX_RATE:.2%} = 3yr IncomeTaxesPaidNet/EBT). "
        f"{CASH_TAX_METHODOLOGY_NOTE}",
        "SEC 10-K FY2025 — tax expense; cash taxes from IncomeTaxesPaidNet",
        URL_FACTS,
    ),
    (
        15,
        "DSO — Accounts Receivable (Days)",
        "Days Sales Outstanding. AR = Revenue × DSO / 365.",
        f"Excel equation: ROUND($J$115,1) = blended segment DSO ≈ {blended_dso():.1f} days "
        f"(SaaS {DSO_SAAS:.0f}d / B2C {DSO_B2C:.0f}d / B2B {DSO_B2B:.0f}d / "
        f"PS {DSO_PROF_SVCS:.0f}d / On-prem {DSO_ONPREM:.0f}d × FY25 mix).",
        f"MODEL11: {WC_METHODOLOGY_NOTE} "
        f"{SEGMENT_METHODOLOGY_NOTE} "
        "AR is organic from blended collections days — no top-down NWC% AR plug. "
        f"{DEFERRED_METHODOLOGY_NOTE}",
        "SEC 10-K FY2025 — Note 9 disaggregated revenue; payment terms 30–60 days",
        URL_10K,
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
        "MODEL11: DPO is an explicit WC driver paired with DSO. "
        "NWC = AR + Inventory − AP − Deferred (output); ΔNWC = NWCt − NWCt−1.",
        "SEC 10-K FY2025 — Accounts payable",
        URL_10K,
    ),
    (
        18,
        "CapEx % of Revenue",
        "Capital investment (PPE + capitalized software) as % of sales.",
        f"Yellow POLICY input: flat {CAPEX_PCT_REVENUE:.2%} "
        "(3yr avg of PPE purchases + PaymentsToDevelopSoftware / Revenue).",
        f"MODEL11: {CAPEX_METHODOLOGY_NOTE}",
        "SEC companyfacts — PPE purchases + PaymentsToDevelopSoftware",
        URL_FACTS,
    ),
    (
        19,
        "Debt Issuance (Repayment)",
        "New borrowing (+) or repayment (−) in the forecast ($000s).",
        f"Yellow POLICY input: {DEBT_NET_RUNRATE_000s:,.0f} each year "
        "(3yr avg senior-note proceeds − line-of-credit repayments).",
        f"MODEL11: {FINANCING_METHODOLOGY_NOTE} Financing is excluded from FCFF.",
        "SEC companyfacts — ProceedsFromIssuanceOfSeniorLongTermDebt / RepaymentsOfLinesOfCredit",
        URL_FACTS,
    ),
    (
        20,
        "Equity Issued (Repaid)",
        "New equity issued (+) or buybacks (−) in the forecast BS/CF.",
        "Equation: −(CFO − CapEx) − DebtIssuance so residual levered FCF funds buybacks "
        f"(hist 3yr avg buybacks ≈ ${BUYBACK_RUNRATE_000s/1000:.0f}M/yr).",
        "Restores financing realism so cash does not artificially stockpile. "
        "Buybacks are financing — excluded from FCFF. "
        f"DCF $/share uses SBC-diluted shares (row 16), not a static count. "
        f"{DILUTION_METHODOLOGY_NOTE}",
        "SEC companyfacts — PaymentsForRepurchaseOfCommonStock",
        URL_FACTS,
    ),
    (
        21,
        "SBC % of Revenue",
        "Stock-based compensation as % of sales; non-cash add-back in CFO and FCFF.",
        f"Yellow POLICY input: flat {SBC_PCT_REVENUE:.2%} "
        "(3yr FY23–25 avg ShareBasedCompensation / Revenue).",
        f"MODEL11: {SBC_METHODOLOGY_NOTE} "
        "SBC$ (3S row 64) also increases DCF diluted shares each year "
        "(ΔShares = SBC$000s / Price).",
        "SEC 10-K cash flow — ShareBasedCompensation",
        URL_FACTS,
    ),
    (
        104,
        "Mix % — SaaS / Platform software",
        "Share of Total Revenue from SaaS / cloud software (incl. FICO Platform cloud).",
        f"Yellow POLICY = {MIX_SAAS:.2%} "
        f"(FY25 SaaS ${FY25_SAAS_000s:,.0f}k / Total ${FY25_REVENUE_000s:,.0f}k). "
        f"Platform ARR KPI was ${FY25_PLATFORM_ARR_000s/1000:.1f}M (35% of software ARR) "
        "— ARR is not added on top of IS revenue.",
        "SaaS is billed largely annually in advance → primary Deferred Revenue driver "
        "with on-prem maintenance. 30–60 day payment terms (model uses 45-day DSO). "
        "Mix offsets within Total Rev — does not create a second revenue line.",
        "SEC 10-K FY2025 — Software SaaS disaggregation + Platform ARR MD&A",
        URL_10K,
    ),
    (
        105,
        "Mix % — B2C Subscriptions (myFICO)",
        "Share of Total Revenue from B2C scoring / myFICO.com subscriptions.",
        f"Yellow POLICY = {MIX_B2C:.2%} "
        f"(FY25 B2C Scores ${FY25_B2C_000s:,.0f}k / Total ${FY25_REVENUE_000s:,.0f}k).",
        "Near-zero DSO (card-settled consumer subscriptions). Minimal Deferred Revenue "
        "vs annual SaaS invoices. High incremental margin; WC-light cash conversion.",
        "SEC 10-K FY2025 — Scores segment B2C / myFICO disaggregation",
        URL_10K,
    ),
    (
        106,
        "Mix % — B2B Scores",
        "Share of Total Revenue from B2B scoring (mortgage, auto, card via CRAs).",
        f"Yellow POLICY = {MIX_B2B:.2%} "
        f"(FY25 B2B Scores ${FY25_B2B_SCORES_000s:,.0f}k / Total ${FY25_REVENUE_000s:,.0f}k).",
        "Transactional volume through Experian/TransUnion/Equifax; ~30-day DSO. "
        "Does not create SaaS-style Deferred Revenue (usage-based recognition).",
        "SEC 10-K FY2025 — Scores segment B2B disaggregation",
        URL_10K,
    ),
    (
        107,
        "Mix % — Professional Services",
        "Share of Total Revenue from implementation / consulting / training fees.",
        f"Yellow POLICY = {MIX_PROF_SVCS:.2%} "
        f"(FY25 PS ${FY25_PROF_SERVICES_000s:,.0f}k / Total ${FY25_REVENUE_000s:,.0f}k).",
        "Upfront integration fees convert cash quickly (low DSO) but carry significantly "
        "lower gross margin than software/scores — pulls blended COGS up vs pure SaaS.",
        "SEC 10-K FY2025 — Software segment Professional services line",
        URL_10K,
    ),
    (
        108,
        "Mix % — On-Premises Software",
        "Share of Total Revenue from on-premises license + maintenance software.",
        f"Yellow POLICY = {MIX_ONPREM:.2%} "
        f"(FY25 on-prem ${FY25_ONPREM_000s:,.0f}k / Total ${FY25_REVENUE_000s:,.0f}k).",
        "With SaaS mix, drives Deferred Revenue (maintenance billed in advance). "
        "Same 45-day DSO policy as SaaS. Completes mix identity to 100% of Total Rev.",
        "SEC 10-K FY2025 — Software on-premises vs SaaS deployment table",
        URL_10K,
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
            "assumption": "Cash Tax Rate (unlevered)",
            "what_it_is": "Cash tax rate used for FCFF unlevered taxes and after-tax cost of debt.",
            "how_set_in_model": (
                f"Yellow POLICY = {CASH_TAX_RATE:.2%} "
                "(3yr avg IncomeTaxesPaidNet ÷ EBT). NOT linked to book tax J13."
            ),
            "why_this_choice": CASH_TAX_METHODOLOGY_NOTE,
            "source": "SEC companyfacts — IncomeTaxesPaidNet / EBT",
            "source_url": URL_FACTS,
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
            "how_set_in_model": "Latest 10-Q bridge totals (more current than FY25 3S balances).",
            "why_this_choice": "Bridge should reflect the latest capital structure; 3S FY25 shown as reference.",
            "source": "SEC 10-Q — Fair Isaac Corp",
            "source_url": URL_10Q,
        },
        {
            "row": "DCF-D12/I16",
            "assumption": "Shares Outstanding (SBC dilution)",
            "what_it_is": (
                "Starting diluted shares (D12) roll forward each year by "
                "SBC$ ÷ share price; Equity Value/Share uses Year-5 diluted shares (I16)."
            ),
            "how_set_in_model": (
                "D16=$D$12; E16:I16 = prior + 3S!{J–N}64 / $D$11; D37=$D$35/$I$16."
            ),
            "why_this_choice": DILUTION_METHODOLOGY_NOTE,
            "source": "3S SBC row 64 (Rev × 8.25%); SEC ShareBasedCompensation",
            "source_url": URL_FACTS,
        },
        {
            "row": "DCF-E25",
            "assumption": "ΔNWC (includes Deferred Revenue)",
            "what_it_is": (
                "Change in NWC = AR + Inventory − AP − Deferred Revenue. "
                "Linked live from 3S row 90 into FCFF."
            ),
            "how_set_in_model": "DCF E25:I25 = 3S!J90:N90 (organic WC schedule).",
            "why_this_choice": DEFERRED_METHODOLOGY_NOTE,
            "source": "SEC 10-K FY2025 — Deferred revenue / contract liabilities",
            "source_url": URL_10K,
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
        Paragraph(SEGMENT_METHODOLOGY_NOTE, body),
        Paragraph(DEFERRED_METHODOLOGY_NOTE, body),
        Paragraph(DA_METHODOLOGY_NOTE, body),
        Paragraph(SBC_METHODOLOGY_NOTE, body),
        Paragraph(DILUTION_METHODOLOGY_NOTE, body),
        Paragraph(CAPEX_METHODOLOGY_NOTE, body),
        Paragraph(CASH_TAX_METHODOLOGY_NOTE, body),
        Paragraph(FINANCING_METHODOLOGY_NOTE, body),
        Paragraph(BS_METHODOLOGY_NOTE, body),
        Paragraph(
            f"Primary filings hub: <link href='{URL_SEC_FILINGS_IR}' color='blue'>"
            f"{URL_SEC_FILINGS_IR}</link> — segmentation derived from the FY2025 10-K "
            f"disaggregated revenue notes "
            f"(<link href='{URL_10K}' color='blue'>{URL_10K}</link>; "
            f"IR copy <link href='{URL_10K_HTM_IR}' color='blue'>{URL_10K_HTM_IR}</link>).",
            small,
        ),
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
    """Write MODEL11_*_ASSUMPTIONS_EXPLAINED.csv, TXT, and PDF list (no charts)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"MODEL11_{ticker}_ASSUMPTIONS_EXPLAINED.csv"
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

    txt_path = out_dir / f"MODEL11_{ticker}_ASSUMPTIONS_EXPLAINED.txt"
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
        SEGMENT_METHODOLOGY_NOTE,
        DEFERRED_METHODOLOGY_NOTE,
        DA_METHODOLOGY_NOTE,
        SBC_METHODOLOGY_NOTE,
        DILUTION_METHODOLOGY_NOTE,
        CAPEX_METHODOLOGY_NOTE,
        CASH_TAX_METHODOLOGY_NOTE,
        FINANCING_METHODOLOGY_NOTE,
        BS_METHODOLOGY_NOTE,
        f"SEC filings hub: {URL_SEC_FILINGS_IR}",
        f"FY2025 10-K (HTML): {URL_10K}",
        f"FY2025 10-K (IR): {URL_10K_HTM_IR}",
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
    md_path = out_dir / f"MODEL11_{ticker}_ASSUMPTIONS_LIST.md"
    md = [
        f"# {_MODEL_NAME} — Assumptions List",
        "",
        f"**PDF download:** {ASSUMPTIONS_PDF_URL}",
        "",
        "Plain list only (no charts). Paste into Google Docs via File → Open if needed.",
        "",
        f"> {WC_METHODOLOGY_NOTE}",
        "",
        f"> {SEGMENT_METHODOLOGY_NOTE}",
        "",
        f"> {DEFERRED_METHODOLOGY_NOTE}",
        "",
        f"> {DA_METHODOLOGY_NOTE}",
        "",
        f"> {SBC_METHODOLOGY_NOTE}",
        "",
        f"> {DILUTION_METHODOLOGY_NOTE}",
        "",
        f"> {CAPEX_METHODOLOGY_NOTE}",
        "",
        f"> {CASH_TAX_METHODOLOGY_NOTE}",
        "",
        f"> {FINANCING_METHODOLOGY_NOTE}",
        "",
        f"> {BS_METHODOLOGY_NOTE}",
        "",
        f"SEC filings hub: {URL_SEC_FILINGS_IR}",
        "",
        f"FY2025 10-K: {URL_10K}",
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
