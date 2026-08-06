"""Plain-English explanations for every 3-statement assumption row (MODEL17).

SOURCE fields always include a clickable URL (col U) so users can open the filing/data.
Every assumption also links to the downloadable Assumptions List PDF (no charts).
Change log format: Updated [Variable] from Model16 (Previous: X) to Model17 (New: Y).
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import List, Tuple

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .model17_assumptions import (
    ASSUMPTIONS_PDF_FILENAME,
    ASSUMPTIONS_PDF_URL,
    ASSUMPTIONS_PDF_VIEW_URL,
    BUYBACK_FCF_MULTIPLE,
    BUYBACK_RUNRATE_000s,
    CAPEX_METHODOLOGY_NOTE,
    CAPEX_PCT_REVENUE,
    CASH_TAX_METHODOLOGY_NOTE,
    CASH_TAX_RATE,
    COGS_FLOOR_PCT,
    DA_METHODOLOGY_NOTE,
    DA_PCT_REVENUE,
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
    EXIT_METHODOLOGY_NOTE,
    BS_METHODOLOGY_NOTE,
    CAPEX_PCT_PATH,
    COGS_IMPROVEMENT_BPS,
    FINANCING_METHODOLOGY_NOTE,
    FY25_B2B_SCORES_000s,
    GROWTH_METHODOLOGY_NOTE,
    HIST_DSO_DAYS,
    MARGIN_METHODOLOGY_NOTE,
    MODEL16_RATING,
    MODEL16_WACC,
    MODEL17_WACC,
    PERPETUAL_GROWTH,
    RD_FLOOR_PCT,
    RD_IMPROVEMENT_BPS,
    REVENUE_GROWTH_PATH,
    SAAS_MIX_SHIFT_BPS,
    SBC_PCT_PATH,
    SGA_FLOOR_PCT,
    SGA_IMPROVEMENT_BPS,
    SHARES_OUTSTANDING_000s,
    SHARE_PRICE,
    TARGET_DSO_DAYS,
    URL_10Q_DEFERRED_DETAIL,
    URL_10Q_Q1_FY26,
    URL_FRED_DGS10,
    URL_YAHOO_BETA,
    WACC_METHODOLOGY_NOTE,
    YACKTMAN_METHODOLOGY_NOTE,
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
        "Yellow POLICY input: "
        + " / ".join(f"{g:.1%}" for g in REVENUE_GROWTH_PATH)
        + ".",
        GROWTH_METHODOLOGY_NOTE,
        "SEC EX-99.1 Q3 FY2026 — Scores +41% / B2B +49%; FY2026 guidance",
        URL_GUIDANCE,
    ),
    (
        8,
        "COGS % of Revenue",
        "Cost of revenues as a % of sales (gross margin = 1 − this).",
        f"Excel: MAX({COGS_FLOOR_PCT:.0%}, segment blended COGS base≈{blended_cogs_pct():.2%} − "
        f"{COGS_IMPROVEMENT_BPS:.0f}bps × year).",
        f"MODEL17 Yacktman: {MARGIN_METHODOLOGY_NOTE} "
        "Mix % offset within Total Revenue (not a second revenue stack).",
        "SEC 10-K FY2025 — Scores +$249M YoY 'primarily attributable to a higher unit price'",
        URL_10K,
    ),
    (
        9,
        "SG&A % of Revenue",
        "Operating opex (SG&A) as % of sales.",
        f"Equation: MAX({SGA_FLOOR_PCT:.0%}, (I28 − 10,922)/I24 − "
        f"{SGA_IMPROVEMENT_BPS:.0f}bps × year). "
        f"Strips FY25 restructuring; then −{SGA_IMPROVEMENT_BPS/100:.2f}% of sales each year.",
        f"Updated SG&A from Model16 (Previous: −125 bps, floor 15%) to Model17 "
        f"(New: −{SGA_IMPROVEMENT_BPS:.0f} bps/yr, floor {SGA_FLOOR_PCT:.0%}) — "
        "Scores near-100% incremental margins drop dollars below the line.",
        "SEC 10-K FY2025 — SG&A + Scores pricing commentary",
        URL_10K,
    ),
    (
        10,
        "R&D % of Revenue",
        "Research & development expense as % of sales (IS label: R&D, not Rent).",
        f"Excel: MAX({RD_FLOOR_PCT:.0%}, I29/I24 − {RD_IMPROVEMENT_BPS:.0f}bps × year).",
        f"Updated R&D from Model16 (Previous: flat ~9.5%) to Model17 "
        f"(New: −{RD_IMPROVEMENT_BPS:.0f} bps/yr, floor {RD_FLOOR_PCT:.0%}) — "
        "royalty price dollars do not require 1:1 R&D reinvestment.",
        "SEC 10-K FY2025 — Research and development",
        URL_10K,
    ),
    (
        11,
        "D&A % of Revenue",
        "Total depreciation & amortization as % of sales (software-industry driver).",
        f"Yellow POLICY input: flat {DA_PCT_REVENUE:.2%} (3yr FY23–25 avg TOTAL D&A). "
        "Forecast DA$ = Revenue × DA%; CF row 63 adds it back once.",
        f"MODEL17: {DA_METHODOLOGY_NOTE}",
        "SEC companyfacts — DepreciationDepletionAndAmortization (+ AmortizationOfIntangibleAssets)",
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
        f"Yellow POLICY path: phases {HIST_DSO_DAYS:.0f}d (FY25 hist) → "
        f"{TARGET_DSO_DAYS:.0f}d over 5 years "
        f"(Y1≈{(HIST_DSO_DAYS*0.8+TARGET_DSO_DAYS*0.2):.1f}d … Y5={TARGET_DSO_DAYS:.0f}d). "
        f"Segment blend ref ≈ {blended_dso():.1f}d on schedule row 115 "
        f"(SaaS {DSO_SAAS:.0f}/B2C {DSO_B2C:.0f}/B2B {DSO_B2B:.0f}/"
        f"PS {DSO_PROF_SVCS:.0f}/On-prem {DSO_ONPREM:.0f}).",
        f"DSO path unchanged vs Model16 (97→{TARGET_DSO_DAYS:.0f}d). "
        f"{WC_METHODOLOGY_NOTE} {DEFERRED_METHODOLOGY_NOTE}",
        "SEC 10-K FY2025 — AR; Q1 FY2026 deferred detail (R45)",
        URL_10Q_DEFERRED_DETAIL,
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
        "MODEL17: DPO pairs with phased DSO for Op NWC = AR+Inv−AP. "
        "Deferred is a SEPARATE CFO cash source (row 81) — never mixed into AR/DPO.",
        "SEC 10-K FY2025 — Accounts payable",
        URL_10K,
    ),
    (
        18,
        "CapEx % of Revenue",
        "Capital investment (PPE + capitalized software) as % of sales.",
        f"Yellow POLICY path: "
        + " / ".join(f"{p:.2%}" for p in CAPEX_PCT_PATH)
        + " (fade with operating leverage).",
        f"MODEL17 Yacktman: {CAPEX_METHODOLOGY_NOTE}",
        "SEC 10-K / companyfacts — PPE purchases + capitalized software; fade to maintenance",
        URL_FACTS,
    ),
    (
        19,
        "Debt Issuance (Repayment)",
        "New borrowing (+) or repayment (−) in the forecast ($000s).",
        f"Equation: −(CFO−CapEx) − EquityCF = ({BUYBACK_FCF_MULTIPLE:.2f}−1)×FCF "
        "(debt funds buybacks above organic FCF).",
        f"MODEL17: {FINANCING_METHODOLOGY_NOTE} Financing is excluded from FCFF.",
        "SEC EX-99.1 Q3 FY2026 — levered buybacks; companyfacts senior notes / LOC",
        URL_FACTS,
    ),
    (
        20,
        "Equity Issued (Repaid)",
        "New equity issued (+) or buybacks (−) in the forecast BS/CF.",
        f"Equation: −{BUYBACK_FCF_MULTIPLE:.2f}×(CFO−CapEx) levered buybacks "
        f"(hist 3yr avg ≈ ${BUYBACK_RUNRATE_000s/1000:.0f}M/yr; Q3 FY26 >> FCF).",
        f"Updated buybacks from Model16 (Previous: residual ≈1.0× FCF) to Model17 "
        f"(New: {BUYBACK_FCF_MULTIPLE:.2f}× FCF). Buybacks are financing — excluded "
        f"from FCFF. DCF share count FALLS with this outflow. {DILUTION_METHODOLOGY_NOTE}",
        "SEC EX-99.1 Q3 FY2026 — buybacks vs FCF; 10-K FY2025 repurchase footnote",
        URL_10K,
    ),
    (
        21,
        "SBC % of Revenue",
        "Stock-based compensation as % of sales; non-cash add-back in CFO and FCFF.",
        "Yellow POLICY fade path: "
        + " / ".join(f"{p:.2%}" for p in SBC_PCT_PATH)
        + f" (starts at 3yr avg {SBC_PCT_REVENUE:.2%}).",
        f"MODEL17 Yacktman: {SBC_METHODOLOGY_NOTE}",
        "SEC 10-K cash flow — ShareBasedCompensation (add-back only; no share dilution)",
        URL_FACTS,
    ),
    (
        104,
        "Mix % — SaaS / Platform software",
        "Share of Total Revenue from SaaS / cloud software (incl. FICO Platform cloud).",
        f"FY25 base {MIX_SAAS:.2%} (SaaS ${FY25_SAAS_000s:,.0f}k / Total "
        f"${FY25_REVENUE_000s:,.0f}k), then +{SAAS_MIX_SHIFT_BPS:.0f} bps/yr "
        f"taken from on-prem. "
        f"Platform ARR KPI ${FY25_PLATFORM_ARR_000s/1000:.1f}M is not additive IS revenue.",
        f"SaaS mix shift unchanged vs Model16 (+{SAAS_MIX_SHIFT_BPS:.0f} bps/yr). "
        f"SaaS billed annually in advance → Deferred Revenue growth → explicit "
        f"CFO cash source (row 81). Source: {URL_10Q_Q1_FY26}.",
        "SEC 10-Q Q1 FY2026 — Deferred revenue / contract liabilities detail",
        URL_10Q_DEFERRED_DETAIL,
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
            "assumption": "WACC (Yacktman-adj CAPM — D6 = R15)",
            "what_it_is": "Weighted average cost of capital for XNPV of FCFF + TV.",
            "how_set_in_model": (
                f"D6 = R15 (Yacktman-adj CAPM) ≈ {MODEL17_WACC:.2%}. "
                f"Updated from Model16 (Previous: raw CAPM {MODEL16_WACC:.2%}) "
                f"to Model17 (New: Yacktman-adj CAPM {MODEL17_WACC:.2%}). "
                f"Rf/β/ERP/Rd inputs in Q6:R15 with clickable sources."
            ),
            "why_this_choice": WACC_METHODOLOGY_NOTE,
            "source": "Blume 1971; FRED DGS10; Yahoo β; Damodaran ERP; SEC 8-K 6.250% Rd",
            "source_url": URL_FRED_DGS10,
        },
        {
            "row": "DCF-D7",
            "assumption": "Perpetual Growth (g)",
            "what_it_is": "Long-run growth used only in the Gordon TV cross-check.",
            "how_set_in_model": (
                f"POLICY input = {PERPETUAL_GROWTH:.1%}. Updated from Model16 "
                f"(Previous: 3.0%) to Model17 (New: {PERPETUAL_GROWTH:.1%})."
            ),
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
            "why_this_choice": EXIT_METHODOLOGY_NOTE,
            "source": "VCP Scanner FICO peer EV/EBITDA comps + Yacktman premium",
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
            "assumption": "Shares Outstanding (buyback-adjusted)",
            "what_it_is": (
                f"Starting shares (D12 ≈ {SHARES_OUTSTANDING_000s:,.1f}k post–Q3 @ "
                f"${SHARE_PRICE:,.0f}) fall each year with 1.4×FCF buybacks "
                "(3S equity CF / price). SBC is added back in FCFF but does NOT "
                "increase shares — avoids double penalty. $/share uses Y5 shares (I16)."
            ),
            "how_set_in_model": (
                "D16=$D$12; E16:I16 = MAX(1000, prior + 3S!{J–N}20 / $D$11); "
                "D37=$D$35/$I$16."
            ),
            "why_this_choice": DILUTION_METHODOLOGY_NOTE,
            "source": (
                f"EX-99.1 Q3 FY2026 — shares ≈ {SHARES_OUTSTANDING_000s:,.1f}k @ "
                f"${SHARE_PRICE:,.0f}; FY25 buybacks $1.415B"
            ),
            "source_url": URL_10K,
        },
        {
            "row": "DCF-E25",
            "assumption": "Net WC investment (ΔOpNWC − ΔDeferred)",
            "what_it_is": (
                "FCFF subtracts Δ Operating NWC (AR+Inv−AP) and adds Increase in "
                "Deferred Revenue. AR and Deferred are strictly separated."
            ),
            "how_set_in_model": (
                "DCF E25:I25 = 3S!{J–N}90 − (3S Def_t − Def_t−1). "
                "Equivalent to −ΔOpNWC + ΔDeferred inside UFCFF."
            ),
            "why_this_choice": (
                f"{WC_METHODOLOGY_NOTE} {DEFERRED_METHODOLOGY_NOTE}"
            ),
            "source": "SEC 10-Q Q1 FY2026 — Deferred revenue roll-forward (R45)",
            "source_url": URL_10Q_DEFERRED_DETAIL,
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
            f"MODEL16 rated {MODEL16_RATING}. "
            f"Updated WACC from Model16 (Previous: raw CAPM {MODEL16_WACC:.2%}) "
            f"to Model17 (New: Yacktman-adj CAPM {MODEL17_WACC:.2%} = D6←R15). "
            f"DSO unchanged (97→{TARGET_DSO_DAYS:.0f}d). "
            f"Exit BASE {EXIT_EV_EBITDA:.1f}x (was 21.0x); g {PERPETUAL_GROWTH:.1%} "
            f"(was 3.0%); peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x; "
            f"bull {EXIT_EV_EBITDA_BULL:.1f}x; bear {EXIT_EV_EBITDA_BEAR:.1f}x).",
            body,
        ),
        Paragraph(
            f"CAPM sources: <link href='{URL_FRED_DGS10}' color='blue'>"
            f"{URL_FRED_DGS10}</link>; "
            f"<link href='{URL_YAHOO_BETA}' color='blue'>{URL_YAHOO_BETA}</link>.",
            small,
        ),
        Paragraph(YACKTMAN_METHODOLOGY_NOTE, body),
        Paragraph(WACC_METHODOLOGY_NOTE, body),
        Paragraph(EXIT_METHODOLOGY_NOTE, body),
        Paragraph(WC_METHODOLOGY_NOTE, body),
        Paragraph(SEGMENT_METHODOLOGY_NOTE, body),
        Paragraph(DEFERRED_METHODOLOGY_NOTE, body),
        Paragraph(DA_METHODOLOGY_NOTE, body),
        Paragraph(SBC_METHODOLOGY_NOTE, body),
        Paragraph(DILUTION_METHODOLOGY_NOTE, body),
        Paragraph(CAPEX_METHODOLOGY_NOTE, body),
        Paragraph(MARGIN_METHODOLOGY_NOTE, body),
        Paragraph(CASH_TAX_METHODOLOGY_NOTE, body),
        Paragraph(FINANCING_METHODOLOGY_NOTE, body),
        Paragraph(BS_METHODOLOGY_NOTE, body),
        Paragraph(
            f"Primary filings hub: <link href='{URL_SEC_FILINGS_IR}' color='blue'>"
            f"{URL_SEC_FILINGS_IR}</link> — segmentation from FY2025 10-K "
            f"(<link href='{URL_10K}' color='blue'>{URL_10K}</link>; "
            f"IR <link href='{URL_10K_HTM_IR}' color='blue'>{URL_10K_HTM_IR}</link>). "
            f"Deferred/contract liabilities: Q1 FY2026 10-Q detail "
            f"(<link href='{URL_10Q_DEFERRED_DETAIL}' color='blue'>"
            f"{URL_10Q_DEFERRED_DETAIL}</link>).",
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
    """Write MODEL17_*_ASSUMPTIONS_EXPLAINED.csv, TXT, and PDF list (no charts)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"MODEL17_{ticker}_ASSUMPTIONS_EXPLAINED.csv"
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

    txt_path = out_dir / f"MODEL17_{ticker}_ASSUMPTIONS_EXPLAINED.txt"
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
        YACKTMAN_METHODOLOGY_NOTE,
        WC_METHODOLOGY_NOTE,
        SEGMENT_METHODOLOGY_NOTE,
        DEFERRED_METHODOLOGY_NOTE,
        DA_METHODOLOGY_NOTE,
        SBC_METHODOLOGY_NOTE,
        DILUTION_METHODOLOGY_NOTE,
        CAPEX_METHODOLOGY_NOTE,
        MARGIN_METHODOLOGY_NOTE,
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
    md_path = out_dir / f"MODEL17_{ticker}_ASSUMPTIONS_LIST.md"
    md = [
        f"# {_MODEL_NAME} — Assumptions List",
        "",
        f"**PDF download:** {ASSUMPTIONS_PDF_URL}",
        "",
        "Plain list only (no charts). Paste into Google Docs via File → Open if needed.",
        "",
        f"> {YACKTMAN_METHODOLOGY_NOTE}",
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
        f"> {MARGIN_METHODOLOGY_NOTE}",
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
