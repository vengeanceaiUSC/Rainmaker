"""vengeanceaiUSCMODEL13.0 — Yacktman AAA-equity / bond-like CF on MODEL12 scaffolding.

MODEL12 rated 6/10 for a high-conviction long-term value investor: good Yacktman
cash framing, but punitive AR DSO cliff, bundled Deferred inside ΔNWC, linear
margins, and CAPM WACC (~9.2%) that misprices a toll-bridge Scores franchise.

Lens (Don Yacktman): treat the equity like a long-duration AAA-equivalent bond —
"what rate of return if owned 20–30 years?" FICO's B2B Scores monopoly + SaaS
upfront billings produce bond-like, inflation-protected cash compounding.

MODEL13.0 fixes (optimistic but sourced):
  1. Op NWC (AR+Inv−AP) SEPARATE from Deferred growth (explicit CFO cash source)
  2. Phased DSO (hist ~97d → ~45d) — no one-year AR cliff
  3. SBC add-back; buybacks REDUCE share count (no SBC$/Price dilution)
  4. Stronger incremental margins (Scores price hikes ≈ 100% incremental GM)
  5. Yacktman WACC 7.8% (AAA-equity band 7.5–8.2%); CAPM ~9.2% kept as reference
  6. Base exit EV/EBITDA 21.0x (bond-like premium vs peer median ~18x)
  7. Hard BS identity plug every year
"""

from __future__ import annotations

from typing import Dict, List, Tuple

from .model3_assumptions import (  # noqa: F401
    CASH_10Q_000s,
    MKT_SECS_10Q_000s,
    PERPETUAL_GROWTH,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
    SHARE_PRICE,
    TOTAL_DEBT_10Q_000s,
)
from .model5_assumptions import SOURCE_LINKS as _M5_LINKS
from .wacc import MODEL3_WACC

MODEL_NAME = "vengeanceaiUSCMODEL13.0"

# Predecessor rating (high-conviction LT value lens)
MODEL12_RATING = "6/10"

# FYE Sep 30, 2025 shares outstanding (10-K cover / equity footnote)
SHARES_OUTSTANDING_000s = 23_764.0

# --- Operating leverage / pricing power (Scores ≈ 100% incremental margin) ---
SGA_IMPROVEMENT_BPS = 125.0  # was 100 — faster opex leverage
SGA_FLOOR_PCT = 0.15
# B2B Scores unit-price hikes drop nearly 1:1 to GM → grind COGS harder than M12
COGS_IMPROVEMENT_BPS = 75.0  # was 40

# --- SBC: 3yr avg ShareBasedCompensation / Revenue (FY23–25) ---
SBC_PCT_REVENUE = 0.08251

# --- D&A: 3yr avg TOTAL D&A / Revenue (includes amort. of intangibles) ---
DA_PCT_REVENUE = 0.008411

# --- CapEx: start near 3yr hist avg, fade with operating leverage ---
CAPEX_PCT_REVENUE = 0.0125  # FY1
CAPEX_PCT_PATH: Tuple[float, ...] = (0.0125, 0.0105, 0.0090, 0.0075, 0.0060)

# --- Cash taxes: 3yr avg IncomeTaxesPaidNet / EBT ---
CASH_TAX_RATE = 0.22873

# --- Financing ($000s) ---
BUYBACK_FY23_000s = 405_526.0
BUYBACK_FY24_000s = 821_702.0
BUYBACK_FY25_000s = 1_414_502.0
BUYBACK_RUNRATE_000s = (
    BUYBACK_FY23_000s + BUYBACK_FY24_000s + BUYBACK_FY25_000s
) / 3.0
DEBT_NET_RUNRATE_000s = -290_917.0

# SaaS mix shift: +200 bps/yr from on-prem → cloud (Deferred WC cash source)
SAAS_MIX_SHIFT_BPS = 200.0

# --- Yacktman AAA-equity discount rate (overrides CAPM D6 after bake) ---
# Band 7.5–8.2%; mid 7.8%. CAPM stack (~9.24%) remains on-sheet as reference (R15).
MODEL13_WACC = 0.0780
WACC_BAND_LOW = 0.0750
WACC_BAND_HIGH = 0.0820

# --- Exit multiples (bond-like premium vs peer median ~18.1x) ---
EXIT_EV_EBITDA = 21.0
EXIT_EV_EBITDA_BULL = 26.0
EXIT_EV_EBITDA_BEAR = 15.5

PEER_EV_EBITDA: List[Tuple[str, str, float]] = [
    ("EFX", "Equifax", 13.88),
    ("VRSK", "Verisk Analytics", 17.42),
    ("SPGI", "S&P Global", 18.11),
    ("MCO", "Moody's", 22.20),
    ("MSCI", "MSCI Inc.", 23.74),
]
PEER_MEDIAN_EV_EBITDA = 18.11
URL_PEER_COMPS = "https://vcpscanner.com/valuation/fico/relative"
URL_FICO_EV_EBITDA = (
    "https://www.alphaspread.com/security/nyse/fico/relative-valuation/"
    "ratio/enterprise-value-to-ebitda"
)

# --- SEC / IR source URLs ---
URL_10K = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"
URL_10K_HTM_IR = (
    "https://investors.fico.com/static-files/663245db-39a2-4db7-8f89-284034bbccbf"
)
URL_SEC_FILINGS_IR = "https://investors.fico.com/financial-information/sec-filings"
URL_FACTS = "https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json"
URL_10K_INDEX = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/"
    "0000814547-25-000030-index.htm"
)
# Q1 FY2026 10-Q — deferred revenue roll-forward / contract liabilities
URL_10Q_Q1_FY26 = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/fico-20251231.htm"
)
URL_10Q_Q1_FY26_IR = (
    "https://investors.fico.com/static-files/935bc755-fe08-44bd-9898-d04a4e3c92f1"
)
URL_10Q_DEFERRED_DETAIL = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm"
)

# ---------------------------------------------------------------------------
# FY2025 10-K disaggregated revenue ($000s)
# ---------------------------------------------------------------------------
FY25_REVENUE_000s = 1_990_869.0
FY25_SAAS_000s = 419_720.0
FY25_ONPREM_000s = 320_425.0
FY25_B2C_000s = 219_980.0
FY25_B2B_SCORES_000s = 948_595.0
FY25_PROF_SERVICES_000s = 82_149.0
FY25_PLATFORM_ARR_000s = 263_600.0
FY25_DEFERRED_REV_000s = 187_372.0  # current deferred at 9/30/2025 (10-K / model I88)

MIX_SAAS = FY25_SAAS_000s / FY25_REVENUE_000s  # 21.08%
MIX_ONPREM = FY25_ONPREM_000s / FY25_REVENUE_000s  # 16.09%
MIX_B2C = FY25_B2C_000s / FY25_REVENUE_000s  # 11.05%
MIX_B2B = FY25_B2B_SCORES_000s / FY25_REVENUE_000s  # 47.65%
MIX_PROF_SVCS = FY25_PROF_SERVICES_000s / FY25_REVENUE_000s  # 4.13%

# Segment cash-conversion DSOs (days) — long-run economics (schedule rows 110–115)
DSO_SAAS = 45.0
DSO_ONPREM = 45.0
DSO_B2C = 2.0
DSO_B2B = 30.0
DSO_PROF_SVCS = 15.0

# MODEL13: phase AR days from FY25 hist (~97) → policy target (45) over 5 years.
# Avoids MODEL12's one-year AR cliff (97 → ~32) that distorted ΔNWC / FCFF.
HIST_DSO_DAYS = 97.0  # ≈ FY25 Gross AR / Rev × 365
TARGET_DSO_DAYS = 45.0  # enterprise SaaS / Scores collection destination

# Segment gross margins (calibrated so blended COGS ≈ FY25 ~17.8%)
GM_SAAS = 0.80
GM_ONPREM = 0.80
GM_B2C = 0.78
GM_B2B = 0.90  # Scores price hikes ≈ 100% incremental margin
GM_PROF_SVCS = 0.28


def blended_dso() -> float:
    return (
        MIX_SAAS * DSO_SAAS
        + MIX_ONPREM * DSO_ONPREM
        + MIX_B2C * DSO_B2C
        + MIX_B2B * DSO_B2B
        + MIX_PROF_SVCS * DSO_PROF_SVCS
    )


def blended_gross_margin() -> float:
    return (
        MIX_SAAS * GM_SAAS
        + MIX_ONPREM * GM_ONPREM
        + MIX_B2C * GM_B2C
        + MIX_B2B * GM_B2B
        + MIX_PROF_SVCS * GM_PROF_SVCS
    )


def blended_cogs_pct() -> float:
    return 1.0 - blended_gross_margin()


def phased_dso(year_index: int) -> float:
    """year_index 0..4 → DSO fading HIST → TARGET over 5 forecast years."""
    w = (year_index + 1) / 5.0
    return HIST_DSO_DAYS * (1.0 - w) + TARGET_DSO_DAYS * w


ASSUMPTIONS_PDF_FILENAME = "MODEL13_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel13-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel13-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)

SOURCE_LINKS: Dict[str, str] = {
    **_M5_LINKS,
    "FICO peer EV/EBITDA comps (VCP Scanner)": URL_PEER_COMPS,
    "FICO EV/EBITDA (Alpha Spread)": URL_FICO_EV_EBITDA,
    "Damodaran EV/EBITDA by sector": (
        "https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/vebitda.htm"
    ),
    "SEC companyfacts XBRL (SBC / cash tax / CapEx / buybacks)": URL_FACTS,
    "SEC 10-K FY2025 (HTML)": URL_10K,
    "SEC 10-K FY2025 filing index": URL_10K_INDEX,
    "FICO IR — SEC filings hub": URL_SEC_FILINGS_IR,
    "FICO IR — FY2025 10-K PDF": URL_10K_HTM_IR,
    "SEC 10-Q Q1 FY2026 (deferred revenue)": URL_10Q_Q1_FY26,
    "SEC 10-Q Q1 FY2026 — deferred revenue detail": URL_10Q_DEFERRED_DETAIL,
    "FICO IR — Q1 FY2026 10-Q PDF": URL_10Q_Q1_FY26_IR,
    "MODEL13 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL13 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "MODEL13 splits working capital for cash-flow clarity (no double count): "
    "Operating NWC = AR + Inventory − AP (excludes Deferred). "
    "Δ Op NWC uses cash when AR/AP/Inv rise. Separately, Increase in Deferred "
    "Revenue is an explicit positive CFO / FCFF line (SaaS & maintenance billed "
    "upfront). DSO phases from FY25 hist ≈ 97 days → policy target 45 days over "
    "5 years (removes MODEL12's punitive AR cliff). AR changes never absorb "
    "Deferred billings."
)

SEGMENT_METHODOLOGY_NOTE = (
    "FY2025 10-K disaggregation ($000s): B2B Scores 948,595; B2C myFICO 219,980; "
    "SaaS software 419,720; on-premises software 320,425; Professional services "
    "82,149 (sum = Total revenues 1,990,869). FICO Platform ARR was $263.6M "
    "(35% of software ARR) at 9/30/2025 — ARR is a KPI, not an incremental IS line. "
    f"MODEL13 shifts +{SAAS_MIX_SHIFT_BPS:.0f} bps/yr from on-prem mix into SaaS "
    "to reflect the cloud transition (drives Deferred Revenue growth)."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Deferred Revenue is a BS liability in Total Liabilities (AP + Debt + Deferred). "
    "It is NOT netted inside Op NWC for CFO. Forecast Deferred = Total Rev × "
    "(SaaS%+OnPrem%) × (FY25 Deferred ÷ FY25 software Rev). FY25 current deferred "
    f"≈ ${FY25_DEFERRED_REV_000s/1000:.1f}M (10-K). Q1 FY2026 10-Q shows deferred "
    "roll-forward and notes maintenance/SaaS billed annually in advance "
    f"({URL_10Q_DEFERRED_DETAIL}). Increase in Deferred = Def_t − Def_t−1 is added "
    "in CFO and FCFF. Scores/myFICO usage fees are NOT forced into Deferred "
    "(no double count with AR)."
)

DA_METHODOLOGY_NOTE = (
    "CFO/FCFF add back TOTAL D&A at ≈ 0.84% of Revenue (3yr avg "
    "DepreciationDepletionAndAmortization / Rev). That XBRL total already includes "
    "AmortizationOfIntangibleAssets — intangibles amortization is fully added back "
    "once (no separate second add-back)."
)

SBC_METHODOLOGY_NOTE = (
    "SBC is added back once in CFO and FCFF at ≈ 8.25% of Revenue "
    "(ShareBasedCompensation). MODEL13 does NOT also dilute shares by SBC$/Price "
    "(that would double-penalize). Share count instead falls with actual buyback "
    "cash outflows from residual levered FCF (3S equity CF row 20)."
)

CAPEX_METHODOLOGY_NOTE = (
    "CapEx % of Revenue fades with operating leverage: "
    f"{CAPEX_PCT_PATH[0]:.2%} → {CAPEX_PCT_PATH[1]:.2%} → {CAPEX_PCT_PATH[2]:.2%} → "
    f"{CAPEX_PCT_PATH[3]:.2%} → {CAPEX_PCT_PATH[4]:.2%}. "
    "FY1 starts near the 3yr hist avg of (PPE + capitalized software)/Rev; "
    "Y5 approaches asset-light maintenance intensity as Scores/SaaS mix dominates."
)

MARGIN_METHODOLOGY_NOTE = (
    f"Pricing power: blended COGS% grinds −{COGS_IMPROVEMENT_BPS:.0f} bps/yr from "
    f"segment GM base; SG&A grinds −{SGA_IMPROVEMENT_BPS:.0f} bps/yr toward a "
    f"{SGA_FLOOR_PCT:.0%} floor. B2B Scores unit-price increases have near-100% "
    "incremental gross margin (little variable COGS) — MODEL12's linear −40/−100 "
    "bps understated that cash conversion. FY25 10-K: Scores revenue +$249M YoY "
    "'primarily attributable to a higher unit price'."
)

CASH_TAX_METHODOLOGY_NOTE = (
    "DCF unlevered cash taxes use the 3yr average cash tax rate "
    "(IncomeTaxesPaidNet ÷ EBT ≈ 22.87%), not the book effective tax rate. "
    "The 3-statement income statement still applies book tax for NI."
)

BS_METHODOLOGY_NOTE = (
    "Balance sheet identity is enforced every year: Equity Capital = "
    "Total Assets − Total Liabilities − Retained Earnings (hard plug). "
    "Total Liabilities = AP + Debt + Deferred Revenue. Row-3 Balance Sheet "
    "Check flags any residual ≠ 0."
)

FINANCING_METHODOLOGY_NOTE = (
    "Debt CF uses the 3yr avg net debt cash flow (≈ −$291M/yr). Equity buybacks "
    "distribute residual levered FCF so cash does not stockpile. 10-K CF buybacks: "
    f"FY23 ${BUYBACK_FY23_000s/1000:.0f}M / FY24 ${BUYBACK_FY24_000s/1000:.0f}M / "
    f"FY25 ${BUYBACK_FY25_000s/1000:.0f}M (avg ≈ ${BUYBACK_RUNRATE_000s/1000:.0f}M). "
    "June 2025 authorization: $1.0B program. Accretive buyback engine: robust FCFF "
    "→ residual equity CF → lower forward share count in $/share."
)

DILUTION_METHODOLOGY_NOTE = (
    "Starting shares = 23,764k (FYE Sep 30, 2025 outstanding). Each forecast year: "
    "Shares_t = Shares_t−1 + EquityIssuanceCF_t / Price. EquityIssuanceCF is the "
    "buyback outflow (negative) from 3S row 20 — so repurchases REDUCE the share "
    "count. SBC is added back in FCFF but does NOT increase shares (avoids double "
    "penalty). $/share uses Year-5 buyback-adjusted shares."
)

WACC_METHODOLOGY_NOTE = (
    f"Yacktman AAA-equity WACC = {MODEL13_WACC:.2%} (policy band "
    f"{WACC_BAND_LOW:.1%}–{WACC_BAND_HIGH:.1%}). FICO's Scores toll-bridge has "
    "ultra-low default risk and bond-like cash predictability; CAPM (~"
    f"{MODEL3_WACC:.2%} with β≈1.32) overstates the long-horizon discount rate for "
    "this franchise. Live CAPM stack remains in DCF!R15 as a reference; D6 is the "
    "Yacktman policy rate used for XNPV."
)

EXIT_METHODOLOGY_NOTE = (
    f"Base exit EV/EBITDA = {EXIT_EV_EBITDA:.1f}x (raised from MODEL12 17.5x). "
    f"A perpetual pricing-power compounder deserves a premium to the peer median "
    f"(~{PEER_MEDIAN_EV_EBITDA:.1f}x); band Bull {EXIT_EV_EBITDA_BULL:.1f}x / "
    f"Bear {EXIT_EV_EBITDA_BEAR:.1f}x. Still below FICO spot (~25–27x)."
)

YACKTMAN_METHODOLOGY_NOTE = (
    f"Yacktman lens (MODEL12 rated {MODEL12_RATING}): think of FICO equity as a "
    "long-duration AAA-equivalent bond — monopolistic B2B Scores pricing, rising "
    "SaaS deferred billings, and persistent buybacks. MODEL13 lowers WACC to "
    f"{MODEL13_WACC:.2%}, lifts exit to {EXIT_EV_EBITDA:.1f}x, separates Deferred "
    "cash inflows from Op NWC/AR, phases DSO, and sharpens incremental margins — "
    "optimistic cash conversion without fake plugs; BS identity always holds."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME} (Yacktman AAA-equity): MODEL12={MODEL12_RATING}; "
        f"OpNWC≠Deferred (+ΔDef in CFO); DSO {HIST_DSO_DAYS:.0f}→{TARGET_DSO_DAYS:.0f}d; "
        f"COGS −{COGS_IMPROVEMENT_BPS:.0f}bps/SG&A −{SGA_IMPROVEMENT_BPS:.0f}bps; "
        f"buybacks cut shares; exit {EXIT_EV_EBITDA:.1f}x; "
        f"WACC={MODEL13_WACC:.2%} (CAPM ref {MODEL3_WACC:.2%})."
    )
