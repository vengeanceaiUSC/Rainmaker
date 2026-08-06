"""vengeanceaiUSCMODEL14.0 — Yacktman AAA toll-road monopoly on MODEL13 scaffolding.

MODEL13 rated 6.5/10: captures SaaS/Deferred WC mechanics but still treats FICO
like a standard software OpCo — linear opex with Scores price hikes, 7.8% WACC
too punitive for mandated GSE cash flows, and DSO fade that traps cash vs
upfront SaaS/Scores collection reality.

Yacktman lens: FICO equity ≈ AAA sovereign bond with a coupon that grows faster
than inflation. Lenders are effectively mandated by GSEs to use FICO scores.
Price hikes on B2B / mortgage royalties have ~100% incremental margin (no COGS,
no incremental SG&A/R&D) and flow ~95%+ to CFO/FCFF.

MODEL14.0 package:
  1. WACC 7.1% (AAA-bond + thin equity premium; band 7.0–7.2%)
  2. Perpetual g = 3.5% (pricing power > baseline inflation)
  3. Mortgage royalty $4.95→$10 pricing + Scores incremental-margin grind
  4. Faster DSO (97→35d) + Deferred CFO source (no AR double-count)
  5. Post–Q3 FY2026 buyback share base; SBC add-back without share dilution
  6. Hard BS identity plug every year
"""

from __future__ import annotations

from typing import Dict, List, Tuple

from .model3_assumptions import (  # noqa: F401
    CASH_10Q_000s,
    MKT_SECS_10Q_000s,
    RESTRUCTURING_NORMALIZE_000s,
    TOTAL_DEBT_10Q_000s,
)
from .model5_assumptions import SOURCE_LINKS as _M5_LINKS
from .wacc import MODEL3_WACC

MODEL_NAME = "vengeanceaiUSCMODEL14.0"

# Predecessor rating (high-conviction LT value / Yacktman lens)
MODEL13_RATING = "6.5/10"

# --- Market / share base (post Q3 FY2026 record buybacks) ---
# Q3 FY2026: repurchased 1.705M shares for $1.96B @ avg $1,149
# Starting DCF shares = diluted/outstanding after that repurchase wave.
SHARE_PRICE = 1_149.0
SHARES_OUTSTANDING_000s = 21_597.635  # post–Q3 FY2026 bridge (vs MODEL13 23,764k FYE25)

BUYBACK_Q3_FY26_000s = 1_960_000.0
BUYBACK_Q3_FY26_SHARES_000s = 1_705.0
BUYBACK_Q3_FY26_AVG_PRICE = 1_149.0
Q3_FY26_FCF_000s = 370_343.0

# --- Revenue growth (MODEL13 was 27/16/13/10/7) ---
# MODEL14 lifts path for mortgage royalty reprice ($4.95→$10) flowing to Scores.
REVENUE_GROWTH_PATH: Tuple[float, ...] = (0.30, 0.18, 0.145, 0.11, 0.085)
REVENUE_GROWTH_PATH_MODEL13: Tuple[float, ...] = (0.27, 0.16, 0.13, 0.10, 0.07)

# --- Operating leverage / 95%+ incremental Scores margin ---
SGA_IMPROVEMENT_BPS = 175.0  # MODEL13: 125 — price hikes need ~0 incremental SG&A
SGA_FLOOR_PCT = 0.14
RD_IMPROVEMENT_BPS = 50.0  # MODEL13: flat R&D% — fade; Scores royalties need no R&D
RD_FLOOR_PCT = 0.07
COGS_IMPROVEMENT_BPS = 120.0  # MODEL13: 75 — royalty hikes ≈ 0 COGS

# --- SBC / D&A / CapEx / cash tax (carry forward; SBC still add-back only) ---
SBC_PCT_REVENUE = 0.08251
DA_PCT_REVENUE = 0.008411
CAPEX_PCT_REVENUE = 0.0125
CAPEX_PCT_PATH: Tuple[float, ...] = (0.0125, 0.0105, 0.0090, 0.0075, 0.0060)
CASH_TAX_RATE = 0.22873

# --- Financing ($000s) ---
BUYBACK_FY23_000s = 405_526.0
BUYBACK_FY24_000s = 821_702.0
BUYBACK_FY25_000s = 1_414_502.0
BUYBACK_RUNRATE_000s = (
    BUYBACK_FY23_000s + BUYBACK_FY24_000s + BUYBACK_FY25_000s
) / 3.0
DEBT_NET_RUNRATE_000s = -290_917.0

# SaaS mix shift (Deferred CFO cash) — slightly faster cloud transition
SAAS_MIX_SHIFT_BPS = 250.0  # MODEL13: 200

# --- Yacktman AAA discount + terminal g ---
# Band 7.0–7.2%; mid 7.1%. CAPM (~9.24%) remains on-sheet as reference (R15).
MODEL14_WACC = 0.0710
WACC_BAND_LOW = 0.0700
WACC_BAND_HIGH = 0.0720
MODEL13_WACC = 0.0780  # predecessor for Branch Process notes

PERPETUAL_GROWTH = 0.035  # MODEL13: 3.0% — pricing power > inflation
PERPETUAL_GROWTH_MODEL13 = 0.030

# --- Exit multiples ---
EXIT_EV_EBITDA = 22.0  # MODEL13: 21.0 — AAA toll-road premium
EXIT_EV_EBITDA_BULL = 27.0
EXIT_EV_EBITDA_BEAR = 16.0

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

# --- SEC / IR / industry source URLs ---
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
URL_10Q_Q1_FY26 = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/fico-20251231.htm"
)
URL_10Q_DEFERRED_DETAIL = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm"
)
URL_Q3_FY26_EX991 = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/"
    "exhibit991erq32026.htm"
)
URL_Q3_FY26_IR = (
    "https://investors.fico.com/news-releases/news-release-details/"
    "fico-announces-earnings-1045-share-third-quarter-fiscal-2026"
)
URL_EQUIFAX_FICO_PRICING = (
    "https://www.equifax.com/newsroom/all-news/-/story/"
    "equifax-statement-on-fico-2x-price-increase-for-2026-and-mortgage-direct-license-program/"
)
URL_HAWLEY_LETTER = (
    "https://www.hawley.senate.gov/wp-content/uploads/2026/03/2026-03-23-Hawley-Oversight-Letter-to-FICO.pdf"
)
URL_HOUSINGWIRE_FICO = "https://www.housingwire.com/articles/hawley-fico-mortgage-pricing/"

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
FY25_DEFERRED_REV_000s = 187_372.0

# Mortgage royalty reprice (industry / oversight disclosures)
MORTGAGE_ROYALTY_2025 = 4.95
MORTGAGE_ROYALTY_2026 = 10.00
MORTGAGE_INDUSTRY_COST_UPLIFT_MID_M = 500.0  # Hawley ~$500M industry; Equifax cites ~$100M FICO-only frame

MIX_SAAS = FY25_SAAS_000s / FY25_REVENUE_000s
MIX_ONPREM = FY25_ONPREM_000s / FY25_REVENUE_000s
MIX_B2C = FY25_B2C_000s / FY25_REVENUE_000s
MIX_B2B = FY25_B2B_SCORES_000s / FY25_REVENUE_000s
MIX_PROF_SVCS = FY25_PROF_SERVICES_000s / FY25_REVENUE_000s

DSO_SAAS = 40.0  # MODEL13: 45 — faster SaaS collections / annual billings
DSO_ONPREM = 40.0
DSO_B2C = 2.0
DSO_B2B = 25.0  # MODEL13: 30 — Scores settle faster than modeled
DSO_PROF_SVCS = 12.0

# Faster AR release than MODEL13 (97→45): hist → 35d over 5 years
HIST_DSO_DAYS = 97.0
TARGET_DSO_DAYS = 35.0  # MODEL13: 45

# Segment GMs — B2B Scores royalty economics ~100% incremental
GM_SAAS = 0.82
GM_ONPREM = 0.80
GM_B2C = 0.80
GM_B2B = 0.95  # MODEL13: 0.90 — pure royalty / data input only
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


ASSUMPTIONS_PDF_FILENAME = "MODEL14_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel14-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel14-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
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
    "SEC EX-99.1 Q3 FY2026 earnings (FCF $370M)": URL_Q3_FY26_EX991,
    "FICO IR — Q3 FY2026 earnings release": URL_Q3_FY26_IR,
    "Equifax statement — FICO 2026 2x mortgage royalty": URL_EQUIFAX_FICO_PRICING,
    "Sen. Hawley oversight letter — FICO mortgage pricing": URL_HAWLEY_LETTER,
    "HousingWire — Hawley probes FICO mortgage pricing": URL_HOUSINGWIRE_FICO,
    "MODEL14 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL14 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "Branch Process MODEL13→MODEL14: Operating NWC = AR+Inv−AP (excludes Deferred) "
    "unchanged in structure, but DSO fade accelerates — MODEL13 targeted 45d; "
    f"MODEL14 targets {TARGET_DSO_DAYS:.0f}d (hist {HIST_DSO_DAYS:.0f}d → "
    f"{TARGET_DSO_DAYS:.0f}d over 5yrs) so cash is not artificially trapped. "
    "Increase in Deferred Revenue remains an explicit CFO/FCFF source (row 81). "
    "AR and unearned revenue are never double-counted."
)

SEGMENT_METHODOLOGY_NOTE = (
    "FY2025 10-K disaggregation ($000s): B2B Scores 948,595; B2C myFICO 219,980; "
    "SaaS 419,720; on-prem 320,425; PS 82,149. "
    f"MODEL14 shifts +{SAAS_MIX_SHIFT_BPS:.0f} bps/yr on-prem→SaaS "
    f"(MODEL13 used +200 bps) to reflect faster cloud/Deferred billings."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Deferred Revenue stays in Total Liabilities and is an explicit +ΔDeferred in "
    "CFO/FCFF (not inside Op NWC). Forecast Deferred = Rev × (SaaS%+OnPrem%) × "
    f"(FY25 Def/FY25 soft Rev). FY25 current deferred ≈ ${FY25_DEFERRED_REV_000s/1000:.1f}M. "
    f"Q1 FY2026 roll-forward: {URL_10Q_DEFERRED_DETAIL}."
)

DA_METHODOLOGY_NOTE = (
    "CFO/FCFF add back TOTAL D&A at ≈ 0.84% of Revenue (3yr avg). "
    "Intangibles amortization is inside that total — added back once."
)

SBC_METHODOLOGY_NOTE = (
    "SBC added back once in CFO/FCFF (~8.25% of Rev). MODEL14 does NOT dilute "
    "shares by SBC$/Price. Branch Process: MODEL13 started shares at 23,764k "
    f"(FYE25); MODEL14 starts at {SHARES_OUTSTANDING_000s:,.3f}k after Q3 FY2026 "
    f"repurchases of {BUYBACK_Q3_FY26_SHARES_000s:,.0f}k shares for "
    f"${BUYBACK_Q3_FY26_000s/1000:.0f}M @ ${BUYBACK_Q3_FY26_AVG_PRICE:,.0f} avg "
    f"(Q3 FCF ${Q3_FY26_FCF_000s/1000:.1f}M). Residual FCF buybacks continue to "
    "shrink the forward share count — offsets SBC dilution economically."
)

CAPEX_METHODOLOGY_NOTE = (
    "CapEx % fades with operating leverage: "
    f"{CAPEX_PCT_PATH[0]:.2%}→{CAPEX_PCT_PATH[-1]:.2%} (same path as MODEL13)."
)

MARGIN_METHODOLOGY_NOTE = (
    f"Branch Process MODEL13→MODEL14: COGS grind −{COGS_IMPROVEMENT_BPS:.0f} bps/yr "
    f"(was −75); SG&A −{SGA_IMPROVEMENT_BPS:.0f} bps/yr (was −125); "
    f"R&D −{RD_IMPROVEMENT_BPS:.0f} bps/yr (MODEL13 held R&D% flat). "
    "B2B / mortgage royalty price hikes have ~100% incremental GM and need ~0 "
    "incremental SG&A/R&D — ~95%+ of incremental Scores revenue flows to CFO. "
    f"2026 mortgage wholesale royalty {MORTGAGE_ROYALTY_2025:.2f}→"
    f"{MORTGAGE_ROYALTY_2026:.2f} (Equifax/Hawley; industry uplift cited up to "
    f"~${MORTGAGE_INDUSTRY_COST_UPLIFT_MID_M:.0f}M)."
)

CASH_TAX_METHODOLOGY_NOTE = (
    "DCF unlevered cash taxes use 3yr IncomeTaxesPaidNet/EBT ≈ 22.87%; "
    "3S NI still uses book tax."
)

BS_METHODOLOGY_NOTE = (
    "Equity Capital = Assets − Liabilities − RE every year (hard plug). "
    "Total Liabilities = AP + Debt + Deferred Revenue. Row-3 BS Check = OK."
)

FINANCING_METHODOLOGY_NOTE = (
    "Debt CF ≈ 3yr avg net debt CF. Equity buybacks = residual levered FCF after "
    "debt (no cash stockpile). Q3 FY2026 proved hyper-aggressive capacity: "
    f"${BUYBACK_Q3_FY26_000s/1000:.0f}M buybacks vs ${Q3_FY26_FCF_000s/1000:.1f}M FCF "
    f"({URL_Q3_FY26_EX991})."
)

DILUTION_METHODOLOGY_NOTE = (
    f"Starting shares = {SHARES_OUTSTANDING_000s:,.3f}k (post–Q3 FY2026). "
    "Shares_t = Shares_t−1 + EquityCF_t / Price; EquityCF is buyback outflow "
    f"(negative). Price policy ${SHARE_PRICE:,.0f} (= Q3 avg repurchase). "
    "SBC add-back does NOT increase shares. $/share uses Y5 buyback-adjusted shares."
)

WACC_METHODOLOGY_NOTE = (
    f"Branch Process: MODEL13 WACC = {MODEL13_WACC:.2%}; MODEL14 WACC = "
    f"{MODEL14_WACC:.2%} (band {WACC_BAND_LOW:.1%}–{WACC_BAND_HIGH:.1%}). "
    "Yacktman AAA-equity: mandated GSE Scores cash flows are closer to a "
    f"sovereign coupon than a β≈1.32 CAPM stock (CAPM ref ≈ {MODEL3_WACC:.2%} "
    "kept at R15 only). D6 is the policy rate used for XNPV."
)

EXIT_METHODOLOGY_NOTE = (
    f"Branch Process: MODEL13 exit = 21.0x; MODEL14 BASE = {EXIT_EV_EBITDA:.1f}x "
    f"(Bull {EXIT_EV_EBITDA_BULL:.1f}x / Bear {EXIT_EV_EBITDA_BEAR:.1f}x). "
    f"AAA toll-road premium vs peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x; "
    "still below FICO spot (~25–27x)."
)

GROWTH_METHODOLOGY_NOTE = (
    "Branch Process: MODEL13 revenue path = 27%/16%/13%/10%/7%. "
    f"MODEL14 = {REVENUE_GROWTH_PATH[0]:.0%}/{REVENUE_GROWTH_PATH[1]:.0%}/"
    f"{REVENUE_GROWTH_PATH[2]:.1%}/{REVENUE_GROWTH_PATH[3]:.0%}/"
    f"{REVENUE_GROWTH_PATH[4]:.1%} — lifts near-term growth for the 2026 "
    f"mortgage royalty reprice ({MORTGAGE_ROYALTY_2025:.2f}→{MORTGAGE_ROYALTY_2026:.2f}) "
    "and ongoing B2B Scores unit-price power. Terminal g raised "
    f"{PERPETUAL_GROWTH_MODEL13:.1%}→{PERPETUAL_GROWTH:.1%} (pricing > inflation)."
)

YACKTMAN_METHODOLOGY_NOTE = (
    f"Yacktman lens (MODEL13 rated {MODEL13_RATING}): treat FICO as an AAA-equivalent "
    "equity bond whose coupon (Scores royalties) is GSE-mandated and grows faster "
    f"than inflation. MODEL14 lowers WACC to {MODEL14_WACC:.2%}, lifts g to "
    f"{PERPETUAL_GROWTH:.1%}, accelerates DSO, and routes ~95%+ of incremental "
    "Scores revenue to CFO via COGS/SG&A/R&D leverage — without double-counting "
    "Deferred vs AR, and with BS identity always held."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME} (Yacktman AAA toll-road): MODEL13={MODEL13_RATING}; "
        f"WACC {MODEL13_WACC:.2%}→{MODEL14_WACC:.2%}; g {PERPETUAL_GROWTH:.1%}; "
        f"mortgage royalty ${MORTGAGE_ROYALTY_2025:.2f}→${MORTGAGE_ROYALTY_2026:.2f}; "
        f"COGS −{COGS_IMPROVEMENT_BPS:.0f}/SG&A −{SGA_IMPROVEMENT_BPS:.0f}/"
        f"R&D −{RD_IMPROVEMENT_BPS:.0f}bps; DSO {HIST_DSO_DAYS:.0f}→{TARGET_DSO_DAYS:.0f}d; "
        f"shares {SHARES_OUTSTANDING_000s/1000:.1f}M post-Q3 buybacks; "
        f"exit {EXIT_EV_EBITDA:.1f}x."
    )
