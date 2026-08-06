"""vengeanceaiUSCMODEL15.0 — Scores toll-road AAA coupon on MODEL14 scaffolding.

MODEL14 rated 7/10: fixed Deferred/OpNWC split and cut WACC to 7.1%, but still
grades the Scores bond with blended Software OpCo cost ratios (COGS floor 10%,
flat SBC%, CapEx on total rev, residual-only buybacks).

Yacktman lens: Scores royalties ≈ AAA GSE-mandated coupon (IR Scores OM ~91%).
Software is optionality — do not let Software CapEx/SBC/opex tax the coupon.

MODEL15.0 package:
  1. WACC 6.65% (AAA + thin ERP; band 6.5–6.8%); MODEL14 was 7.1%
  2. g = 3.75%; exit 23.5x
  3. Scores-led incremental margins (GM_B2B 97%, COGS floor 8%, stronger opex fade)
  4. CapEx path Scores-light; SBC% fades as Scores mix/price rises
  5. Levered buybacks (1.4× FCF) funded by incremental debt — matches Q3 aggression
  6. Op NWC ≠ Deferred (no double-count); hard BS plug
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

MODEL_NAME = "vengeanceaiUSCMODEL15.0"

# Predecessor rating
MODEL14_RATING = "7/10"

# --- Market / share base (post Q3 FY2026) ---
SHARE_PRICE = 1_149.0
SHARES_OUTSTANDING_000s = 21_597.635

BUYBACK_Q3_FY26_000s = 1_960_000.0
BUYBACK_Q3_FY26_SHARES_000s = 1_705.0
BUYBACK_Q3_FY26_AVG_PRICE = 1_149.0
Q3_FY26_FCF_000s = 370_343.0

# --- Revenue growth (MODEL14 was 30/18/14.5/11/8.5) ---
# Lift for multi-year mortgage royalty / Direct License runway (Scores +41% Q3).
REVENUE_GROWTH_PATH: Tuple[float, ...] = (0.32, 0.20, 0.16, 0.12, 0.09)
REVENUE_GROWTH_PATH_MODEL14: Tuple[float, ...] = (0.30, 0.18, 0.145, 0.11, 0.085)

# --- Operating leverage / Scores coupon economics ---
SGA_IMPROVEMENT_BPS = 200.0  # MODEL14: 175
SGA_FLOOR_PCT = 0.13  # MODEL14: 0.14
RD_IMPROVEMENT_BPS = 75.0  # MODEL14: 50
RD_FLOOR_PCT = 0.065  # MODEL14: 0.07
COGS_IMPROVEMENT_BPS = 150.0  # MODEL14: 120
COGS_FLOOR_PCT = 0.08  # MODEL14 bake used 10% — too software-like for Scores mix

# Scores segment OM (IR) — used in methodology / mix schedule notes
SCORES_SEGMENT_OM = 0.91

# --- SBC: fade as Scores price dollars rise (MODEL14 flat 8.251%) ---
SBC_PCT_REVENUE = 0.08251  # FY1
SBC_PCT_PATH: Tuple[float, ...] = (0.08251, 0.0770, 0.0720, 0.0680, 0.0640)

# --- D&A / CapEx / cash tax ---
DA_PCT_REVENUE = 0.008411
# MODEL14 CapEx 1.25%→0.60%; MODEL15 Scores-light fade (royalties need ~0 CapEx)
CAPEX_PCT_REVENUE = 0.0100
CAPEX_PCT_PATH: Tuple[float, ...] = (0.0100, 0.0080, 0.0060, 0.0045, 0.0035)
# Cash tax: keep 3yr cash tax; note guided ETR ~24% is book — do not silently cut
CASH_TAX_RATE = 0.22873

# --- Financing ---
BUYBACK_FY23_000s = 405_526.0
BUYBACK_FY24_000s = 821_702.0
BUYBACK_FY25_000s = 1_414_502.0
BUYBACK_RUNRATE_000s = (
    BUYBACK_FY23_000s + BUYBACK_FY24_000s + BUYBACK_FY25_000s
) / 3.0
DEBT_NET_RUNRATE_000s = -290_917.0  # baseline; overridden by levered-buyback math
# Buybacks = MULT × (CFO − CapEx); incremental debt funds MULT−1 (cash stays flat)
BUYBACK_FCF_MULTIPLE = 1.40  # MODEL14 was 1.0× residual after fixed debt run-rate

SAAS_MIX_SHIFT_BPS = 250.0

# --- Yacktman AAA discount + terminal ---
MODEL15_WACC = 0.0665  # mid of 6.5–6.8%
WACC_BAND_LOW = 0.0650
WACC_BAND_HIGH = 0.0680
MODEL14_WACC = 0.0710

PERPETUAL_GROWTH = 0.0375  # MODEL14: 3.5%
PERPETUAL_GROWTH_MODEL14 = 0.035

EXIT_EV_EBITDA = 23.5  # MODEL14: 22.0
EXIT_EV_EBITDA_BULL = 28.0
EXIT_EV_EBITDA_BEAR = 17.0

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

# --- Sources ---
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
URL_IR_DECK = "https://investors.fico.com/static-files/66ef723c-1f2e-4501-a452-ab2f3d02ae85"
URL_EQUIFAX_FICO_PRICING = (
    "https://www.equifax.com/newsroom/all-news/-/story/"
    "equifax-statement-on-fico-2x-price-increase-for-2026-and-mortgage-direct-license-program/"
)
URL_HAWLEY_LETTER = (
    "https://www.hawley.senate.gov/wp-content/uploads/2026/03/2026-03-23-Hawley-Oversight-Letter-to-FICO.pdf"
)
URL_HOUSINGWIRE_FICO = "https://www.housingwire.com/articles/hawley-fico-mortgage-pricing/"

FY25_REVENUE_000s = 1_990_869.0
FY25_SAAS_000s = 419_720.0
FY25_ONPREM_000s = 320_425.0
FY25_B2C_000s = 219_980.0
FY25_B2B_SCORES_000s = 948_595.0
FY25_PROF_SERVICES_000s = 82_149.0
FY25_PLATFORM_ARR_000s = 263_600.0
FY25_DEFERRED_REV_000s = 187_372.0

MORTGAGE_ROYALTY_2025 = 4.95
MORTGAGE_ROYALTY_2026 = 10.00
MORTGAGE_INDUSTRY_COST_UPLIFT_MID_M = 500.0

MIX_SAAS = FY25_SAAS_000s / FY25_REVENUE_000s
MIX_ONPREM = FY25_ONPREM_000s / FY25_REVENUE_000s
MIX_B2C = FY25_B2C_000s / FY25_REVENUE_000s
MIX_B2B = FY25_B2B_SCORES_000s / FY25_REVENUE_000s
MIX_PROF_SVCS = FY25_PROF_SERVICES_000s / FY25_REVENUE_000s

DSO_SAAS = 40.0
DSO_ONPREM = 40.0
DSO_B2C = 2.0
DSO_B2B = 25.0
DSO_PROF_SVCS = 12.0

HIST_DSO_DAYS = 97.0
TARGET_DSO_DAYS = 32.0  # MODEL14: 35 — faster Scores/SaaS cash conversion

GM_SAAS = 0.84
GM_ONPREM = 0.80
GM_B2C = 0.82
GM_B2B = 0.97  # MODEL14: 0.95 — Scores OM ~91% supports near-pure royalty GM
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
    w = (year_index + 1) / 5.0
    return HIST_DSO_DAYS * (1.0 - w) + TARGET_DSO_DAYS * w


ASSUMPTIONS_PDF_FILENAME = "MODEL15_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel15-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel15-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
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
    "SEC EX-99.1 Q3 FY2026 earnings (Scores +41%, FCF $370M)": URL_Q3_FY26_EX991,
    "FICO IR — Q3 FY2026 earnings release": URL_Q3_FY26_IR,
    "FICO IR investor presentation — Scores OM ~91%": URL_IR_DECK,
    "Equifax statement — FICO 2026 2x mortgage royalty": URL_EQUIFAX_FICO_PRICING,
    "Sen. Hawley oversight letter — FICO mortgage pricing": URL_HAWLEY_LETTER,
    "HousingWire — Hawley probes FICO mortgage pricing": URL_HOUSINGWIRE_FICO,
    "MODEL15 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL15 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "Branch Process MODEL14→MODEL15: Op NWC = AR+Inv−AP (excludes Deferred) unchanged. "
    f"DSO target {TARGET_DSO_DAYS:.0f}d (MODEL14 was 35d) — faster Scores/SaaS cash. "
    "ΔDeferred remains an explicit CFO/FCFF source (row 81). No AR/unearned double-count."
)

SEGMENT_METHODOLOGY_NOTE = (
    f"Scores segment operating margin ≈ {SCORES_SEGMENT_OM:.0%} (FICO IR deck "
    f"{URL_IR_DECK}). Q3 FY2026 Scores $458.9M (+41%); B2B +49% on mortgage unit "
    f"price ({URL_Q3_FY26_EX991}). MODEL15 raises GM_B2B to {GM_B2B:.0%} and lowers "
    f"COGS floor to {COGS_FLOOR_PCT:.0%} so Software/PS no longer grade the Scores coupon."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Deferred stays in Total Liab; +ΔDeferred is explicit in CFO/FCFF (not in Op NWC). "
    f"FY25 deferred ≈ ${FY25_DEFERRED_REV_000s/1000:.1f}M. Detail: {URL_10Q_DEFERRED_DETAIL}."
)

DA_METHODOLOGY_NOTE = (
    "TOTAL D&A add-back ≈ 0.84% of Rev (3yr avg). Intangibles amort included once."
)

SBC_METHODOLOGY_NOTE = (
    "Branch Process: MODEL14 held SBC flat at 8.251% of total Rev. MODEL15 fades "
    f"SBC% {' / '.join(f'{p:.2%}' for p in SBC_PCT_PATH)} because royalty price "
    "dollars do not require 1:1 SBC. Add-back once in CFO/FCFF; shares fall only via "
    f"buybacks (start {SHARES_OUTSTANDING_000s:,.1f}k post–Q3)."
)

CAPEX_METHODOLOGY_NOTE = (
    "Branch Process: MODEL14 CapEx 1.25%→0.60% of TOTAL rev. MODEL15 Scores-light "
    f"path {' → '.join(f'{p:.2%}' for p in CAPEX_PCT_PATH)} — mortgage/B2B royalty "
    "hikes need ~0 CapEx; do not tax the Scores coupon with Software capitalization."
)

MARGIN_METHODOLOGY_NOTE = (
    f"Branch Process: MODEL14 COGS −120 / SG&A −175 / R&D −50 bps (COGS floor 10%). "
    f"MODEL15 COGS −{COGS_IMPROVEMENT_BPS:.0f} / SG&A −{SGA_IMPROVEMENT_BPS:.0f} / "
    f"R&D −{RD_IMPROVEMENT_BPS:.0f} bps; COGS floor {COGS_FLOOR_PCT:.0%}; "
    f"GM_B2B {GM_B2B:.0%} (Scores OM ~{SCORES_SEGMENT_OM:.0%}). "
    f"Mortgage royalty ${MORTGAGE_ROYALTY_2025:.2f}→${MORTGAGE_ROYALTY_2026:.2f}."
)

CASH_TAX_METHODOLOGY_NOTE = (
    "Keep 3yr cash tax IncomeTaxesPaidNet/EBT ≈ 22.87% (not silently cut). "
    "Guided book ETR ~24–26% is not a cash-tax shortcut without XBRL support."
)

BS_METHODOLOGY_NOTE = (
    "Equity = Assets − Liab − RE every year. Liab = AP + Debt + Deferred. "
    "Row-3 OK; cash check ≈ 0. Levered buybacks raise Debt and cut Equity capital "
    "via the plug — statements stay balanced."
)

FINANCING_METHODOLOGY_NOTE = (
    f"Branch Process: MODEL14 used residual FCF after fixed debt run-rate (≈1.0×). "
    f"MODEL15 buybacks = {BUYBACK_FCF_MULTIPLE:.2f}× (CFO−CapEx); incremental debt "
    f"funds the (MULT−1) portion so ΔCash≈0. Anchored to Q3 FY2026: buybacks "
    f"${BUYBACK_Q3_FY26_000s/1000:.0f}M vs FCF ${Q3_FY26_FCF_000s/1000:.1f}M "
    f"({URL_Q3_FY26_EX991})."
)

DILUTION_METHODOLOGY_NOTE = (
    f"Start shares {SHARES_OUTSTANDING_000s:,.3f}k; price ${SHARE_PRICE:,.0f}. "
    "Shares fall with EquityCF (buybacks). SBC add-back does NOT dilute. "
    "Levered buybacks accelerate share shrink vs MODEL14 residual-only."
)

WACC_METHODOLOGY_NOTE = (
    f"Branch Process: MODEL14 WACC = {MODEL14_WACC:.2%}; MODEL15 WACC = "
    f"{MODEL15_WACC:.2%} (band {WACC_BAND_LOW:.1%}–{WACC_BAND_HIGH:.1%}). "
    f"AAA + thin ERP for GSE-mandated Scores coupons; CAPM ≈ {MODEL3_WACC:.2%} "
    "at R15 reference only."
)

EXIT_METHODOLOGY_NOTE = (
    f"Branch Process: MODEL14 exit = 22.0x; MODEL15 BASE = {EXIT_EV_EBITDA:.1f}x "
    f"(Bull {EXIT_EV_EBITDA_BULL:.1f}x / Bear {EXIT_EV_EBITDA_BEAR:.1f}x). "
    f"Perpetual pricing-power premium vs peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x."
)

GROWTH_METHODOLOGY_NOTE = (
    "Branch Process: MODEL14 revenue path = 30%/18%/14.5%/11%/8.5%. "
    f"MODEL15 = {REVENUE_GROWTH_PATH[0]:.0%}/{REVENUE_GROWTH_PATH[1]:.0%}/"
    f"{REVENUE_GROWTH_PATH[2]:.0%}/{REVENUE_GROWTH_PATH[3]:.0%}/"
    f"{REVENUE_GROWTH_PATH[4]:.0%} — extends Q3 Scores +41% / B2B +49% mortgage "
    f"unit-price momentum ({URL_Q3_FY26_EX991}). Terminal g "
    f"{PERPETUAL_GROWTH_MODEL14:.1%}→{PERPETUAL_GROWTH:.2%}."
)

YACKTMAN_METHODOLOGY_NOTE = (
    f"Yacktman lens (MODEL14 rated {MODEL14_RATING}): Scores royalties are an "
    f"AAA-equivalent coupon (IR OM ~{SCORES_SEGMENT_OM:.0%}); Software is optionality. "
    f"MODEL15 lowers WACC to {MODEL15_WACC:.2%}, lifts exit to {EXIT_EV_EBITDA:.1f}x, "
    "Stops Software CapEx/SBC/opex from taxing royalty dollars, and allows levered "
    "buybacks so the forward share count reflects proven capital returns."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME} (Scores AAA coupon): MODEL14={MODEL14_RATING}; "
        f"WACC {MODEL14_WACC:.2%}→{MODEL15_WACC:.2%}; g {PERPETUAL_GROWTH:.2%}; "
        f"exit {EXIT_EV_EBITDA:.1f}x; Scores OM~{SCORES_SEGMENT_OM:.0%}; "
        f"COGS floor {COGS_FLOOR_PCT:.0%}; CapEx→{CAPEX_PCT_PATH[-1]:.2%}; "
        f"buybacks {BUYBACK_FCF_MULTIPLE:.1f}×FCF; SBC fade; "
        f"DSO→{TARGET_DSO_DAYS:.0f}d."
    )
