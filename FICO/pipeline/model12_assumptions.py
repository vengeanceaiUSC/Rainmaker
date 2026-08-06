"""VengeanceUSCModel12.0 — Yacktman-style cash-flow optimism on MODEL11 scaffolding.

Lens: treat FICO as a long-duration, high-quality cash compounder (monopoly Scores
pricing + SaaS deferred billings + aggressive buybacks). Optimistic but sourced.

Working Capital:
  Bottom-up blended DSO; NWC = AR + Inv − AP − Deferred Revenue.
  SaaS mix shifts up over the forecast so Deferred grows → positive CFO WC swing.

MODEL12.0 Yacktman cash-flow package:
  1. Segment mix (offsets within Total Rev) + SaaS mix shift → Deferred WC cash
  2. SBC add-back in CFO/FCFF; buybacks REDUCE share count (no SBC double-penalty)
  3. CapEx % fades with operating leverage (not flat hist)
  4. Total D&A (incl. amort. of intangibles) fully added back
  5. Margin expansion from pricing power (COGS grind + faster SG&A leverage)
  6. Hard BS identity plug
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

MODEL_NAME = "VengeanceUSCModel12.0"

# FYE Sep 30, 2025 shares outstanding (10-K cover / equity footnote)
SHARES_OUTSTANDING_000s = 23_764.0

# --- Operating leverage / pricing power ---
SGA_IMPROVEMENT_BPS = 100.0  # was 75 — faster opex leverage
SGA_FLOOR_PCT = 0.15
# Scores price hikes drop to GM; grind blended COGS −40 bps/yr
COGS_IMPROVEMENT_BPS = 40.0

# --- SBC: 3yr avg ShareBasedCompensation / Revenue (FY23–25) ---
SBC_PCT_REVENUE = 0.08251

# --- D&A: 3yr avg TOTAL D&A / Revenue (includes amort. of intangibles) ---
# DepreciationDepletionAndAmortization embeds AmortizationOfIntangibleAssets
DA_PCT_REVENUE = 0.008411

# --- CapEx: start near 3yr hist avg, fade with operating leverage ---
# FY23 0.28% | FY24 1.49% | FY25 1.98% (PPE + capitalized software) → start 1.25%
# Fade toward maintenance ~0.60% by Y5 (asset-light software / Scores mix).
CAPEX_PCT_REVENUE = 0.0125  # FY1
CAPEX_PCT_PATH: Tuple[float, ...] = (0.0125, 0.0105, 0.0090, 0.0075, 0.0060)

# --- Cash taxes: 3yr avg IncomeTaxesPaidNet / EBT ---
CASH_TAX_RATE = 0.22873

# --- Financing ($000s) ---
# 10-K CF: buybacks FY23 $405.5M | FY24 $821.7M | FY25 $1,414.5M → avg ≈ $880.6M
BUYBACK_FY23_000s = 405_526.0
BUYBACK_FY24_000s = 821_702.0
BUYBACK_FY25_000s = 1_414_502.0
BUYBACK_RUNRATE_000s = (
    BUYBACK_FY23_000s + BUYBACK_FY24_000s + BUYBACK_FY25_000s
) / 3.0
DEBT_NET_RUNRATE_000s = -290_917.0

# SaaS mix shift: +150 bps/yr from on-prem → cloud (Deferred WC cash source)
SAAS_MIX_SHIFT_BPS = 150.0

# --- Exit multiples ---
EXIT_EV_EBITDA = 17.5
EXIT_EV_EBITDA_BULL = 25.0
EXIT_EV_EBITDA_BEAR = 12.8

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

# --- SEC / IR source URLs (segment mix + contract liabilities) ---
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

# ---------------------------------------------------------------------------
# FY2025 10-K disaggregated revenue ($000s) — Note 9 / consolidated statements
# Total Rev 1,990,869 = Scores 1,168,575 + Software (SaaS+on-prem) 740,145
#                        + Professional services 82,149
# Scores: B2B 948,595 + B2C (myFICO) 219,980
# Software: SaaS 419,720 + On-premises 320,425
# Platform ARR (disclosure, not FY revenue): $263.6M = 35% of software ARR
# ---------------------------------------------------------------------------
FY25_REVENUE_000s = 1_990_869.0
FY25_SAAS_000s = 419_720.0
FY25_ONPREM_000s = 320_425.0
FY25_B2C_000s = 219_980.0
FY25_B2B_SCORES_000s = 948_595.0
FY25_PROF_SERVICES_000s = 82_149.0
FY25_PLATFORM_ARR_000s = 263_600.0  # ARR metric at 9/30/2025 (not income-stmt line)

MIX_SAAS = FY25_SAAS_000s / FY25_REVENUE_000s  # 21.08%
MIX_ONPREM = FY25_ONPREM_000s / FY25_REVENUE_000s  # 16.09%
MIX_B2C = FY25_B2C_000s / FY25_REVENUE_000s  # 11.05%
MIX_B2B = FY25_B2B_SCORES_000s / FY25_REVENUE_000s  # 47.65%
MIX_PROF_SVCS = FY25_PROF_SERVICES_000s / FY25_REVENUE_000s  # 4.13%

# Segment cash-conversion DSOs (days) — policy from 10-K payment terms (30–60)
# and product economics. B2C ≈ card-settled; B2B Scores standard trade terms.
DSO_SAAS = 45.0  # midpoint of 30–60; SaaS billed annually in advance → Deferred
DSO_ONPREM = 45.0
DSO_B2C = 2.0  # near-zero (instant card / subscription collection)
DSO_B2B = 30.0
DSO_PROF_SVCS = 15.0  # implementation fees; faster cash than software AR

# Segment gross margins (calibrated so blended COGS ≈ FY25 ~17.8%)
GM_SAAS = 0.80
GM_ONPREM = 0.80
GM_B2C = 0.78
GM_B2B = 0.90
GM_PROF_SVCS = 0.28  # services carry significantly lower GM than software/scores


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


ASSUMPTIONS_PDF_FILENAME = "MODEL12_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceuscmodel12-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceuscmodel12-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
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
    "MODEL12 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL12 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "Working Capital uses segment cash-conversion DSOs blended by FY2025 10-K "
    "revenue mix (SaaS, on-prem software, B2C myFICO, B2B Scores, Professional "
    "services). NWC = AR + Inventory − AP − Deferred Revenue (output); "
    "ΔNWC = NWCt − NWCt−1. Rising Deferred (SaaS mix shift) is a cash SOURCE in "
    "CFO because it reduces NWC. Mix % offset within Total Revenue — no second "
    "revenue stack."
)

SEGMENT_METHODOLOGY_NOTE = (
    "FY2025 10-K disaggregation ($000s): B2B Scores 948,595; B2C myFICO 219,980; "
    "SaaS software 419,720; on-premises software 320,425; Professional services "
    "82,149 (sum = Total revenues 1,990,869). FICO Platform ARR was $263.6M "
    "(35% of software ARR) at 9/30/2025 — ARR is a KPI, not an incremental IS line. "
    "MODEL12 shifts +150 bps/yr from on-prem mix into SaaS to reflect the cloud "
    "transition (drives Deferred Revenue growth and CFO WC cash)."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Deferred Revenue is a BS liability in Total Liabilities and in NWC "
    "(= AR+Inv−AP−Deferred). Forecast Deferred = Total Rev × (SaaS%+OnPrem%) × "
    "(FY25 Deferred ÷ FY25 software Rev). As SaaS% rises, Deferred rises faster "
    "than revenue → negative ΔNWC → positive CFO. Scores/myFICO are NOT added "
    "into Deferred (no double count). Hist Deferred/Rev rose 8.0%→9.4% (FY21–25)."
)

DA_METHODOLOGY_NOTE = (
    "CFO/FCFF add back TOTAL D&A at ≈ 0.84% of Revenue (3yr avg "
    "DepreciationDepletionAndAmortization / Rev). That XBRL total already includes "
    "AmortizationOfIntangibleAssets — intangibles amortization is fully added back "
    "once (no separate second add-back)."
)

SBC_METHODOLOGY_NOTE = (
    "SBC is added back once in CFO and FCFF at ≈ 8.25% of Revenue "
    "(ShareBasedCompensation). MODEL12 does NOT also dilute shares by SBC$/Price "
    "(that would double-penalize). Share count instead falls with buybacks."
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
    f"segment GM base (B2B Scores price/volume mix); SG&A grinds "
    f"−{SGA_IMPROVEMENT_BPS:.0f} bps/yr toward a {SGA_FLOOR_PCT:.0%} floor. "
    "FY25 10-K: Scores revenue +$249M YoY 'primarily attributable to a higher unit price'."
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
    "June 2025 authorization: $1.0B program."
)

DILUTION_METHODOLOGY_NOTE = (
    "Starting shares = 23,764k (FYE Sep 30, 2025 outstanding). Each forecast year: "
    "Shares_t = Shares_t−1 + EquityIssuanceCF_t / Price. EquityIssuanceCF is the "
    "buyback outflow (negative) from 3S row 20 — so repurchases REDUCE the share "
    "count. SBC is added back in FCFF but does NOT increase shares (avoids double "
    "penalty). $/share uses Year-5 buyback-adjusted shares."
)

YACKTMAN_METHODOLOGY_NOTE = (
    "Yacktman lens: FICO is valued as a long-duration, high-quality cash stream — "
    "monopolistic Scores pricing, rising SaaS deferred billings, and persistent "
    "buybacks. Model bias is optimistic on cash conversion (Deferred WC, CapEx fade, "
    "margin grind) while keeping BS identity and no double-counted add-backs."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME} (Yacktman CF): Deferred WC + SBC add-back; "
        f"buybacks cut shares (no SBC dilution); "
        f"CapEx fade {CAPEX_PCT_PATH[0]:.2%}→{CAPEX_PCT_PATH[-1]:.2%}; "
        f"COGS −{COGS_IMPROVEMENT_BPS:.0f}bps/SG&A −{SGA_IMPROVEMENT_BPS:.0f}bps; "
        f"DA+amort {DA_PCT_REVENUE:.2%}; exit {EXIT_EV_EBITDA:.1f}x; "
        f"WACC={MODEL3_WACC:.2%}."
    )
