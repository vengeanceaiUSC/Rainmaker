"""vengeanceaiUSCMODEL11 — segment cash conversion + hard BS plug + SBC dilution.

Grounded in SEC XBRL companyfacts (CIK 0000814547) and FY2025 10-K Note 9 / MD&A
disaggregated revenue (Scores B2B/B2C, Software SaaS/on-prem, Professional services).

Working Capital:
  Bottom-up blended DSO from segment cash-conversion days;
  NWC = AR + Inv − AP − Deferred Revenue (output).
  Deferred Revenue is a BS liability (rolled into Total Liabilities) driven by
  SaaS + on-prem software mix (annual-in-advance billings).

MODEL11 additions vs MODEL10:
  1. Revenue mix disaggregation (offsets within total Rev — no double count)
  2. Segment DSO / GM → blended AR days + COGS%
  3. SaaS-driven Deferred Revenue in NWC / FCFF
  4. Hard BS identity plug (Equity = Assets − Liab − RE)
  5. SBC-linked share dilution schedule in DCF $/share

Keeps MODEL10 cash-flow realism:
  SBC add-back, DA% sales, flat CapEx%, cash tax, financing run-rates
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
    SHARES_OUTSTANDING_000s,
    TOTAL_DEBT_10Q_000s,
)
from .model5_assumptions import SOURCE_LINKS as _M5_LINKS
from .wacc import MODEL3_WACC

MODEL_NAME = "vengeanceaiUSCMODEL11"

# --- SGA (unchanged operating leverage path) ---
SGA_IMPROVEMENT_BPS = 75.0
SGA_FLOOR_PCT = 0.15

# --- SBC: 3yr avg ShareBasedCompensation / Revenue (FY23–25) ---
# FY23 8.18% | FY24 8.70% | FY25 7.87% → 8.25%
SBC_PCT_REVENUE = 0.08251

# --- D&A: 3yr avg total D&A / Revenue (not PP&E-only) ---
DA_PCT_REVENUE = 0.008411

# --- CapEx: 3yr avg (PPE purchases + software capitalization) / Revenue ---
CAPEX_PCT_REVENUE = 0.01249

# --- Cash taxes: 3yr avg IncomeTaxesPaidNet / EBT ---
CASH_TAX_RATE = 0.22873

# --- Financing run-rates ($000s, 3yr absolute averages) ---
BUYBACK_RUNRATE_000s = 880_577.0
DEBT_NET_RUNRATE_000s = -290_917.0

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


ASSUMPTIONS_PDF_FILENAME = "MODEL11_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel11-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel11-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
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
    "MODEL11 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL11 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "Working Capital uses segment cash-conversion DSOs blended by FY2025 10-K "
    "revenue mix (SaaS, on-prem software, B2C myFICO, B2B Scores, Professional "
    "services). NWC = AR + Inventory − AP − Deferred Revenue (output); "
    "ΔNWC = NWCt − NWCt−1. Mix percentages offset within Total Revenue — they "
    "do not add a second revenue stack."
)

SEGMENT_METHODOLOGY_NOTE = (
    "FY2025 10-K disaggregation ($000s): B2B Scores 948,595; B2C myFICO 219,980; "
    "SaaS software 419,720; on-premises software 320,425; Professional services "
    "82,149 (sum = Total revenues 1,990,869). FICO Platform ARR was $263.6M "
    "(35% of software ARR) at 9/30/2025 — ARR is a KPI, not an incremental IS line. "
    "SaaS/on-prem billings create Deferred Revenue (contract liability); B2C is "
    "near-zero DSO (card-settled); B2B Scores use ~30-day trade terms; Professional "
    "services convert cash quickly but at much lower gross margin (~28% vs "
    "software/scores)."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Deferred Revenue is an explicit Balance Sheet liability included in Total "
    "Liabilities and in NWC (= AR+Inv−AP−Deferred). Forecast Deferred = "
    "Total Rev × (SaaS% + OnPrem%) × (FY25 Deferred ÷ FY25 software revenue). "
    "Only software subscription / maintenance billings drive the liability — "
    "Scores and myFICO are not double-counted into Deferred."
)

DA_METHODOLOGY_NOTE = (
    "D&A is forecast as a % of Revenue (3yr FY23–25 average of total "
    "DepreciationDepletionAndAmortization / Revenue ≈ 0.84%), not a PP&E-only "
    "rate — capturing software amortization and intangibles."
)

SBC_METHODOLOGY_NOTE = (
    "Stock-Based Compensation is added back in CFO and FCFF at the 3yr average "
    "SBC/Revenue (≈ 8.25%) from ShareBasedCompensation in the 10-K cash flow. "
    "DCF shares outstanding grow each year by SBC$ / share price so Equity "
    "Value/Share reflects dilution from the SBC add-back."
)

CAPEX_METHODOLOGY_NOTE = (
    "CapEx uses a flat 3yr average of (PPE purchases + capitalized software) / "
    "Revenue (≈ 1.25%). The prior hardcoded fade to 1% was removed so reinvestment "
    "reflects actual history rather than forcing CapEx ≈ D&A by Year 5."
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
    "Financing CF restores realism outside FCFF: debt uses the 3yr average net debt "
    "cash flow (senior-note proceeds − line-of-credit repayments ≈ −$291M/yr). "
    "Equity buybacks distribute residual levered FCF after that debt CF so cash does "
    "not artificially stockpile (historical buybacks averaged ≈ $881M/yr in FY23–25)."
)

DILUTION_METHODOLOGY_NOTE = (
    "Starting diluted shares (D12) roll forward each forecast year: "
    "Shares_t = Shares_t−1 + SBC_t($000s) / Current Price. SBC$ comes from "
    "3-statement row 64 (Rev × 8.25% SBC). Terminal Equity Value/Share uses "
    "Year-5 diluted shares, not the static starting share count."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: segment WC (SaaS/B2C/B2B/PS); "
        f"blended DSO≈{blended_dso():.0f}d; Deferred in NWC; "
        f"SBC {SBC_PCT_REVENUE:.1%} + share dilution; "
        f"DA={DA_PCT_REVENUE:.2%}; CapEx={CAPEX_PCT_REVENUE:.2%} flat; "
        f"cash tax={CASH_TAX_RATE:.1%}; exit {EXIT_EV_EBITDA:.1f}x; "
        f"WACC={MODEL3_WACC:.2%}."
    )
