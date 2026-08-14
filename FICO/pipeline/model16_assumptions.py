"""vengeanceaiUSCMODEL16.0 — CAPM-primary upgrade directly from MODEL13.

MODEL13 rated 6.5/10 under Yacktman yield lens: Deferred/OpNWC split and SBC
add-back exist, but DSO still floors at 45d (harsh vs Scores cash) and D6 uses a
hardcoded Yacktman 7.8% instead of the live CAPM bond-yield stack.

Yacktman: "Think of stocks as though they were bonds… What credit rating would
you give it?" → discount at a mathematically justified CAPM WACC (Rf+β×ERP),
while the cash-flow side captures SaaS Deferred + SBC + Scores conversion.

MODEL16.0 package (Model13 → Model16):
  1. D6 = live CAPM (D6←R15); retire MODEL13's 7.8% policy override
  2. Deferred growth tied to SaaS/on-prem mix; +ΔDeferred strict CFO add-back
  3. SBC % of Rev from companyfacts; explicit CFO/FCFF add-back (≠ D&A)
  4. DSO 97→28d (was 97→45d) for Scores near-instant cash
  5. Faster SaaS mix shift (+250 bps) to grow Deferred cash source
  6. Hard BS plug every year; cash check ≈ 0
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
from .wacc import MODEL3_WACC, WaccInputs

MODEL_NAME = "vengeanceaiUSCMODEL16.0"

# Predecessor
MODEL13_RATING = "6.5/10"
MODEL13_WACC = 0.0780  # MODEL13 Yacktman policy (retired as D6)

SHARES_OUTSTANDING_000s = 23_764.0

# --- Operating leverage (carry MODEL13; document in Branch Process) ---
SGA_IMPROVEMENT_BPS = 125.0
SGA_FLOOR_PCT = 0.15
COGS_IMPROVEMENT_BPS = 75.0
RD_IMPROVEMENT_BPS = 0.0  # MODEL13 held R&D flat via hist ratio equation
RD_FLOOR_PCT = 0.08

# --- SBC: 3yr avg ShareBasedCompensation / Revenue (FY23–25) ---
SBC_PCT_REVENUE = 0.08251

# --- D&A / CapEx / cash tax ---
DA_PCT_REVENUE = 0.008411
CAPEX_PCT_REVENUE = 0.0125
CAPEX_PCT_PATH: Tuple[float, ...] = (0.0125, 0.0105, 0.0090, 0.0075, 0.0060)
CASH_TAX_RATE = 0.22873
BOOK_TAX_RATE_FY25 = 0.1877  # ~18.8% for CAPM Rd(1−T) reference

# --- Financing ---
BUYBACK_FY23_000s = 405_526.0
BUYBACK_FY24_000s = 821_702.0
BUYBACK_FY25_000s = 1_414_502.0
BUYBACK_RUNRATE_000s = (
    BUYBACK_FY23_000s + BUYBACK_FY24_000s + BUYBACK_FY25_000s
) / 3.0
DEBT_NET_RUNRATE_000s = -290_917.0

# SaaS mix shift: faster cloud → more Deferred CFO cash
SAAS_MIX_SHIFT_BPS = 250.0  # MODEL13: 200

# --- CAPM-primary WACC (D6 = R15; live stack) ---
_WACC_IN = WaccInputs()
MODEL16_WACC = MODEL3_WACC  # ≈ 9.24%; Excel D6 formula-linked to R15
WACC_BAND_LOW = round(MODEL16_WACC - 0.005, 4)
WACC_BAND_HIGH = round(MODEL16_WACC + 0.005, 4)

RF_RATE = _WACC_IN.risk_free_rate
ERP_RATE = _WACC_IN.equity_risk_premium
BETA = _WACC_IN.beta
PRE_TAX_RD = _WACC_IN.pre_tax_cost_of_debt
COST_OF_EQUITY = _WACC_IN.cost_of_equity
AFTER_TAX_RD = _WACC_IN.after_tax_cost_of_debt

# --- Exit multiples (carry MODEL13 base) ---
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
URL_10Q_Q1_FY26_IR = (
    "https://investors.fico.com/static-files/935bc755-fe08-44bd-9898-d04a4e3c92f1"
)
URL_10Q_DEFERRED_DETAIL = (
    "https://www.sec.gov/Archives/edgar/data/814547/000081454726000011/R45.htm"
)
URL_FRED_DGS10 = "https://fred.stlouisfed.org/series/DGS10"
URL_YAHOO_BETA = "https://finance.yahoo.com/quote/FICO/key-statistics/"
URL_YAHOO_TNX = "https://finance.yahoo.com/quote/%5ETNX/"
URL_DAMODARAN_ERP = "https://pages.stern.nyu.edu/adamodar/New_Home_Page/home.htm"
URL_8K_NOTES = "https://www.sec.gov/Archives/edgar/data/814547/000119312526117936/d56220d8k.htm"
URL_TREASURY_CURVE = (
    "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
    "TextView?type=daily_treasury_yield_curve"
)

FY25_REVENUE_000s = 1_990_869.0
FY25_SAAS_000s = 419_720.0
FY25_ONPREM_000s = 320_425.0
FY25_B2C_000s = 219_980.0
FY25_B2B_SCORES_000s = 948_595.0
FY25_PROF_SERVICES_000s = 82_149.0
FY25_PLATFORM_ARR_000s = 263_600.0
FY25_DEFERRED_REV_000s = 187_372.0

MIX_SAAS = FY25_SAAS_000s / FY25_REVENUE_000s
MIX_ONPREM = FY25_ONPREM_000s / FY25_REVENUE_000s
MIX_B2C = FY25_B2C_000s / FY25_REVENUE_000s
MIX_B2B = FY25_B2B_SCORES_000s / FY25_REVENUE_000s
MIX_PROF_SVCS = FY25_PROF_SERVICES_000s / FY25_REVENUE_000s

DSO_SAAS = 40.0  # MODEL13: 45
DSO_ONPREM = 40.0
DSO_B2C = 2.0
DSO_B2B = 25.0  # MODEL13: 30 — Scores settle faster
DSO_PROF_SVCS = 12.0

HIST_DSO_DAYS = 97.0
TARGET_DSO_DAYS = 28.0  # MODEL13: 45 — Scores near-instant cash floor

GM_SAAS = 0.80
GM_ONPREM = 0.80
GM_B2C = 0.78
GM_B2B = 0.90
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


ASSUMPTIONS_PDF_FILENAME = "MODEL16_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel16-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel16-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
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
    "FRED DGS10 — US 10Y risk-free rate (Rf)": URL_FRED_DGS10,
    "Yahoo ^TNX — 10Y yield cross-check": URL_YAHOO_TNX,
    "US Treasury daily par yield curve": URL_TREASURY_CURVE,
    "Yahoo FICO key statistics — 5Y monthly Beta": URL_YAHOO_BETA,
    "Damodaran Implied ERP": URL_DAMODARAN_ERP,
    "SEC 8-K — 6.250% Senior Notes due 2034 (Rd)": URL_8K_NOTES,
    "MODEL16 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL16 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "Updated DSO from Model13 (Previous: 97→45d) to Model16 (New: 97→28d). "
    "Op NWC = AR+Inv−AP ONLY. Increase in Deferred Revenue is a SEPARATE positive "
    f"CFO/FCFF add-back tied to SaaS/on-prem mix (+{SAAS_MIX_SHIFT_BPS:.0f}bps/yr). "
    f"No AR/unearned double-count. Source: {URL_10Q_DEFERRED_DETAIL}."
)

SEGMENT_METHODOLOGY_NOTE = (
    "FY2025 10-K disaggregation drives mix rows 104–108. "
    f"Updated SaaS mix shift from Model13 (Previous: +200 bps/yr) to Model16 "
    f"(New: +{SAAS_MIX_SHIFT_BPS:.0f} bps/yr) to grow Deferred cash source. "
    f"Platform ARR ${FY25_PLATFORM_ARR_000s/1000:.1f}M is KPI only — not additive IS revenue."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Updated Deferred from Model13 (Previous: present but under-weighted vs Scores cash) "
    "to Model16 (New: explicit BS liability + CFO row 81). "
    "Deferred_t = Rev × (SaaS%+OnPrem%) × (FY25 Def / FY25 soft Rev). "
    f"FY25 deferred ≈ ${FY25_DEFERRED_REV_000s/1000:.1f}M. "
    f"Detail: {URL_10Q_DEFERRED_DETAIL}."
)

DA_METHODOLOGY_NOTE = (
    "TOTAL D&A add-back ≈ 0.84% of Rev (3yr avg). SBC is a SEPARATE add-back — "
    "no double-count with D&A / intangibles amort."
)

SBC_METHODOLOGY_NOTE = (
    "Updated SBC from Model13 (Previous: 8.251% present but user CSV inspection "
    "showed weak CFO integration clarity) to Model16 (New: yellow POLICY row 21 = "
    f"{SBC_PCT_REVENUE:.2%} from SEC companyfacts ShareBasedCompensation/Rev; "
    f"CFO row 64 = Rev×SBC%; FCFF add-back once). Source: {URL_FACTS}. "
    "Shares fall only via buybacks — no SBC$/Price dilution."
)

CAPEX_METHODOLOGY_NOTE = (
    "Updated CapEx path documentation (unchanged levels vs Model13): "
    f"{' → '.join(f'{p:.2%}' for p in CAPEX_PCT_PATH)}."
)

MARGIN_METHODOLOGY_NOTE = (
    f"COGS −{COGS_IMPROVEMENT_BPS:.0f} bps/yr from segment GM; SG&A −"
    f"{SGA_IMPROVEMENT_BPS:.0f} bps toward {SGA_FLOOR_PCT:.0%} floor "
    "(carry Model13 Scores pricing-power grind)."
)

CASH_TAX_METHODOLOGY_NOTE = (
    f"DCF cash tax {CASH_TAX_RATE:.2%} (3yr IncomeTaxesPaidNet/EBT). "
    f"CAPM after-tax Rd uses book ETR ~{BOOK_TAX_RATE_FY25:.1%} in WaccInputs; "
    "Excel R11 links to D5 for Rd(1−t) consistency with FCFF cash taxes."
)

BS_METHODOLOGY_NOTE = (
    "Equity = Assets − Liab − RE every year. Liab = AP + Debt + Deferred. "
    "Row-3 OK; cash check ≈ 0."
)

FINANCING_METHODOLOGY_NOTE = (
    "Debt = 3yr avg net debt CF; equity buybacks = residual (CFO−CapEx−Debt) so "
    f"ΔCash≈0. Hist buybacks avg ≈ ${BUYBACK_RUNRATE_000s/1000:.0f}M/yr."
)

DILUTION_METHODOLOGY_NOTE = (
    f"Start shares {SHARES_OUTSTANDING_000s:,.0f}k (FYE25). "
    "Shares_t += EquityCF_t/Price; EquityCF is buybacks (<0). "
    "SBC add-back does NOT dilute."
)

WACC_METHODOLOGY_NOTE = (
    f"Updated WACC from Model13 (Previous: Yacktman policy {MODEL13_WACC:.2%}) "
    f"to Model16 (New: live CAPM {MODEL16_WACC:.2%} = D6←R15). "
    f"Rf={RF_RATE:.2%} (FRED DGS10 / ^TNX), β={BETA:.2f} (Yahoo 5Y monthly), "
    f"ERP={ERP_RATE:.2%} (Damodaran), Rd_pre={PRE_TAX_RD:.3%} (8-K 6.250% notes), "
    f"T≈{BOOK_TAX_RATE_FY25:.1%}, Ke={COST_OF_EQUITY:.2%}, "
    f"Rd_at={AFTER_TAX_RD:.2%}. "
    f"Sources: {URL_FRED_DGS10} | {URL_YAHOO_BETA} | {URL_DAMODARAN_ERP} | {URL_8K_NOTES}."
)

EXIT_METHODOLOGY_NOTE = (
    f"Exit BASE {EXIT_EV_EBITDA:.1f}x unchanged vs Model13 "
    f"(Bull {EXIT_EV_EBITDA_BULL:.1f}x / Bear {EXIT_EV_EBITDA_BEAR:.1f}x)."
)

GROWTH_METHODOLOGY_NOTE = (
    "Revenue growth path unchanged vs Model13: "
    + " / ".join(f"{g:.0%}" for g in REVENUE_GROWTH_PATH)
    + f"; g={PERPETUAL_GROWTH:.1%}."
)

YACKTMAN_METHODOLOGY_NOTE = (
    f"Yacktman lens (MODEL13 rated {MODEL13_RATING}): compare FICO as a bond via "
    f"CAPM yield (WACC {MODEL16_WACC:.2%}), not a hardcoded override. "
    "Cash side: Deferred + SBC + lower DSO so Scores/SaaS cash is not trapped."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME} (CAPM-primary from MODEL13): "
        f"MODEL13={MODEL13_RATING}; WACC {MODEL13_WACC:.2%}→CAPM {MODEL16_WACC:.2%} "
        f"(D6=R15); DSO {HIST_DSO_DAYS:.0f}→{TARGET_DSO_DAYS:.0f}d; "
        f"+ΔDeferred in CFO; SBC {SBC_PCT_REVENUE:.2%} add-back; "
        f"exit {EXIT_EV_EBITDA:.1f}x; g={PERPETUAL_GROWTH:.1%}."
    )
