"""vengeanceaiUSCMODEL18.0 — Stub-period DCF + grounded SBC on MODEL17.

MODEL17 rated 7.0/10: Yacktman-adj CAPM + Scores CF plumbing work, but
(1) DCF still discounts a FULL Year-1 FCFF from hist FYE 9/30/2025 mid-year
dates even though ~35% of the forward year from 3/30/2026 has already elapsed
as of 8/6/2026 — overstating near-term PV; (2) SBC still starts at 8.25% and
taxes Scores royalty dollars; (3) Platform ARR +62% / SaaS +21% (Q3) under-
recognized in mix shift vs on-prem runoff.

Yacktman: treat Scores as AAA coupons; only discount cash that has NOT yet
occurred; fade SBC toward a maintenance grant rate; accelerate SaaS/Platform
Deferred cash.

MODEL18.0 package:
  1. Valuation date 2026-08-06; FY stub from 3/30/2026 (129/365 elapsed)
  2. Y1 FCFF × 64.66% remaining; mid-stub date for XNPV (not full-year mid)
  3. SBC fade 7.00%→5.00% (suspicious but grounded)
  4. SaaS mix +300 bps/yr (Platform ARR $413M, +62% Q3)
  5. CapEx 0.85%→0.35%; keep Yacktman-adj CAPM D6=R15
  6. Op NWC ≠ Deferred; BS identity Equity = A − L − RE
"""

from __future__ import annotations

from datetime import date, timedelta
from typing import Dict, List, Tuple

from .model3_assumptions import (  # noqa: F401
    CASH_10Q_000s,
    MKT_SECS_10Q_000s,
    RESTRUCTURING_NORMALIZE_000s,
    TOTAL_DEBT_10Q_000s,
)
from .model5_assumptions import SOURCE_LINKS as _M5_LINKS
from .wacc import (
    BETA_AAA_COUPON,
    BETA_BLUME,
    BETA_RAW,
    BETA_YACKTMAN,
    MODEL3_WACC,
    MODEL17_CAPM_WACC,
    WaccInputs,
    YACKTMAN_WACC_INPUTS,
)

MODEL_NAME = "vengeanceaiUSCMODEL18.0"

# --- MODEL17 critique (Phase 1) ---
MODEL17_RATING = "7.0/10"
MODEL17_CRITIQUE = (
    "MODEL17 flaws affecting CF→DCF: (a) full-year FY1 FCFF still in XNPV while "
    "valuation is mid-stub (overstates PV); (b) SBC path starts at 8.25% and "
    "penalizes Scores price-ups that need little incremental grant cost; "
    "(c) SaaS/Platform mix shift +250 bps understates Q3 Platform ARR +62% / "
    "SaaS +21% Deferred cash; (d) CapEx path still slightly taxes coupon cash. "
    "IS interest grows with levered buyback debt but FCFF correctly starts at "
    "EBIT — not a double-count. Equity=A−L−RE is BS identity, not a soft plug."
)

MODEL17_WACC = MODEL17_CAPM_WACC  # predecessor D6 ≈ 7.68%
MODEL16_WACC = MODEL3_WACC  # raw CAPM ≈ 9.27% reference

# --- Market / share base (post Q3 FY2026) ---
SHARE_PRICE = 1_149.0
SHARES_OUTSTANDING_000s = 21_597.635

BUYBACK_Q3_FY26_000s = 1_960_000.0
BUYBACK_Q3_FY26_SHARES_000s = 1_705.0
BUYBACK_Q3_FY26_AVG_PRICE = 1_149.0
Q3_FY26_FCF_000s = 370_343.0

# --- Stub-period temporal (Phase 3) — valuation date = 2026-08-06 ---
VALUATION_DATE = date(2026, 8, 6)
FY_STUB_START = date(2026, 3, 30)
FY_STUB_END = date(2027, 3, 29)
STUB_TOTAL_DAYS = 365
STUB_ELAPSED_DAYS = (VALUATION_DATE - FY_STUB_START).days  # 129
STUB_REMAINING_DAYS = STUB_TOTAL_DAYS - STUB_ELAPSED_DAYS  # 236
STUB_FRACTION_ELAPSED = STUB_ELAPSED_DAYS / STUB_TOTAL_DAYS  # ≈ 0.3534
STUB_FRACTION_REMAINING = 1.0 - STUB_FRACTION_ELAPSED  # ≈ 0.6466
STUB_MID_OFFSET_DAYS = STUB_REMAINING_DAYS // 2  # 118 — mid of remaining stub
STUB_MID_DATE = VALUATION_DATE + timedelta(days=STUB_MID_OFFSET_DAYS)

# --- Revenue growth (MODEL17 was 30/18/14.5/11/8.5) ---
REVENUE_GROWTH_PATH: Tuple[float, ...] = (0.30, 0.18, 0.145, 0.11, 0.085)
REVENUE_GROWTH_PATH_MODEL17: Tuple[float, ...] = (0.30, 0.18, 0.145, 0.11, 0.085)

# --- Operating leverage / Scores coupon ---
SGA_IMPROVEMENT_BPS = 175.0
SGA_FLOOR_PCT = 0.14
RD_IMPROVEMENT_BPS = 50.0
RD_FLOOR_PCT = 0.07
COGS_IMPROVEMENT_BPS = 120.0
COGS_FLOOR_PCT = 0.08
SCORES_SEGMENT_OM = 0.91

# --- SBC fade (MODEL17: 8.25→6.60; MODEL18 grounded lower) ---
SBC_PCT_REVENUE = 0.0700  # Y1 policy (was 8.251%)
SBC_PCT_PATH: Tuple[float, ...] = (0.0700, 0.0650, 0.0600, 0.0550, 0.0500)
SBC_PCT_PATH_MODEL17: Tuple[float, ...] = (0.08251, 0.0780, 0.0740, 0.0700, 0.0660)

# --- D&A / CapEx / cash tax ---
DA_PCT_REVENUE = 0.008411
CAPEX_PCT_REVENUE = 0.0085  # MODEL17 FY1 was 1.00%
CAPEX_PCT_PATH: Tuple[float, ...] = (0.0085, 0.0070, 0.0055, 0.0045, 0.0035)
CAPEX_PCT_PATH_MODEL17: Tuple[float, ...] = (0.0100, 0.0085, 0.0070, 0.0055, 0.0040)
CASH_TAX_RATE = 0.22873
BOOK_TAX_RATE_FY25 = 0.1877

# --- Financing ---
BUYBACK_FY23_000s = 405_526.0
BUYBACK_FY24_000s = 821_702.0
BUYBACK_FY25_000s = 1_414_502.0
BUYBACK_RUNRATE_000s = (
    BUYBACK_FY23_000s + BUYBACK_FY24_000s + BUYBACK_FY25_000s
) / 3.0
DEBT_NET_RUNRATE_000s = -290_917.0
BUYBACK_FCF_MULTIPLE = 1.40

SAAS_MIX_SHIFT_BPS = 300.0  # MODEL17: 250 — Platform ARR +62% Q3

# --- Yacktman-adjusted CAPM (primary D6) — same algebra as MODEL17 ---
MODEL18_WACC = MODEL17_CAPM_WACC
WACC_BAND_LOW = 0.0650
WACC_BAND_HIGH = 0.0780
RF_RATE = YACKTMAN_WACC_INPUTS.risk_free_rate
ERP_RATE = YACKTMAN_WACC_INPUTS.equity_risk_premium
BETA = BETA_YACKTMAN
PRE_TAX_RD = YACKTMAN_WACC_INPUTS.pre_tax_cost_of_debt
COST_OF_EQUITY = YACKTMAN_WACC_INPUTS.cost_of_equity
AFTER_TAX_RD = YACKTMAN_WACC_INPUTS.after_tax_cost_of_debt

PERPETUAL_GROWTH = 0.035
PERPETUAL_GROWTH_MODEL17 = 0.035

EXIT_EV_EBITDA = 23.0
EXIT_EV_EBITDA_BULL = 27.0
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
URL_10Q_Q1_FY26_IR = (
    "https://investors.fico.com/static-files/935bc755-fe08-44bd-9898-d04a4e3c92f1"
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
URL_Q3_FY26_TRANSCRIPT = "https://stockanalysis.com/stocks/fico/transcripts/656443-q3-2026/"
URL_IR_DECK = "https://investors.fico.com/static-files/66ef723c-1f2e-4501-a452-ab2f3d02ae85"
URL_EQUIFAX_FICO_PRICING = (
    "https://www.equifax.com/newsroom/all-news/-/story/"
    "equifax-statement-on-fico-2x-price-increase-for-2026-and-mortgage-direct-license-program/"
)
URL_HAWLEY_LETTER = (
    "https://www.hawley.senate.gov/wp-content/uploads/2026/03/2026-03-23-Hawley-Oversight-Letter-to-FICO.pdf"
)
URL_FRED_DGS10 = "https://fred.stlouisfed.org/series/DGS10"
URL_YAHOO_BETA = "https://finance.yahoo.com/quote/FICO/key-statistics/"
URL_YAHOO_TNX = "https://finance.yahoo.com/quote/%5ETNX/"
URL_DAMODARAN_ERP = "https://pages.stern.nyu.edu/adamodar/New_Home_Page/home.htm"
URL_8K_NOTES = "https://www.sec.gov/Archives/edgar/data/814547/000119312526117936/d56220d8k.htm"
URL_BLUME_BETA = "https://www.jstor.org/stable/2329867"

FY25_REVENUE_000s = 1_990_869.0
FY25_SAAS_000s = 419_720.0
FY25_ONPREM_000s = 320_425.0
FY25_B2C_000s = 219_980.0
FY25_B2B_SCORES_000s = 948_595.0
FY25_PROF_SERVICES_000s = 82_149.0
FY25_PLATFORM_ARR_000s = 263_600.0
Q3_FY26_PLATFORM_ARR_000s = 413_000.0  # Platform ARR $413M (+62% YoY)
Q3_FY26_SOFTWARE_ARR_000s = 816_000.0  # Total Software ARR $816M (+10%)
FY25_DEFERRED_REV_000s = 187_372.0

MORTGAGE_ROYALTY_2025 = 4.95
MORTGAGE_ROYALTY_2026 = 10.00

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
TARGET_DSO_DAYS = 28.0

GM_SAAS = 0.84
GM_ONPREM = 0.80
GM_B2C = 0.82
GM_B2B = 0.97
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


ASSUMPTIONS_PDF_FILENAME = "MODEL18_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel18-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel18-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
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
    "SEC EX-99.1 Q3 FY2026 earnings (Scores +41%, Platform ARR +62%)": URL_Q3_FY26_EX991,
    "FICO IR — Q3 FY2026 earnings release": URL_Q3_FY26_IR,
    "FICO Q3 FY2026 earnings call transcript (Platform ARR $413M)": URL_Q3_FY26_TRANSCRIPT,
    "FICO IR investor presentation — Scores OM ~91%": URL_IR_DECK,
    "Equifax statement — FICO 2026 2x mortgage royalty": URL_EQUIFAX_FICO_PRICING,
    "Sen. Hawley oversight letter — FICO mortgage pricing": URL_HAWLEY_LETTER,
    "FRED DGS10 — US 10Y risk-free rate (Rf)": URL_FRED_DGS10,
    "Yahoo ^TNX — 10Y yield cross-check": URL_YAHOO_TNX,
    "Yahoo FICO key statistics — 5Y monthly Beta (raw)": URL_YAHOO_BETA,
    "Damodaran Implied ERP": URL_DAMODARAN_ERP,
    "SEC 8-K — 6.250% Senior Notes due 2034 (Rd)": URL_8K_NOTES,
    "Blume (1971) — beta adjustment toward 1.0": URL_BLUME_BETA,
    "MODEL18 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL18 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    f"DSO path unchanged vs Model17 (97→{TARGET_DSO_DAYS:.0f}d). "
    "Op NWC = AR+Inv−AP ONLY. +ΔDeferred is a SEPARATE CFO/FCFF source. "
    f"No AR/unearned double-count. Source: {URL_10Q_DEFERRED_DETAIL}."
)

SEGMENT_METHODOLOGY_NOTE = (
    f"Scores OM ≈ {SCORES_SEGMENT_OM:.0%} (IR {URL_IR_DECK}). GM_B2B {GM_B2B:.0%} "
    f"unchanged vs Model17. Q3 Platform ARR ${Q3_FY26_PLATFORM_ARR_000s/1000:.0f}M "
    f"(+62%) / Software ARR ${Q3_FY26_SOFTWARE_ARR_000s/1000:.0f}M (+10%) "
    f"({URL_Q3_FY26_EX991}; transcript {URL_Q3_FY26_TRANSCRIPT})."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Deferred = Rev × (SaaS%+OnPrem%) × (FY25 Def/FY25 soft Rev); "
    f"+ΔDeferred in CFO row 81. Faster SaaS mix (+{SAAS_MIX_SHIFT_BPS:.0f} bps/yr) "
    f"raises Deferred cash vs Model17 (+250). Detail: {URL_10Q_DEFERRED_DETAIL}."
)

DA_METHODOLOGY_NOTE = (
    "TOTAL D&A ≈ 0.84% of Rev. SBC is a SEPARATE add-back — no double-count."
)

SBC_METHODOLOGY_NOTE = (
    f"Updated SBC from Model17 (Previous: fade "
    f"{' / '.join(f'{p:.2%}' for p in SBC_PCT_PATH_MODEL17)}) to Model18 "
    f"(New: fade {' / '.join(f'{p:.2%}' for p in SBC_PCT_PATH)}). "
    "Treat SBC with suspicion but not so hard that royalty price-ups are taxed "
    f"1:1 with grants. Source: {URL_FACTS}; Scores pricing {URL_Q3_FY26_EX991}."
)

CAPEX_METHODOLOGY_NOTE = (
    f"Updated CapEx from Model17 (Previous: "
    f"{' → '.join(f'{p:.2%}' for p in CAPEX_PCT_PATH_MODEL17)}) to Model18 "
    f"(New: {' → '.join(f'{p:.2%}' for p in CAPEX_PCT_PATH)}) — Scores coupons "
    "need ~0 CapEx; Platform is mostly cloud opex not PPE."
)

MARGIN_METHODOLOGY_NOTE = (
    f"Margins unchanged vs Model17 (COGS −{COGS_IMPROVEMENT_BPS:.0f} / "
    f"SG&A −{SGA_IMPROVEMENT_BPS:.0f} / R&D −{RD_IMPROVEMENT_BPS:.0f}; "
    f"floor {COGS_FLOOR_PCT:.0%}; GM_B2B {GM_B2B:.0%})."
)

CASH_TAX_METHODOLOGY_NOTE = (
    f"Keep cash tax {CASH_TAX_RATE:.2%} (3yr IncomeTaxesPaidNet/EBT). "
    f"Book ETR ~{BOOK_TAX_RATE_FY25:.1%} used in CAPM Rd(1−T)."
)

BS_METHODOLOGY_NOTE = (
    "Equity = Assets − Liab − RE every year (accounting identity, not a soft plug). "
    "Liab = AP + Debt + Deferred. Row-3 OK; cash check ≈ 0."
)

FINANCING_METHODOLOGY_NOTE = (
    f"Buybacks unchanged vs Model17 ({BUYBACK_FCF_MULTIPLE:.2f}×(CFO−CapEx); "
    f"debt funds MULT−1). Q3 FY2026 buybacks ${BUYBACK_Q3_FY26_000s/1000:.0f}M vs "
    f"FCF ${Q3_FY26_FCF_000s/1000:.1f}M ({URL_Q3_FY26_EX991})."
)

DILUTION_METHODOLOGY_NOTE = (
    f"Shares unchanged vs Model17 ({SHARES_OUTSTANDING_000s:,.3f}k @ "
    f"${SHARE_PRICE:,.0f} post–Q3). SBC add-back does NOT dilute; Y1 share "
    f"reduction uses stub-scaled equity CF ({STUB_FRACTION_REMAINING:.2%} of year)."
)

WACC_METHODOLOGY_NOTE = (
    f"WACC unchanged vs Model17 algebra: Yacktman-adj CAPM {MODEL18_WACC:.2%} "
    f"(D6=R15). β_Blume=(2/3)×{BETA_RAW:.2f}+(1/3)×1.0={BETA_BLUME:.3f}; "
    f"β_AAA={BETA_AAA_COUPON:.2f}; β_Yacktman=0.30×β_Blume+0.70×β_AAA="
    f"{BETA_YACKTMAN:.3f}; Ke=Rf({RF_RATE:.2%})+β×ERP({ERP_RATE:.2%})="
    f"{COST_OF_EQUITY:.2%}; Rd_at={PRE_TAX_RD:.3%}×(1−{BOOK_TAX_RATE_FY25:.2%})="
    f"{AFTER_TAX_RD:.2%}; WACC=80%×Ke+20%×Rd_at={MODEL18_WACC:.2%}. "
    f"No naked hardcoded discount rate. Sources: {URL_YAHOO_BETA}; "
    f"{URL_BLUME_BETA}; {URL_FRED_DGS10}; {URL_8K_NOTES}."
)

EXIT_METHODOLOGY_NOTE = (
    f"Exit unchanged vs Model17 (BASE {EXIT_EV_EBITDA:.1f}x; Bull "
    f"{EXIT_EV_EBITDA_BULL:.1f}x / Bear {EXIT_EV_EBITDA_BEAR:.1f}x)."
)

GROWTH_METHODOLOGY_NOTE = (
    f"Revenue growth path unchanged vs Model17 ("
    f"{REVENUE_GROWTH_PATH[0]:.0%}/{REVENUE_GROWTH_PATH[1]:.0%}/"
    f"{REVENUE_GROWTH_PATH[2]:.1%}/{REVENUE_GROWTH_PATH[3]:.0%}/"
    f"{REVENUE_GROWTH_PATH[4]:.1%}). DCF applies stub scaling to Y1 FCFF only — "
    "3S still builds a full fiscal year for BS integrity."
)

STUB_METHODOLOGY_NOTE = (
    f"Updated temporal discounting from Model17 (Previous: D9=hist FYE 2025-09-30; "
    f"full-year Y1 FCFF at mid-year EDATE) to Model18 (New: valuation date "
    f"{VALUATION_DATE.isoformat()}; FY stub start {FY_STUB_START.isoformat()}; "
    f"{STUB_ELAPSED_DAYS}/{STUB_TOTAL_DAYS} elapsed = {STUB_FRACTION_ELAPSED:.2%}; "
    f"Y1 FCFF × {STUB_FRACTION_REMAINING:.2%} remaining; E18 = mid-stub "
    f"{STUB_MID_DATE.isoformat()} = D9+{STUB_MID_OFFSET_DAYS}d; Y2–Y5 midpoints "
    f"via EDATE from stub FY calendar). Only cash yet to occur is valued."
)

YACKTMAN_METHODOLOGY_NOTE = (
    f"Yacktman lens (MODEL17 rated {MODEL17_RATING}): {MODEL17_CRITIQUE} "
    f"MODEL18 keeps AAA-coupon β ({BETA_YACKTMAN:.3f}), grounds SBC, accelerates "
    "Platform/SaaS Deferred cash, and discounts only the remaining stub. "
    "Every WACC term is algebraic — no naked hardcoded discount rate."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: MODEL17={MODEL17_RATING}; "
        f"stub {STUB_FRACTION_REMAINING:.1%} of Y1 FCFF from {VALUATION_DATE.isoformat()}; "
        f"WACC Yacktman-adj {MODEL18_WACC:.2%} (β {BETA_YACKTMAN:.3f}); "
        f"SBC {'/'.join(f'{p:.0%}' for p in SBC_PCT_PATH)}; "
        f"SaaS +{SAAS_MIX_SHIFT_BPS:.0f}bps; exit {EXIT_EV_EBITDA:.1f}x; "
        f"shares {SHARES_OUTSTANDING_000s/1000:.1f}M @ ${SHARE_PRICE:,.0f}."
    )
