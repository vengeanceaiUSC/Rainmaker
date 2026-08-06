"""vengeanceaiUSCMODEL17.0 — Yacktman credit-adjusted CAPM on MODEL16 scaffolding.

MODEL16 rated 6.5/10: CAPM primary + Deferred/SBC/DSO plumbing work, but
Yahoo β=1.32 prices FICO like cyclical software, growth stays 27/16/13/10/7,
and blended CapEx/opex/SBC still tax Scores royalty cash.

Yacktman: "What credit rating would you give it?" Scores royalties ≈ AAA
GSE-mandated coupon. Adjust CAPM β (math on-sheet) — never hardcode a naked WACC%.

MODEL17.0 package:
  1. D6 = Yacktman-adjusted CAPM (Blume + AAA-coupon β blend); raw CAPM at R15 ref
  2. Growth 30/18/14.5/11/8.5 (mortgage royalty / Q3 Scores +41%)
  3. Scores-led margins (GM_B2B 97%, COGS floor 8%, stronger opex/R&D fade)
  4. CapEx Scores-light; SBC% fade; levered buybacks 1.4× FCF
  5. Post–Q3 shares @ $1,149; exit 23.0x; g 3.5%
  6. Op NWC ≠ Deferred; hard BS plug
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

MODEL_NAME = "vengeanceaiUSCMODEL17.0"

MODEL16_RATING = "6.5/10"
MODEL16_WACC = MODEL3_WACC  # raw CAPM ≈ 9.27% (predecessor D6)

# --- Market / share base (post Q3 FY2026) ---
SHARE_PRICE = 1_149.0  # MODEL16: 1046.23
SHARES_OUTSTANDING_000s = 21_597.635  # MODEL16: 23,764 FYE25

BUYBACK_Q3_FY26_000s = 1_960_000.0
BUYBACK_Q3_FY26_SHARES_000s = 1_705.0
BUYBACK_Q3_FY26_AVG_PRICE = 1_149.0
Q3_FY26_FCF_000s = 370_343.0

# --- Revenue growth (MODEL16 was 27/16/13/10/7) ---
REVENUE_GROWTH_PATH: Tuple[float, ...] = (0.30, 0.18, 0.145, 0.11, 0.085)
REVENUE_GROWTH_PATH_MODEL16: Tuple[float, ...] = (0.27, 0.16, 0.13, 0.10, 0.07)

# --- Operating leverage / Scores coupon ---
SGA_IMPROVEMENT_BPS = 175.0  # MODEL16: 125
SGA_FLOOR_PCT = 0.14  # MODEL16: 0.15
RD_IMPROVEMENT_BPS = 50.0  # MODEL16: flat
RD_FLOOR_PCT = 0.07
COGS_IMPROVEMENT_BPS = 120.0  # MODEL16: 75
COGS_FLOOR_PCT = 0.08  # MODEL16 bake used 10%
SCORES_SEGMENT_OM = 0.91

# --- SBC fade (MODEL16 flat 8.251%) ---
SBC_PCT_REVENUE = 0.08251
SBC_PCT_PATH: Tuple[float, ...] = (0.08251, 0.0780, 0.0740, 0.0700, 0.0660)

# --- D&A / CapEx / cash tax ---
DA_PCT_REVENUE = 0.008411
CAPEX_PCT_REVENUE = 0.0100  # MODEL16 FY1 was 1.25%
CAPEX_PCT_PATH: Tuple[float, ...] = (0.0100, 0.0085, 0.0070, 0.0055, 0.0040)
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
BUYBACK_FCF_MULTIPLE = 1.40  # MODEL16 was residual ≈1.0×

SAAS_MIX_SHIFT_BPS = 250.0

# --- Yacktman-adjusted CAPM (primary D6) ---
MODEL17_WACC = MODEL17_CAPM_WACC
WACC_BAND_LOW = 0.0650
WACC_BAND_HIGH = 0.0780
RF_RATE = YACKTMAN_WACC_INPUTS.risk_free_rate
ERP_RATE = YACKTMAN_WACC_INPUTS.equity_risk_premium
BETA = BETA_YACKTMAN  # primary β used in D6
PRE_TAX_RD = YACKTMAN_WACC_INPUTS.pre_tax_cost_of_debt
COST_OF_EQUITY = YACKTMAN_WACC_INPUTS.cost_of_equity
AFTER_TAX_RD = YACKTMAN_WACC_INPUTS.after_tax_cost_of_debt

PERPETUAL_GROWTH = 0.035  # MODEL16: 3.0%
PERPETUAL_GROWTH_MODEL16 = 0.030

EXIT_EV_EBITDA = 23.0  # MODEL16: 21.0
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
URL_BLUME_BETA = "https://www.jstor.org/stable/2329867"  # Blume 1971 beta adjustment

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
TARGET_DSO_DAYS = 28.0  # unchanged vs MODEL16

GM_SAAS = 0.84
GM_ONPREM = 0.80
GM_B2C = 0.82
GM_B2B = 0.97  # MODEL16: 0.90 — Scores OM ~91%
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


ASSUMPTIONS_PDF_FILENAME = "MODEL17_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel17-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel17-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
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
    "SEC EX-99.1 Q3 FY2026 earnings (Scores +41%, FCF $370M)": URL_Q3_FY26_EX991,
    "FICO IR — Q3 FY2026 earnings release": URL_Q3_FY26_IR,
    "FICO IR investor presentation — Scores OM ~91%": URL_IR_DECK,
    "Equifax statement — FICO 2026 2x mortgage royalty": URL_EQUIFAX_FICO_PRICING,
    "Sen. Hawley oversight letter — FICO mortgage pricing": URL_HAWLEY_LETTER,
    "FRED DGS10 — US 10Y risk-free rate (Rf)": URL_FRED_DGS10,
    "Yahoo ^TNX — 10Y yield cross-check": URL_YAHOO_TNX,
    "Yahoo FICO key statistics — 5Y monthly Beta (raw)": URL_YAHOO_BETA,
    "Damodaran Implied ERP": URL_DAMODARAN_ERP,
    "SEC 8-K — 6.250% Senior Notes due 2034 (Rd)": URL_8K_NOTES,
    "Blume (1971) — beta adjustment toward 1.0": URL_BLUME_BETA,
    "MODEL17 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL17 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    f"DSO path unchanged vs Model16 (97→{TARGET_DSO_DAYS:.0f}d). "
    "Op NWC = AR+Inv−AP ONLY. +ΔDeferred is a SEPARATE CFO/FCFF source. "
    f"No AR/unearned double-count. Source: {URL_10Q_DEFERRED_DETAIL}."
)

SEGMENT_METHODOLOGY_NOTE = (
    f"Scores OM ≈ {SCORES_SEGMENT_OM:.0%} (IR {URL_IR_DECK}). "
    f"Updated GM_B2B from Model16 (Previous: 90%) to Model17 (New: {GM_B2B:.0%}). "
    f"Q3 Scores +41% / B2B +49% ({URL_Q3_FY26_EX991})."
)

DEFERRED_METHODOLOGY_NOTE = (
    "Deferred = Rev × (SaaS%+OnPrem%) × (FY25 Def/FY25 soft Rev); "
    f"+ΔDeferred in CFO row 81. FY25 deferred ≈ ${FY25_DEFERRED_REV_000s/1000:.1f}M. "
    f"Detail: {URL_10Q_DEFERRED_DETAIL}."
)

DA_METHODOLOGY_NOTE = (
    "TOTAL D&A ≈ 0.84% of Rev. SBC is a SEPARATE add-back — no double-count."
)

SBC_METHODOLOGY_NOTE = (
    f"Updated SBC from Model16 (Previous: flat {SBC_PCT_REVENUE:.2%}) to Model17 "
    f"(New: fade {' / '.join(f'{p:.2%}' for p in SBC_PCT_PATH)}) because royalty "
    f"price dollars do not require 1:1 SBC. Source: {URL_FACTS}."
)

CAPEX_METHODOLOGY_NOTE = (
    f"Updated CapEx from Model16 (Previous: 1.25%→0.60%) to Model17 "
    f"(New: {' → '.join(f'{p:.2%}' for p in CAPEX_PCT_PATH)}) — Scores royalties "
    "need ~0 CapEx; do not tax the coupon with Software capitalization."
)

MARGIN_METHODOLOGY_NOTE = (
    f"Updated margins from Model16 (Previous: COGS −75 / SG&A −125 / R&D flat, "
    f"COGS floor 10%) to Model17 (New: COGS −{COGS_IMPROVEMENT_BPS:.0f} / "
    f"SG&A −{SGA_IMPROVEMENT_BPS:.0f} / R&D −{RD_IMPROVEMENT_BPS:.0f}; "
    f"floor {COGS_FLOOR_PCT:.0%}; GM_B2B {GM_B2B:.0%}). "
    f"Mortgage royalty ${MORTGAGE_ROYALTY_2025:.2f}→${MORTGAGE_ROYALTY_2026:.2f}."
)

CASH_TAX_METHODOLOGY_NOTE = (
    f"Keep cash tax {CASH_TAX_RATE:.2%} (3yr IncomeTaxesPaidNet/EBT). "
    f"Book ETR ~{BOOK_TAX_RATE_FY25:.1%} used in CAPM Rd(1−T)."
)

BS_METHODOLOGY_NOTE = (
    "Equity = Assets − Liab − RE every year. Liab = AP + Debt + Deferred. "
    "Row-3 OK; cash check ≈ 0."
)

FINANCING_METHODOLOGY_NOTE = (
    f"Updated buybacks from Model16 (Previous: residual ≈1.0× FCF) to Model17 "
    f"(New: {BUYBACK_FCF_MULTIPLE:.2f}×(CFO−CapEx); debt funds MULT−1). "
    f"Q3 FY2026: buybacks ${BUYBACK_Q3_FY26_000s/1000:.0f}M vs FCF "
    f"${Q3_FY26_FCF_000s/1000:.1f}M ({URL_Q3_FY26_EX991})."
)

DILUTION_METHODOLOGY_NOTE = (
    f"Updated shares from Model16 (Previous: 23,764k @ $1,046) to Model17 "
    f"(New: {SHARES_OUTSTANDING_000s:,.3f}k @ ${SHARE_PRICE:,.0f} post–Q3). "
    "SBC add-back does NOT dilute; shares fall via buybacks only."
)

WACC_METHODOLOGY_NOTE = (
    f"Updated WACC from Model16 (Previous: raw CAPM {MODEL16_WACC:.2%} with "
    f"Yahoo β={BETA_RAW:.2f}) to Model17 (New: Yacktman-adjusted CAPM "
    f"{MODEL17_WACC:.2%}). Algebra: β_Blume=(2/3)×{BETA_RAW:.2f}+(1/3)×1.0="
    f"{BETA_BLUME:.3f}; β_AAA={BETA_AAA_COUPON:.2f} (GSE Scores coupon credit); "
    f"β_Yacktman=0.30×β_Blume+0.70×β_AAA={BETA_YACKTMAN:.3f}; "
    f"Ke=Rf({RF_RATE:.2%})+β×ERP({ERP_RATE:.2%})={COST_OF_EQUITY:.2%}; "
    f"Rd_at={PRE_TAX_RD:.3%}×(1−{BOOK_TAX_RATE_FY25:.2%})={AFTER_TAX_RD:.2%}; "
    f"WACC=80%×Ke+20%×Rd_at={MODEL17_WACC:.2%}. "
    f"Raw CAPM kept at R15 as reference. Sources: {URL_YAHOO_BETA}; "
    f"{URL_BLUME_BETA}; {URL_IR_DECK}; {URL_FRED_DGS10}; {URL_8K_NOTES}."
)

EXIT_METHODOLOGY_NOTE = (
    f"Updated exit from Model16 (Previous: 21.0x) to Model17 "
    f"(New: BASE {EXIT_EV_EBITDA:.1f}x; Bull {EXIT_EV_EBITDA_BULL:.1f}x / "
    f"Bear {EXIT_EV_EBITDA_BEAR:.1f}x). Still below spot."
)

GROWTH_METHODOLOGY_NOTE = (
    "Updated revenue growth from Model16 (Previous: 27%/16%/13%/10%/7%) "
    f"to Model17 (New: {REVENUE_GROWTH_PATH[0]:.0%}/{REVENUE_GROWTH_PATH[1]:.0%}/"
    f"{REVENUE_GROWTH_PATH[2]:.1%}/{REVENUE_GROWTH_PATH[3]:.0%}/"
    f"{REVENUE_GROWTH_PATH[4]:.1%}) — Q3 Scores +41% / B2B +49% mortgage unit price "
    f"({URL_Q3_FY26_EX991}); royalty ${MORTGAGE_ROYALTY_2025:.2f}→"
    f"${MORTGAGE_ROYALTY_2026:.2f}. Terminal g {PERPETUAL_GROWTH_MODEL16:.1%}→"
    f"{PERPETUAL_GROWTH:.1%}."
)

YACKTMAN_METHODOLOGY_NOTE = (
    f"Yacktman lens (MODEL16 rated {MODEL16_RATING}): assign AAA credit to Scores "
    f"coupons → lower β ({BETA_RAW:.2f}→{BETA_YACKTMAN:.3f}), raise FCFF via "
    "Scores-aware margins/CapEx/SBC, and shrink shares with levered buybacks. "
    "Every WACC term is algebraic — no naked hardcoded discount rate."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME} (Yacktman CAPM): MODEL16={MODEL16_RATING}; "
        f"WACC raw {MODEL16_WACC:.2%}→adj {MODEL17_WACC:.2%} "
        f"(β {BETA_RAW:.2f}→{BETA_YACKTMAN:.3f}); g {PERPETUAL_GROWTH:.1%}; "
        f"exit {EXIT_EV_EBITDA:.1f}x; growth 30/18/14.5/11/8.5; "
        f"COGS floor {COGS_FLOOR_PCT:.0%}; buybacks {BUYBACK_FCF_MULTIPLE:.1f}×FCF; "
        f"shares {SHARES_OUTSTANDING_000s/1000:.1f}M post–Q3."
    )
