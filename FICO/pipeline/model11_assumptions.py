"""vengeanceaiUSCMODEL11 — MODEL10 cash-flow realism + BS balance (SBC→APIC) fixes on MODEL9 scaffolding.

Grounded in SEC XBRL companyfacts (CIK 0000814547), FY2023–FY2025 10-K:

Working Capital (kept from MODEL9):
  Bottom-up DSO / DPO; NWC = AR + Inv − AP − Deferred (output).

MODEL10 cash-flow fixes:
  1. SBC add-back — 3yr avg SBC/Revenue from ShareBasedCompensation
  2. D&A % of Revenue — total DepreciationDepletionAndAmortization / Rev
     (software/intangibles, not PP&E-only rate)
  3. CapEx — flat 3yr avg (PPE + PaymentsToDevelopSoftware) / Rev
     (no hardcoded fade to 1%)
  4. Cash tax rate — IncomeTaxesPaidNet / EBT (DCF unlevered taxes)
  5. Financing — normalized 3yr avg buybacks + net debt CF (3S realism)

Keeps:
  • SG&A floor 15% / −75 bps
  • Exit base 17.5x (peer median); bull 25x; bear 12.8x
  • IS labels SG&A / R&D
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
# FY23 0.97% | FY24 0.81% | FY25 0.75% → 0.841%
DA_PCT_REVENUE = 0.008411

# --- CapEx: 3yr avg (PPE purchases + software capitalization) / Revenue ---
# FY23 0.28% | FY24 1.49% | FY25 1.98% → 1.249%
# Removes MODEL8/9 "fast fade → 1%" policy.
CAPEX_PCT_REVENUE = 0.01249

# --- Cash taxes: 3yr avg IncomeTaxesPaidNet / EBT ---
# FY23 27.60% | FY24 20.83% | FY25 20.20% → 22.87%
CASH_TAX_RATE = 0.22873

# --- Financing run-rates ($000s, 3yr absolute averages) ---
# Buybacks: PaymentsForRepurchaseOfCommonStock avg FY23–25
BUYBACK_RUNRATE_000s = 880_577.0  # equity issuance input = −this
# Net debt CF ≈ ProceedsFromIssuanceOfSeniorLongTermDebt − RepaymentsOfLinesOfCredit
DEBT_NET_RUNRATE_000s = -290_917.0

# --- Exit multiples (comps-sourced base; unchanged) ---
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

URL_10K = "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm"
URL_FACTS = "https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json"

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
    "SEC 10-K FY2025": URL_10K,
    "MODEL11 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL11 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "Working Capital uses bottom-up DSO and DPO historical drivers. "
    "NWC = AR + Inventory − AP − Deferred Revenue (output); "
    "ΔNWC = NWCt − NWCt−1."
)

DA_METHODOLOGY_NOTE = (
    "D&A is forecast as a % of Revenue (3yr FY23–25 average of total "
    "DepreciationDepletionAndAmortization / Revenue ≈ 0.84%), not a PP&E-only "
    "rate — capturing software amortization and intangibles."
)

SBC_METHODOLOGY_NOTE = (
    "Stock-Based Compensation is added back in CFO and FCFF at the 3yr average "
    "SBC/Revenue (≈ 8.25%) from ShareBasedCompensation in the 10-K cash flow."
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
    "Balance sheet identity is enforced: Equity Capital rolls Prior + Equity Issuance "
    "+ SBC (APIC credit). Without the SBC equity credit, Assets − (L+E) equals "
    "cumulative SBC because SBC is added back in CFO/FCFF but is non-cash."
)

FINANCING_METHODOLOGY_NOTE = (
    "Financing CF restores realism outside FCFF: debt uses the 3yr average net debt "
    "cash flow (senior-note proceeds − line-of-credit repayments ≈ −$291M/yr). "
    "Equity buybacks distribute residual levered FCF after that debt CF so cash does "
    "not artificially stockpile (historical buybacks averaged ≈ $881M/yr in FY23–25)."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: SBC add-back {SBC_PCT_REVENUE:.1%} of sales; "
        f"DA={DA_PCT_REVENUE:.2%} of sales; CapEx={CAPEX_PCT_REVENUE:.2%} flat; "
        f"cash tax={CASH_TAX_RATE:.1%}; buybacks ${BUYBACK_RUNRATE_000s/1000:.0f}M/yr avg; "
        f"SG&A floor {SGA_FLOOR_PCT:.0%}/−{SGA_IMPROVEMENT_BPS:.0f}bps; "
        f"exit base {EXIT_EV_EBITDA:.1f}x. WACC={MODEL3_WACC:.2%}."
    )
