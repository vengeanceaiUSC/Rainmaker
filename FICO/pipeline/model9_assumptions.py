"""vengeanceaiUSCMODEL9 — bottom-up WC (DSO/DPO) + corrected D&A.

Working Capital:
  Transitioned from top-down NWC % of sales (AR plug) to bottom-up
  DSO / DPO historical drivers. NWC = AR + Inv − AP − Deferred (output).

Depreciation:
  DA$ = (Opening PPE + CapEx/2) × DA%   [non-circular avg-PPE convention]

Keeps MODEL8 ops elsewhere:
  • SG&A floor 15% / −75 bps
  • CapEx fade → 1.0% by Y5
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

MODEL_NAME = "vengeanceaiUSCMODEL9"

# --- SGA ---
SGA_IMPROVEMENT_BPS = 75.0
SGA_FLOOR_PCT = 0.15

# --- CapEx ---
CAPEX_STEADY_PCT = 0.010
CAPEX_FADE_WEIGHTS: Tuple[float, ...] = (0.40, 0.20, 0.10, 0.0, 0.0)

# --- Exit multiples (comps-sourced base; unchanged from MODEL8) ---
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

ASSUMPTIONS_PDF_FILENAME = "MODEL9_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel9-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel9-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)

SOURCE_LINKS: Dict[str, str] = {
    **_M5_LINKS,
    "FICO peer EV/EBITDA comps (VCP Scanner)": URL_PEER_COMPS,
    "FICO EV/EBITDA (Alpha Spread)": URL_FICO_EV_EBITDA,
    "Damodaran EV/EBITDA by sector": (
        "https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/vebitda.htm"
    ),
    "MODEL9 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL9 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}

WC_METHODOLOGY_NOTE = (
    "Working Capital forecasting has been transitioned from a top-down NWC % of "
    "sales methodology to a bottom-up approach using DSO and DPO historical "
    "drivers to eliminate artificial cash flow distortions."
)

DA_METHODOLOGY_NOTE = (
    "Depreciation uses (Opening PPE + CapEx/2) × DA%, the non-circular form of "
    "average net PP&E depreciation (equivalent to mid-year CapEx convention)."
)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: bottom-up WC (DSO/DPO); "
        f"DA=(Open+CapEx/2)×DA%; "
        f"SG&A floor {SGA_FLOOR_PCT:.0%}/−{SGA_IMPROVEMENT_BPS:.0f}bps; "
        f"CapEx→{CAPEX_STEADY_PCT:.0%}; exit base {EXIT_EV_EBITDA:.1f}x "
        f"(peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x). WACC={MODEL3_WACC:.2%}."
    )
