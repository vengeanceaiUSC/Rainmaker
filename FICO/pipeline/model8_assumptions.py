"""vengeanceaiUSCMODEL8 — comps-sourced exit multiple + audit cash-flow fixes.

Public comps (EV/EBITDA, VCP Scanner peer set, 2026):
  VRSK 17.4x | SPGI 18.1x | MCO 22.2x | MSCI 23.7x | EFX 13.9x
  → peer median ≈ 18.1x
  Source: https://vcpscanner.com/valuation/fico/relative

Base exit = 17.5x (audit target; near peer median, between Gordon ~12.8x and bull/spot ~25x).
Bull = 25x | Bear = 12.8x (Gordon-implied).

Also locks checklist ops:
  • NWC = 2.5% × Revenue (flat)
  • SG&A floor 15% / −75 bps
  • CapEx fade → 1.0% by Y5
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

MODEL_NAME = "vengeanceaiUSCMODEL8"

# --- SGA ---
SGA_IMPROVEMENT_BPS = 75.0
SGA_FLOOR_PCT = 0.15

# --- CapEx ---
CAPEX_STEADY_PCT = 0.010
CAPEX_FADE_WEIGHTS: Tuple[float, ...] = (0.40, 0.20, 0.10, 0.0, 0.0)

# --- NWC: flat 2.5% of sales (checklist) ---
NWC_STEADY_PCT = 0.025
NWC_FADE_WEIGHTS: Tuple[float, ...] = (0.0, 0.0, 0.0, 0.0, 0.0)  # unused; flat

# --- Exit multiples (comps-sourced base) ---
EXIT_EV_EBITDA = 17.5  # base — near peer median ~18.1x
EXIT_EV_EBITDA_BULL = 25.0  # spot / bull
EXIT_EV_EBITDA_BEAR = 12.8  # Gordon-implied

# Peer comps table for on-sheet documentation
PEER_EV_EBITDA: List[Tuple[str, str, float]] = [
    ("EFX", "Equifax", 13.88),
    ("VRSK", "Verisk Analytics", 17.42),
    ("SPGI", "S&P Global", 18.11),
    ("MCO", "Moody's", 22.20),
    ("MSCI", "MSCI Inc.", 23.74),
]
PEER_MEDIAN_EV_EBITDA = 18.11  # SPGI = median of sorted peers
URL_PEER_COMPS = "https://vcpscanner.com/valuation/fico/relative"
URL_FICO_EV_EBITDA = "https://www.alphaspread.com/security/nyse/fico/relative-valuation/ratio/enterprise-value-to-ebitda"

# Downloadable assumptions list (PDF, no charts) — updated on each MODEL8 publish
ASSUMPTIONS_PDF_FILENAME = "MODEL8_FICO_ASSUMPTIONS_LIST.pdf"
ASSUMPTIONS_PDF_URL = (
    "https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
    f"cursor/vengeanceaiusmodel8-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)
# Browser-viewable GitHub page (same file)
ASSUMPTIONS_PDF_VIEW_URL = (
    "https://github.com/vengeanceaiUSC/Rainmaker/blob/"
    f"cursor/vengeanceaiusmodel8-d3ac/FICO/output/{ASSUMPTIONS_PDF_FILENAME}"
)

SOURCE_LINKS: Dict[str, str] = {
    **_M5_LINKS,
    "FICO peer EV/EBITDA comps (VCP Scanner)": URL_PEER_COMPS,
    "FICO EV/EBITDA (Alpha Spread)": URL_FICO_EV_EBITDA,
    "Damodaran EV/EBITDA by sector": (
        "https://pages.stern.nyu.edu/adamodar/New_Home_Page/datafile/vebitda.htm"
    ),
    "MODEL8 Assumptions List (PDF download)": ASSUMPTIONS_PDF_URL,
    "MODEL8 Assumptions List (PDF view)": ASSUMPTIONS_PDF_VIEW_URL,
}


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: NWC={NWC_STEADY_PCT:.1%} of sales; "
        f"SG&A floor {SGA_FLOOR_PCT:.0%}/−{SGA_IMPROVEMENT_BPS:.0f}bps; "
        f"CapEx→{CAPEX_STEADY_PCT:.0%}; exit base {EXIT_EV_EBITDA:.1f}x "
        f"(peer median ~{PEER_MEDIAN_EV_EBITDA:.1f}x; bull {EXIT_EV_EBITDA_BULL:.0f}x; "
        f"bear {EXIT_EV_EBITDA_BEAR:.1f}x). WACC={MODEL3_WACC:.2%}."
    )
