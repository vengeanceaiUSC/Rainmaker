"""vengeanceaiUSCMODEL7 — audit fixes for NWC, SGA leverage, CapEx, exit multiple.

Corrections vs MODEL6 (per model audit):
1. Operating NWC as one % of revenue (fade hist → 3% steady) — kills gross-AR/deferred double-count drain
2. SGA floor 15%, grind −75 bps/yr — software operating leverage
3. CapEx fades faster to 1.0% steady so CapEx ≈ D&A by late forecast
4. Primary exit multiple 16x (blended); 25x kept as bull sensitivity
5. IS labels: SG&A / R&D (not Salaries / Rent)
"""

from __future__ import annotations

from typing import Dict, Tuple

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

MODEL_NAME = "vengeanceaiUSCMODEL7"

# --- SGA: software leverage ---
SGA_IMPROVEMENT_BPS = 75.0  # −75 bps of sales per year
SGA_FLOOR_PCT = 0.15  # 15% floor

# --- CapEx: faster fade to maintenance ---
CAPEX_STEADY_PCT = 0.010  # 1.0% of sales
# Weight on FY25 peak CapEx/Sales; rest is steady (hits 1% by Y4)
CAPEX_FADE_WEIGHTS: Tuple[float, ...] = (0.40, 0.20, 0.10, 0.0, 0.0)

# --- NWC: single operating ratio (not gross AR days − deferred) ---
NWC_STEADY_PCT = 0.03  # 3% of sales — asset-light software
# Weight on FY25 operating NWC/Sales; fade to steady
NWC_FADE_WEIGHTS: Tuple[float, ...] = (0.70, 0.50, 0.30, 0.15, 0.0)

# --- Terminal value ---
EXIT_EV_EBITDA = 16.0  # primary blended (was 25x bull)
EXIT_EV_EBITDA_BULL = 25.0  # sensitivity / bull case only

SOURCE_LINKS: Dict[str, str] = dict(_M5_LINKS)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: NWC%=fade→{NWC_STEADY_PCT:.0%} of sales; "
        f"SGA floor {SGA_FLOOR_PCT:.0%}/−{SGA_IMPROVEMENT_BPS:.0f}bps; "
        f"CapEx fade→{CAPEX_STEADY_PCT:.0%}; primary exit {EXIT_EV_EBITDA:.0f}x "
        f"(bull {EXIT_EV_EBITDA_BULL:.0f}x). WACC={MODEL3_WACC:.2%}."
    )
