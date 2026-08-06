"""vengeanceaiUSCMODEL6 — fixes to D&A, AR/Deferred, and SGA floor.

Inherits MODEL5 (clickable sources + hist-linked Excel drivers) with three
policy corrections called out in review:

1. D&A applies to *average* PP&E (Open + Open+CapEx)/2 so new CapEx is depreciated.
2. AR days use *gross* AR; Deferred Revenue is a separate liability / NWC credit.
3. SGA floor raised to 20% (enterprise software) with a milder −25 bps/yr grind.
"""

from __future__ import annotations

from typing import Dict

from .model3_assumptions import (  # noqa: F401
    CAPEX_STEADY_PCT,
    CASH_10Q_000s,
    EXIT_EV_EBITDA,
    MKT_SECS_10Q_000s,
    PERPETUAL_GROWTH,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
    SHARE_PRICE,
    SHARES_OUTSTANDING_000s,
    TOTAL_DEBT_10Q_000s,
)
from .model4_assumptions import CAPEX_FADE_WEIGHTS  # noqa: F401
from .model5_assumptions import SOURCE_LINKS as _M5_LINKS
from .wacc import MODEL3_WACC

MODEL_NAME = "vengeanceaiUSCMODEL6"

# MODEL6 SGA policy — replaces model3’s 50 bps / 5% floor
SGA_IMPROVEMENT_BPS = 25.0  # −25 bps of sales per forecast year (milder leverage)
SGA_FLOOR_PCT = 0.20  # 20% floor — realistic for enterprise software / scoring

SOURCE_LINKS: Dict[str, str] = dict(_M5_LINKS)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: D&A on avg PP&E (new CapEx depreciated); "
        f"gross AR days + separate Deferred Revenue in NWC; "
        f"SGA floor {SGA_FLOOR_PCT:.0%} / grind −{SGA_IMPROVEMENT_BPS:.0f}bps. "
        f"WACC={MODEL3_WACC:.2%}, Exit {EXIT_EV_EBITDA:.0f}x. "
        f"Every Source is a clickable URL."
    )
