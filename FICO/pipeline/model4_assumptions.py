"""vengeanceaiUSCMODEL4 — assumption policy.

Same researched economics as MODEL3, but forecast *drivers in Excel* are
equations linked to the last historical year (col I), not pasted dollars.

Policy (what we keep as explicit inputs vs hist-linked formulas):
---------------------------------------------------------------------------
EXPLICIT INPUTS (yellow — forward-looking judgment)
  • Revenue growth path: 27% / 16% / 13% / 10% / 7%
      Y1 ≈ FY2026 guidance ~$2.53B / FY25; then fade toward g=3%
  • SGA efficiency grind: −50 bps of sales per year (applied in SGA $ formula)
  • CapEx fade weights: blend FY25 CapEx/Sales → 1.0% steady
  • CAPM stack (Rf, ERP, β, Rd, We/Wd) — researched; live in DCF!R
  • Exit EV/EBITDA = 25x; perpetual g = 3%
  • Debt / equity issuance = 0 (hold capital structure)

HIST-LINKED EQUATIONS (green — must update when hist changes)
  • COGS%     = I25/I24          (FY25 COGS / Revenue)
  • SGA $     = Rev_t × ( (I28−10922)/I24 − 0.005×year )
  • R&D $     = Rev_t × (I29/I24)
  • D&A% PPE  = I30/I92          (or I44 if open PPE)
  • Interest% = I101/I98
  • Tax%      = I35/I33
  • AR days   = ROUND(I42/I24×365, 0)   (I42 is net AR in MODEL4)
  • Inv days  = 0
  • AP days   = ROUND(I48/I25×365, 0)
  • CapEx $   = Rev_t × (peak×w + 1%×(1−w)) with w fade 0.85→0

IS/BS/CF forecast rows stay as CFI template formulas
  (Rev=prior×(1+g), COGS=Rev×%, AR=Rev×days/365, etc.)
"""

from __future__ import annotations

from .model3_assumptions import (  # noqa: F401 — shared researched stack
    CAPEX_STEADY_PCT,
    CASH_10Q_000s,
    EXIT_EV_EBITDA,
    MKT_SECS_10Q_000s,
    PERPETUAL_GROWTH,
    RESTRUCTURING_NORMALIZE_000s,
    REVENUE_GROWTH_PATH,
    SGA_IMPROVEMENT_BPS,
    SHARE_PRICE,
    SHARES_OUTSTANDING_000s,
    SOURCE_LINKS,
    TOTAL_DEBT_10Q_000s,
    cover_blurb as _m3_blurb,
)
from .wacc import MODEL3_WACC

MODEL_NAME = "vengeanceaiUSCMODEL4"

# CapEx fade weights on (peak − steady): year1 keeps most of peak
CAPEX_FADE_WEIGHTS = (0.85, 0.65, 0.45, 0.30, 0.0)


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: forecast drivers are Excel equations off FY25 "
        f"(COGS%=I25/I24, SGA/RD/CapEx=Rev×hist%, days=ROUND(I…)). "
        f"Growth path + CAPM WACC={MODEL3_WACC:.2%} + Exit {EXIT_EV_EBITDA:.0f}x "
        f"are the only forward policy inputs. Open in Excel to recalculate."
    )
