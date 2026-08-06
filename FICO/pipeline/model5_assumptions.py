"""vengeanceaiUSCMODEL5 — assumption policy + mandatory clickable source URLs.

Inherits MODEL4 economics (hist-linked Excel drivers + CAPM DCF). MODEL5’s
product rule: every Source field carries a real hyperlink (3-statement col U/V,
DCF col U/V, Cover Page source index, CSVs with source_url).
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
    SGA_IMPROVEMENT_BPS,
    SHARE_PRICE,
    SHARES_OUTSTANDING_000s,
    SOURCE_LINKS as _M3_LINKS,
    TOTAL_DEBT_10Q_000s,
)
from .model4_assumptions import CAPEX_FADE_WEIGHTS  # noqa: F401
from .wacc import MODEL3_WACC

MODEL_NAME = "vengeanceaiUSCMODEL5"

# Extended registry — every researched input has a clickable URL
SOURCE_LINKS: Dict[str, str] = {
    **_M3_LINKS,
    "FY2026 guidance EX-99.1": (
        "https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/"
        "exhibit991erq32026.htm"
    ),
    "Damodaran FCFF": (
        "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histfcff.html"
    ),
    "Damodaran implied ERP table": (
        "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histimpl.html"
    ),
    "US Treasury yield curve": (
        "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
        "TextView?type=daily_treasury_yield_curve"
    ),
    "FICO IR": "https://www.fico.com/en/investors",
    "Yahoo FICO quote": "https://finance.yahoo.com/quote/FICO/",
}


def cover_blurb() -> str:
    return (
        f"{MODEL_NAME}: same hist-linked Excel drivers as MODEL4, plus every "
        f"Source is a clickable URL (3-statement U/V, DCF U/V, Cover source index). "
        f"WACC={MODEL3_WACC:.2%}, Exit {EXIT_EV_EBITDA:.0f}x. Open in Excel to recalculate."
    )
