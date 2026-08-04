"""
DCF valuation engine — pure Python math on 3-statement forecast outputs.

DCF = Σ CF_t / (1+WACC)^t  +  TV / (1+WACC)^n

FCFF = EBIT*(1-t) + D&A - CapEx - ΔNWC
"""

from __future__ import annotations

from typing import List, Optional

import pandas as pd

from .models import DCFResult, ForecastAssumptions


def _fcff_from_forecast(
    income: pd.DataFrame,
    cashflow: pd.DataFrame,
    proj_years: List[int],
    tax_rate: float,
) -> pd.Series:
    """Build FCFF series for projection years. Frames are year-indexed."""
    ebit = income.loc[proj_years, "operating_income"]
    da = cashflow.loc[proj_years, "da"]
    # CapEx stored as positive outflow in 3-statement forecast
    capex = cashflow.loc[proj_years, "capex"].abs()
    dnwc = cashflow.loc[proj_years, "change_in_nwc"]
    nopat = ebit * (1.0 - tax_rate)
    return nopat + da - capex - dnwc


def run_dcf(
    income: pd.DataFrame,
    cashflow: pd.DataFrame,
    assumptions: ForecastAssumptions,
    *,
    share_price: float,
    diluted_shares_000s: float,
    cash: float,
    marketable_securities: float = 0.0,
    total_debt: Optional[float] = None,
    exit_ev_ebitda: float = 22.0,
    tv_method: str = "exit",
    proj_years: Optional[List[int]] = None,
) -> DCFResult:
    """Discount forecast FCFF; bridge EV -> equity with net debt.

    tv_method:
      - "gordon": TV = FCFF_n*(1+g)/(WACC-g)
      - "exit":   TV = EBITDA_n * exit_ev_ebitda  (default for software comps)
    Both TVs are always computed; primary method drives EV / per-share.
    """
    if proj_years is None:
        avail = [int(y) for y in income.index]
        proj_years = avail[-assumptions.forecast_years :]

    wacc = assumptions.wacc
    g = assumptions.perpetual_growth
    if wacc <= g:
        raise ValueError(f"WACC ({wacc}) must exceed terminal growth ({g})")

    tax = assumptions.tax_rate
    fcff = _fcff_from_forecast(income, cashflow, proj_years, tax)

    discount_factors: List[float] = []
    pv_explicit: List[float] = []
    explicit: List[float] = []
    for i, y in enumerate(proj_years):
        cf = float(fcff.loc[y])
        df = 1.0 / ((1.0 + wacc) ** (i + 1))
        explicit.append(cf)
        discount_factors.append(df)
        pv_explicit.append(cf * df)

    final_fcff = explicit[-1]
    tv_gordon = final_fcff * (1.0 + g) / (wacc - g)
    ebitda_final = float(
        income.loc[proj_years[-1], "operating_income"] + cashflow.loc[proj_years[-1], "da"]
    )
    tv_exit = ebitda_final * exit_ev_ebitda

    method = tv_method.lower().strip()
    if method not in ("gordon", "exit"):
        raise ValueError("tv_method must be 'gordon' or 'exit'")
    terminal_value = tv_exit if method == "exit" else tv_gordon
    pv_terminal = terminal_value * discount_factors[-1]
    enterprise_value = sum(pv_explicit) + pv_terminal

    if total_debt is None:
        net_debt = assumptions.net_debt_thousands
    else:
        net_debt = float(total_debt) - float(cash) - float(marketable_securities)

    equity_value = enterprise_value - net_debt
    shares = diluted_shares_000s or assumptions.shares_thousands
    if shares <= 0:
        raise ValueError("diluted_shares_000s / shares_thousands must be positive")
    per_share = equity_value / shares

    # Alternate method equity/share for transparency
    alt_tv = tv_gordon if method == "exit" else tv_exit
    alt_ev = sum(pv_explicit) + alt_tv * discount_factors[-1]
    alt_eq = alt_ev - net_debt
    alt_ps = alt_eq / shares

    notes = [
        f"Primary TV method: {method}",
        f"Gordon TV = FCFF_n*(1+g)/(WACC-g) = {tv_gordon:,.1f}",
        f"Exit TV = EBITDA_n × {exit_ev_ebitda:.1f}x = {tv_exit:,.1f}",
        f"Alternate ({'gordon' if method == 'exit' else 'exit'}) value/share: ${alt_ps:,.2f}",
        f"Share price (market): ${share_price:,.2f}",
        f"Upside vs price: {(per_share / share_price - 1.0) if share_price else 0.0:.2%}",
        "DCF = Σ CF_t/(1+WACC)^t + TV/(1+WACC)^n",
        "FCFF = EBIT(1-t) + D&A - CapEx - ΔNWC from 3-statement forecast",
    ]

    return DCFResult(
        ticker="",  # filled by runner
        explicit_fcff=explicit,
        discount_factors=discount_factors,
        pv_explicit_fcff=pv_explicit,
        terminal_value=terminal_value,
        pv_terminal_value=pv_terminal,
        enterprise_value=enterprise_value,
        equity_value=equity_value,
        equity_value_per_share=per_share,
        wacc=wacc,
        perpetual_growth=g,
        notes=notes,
    )


def dcf_annual_frame(
    result: DCFResult,
    income: pd.DataFrame,
    cashflow: pd.DataFrame,
    proj_years: List[int],
    tax_rate: float,
) -> pd.DataFrame:
    """Wide annual table for Excel export."""
    rows = {
        "EBIT": [float(income.loc[y, "operating_income"]) for y in proj_years],
        "D&A": [float(cashflow.loc[y, "da"]) for y in proj_years],
        "EBITDA": [
            float(income.loc[y, "operating_income"] + cashflow.loc[y, "da"]) for y in proj_years
        ],
        "NOPAT": [float(income.loc[y, "operating_income"]) * (1.0 - tax_rate) for y in proj_years],
        "CapEx": [float(abs(cashflow.loc[y, "capex"])) for y in proj_years],
        "ΔNWC": [float(cashflow.loc[y, "change_in_nwc"]) for y in proj_years],
        "FCFF": result.explicit_fcff,
        "Discount factor": result.discount_factors,
        "PV of FCFF": result.pv_explicit_fcff,
    }
    df = pd.DataFrame(rows, index=[f"FY{y}" for y in proj_years]).T
    # Terminal column note on last year
    df.loc["Terminal value (primary)", f"FY{proj_years[-1]}"] = result.terminal_value
    df.loc["PV of terminal value", f"FY{proj_years[-1]}"] = result.pv_terminal_value
    return df
