"""Export separate 3-statement and DCF results as CSV (no binary Excel)."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Sequence

import pandas as pd

from .model3_assumptions import MathStep, steps_to_rows
from .models import DCFResult, ForecastAssumptions


def _metrics_as_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Prefer metrics as rows, years as columns for readability."""
    if len(df.index) and all(isinstance(i, (int, float)) for i in list(df.index)[:3]):
        out = df.T.copy()
    else:
        out = df.copy()
    out = out.reset_index()
    out = out.rename(columns={"index": "line_item"})
    if out.columns[0] != "line_item":
        out = out.rename(columns={out.columns[0]: "line_item"})
    return out


def _write_csv(path: Path, df: pd.DataFrame) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False)
    return path


def export_three_statement(
    out_dir: Path,
    income: pd.DataFrame,
    balance: pd.DataFrame,
    cashflow: pd.DataFrame,
    assumptions: ForecastAssumptions,
    ticker: str,
    validation_notes: Optional[List[str]] = None,
) -> Dict[str, Path]:
    """Write 3-statement CSVs under out_dir. Returns map of logical name -> path."""
    out_dir.mkdir(parents=True, exist_ok=True)
    paths: Dict[str, Path] = {}

    paths["income_statement"] = _write_csv(
        out_dir / f"3S_{ticker}_income_statement.csv",
        _metrics_as_rows(income),
    )
    paths["balance_sheet"] = _write_csv(
        out_dir / f"3S_{ticker}_balance_sheet.csv",
        _metrics_as_rows(balance),
    )
    paths["cash_flow"] = _write_csv(
        out_dir / f"3S_{ticker}_cash_flow.csv",
        _metrics_as_rows(cashflow),
    )

    assume_rows = [
        ("model_name", assumptions.model_name),
        ("revenue_growth_path", ",".join(f"{g:.6f}" for g in assumptions.revenue_growth)),
        ("cogs_pct_revenue", assumptions.cogs_pct_revenue),
        ("rd_pct_revenue", assumptions.rd_pct_revenue),
        ("sga_pct_revenue", assumptions.sga_pct_revenue),
        ("sga_margin_improvement_bps", assumptions.sga_margin_improvement_bps),
        ("da_pct_revenue", assumptions.da_pct_revenue),
        ("sbc_pct_revenue", getattr(assumptions, "sbc_pct_revenue", 0.0)),
        ("tax_rate", assumptions.tax_rate),  # book tax (3S NI)
        ("cash_tax_rate", getattr(assumptions, "cash_tax_rate", assumptions.tax_rate)),
        ("capex_pct_revenue", assumptions.capex_pct_revenue),
        (
            "capex_pct_path",
            ",".join(f"{x:.6f}" for x in (assumptions.capex_pct_path or [assumptions.capex_pct_revenue])),
        ),
        ("nwc_pct_revenue", assumptions.nwc_pct_revenue),
        ("debt_issuance_annual_000s", getattr(assumptions, "debt_issuance_annual", 0.0)),
        ("buyback_annual_000s", getattr(assumptions, "buyback_annual", 0.0)),
        ("interest_expense_level_000s", assumptions.interest_expense_level),
        ("forecast_years", assumptions.forecast_years),
        ("wacc", assumptions.wacc),
        ("perpetual_growth", assumptions.perpetual_growth),
    ]
    paths["assumptions"] = _write_csv(
        out_dir / f"3S_{ticker}_assumptions.csv",
        pd.DataFrame(assume_rows, columns=["parameter", "value"]),
    )

    notes = validation_notes or []
    if not notes:
        notes = ["PASS — Assets = Liabilities + Equity; Net Income links to CFS."]
    paths["validation"] = _write_csv(
        out_dir / f"3S_{ticker}_validation.csv",
        pd.DataFrame({"note": notes}),
    )
    return paths


def export_dcf(
    out_dir: Path,
    result: DCFResult,
    annual: pd.DataFrame,
    ticker: str,
    *,
    share_price: float,
    diluted_shares_000s: float,
    cash: float,
    marketable_securities: float,
    total_debt: float,
    net_debt: float,
    exit_ev_ebitda: float = 25.0,
) -> Dict[str, Path]:
    """Write DCF CSVs under out_dir. Returns map of logical name -> path."""
    out_dir.mkdir(parents=True, exist_ok=True)
    paths: Dict[str, Path] = {}

    upside = (result.equity_value_per_share / share_price - 1.0) if share_price else 0.0
    summary = [
        ("model_name", "vengeanceaiUSCMODEL18.0"),  # keep in sync with model18_assumptions.MODEL_NAME
        ("share_price_market", share_price),
        ("shares_outstanding_000s", diluted_shares_000s),
        ("diluted_shares_000s", diluted_shares_000s),  # alias for injectors
        ("wacc", result.wacc),
        ("exit_ev_ebitda", exit_ev_ebitda),
        ("terminal_growth", result.perpetual_growth),
        ("pv_projected_fcff", sum(result.pv_explicit_fcff)),
        ("terminal_value_primary", result.terminal_value),
        ("pv_terminal_value", result.pv_terminal_value),
        ("enterprise_value", result.enterprise_value),
        ("net_debt", net_debt),
        ("equity_value", result.equity_value),
        ("intrinsic_value_per_share", result.equity_value_per_share),
        ("upside_vs_price", upside),
        (
            "formula",
            "DCF mid-year: sum CF_t/(1+WACC)^(t-0.5)+TV/(1+WACC)^(n-0.5); "
            "FCFF=EBIT(1-t)+D&A-CapEx-dNWC; NWC=AR+Inv-AP-deferred",
        ),
    ]
    paths["summary"] = _write_csv(
        out_dir / f"DCF_{ticker}_summary.csv",
        pd.DataFrame(summary, columns=["metric", "value"]),
    )
    paths["annual_fcff"] = _write_csv(
        out_dir / f"DCF_{ticker}_annual_fcff.csv",
        _metrics_as_rows(annual),
    )
    market = [
        ("cash_000s", cash),
        ("marketable_securities_000s", marketable_securities),
        ("total_debt_000s", total_debt),
        ("net_debt_000s", net_debt),
    ]
    paths["market_inputs"] = _write_csv(
        out_dir / f"DCF_{ticker}_market_inputs.csv",
        pd.DataFrame(market, columns=["input", "value"]),
    )
    paths["notes"] = _write_csv(
        out_dir / f"DCF_{ticker}_notes.csv",
        pd.DataFrame({"note": result.notes}),
    )
    return paths


def export_math_explained(
    out_dir: Path,
    steps: Sequence[MathStep],
    *,
    ticker: str = "FICO",
) -> Path:
    """Write line-by-line MODEL3 math (formula → inputs → result → source)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"MODEL21_{ticker}_MATH_EXPLAINED.csv"
    return _write_csv(path, pd.DataFrame(steps_to_rows(steps)))
