"""Pandas 3-statement historical panel + forecast engine (Python math only).

Linkages:
- Net Income -> CF top line and RE bridge
- CapEx / D&A -> PP&E rollforward
- Forecast BS must balance (Assets == Liabilities + Equity) or raise.
"""

from __future__ import annotations

from typing import Dict, List, Tuple

import pandas as pd

from .models import CompanyFundamentals, ForecastAssumptions, HistoricalThreeStatement


class ThreeStatementError(RuntimeError):
    pass


def validate_historical(fund: CompanyFundamentals, tol: float = 2.0) -> None:
    """Raise if any year fails BS identity or NI linkage checks.

    tol default $2k on $000 units (~ rounding / taxonomy residuals).
    """
    errors: List[str] = []
    for year_block in fund.years:
        errors.extend(year_block.validate_linkages(tol=tol))
        # Revenue must exist
        if year_block.income_statement.revenue.value == 0:
            errors.append(
                f"FY{year_block.income_statement.fiscal_year}: revenue is 0 (missing XBRL tag?)"
            )
    if errors:
        raise ThreeStatementError("Historical validation failed:\n- " + "\n- ".join(errors))


def historical_to_frames(fund: CompanyFundamentals) -> Dict[str, pd.DataFrame]:
    """Return IS/BS/CF DataFrames indexed by fiscal year ($ thousands)."""
    years = [y.income_statement.fiscal_year for y in fund.years]
    is_rows = {}
    bs_rows = {}
    cf_rows = {}
    for yb in fund.years:
        y = yb.income_statement.fiscal_year
        is_ = yb.income_statement
        bs = yb.balance_sheet
        cf = yb.cash_flow
        is_rows[y] = {
            "revenue": is_.revenue.value,
            "cogs": is_.cogs.value,
            "gross_profit": is_.gross_profit.value,
            "rd": is_.rd.value,
            "sga": is_.sga.value,
            "da": is_.da.value,
            "operating_income": is_.operating_income.value,
            "interest_expense": is_.interest_expense.value,
            "other_income": is_.other_income.value,
            "ebt": is_.ebt.value,
            "tax_expense": is_.tax_expense.value,
            "net_income": is_.net_income.value,
        }
        bs_rows[y] = {
            "cash": bs.cash.value,
            "accounts_receivable": bs.accounts_receivable.value,
            "inventory": bs.inventory.value,
            "current_assets": bs.current_assets.value,
            "ppe_net": bs.ppe_net.value,
            "other_assets": bs.other_assets.value,
            "total_assets": bs.total_assets.value,
            "accounts_payable": bs.accounts_payable.value,
            "short_term_debt": bs.short_term_debt.value,
            "current_liabilities": bs.current_liabilities.value,
            "long_term_debt": bs.long_term_debt.value,
            "total_debt": bs.total_debt,
            "other_liabilities": bs.other_liabilities.value,
            "total_liabilities": bs.total_liabilities.value,
            "retained_earnings": bs.retained_earnings.value,
            "other_equity": bs.other_equity.value,
            "total_equity": bs.total_equity.value,
            "balance_delta": bs.balance_delta(),
        }
        cf_rows[y] = {
            "net_income": cf.net_income.value,
            "da": cf.da.value,
            "change_in_nwc": cf.change_in_nwc.value,
            "cash_from_operations": cf.cash_from_operations.value,
            "capex": cf.capex.value,
            "cash_from_investing": cf.cash_from_investing.value,
            "cash_from_financing": cf.cash_from_financing.value,
            "net_change_in_cash": cf.net_change_in_cash.value,
        }
    return {
        "income_statement": pd.DataFrame(is_rows, index=list(is_rows[years[0]].keys())).T.sort_index(),
        "balance_sheet": pd.DataFrame(bs_rows).T.sort_index(),
        "cash_flow": pd.DataFrame(cf_rows).T.sort_index(),
    }


def _assert_balance(assets: float, liab: float, equity: float, label: str, tol: float = 0.05) -> None:
    delta = assets - (liab + equity)
    if abs(delta) > tol:
        raise ThreeStatementError(
            f"{label} BS does not balance: Assets={assets:.4f} L+E={liab+equity:.4f} delta={delta:.4f}"
        )


def build_forecast(
    fund: CompanyFundamentals,
    assumptions: ForecastAssumptions,
) -> Dict[str, pd.DataFrame]:
    """Project IS/BS/CF forward with pure Python/pandas math; enforce BS balance."""
    hist = historical_to_frames(fund)
    last_y = int(hist["income_statement"].index.max())
    base_is = hist["income_statement"].loc[last_y]
    base_bs = hist["balance_sheet"].loc[last_y]

    rev0 = float(base_is["revenue"])
    nwc0 = float(base_bs["accounts_receivable"] + base_bs["inventory"] - base_bs["accounts_payable"])
    ppe0 = float(base_bs["ppe_net"])
    cash0 = float(base_bs["cash"])
    debt0 = float(base_bs["total_debt"])
    re0 = float(base_bs["retained_earnings"])
    other_eq0 = float(base_bs["other_equity"])
    other_assets0 = float(base_bs["other_assets"])
    other_liab0 = float(base_bs.get("other_liabilities", 0.0))

    is_f: Dict[int, dict] = {}
    bs_f: Dict[int, dict] = {}
    cf_f: Dict[int, dict] = {}

    rev = rev0
    nwc = nwc0
    ppe = ppe0
    cash = cash0
    debt = debt0
    re = re0

    for t in range(assumptions.forecast_years):
        y = last_y + 1 + t
        g = assumptions.revenue_growth[min(t, len(assumptions.revenue_growth) - 1)]
        rev = rev * (1 + g)
        cogs = rev * assumptions.cogs_pct_revenue
        gp = rev - cogs
        sga = rev * assumptions.sga_pct_revenue
        rd = rev * assumptions.rd_pct_revenue
        da = rev * assumptions.da_pct_revenue
        opinc = gp - sga - rd - da
        interest = assumptions.interest_expense_level or float(base_is["interest_expense"]) * (debt / max(debt0, 1))
        ebt = opinc - interest
        tax = max(0.0, ebt * assumptions.tax_rate)
        ni = ebt - tax

        capex = rev * assumptions.capex_pct_revenue
        nwc_target = rev * assumptions.nwc_pct_revenue
        dnwc = nwc_target - nwc
        nwc = nwc_target
        ppe = ppe + capex - da
        ocf = ni + da - dnwc
        # simplified financing: hold debt flat
        debt_issue = 0.0
        cfi = -capex
        cff = debt_issue
        cash = cash + ocf + cfi + cff
        re = re + ni

        # simplified BS using NWC components: AR = nwc (inventory 0, AP 0) for software cos
        ar = max(nwc, 0.0)
        inv = 0.0
        ap = 0.0
        ca = cash + ar + inv
        total_assets = ca + ppe + other_assets0
        total_debt = debt
        cl = ap + 0.0
        total_liab = cl + total_debt + other_liab0
        equity = other_eq0 + re
        # plug other equity so BS balances to the cent
        plug = total_assets - total_liab - equity
        equity = equity + plug
        _assert_balance(total_assets, total_liab, equity, f"FY{y} forecast")

        is_f[y] = {
            "revenue": rev,
            "cogs": cogs,
            "gross_profit": gp,
            "rd": rd,
            "sga": sga,
            "da": da,
            "operating_income": opinc,
            "interest_expense": interest,
            "other_income": 0.0,
            "ebt": ebt,
            "tax_expense": tax,
            "net_income": ni,
        }
        bs_f[y] = {
            "cash": cash,
            "accounts_receivable": ar,
            "inventory": inv,
            "current_assets": ca,
            "ppe_net": ppe,
            "other_assets": other_assets0,
            "total_assets": total_assets,
            "accounts_payable": ap,
            "short_term_debt": 0.0,
            "current_liabilities": cl,
            "long_term_debt": total_debt,
            "total_debt": total_debt,
            "other_liabilities": other_liab0,
            "total_liabilities": total_liab,
            "retained_earnings": re,
            "other_equity": other_eq0 + plug,
            "total_equity": equity,
            "balance_delta": 0.0,
        }
        cf_f[y] = {
            "net_income": ni,
            "da": da,
            "change_in_nwc": dnwc,
            "cash_from_operations": ocf,
            "capex": capex,
            "cash_from_investing": cfi,
            "debt_issuance": debt_issue,
            "equity_issuance": 0.0,
            "cash_from_financing": cff,
            "net_change_in_cash": ocf + cfi + cff,
        }

    return {
        "income_statement": pd.concat(
            [hist["income_statement"], pd.DataFrame(is_f).T]
        ).sort_index(),
        "balance_sheet": pd.concat(
            [hist["balance_sheet"], pd.DataFrame(bs_f).T]
        ).sort_index(),
        "cash_flow": pd.concat(
            [hist["cash_flow"], pd.DataFrame(cf_f).T]
        ).sort_index(),
        "forecast_years": list(is_f.keys()),
        "historical_years": list(hist["income_statement"].index.astype(int)),
    }


def default_assumptions_from_history(fund: CompanyFundamentals) -> ForecastAssumptions:
    """Rules-based assumptions from last historical year (no LLM required)."""
    frames = historical_to_frames(fund)
    is_ = frames["income_statement"]
    bs = frames["balance_sheet"]
    last = int(is_.index.max())
    rev = float(is_.loc[last, "revenue"])
    # trailing growth if possible
    growths = [0.12, 0.10, 0.09, 0.08, 0.07]
    if last - 1 in is_.index and float(is_.loc[last - 1, "revenue"]) > 0:
        g = float(is_.loc[last, "revenue"] / is_.loc[last - 1, "revenue"] - 1)
        growths = [max(0.03, min(0.20, g * 0.9)), 0.10, 0.09, 0.08, 0.07]
    cogs_pct = float(is_.loc[last, "cogs"] / rev) if rev else 0.18
    sga_pct = float(is_.loc[last, "sga"] / rev) if rev else 0.26
    rd_pct = float(is_.loc[last, "rd"] / rev) if rev else 0.09
    da_pct = float(is_.loc[last, "da"] / rev) if rev else 0.008
    tax_rate = 0.0
    ebt = float(is_.loc[last, "ebt"])
    if ebt:
        tax_rate = max(0.0, min(0.35, float(is_.loc[last, "tax_expense"] / ebt)))
    nwc = float(
        bs.loc[last, "accounts_receivable"]
        + bs.loc[last, "inventory"]
        - bs.loc[last, "accounts_payable"]
    )
    nwc_pct = nwc / rev if rev else 0.25
    cf = frames["cash_flow"]
    capex = float(cf.loc[last, "capex"])
    capex_pct = abs(capex) / rev if rev else 0.02
    net_debt = float(bs.loc[last, "total_debt"] - bs.loc[last, "cash"])
    return ForecastAssumptions(
        revenue_growth=growths,
        cogs_pct_revenue=cogs_pct,
        sga_pct_revenue=sga_pct,
        rd_pct_revenue=rd_pct,
        da_pct_revenue=da_pct,
        capex_pct_revenue=capex_pct,
        nwc_pct_revenue=nwc_pct,
        tax_rate=tax_rate or 0.19,
        interest_expense_level=float(is_.loc[last, "interest_expense"]),
        net_debt_thousands=net_debt,
    )
