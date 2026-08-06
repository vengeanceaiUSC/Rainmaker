"""Strict Pydantic models for a standard 3-statement layout + DCF inputs.

Numbers come ONLY from XBRL/mapping/math — never from an LLM guess.
Missing optional tags default to 0.0 with an explicit provenance note.
"""

from __future__ import annotations

from typing import Dict, List, Optional

from pydantic import BaseModel, Field, field_validator, model_validator


class LineItem(BaseModel):
    """A single financial line with optional XBRL provenance."""

    value: float = 0.0
    xbrl_tag: Optional[str] = None
    source: str = "missing_default_zero"

    @field_validator("value", mode="before")
    @classmethod
    def none_to_zero(cls, v):
        return 0.0 if v is None else float(v)


class IncomeStatement(BaseModel):
    fiscal_year: int
    period_end: str  # YYYY-MM-DD

    revenue: LineItem
    cogs: LineItem
    gross_profit: LineItem = Field(default_factory=LineItem)
    rd: LineItem = Field(default_factory=LineItem)
    sga: LineItem = Field(default_factory=LineItem)
    da: LineItem = Field(default_factory=LineItem)
    operating_income: LineItem = Field(default_factory=LineItem)
    interest_expense: LineItem = Field(default_factory=LineItem)
    other_income: LineItem = Field(default_factory=LineItem)
    ebt: LineItem = Field(default_factory=LineItem)
    tax_expense: LineItem = Field(default_factory=LineItem)
    net_income: LineItem

    @model_validator(mode="after")
    def derive_subtotals(self):
        # Deterministic identities when subtotals missing/zero
        if self.gross_profit.value == 0.0 and (self.revenue.value or self.cogs.value):
            self.gross_profit = LineItem(
                value=self.revenue.value - self.cogs.value,
                source="derived: revenue - cogs",
            )
        if self.operating_income.value == 0.0 and self.gross_profit.value:
            opex = self.rd.value + self.sga.value + self.da.value
            self.operating_income = LineItem(
                value=self.gross_profit.value - opex,
                source="derived: gp - rd - sga - da",
            )
        if self.ebt.value == 0.0 and self.operating_income.value:
            self.ebt = LineItem(
                value=self.operating_income.value - self.interest_expense.value + self.other_income.value,
                source="derived: opinc - interest + other",
            )
        return self


class BalanceSheet(BaseModel):
    fiscal_year: int
    period_end: str

    cash: LineItem
    accounts_receivable: LineItem
    inventory: LineItem = Field(default_factory=LineItem)
    other_current_assets: LineItem = Field(default_factory=LineItem)
    current_assets: LineItem = Field(default_factory=LineItem)
    ppe_net: LineItem
    other_assets: LineItem = Field(default_factory=LineItem)
    total_assets: LineItem

    accounts_payable: LineItem
    deferred_revenue: LineItem = Field(default_factory=LineItem)
    short_term_debt: LineItem = Field(default_factory=LineItem)
    other_current_liabilities: LineItem = Field(default_factory=LineItem)
    current_liabilities: LineItem = Field(default_factory=LineItem)
    long_term_debt: LineItem = Field(default_factory=LineItem)
    other_liabilities: LineItem = Field(default_factory=LineItem)
    total_liabilities: LineItem = Field(default_factory=LineItem)

    retained_earnings: LineItem = Field(default_factory=LineItem)
    other_equity: LineItem = Field(default_factory=LineItem)
    total_equity: LineItem

    @property
    def total_debt(self) -> float:
        return self.short_term_debt.value + self.long_term_debt.value

    def balance_delta(self) -> float:
        return self.total_assets.value - (self.total_liabilities.value + self.total_equity.value)


class CashFlowStatement(BaseModel):
    fiscal_year: int
    period_end: str

    net_income: LineItem
    da: LineItem = Field(default_factory=LineItem)
    change_in_nwc: LineItem = Field(default_factory=LineItem)
    cash_from_operations: LineItem
    capex: LineItem = Field(default_factory=LineItem)
    cash_from_investing: LineItem = Field(default_factory=LineItem)
    debt_issuance: LineItem = Field(default_factory=LineItem)
    equity_issuance: LineItem = Field(default_factory=LineItem)
    cash_from_financing: LineItem = Field(default_factory=LineItem)
    net_change_in_cash: LineItem = Field(default_factory=LineItem)


class HistoricalThreeStatement(BaseModel):
    """One fiscal year of linked 3-statement actuals."""

    income_statement: IncomeStatement
    balance_sheet: BalanceSheet
    cash_flow: CashFlowStatement

    def validate_linkages(self, tol: float = 1.0) -> List[str]:
        """Return list of validation errors (empty if OK). Units: $ thousands."""
        errors: List[str] = []
        y = self.income_statement.fiscal_year
        # Assets = L + E
        delta = self.balance_sheet.balance_delta()
        if abs(delta) > tol:
            errors.append(
                f"FY{y} BS imbalance: Assets - (L+E) = {delta:.3f} (tol={tol})"
            )
        # NI on IS vs CF
        ni_is = self.income_statement.net_income.value
        ni_cf = self.cash_flow.net_income.value
        if ni_cf != 0.0 and abs(ni_is - ni_cf) > tol:
            errors.append(
                f"FY{y} NI mismatch IS({ni_is:.1f}) vs CF({ni_cf:.1f})"
            )
        return errors


class CompanyFundamentals(BaseModel):
    ticker: str
    cik: str
    entity_name: str
    currency_unit: str = "USD_thousands"
    years: List[HistoricalThreeStatement]

    def by_year(self) -> Dict[int, HistoricalThreeStatement]:
        return {y.income_statement.fiscal_year: y for y in self.years}


class ForecastAssumptions(BaseModel):
    """Rules-based / MD&A-informed assumptions applied with Python math only.

    vengeanceaiUSCMODEL3 defaults: CAPM WACC (~9.24%), operating NWC
    (AR+Inv−AP−deferred revenue), CapEx fade, mild margin expansion.
    """

    revenue_growth: List[float] = Field(default_factory=lambda: [0.12, 0.10, 0.09, 0.08, 0.07])
    cogs_pct_revenue: float = 0.18
    sga_pct_revenue: float = 0.26
    rd_pct_revenue: float = 0.09
    da_pct_revenue: float = 0.008
    # Per-year CapEx/Sales (length forecast_years); falls back to scalar path
    capex_pct_revenue: float = 0.02
    capex_pct_path: List[float] = Field(default_factory=list)
    nwc_pct_revenue: float = 0.15
    # Optional per-year opex margin grind (subtracted from sga_pct each year)
    sga_margin_improvement_bps: float = 0.0  # e.g. 50 = −50bps SGA/Sales per year
    # MODEL10 cash-flow realism drivers
    sbc_pct_revenue: float = 0.0  # stock-based compensation / sales (CF & FCFF add-back)
    cash_tax_rate: float = 0.0  # IncomeTaxesPaid/EBT; 0 → fall back to tax_rate
    debt_issuance_annual: float = 0.0  # $000s net debt CF per forecast year
    buyback_annual: float = 0.0  # $000s hist avg buybacks (docs / residual anchor)
    tax_rate: float = 0.19
    interest_expense_level: float = 0.0  # absolute $000s if needed
    wacc: float = 0.0924  # vengeanceaiUSCMODEL3 CAPM default
    perpetual_growth: float = 0.03
    shares_thousands: float = 0.0
    net_debt_thousands: float = 0.0  # debt - cash for equity bridge
    forecast_years: int = 5
    model_name: str = "vengeanceaiUSCMODEL3"


class DCFResult(BaseModel):
    ticker: str
    explicit_fcff: List[float]
    discount_factors: List[float]
    pv_explicit_fcff: List[float]
    terminal_value: float
    pv_terminal_value: float
    enterprise_value: float
    equity_value: float
    equity_value_per_share: float
    wacc: float
    perpetual_growth: float
    notes: List[str] = Field(default_factory=list)
