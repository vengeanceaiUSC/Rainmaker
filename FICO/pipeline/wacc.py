"""CAPM / WACC stack for vengeanceaiUSCMODEL3 — market + SEC inputs, no LLM.

Sources (Aug 2026 research):
- Rf: US 10Y Treasury / FRED DGS10 ≈ 4.63%
- ERP: Damodaran implied ERP (Aug 1, 2026) ≈ 4.28%
- Beta: Yahoo Finance 5Y monthly ≈ 1.32
- Pre-tax kd: FICO 6.250% Senior Notes due 2034 (SEC 8-K)
- Tax: FY2025 effective 150,649 / 802,595 ≈ 18.77% (SEC 10-K)
- Weights: target ~80% equity / 20% debt at market
"""

from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass(frozen=True)
class WaccInputs:
    risk_free_rate: float = 0.0463
    equity_risk_premium: float = 0.0428
    beta: float = 1.32
    pre_tax_cost_of_debt: float = 0.0625
    tax_rate: float = 0.1877
    equity_weight: float = 0.80
    debt_weight: float = 0.20

    @property
    def cost_of_equity(self) -> float:
        return self.risk_free_rate + self.beta * self.equity_risk_premium

    @property
    def after_tax_cost_of_debt(self) -> float:
        return self.pre_tax_cost_of_debt * (1.0 - self.tax_rate)

    @property
    def wacc(self) -> float:
        return (
            self.equity_weight * self.cost_of_equity
            + self.debt_weight * self.after_tax_cost_of_debt
        )

    def notes(self) -> list[str]:
        return [
            f"Rf={self.risk_free_rate:.4%} (US 10Y)",
            f"ERP={self.equity_risk_premium:.4%} (Damodaran implied)",
            f"Beta={self.beta:.3f} (Yahoo 5Y monthly)",
            f"Ke=Rf+β·ERP={self.cost_of_equity:.4%}",
            f"Rd_pre={self.pre_tax_cost_of_debt:.4%} (6.250% notes due 2034, SEC 8-K)",
            f"Tax={self.tax_rate:.4%} (FY25 effective, SEC 10-K)",
            f"Weights E/D={self.equity_weight:.0%}/{self.debt_weight:.0%}",
            f"WACC={self.wacc:.4%}",
        ]

    def as_dict(self) -> dict:
        d = asdict(self)
        d["cost_of_equity"] = self.cost_of_equity
        d["after_tax_cost_of_debt"] = self.after_tax_cost_of_debt
        d["wacc"] = self.wacc
        return d


# Rounded to 1bp for model injection stability
MODEL3_WACC = round(WaccInputs().wacc, 4)  # ≈ 0.0924
