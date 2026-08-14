"""CAPM / WACC stack — raw market CAPM + Yacktman credit-adjusted CAPM (MODEL17).

Raw CAPM (MODEL16-style):
- Rf: US 10Y ≈ 4.67% (Yahoo ^TNX / FRED DGS10)
- ERP: Damodaran implied ERP ≈ 4.28%
- Beta_raw: Yahoo Finance 5Y monthly ≈ 1.32
- Pre-tax kd: FICO 6.250% Senior Notes due 2034 (SEC 8-K)
- Tax: FY2025 effective ≈ 18.77% (SEC 10-K)
- Weights: ~80% equity / 20% debt

Yacktman-adjusted CAPM (MODEL17 primary D6):
- Blume-adjust Yahoo β toward 1.0, then blend with AAA-coupon β for
  GSE-mandated Scores (utility-like credit) — every term stays algebraic.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass


@dataclass(frozen=True)
class WaccInputs:
    risk_free_rate: float = 0.0467
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
            f"Beta={self.beta:.3f}",
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


# Raw CAPM (Yahoo β) — MODEL16 reference
MODEL3_WACC = round(WaccInputs().wacc, 4)  # ≈ 0.0927

# --- Yacktman credit-adjusted CAPM (MODEL17 primary) ---
# Blume (1971): β_adj = (2/3)·β_raw + (1/3)·1.0 — mean-reversion toward market
BETA_RAW = 1.32
BETA_BLUME = (2.0 / 3.0) * BETA_RAW + (1.0 / 3.0) * 1.0  # ≈ 1.213
# AAA-coupon β: GSE-mandated Scores ≈ regulated utility / high-grade data franchise
# (not Yahoo 5Y software noise). Credit rating lens → lower equity risk.
BETA_AAA_COUPON = 0.70
# Blend: 30% Blume market signal + 70% AAA Scores credit
BETA_YACKTMAN = 0.30 * BETA_BLUME + 0.70 * BETA_AAA_COUPON  # ≈ 0.854

YACKTMAN_WACC_INPUTS = WaccInputs(beta=BETA_YACKTMAN)
MODEL17_CAPM_WACC = round(YACKTMAN_WACC_INPUTS.wacc, 4)  # ≈ 0.076–0.078
