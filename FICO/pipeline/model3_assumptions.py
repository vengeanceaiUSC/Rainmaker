"""vengeanceaiUSCMODEL3 — baked-in assumptions with explainable math.

Every driver used by the forecast / DCF lives here (or is derived from SEC
history via formulas documented below). No LLM guesses.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .wacc import MODEL3_WACC, WaccInputs


MODEL_NAME = "vengeanceaiUSCMODEL3"

# ---------------------------------------------------------------------------
# Baked policy constants (override only via CLI where exposed)
# ---------------------------------------------------------------------------
REVENUE_GROWTH_PATH: Tuple[float, ...] = (0.27, 0.16, 0.13, 0.10, 0.07)
# Year-1 ≈ FY2026 company guidance ~$2.53B / FY25 $1.991B − 1 ≈ 27%
GROWTH_RATIONALE = (
    "Y1 = FY2026 revenue guidance (~$2.53B) / FY2025 revenue ($1,990,869k) − 1 ≈ 27%; "
    "then fade 16% → 13% → 10% → 7% toward terminal g=3% (not straight-lined)."
)

CAPEX_STEADY_PCT = 0.010  # long-run maintenance + steady software capitalization
SGA_IMPROVEMENT_BPS = 50.0  # −50 bps of sales per forecast year
RESTRUCTURING_NORMALIZE_000s = 10_922.0  # FY25 one-time (SEC 10-K)
EXIT_EV_EBITDA = 25.0
PERPETUAL_GROWTH = 0.03
EQUITY_WEIGHT = 0.80
DEBT_WEIGHT = 0.20

# Market bridge (Q3 FY2026 10-Q + Yahoo)
SHARE_PRICE = 1046.23
SHARES_OUTSTANDING_000s = 21_597.635
CASH_10Q_000s = 248_444.0
MKT_SECS_10Q_000s = 56_093.0
TOTAL_DEBT_10Q_000s = 5_582_389.0

SOURCE_LINKS: Dict[str, str] = {
    "10-K FY2025": "https://www.sec.gov/Archives/edgar/data/814547/000081454725000030/fico-20250930.htm",
    "10-Q Q3 FY2026": "https://www.sec.gov/Archives/edgar/data/814547/000081454726000030/fico-20260630.htm",
    "8-K 6.250% notes": "https://www.sec.gov/Archives/edgar/data/814547/000119312526117936/d56220d8k.htm",
    "companyfacts": "https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json",
    "Damodaran ERP": "https://pages.stern.nyu.edu/adamodar/New_Home_Page/home.htm",
    "FRED DGS10": "https://fred.stlouisfed.org/series/DGS10",
    "Yahoo FICO": "https://finance.yahoo.com/quote/FICO/key-statistics/",
    "FY2026 guidance EX-99.1": (
        "https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/"
        "exhibit991erq32026.htm"
    ),
}


@dataclass
class MathStep:
    section: str
    formula: str
    inputs: str
    result: str
    source: str = ""


def _pct(x: float) -> str:
    return f"{x:.4%}"


def _num(x: float, decimals: int = 1) -> str:
    return f"{x:,.{decimals}f}"


def explain_wacc(tax_rate: Optional[float] = None) -> List[MathStep]:
    # Prefer Yacktman-adj CAPM (MODEL17/18 primary) when available.
    try:
        from .wacc import YACKTMAN_WACC_INPUTS, MODEL17_CAPM_WACC

        w = WaccInputs(
            risk_free_rate=YACKTMAN_WACC_INPUTS.risk_free_rate,
            equity_risk_premium=YACKTMAN_WACC_INPUTS.equity_risk_premium,
            beta=YACKTMAN_WACC_INPUTS.beta,
            pre_tax_cost_of_debt=YACKTMAN_WACC_INPUTS.pre_tax_cost_of_debt,
            tax_rate=tax_rate if tax_rate is not None else YACKTMAN_WACC_INPUTS.tax_rate,
            equity_weight=YACKTMAN_WACC_INPUTS.equity_weight,
            debt_weight=YACKTMAN_WACC_INPUTS.debt_weight,
        )
        model_wacc = MODEL17_CAPM_WACC
        beta_note = "Yacktman-adj (Blume+AAA)"
    except Exception:
        w = WaccInputs(tax_rate=tax_rate if tax_rate is not None else WaccInputs().tax_rate)
        model_wacc = MODEL3_WACC
        beta_note = "Yahoo 5Y"
    ke = w.cost_of_equity
    rd_at = w.after_tax_cost_of_debt
    steps = [
        MathStep(
            "WACC / CAPM",
            "Ke = Rf + β × ERP",
            f"Rf={_pct(w.risk_free_rate)} (US 10Y), β={w.beta:.3f} ({beta_note}), "
            f"ERP={_pct(w.equity_risk_premium)} (Damodaran Aug-2026 implied)",
            _pct(ke),
            SOURCE_LINKS["FRED DGS10"] + " | " + SOURCE_LINKS["Damodaran ERP"],
        ),
        MathStep(
            "WACC / CAPM",
            "Rd_aftertax = Rd_pre × (1 − t)",
            f"Rd_pre={_pct(w.pre_tax_cost_of_debt)} (6.250% Senior Notes due 2034), "
            f"t={_pct(w.tax_rate)} (FY25 tax/EBT = 150,649/802,595)",
            _pct(rd_at),
            SOURCE_LINKS["8-K 6.250% notes"] + " | " + SOURCE_LINKS["10-K FY2025"],
        ),
        MathStep(
            "WACC / CAPM",
            "WACC = We×Ke + Wd×Rd_aftertax",
            f"We={_pct(w.equity_weight)}, Wd={_pct(w.debt_weight)}, "
            f"Ke={_pct(ke)}, Rd_at={_pct(rd_at)}",
            f"{_pct(w.wacc)} (model uses {model_wacc:.4%} rounded to 1bp)",
            "Target capital structure at market (~80/20)",
        ),
    ]
    return steps


def explain_nwc(
    ar: float,
    inventory: float,
    ap: float,
    deferred: float,
    revenue: float,
) -> List[MathStep]:
    nwc = ar + inventory - ap - deferred
    pct = nwc / revenue if revenue else 0.0
    return [
        MathStep(
            "Operating NWC",
            "NWC = AR + Inventory − AP − DeferredRevenue",
            f"AR={_num(ar)} + Inv={_num(inventory)} − AP={_num(ap)} − Def={_num(deferred)}",
            f"{_num(nwc)} ($000s)",
            "SEC XBRL DeferredRevenueCurrent / 10-K",
        ),
        MathStep(
            "Operating NWC",
            "nwc_pct_revenue = NWC / Revenue",
            f"{_num(nwc)} / {_num(revenue)}",
            _pct(pct),
            "Held flat in forecast → ΔNWC_t = nwc%×Rev_t − nwc%×Rev_(t−1)",
        ),
    ]


def explain_fcff_and_dcf(
    *,
    ebit: Sequence[float],
    da: Sequence[float],
    capex: Sequence[float],
    dnwc: Sequence[float],
    tax_rate: float,
    wacc: float,
    g: float,
    exit_multiple: float,
    net_debt: float,
    shares_000s: float,
    share_price: float,
    mid_year: bool = True,
) -> List[MathStep]:
    steps: List[MathStep] = []
    fcffs: List[float] = []
    pvs: List[float] = []
    for i, (e, d, c, n) in enumerate(zip(ebit, da, capex, dnwc)):
        nopat = e * (1.0 - tax_rate)
        fcff = nopat + d - c - n
        fcffs.append(fcff)
        exp = (i + 0.5) if mid_year else (i + 1.0)
        df = 1.0 / ((1.0 + wacc) ** exp)
        pv = fcff * df
        pvs.append(pv)
        steps.append(
            MathStep(
                f"FCFF FY+{i+1}",
                "FCFF = EBIT×(1−t) + D&A − CapEx − ΔNWC",
                f"EBIT={_num(e)}, t={_pct(tax_rate)}, D&A={_num(d)}, "
                f"CapEx={_num(c)}, ΔNWC={_num(n)}",
                f"FCFF={_num(fcff)}; DF=(1+{_pct(wacc)})^(-{exp:g})={df:.6f}; PV={_num(pv)}",
                "Unlevered tax on EBIT (capital-structure neutral)",
            )
        )

    ebitda_n = ebit[-1] + da[-1]
    tv_exit = ebitda_n * exit_multiple
    tv_gordon = fcffs[-1] * (1.0 + g) / (wacc - g)
    df_n = 1.0 / ((1.0 + wacc) ** ((len(fcffs) - 0.5) if mid_year else len(fcffs)))
    pv_tv = tv_exit * df_n
    ev = sum(pvs) + pv_tv
    equity = ev - net_debt
    per_share = equity / shares_000s if shares_000s else 0.0
    upside = per_share / share_price - 1.0 if share_price else 0.0

    steps.extend(
        [
            MathStep(
                "Terminal value",
                "TV_exit = EBITDA_n × ExitMultiple   [PRIMARY]",
                f"EBITDA_n={_num(ebitda_n)} × {exit_multiple:.1f}x",
                _num(tv_exit),
                "Exit 25x = fade from current ~23–30x trading range toward mature multiple",
            ),
            MathStep(
                "Terminal value",
                "TV_gordon = FCFF_n×(1+g)/(WACC−g)   [CROSS-CHECK]",
                f"FCFF_n={_num(fcffs[-1])}, g={_pct(g)}, WACC={_pct(wacc)}",
                _num(tv_gordon),
                "Not primary: Y5 growth still > g so Gordon understates without normalization",
            ),
            MathStep(
                "Enterprise → Equity",
                "EV = Σ PV(FCFF) + PV(TV); Equity = EV − NetDebt; $/sh = Equity / Shares",
                f"ΣPV(FCFF)={_num(sum(pvs))}, PV(TV)={_num(pv_tv)}, "
                f"NetDebt={_num(net_debt)} (= Debt {_num(TOTAL_DEBT_10Q_000s)} "
                f"− Cash {_num(CASH_10Q_000s)} − MktSecs {_num(MKT_SECS_10Q_000s)}), "
                f"Shares={_num(shares_000s, 3)}k",
                f"EV={_num(ev)}; Equity={_num(equity)}; "
                f"${per_share:,.2f}/sh ({upside:+.1%} vs ${share_price:,.2f})",
                SOURCE_LINKS["10-Q Q3 FY2026"],
            ),
        ]
    )
    return steps


def explain_growth_and_margins(
    *,
    cogs_pct: float,
    sga_pct: float,
    rd_pct: float,
    da_pct: float,
    capex_path: Sequence[float],
    sga_improv_bps: float,
) -> List[MathStep]:
    steps = [
        MathStep(
            "Revenue growth",
            "Rev_t = Rev_(t−1) × (1 + g_t)",
            GROWTH_RATIONALE + f" Path={list(REVENUE_GROWTH_PATH)}",
            ", ".join(_pct(g) for g in REVENUE_GROWTH_PATH),
            SOURCE_LINKS.get(
                "FY2026 guidance EX-99.1",
                "https://www.sec.gov/Archives/edgar/data/814547/000081454726000031/exhibit991erq32026.htm",
            ),
        ),
        MathStep(
            "Margins",
            "COGS/SGA/R&D/DA = % of revenue; SGA_t% = SGA0% − bps×t/10000",
            f"COGS={_pct(cogs_pct)}, SGA0={_pct(sga_pct)} (FY25 SGA − ${_num(RESTRUCTURING_NORMALIZE_000s)} restructuring), "
            f"RD={_pct(rd_pct)}, DA={_pct(da_pct)}, grind={sga_improv_bps:.0f}bps/yr",
            f"EBIT_t = Rev×(1 − COGS% − SGA_t% − RD% − DA%)",
            SOURCE_LINKS["10-K FY2025"],
        ),
        MathStep(
            "CapEx fade",
            "CapEx_t = Rev_t × capex%_t; path fades peak → steady",
            f"Path={[round(x, 6) for x in capex_path]}; steady={_pct(CAPEX_STEADY_PCT)} "
            f"(FY25 CapEx peak includes capitalized software)",
            "Avoid locking peak software capitalization forever",
            SOURCE_LINKS["10-K FY2025"],
        ),
    ]
    return steps


def build_math_explained(
    *,
    assumptions: Any,
    income,
    cashflow,
    proj_years: Sequence[int],
    net_debt: float,
    share_price: float,
    shares_000s: float,
    exit_ev_ebitda: float,
    hist_ar: float,
    hist_inv: float,
    hist_ap: float,
    hist_deferred: float,
    hist_revenue: float,
) -> List[MathStep]:
    """Full MODEL3 math walkthrough for export / console."""
    steps: List[MathStep] = [
        MathStep(
            "Identity",
            "Model",
            MODEL_NAME,
            "Deterministic SEC XBRL → 3-statement → FCFF DCF (no LLM numbers)",
            SOURCE_LINKS["companyfacts"],
        )
    ]
    steps.extend(explain_wacc(tax_rate=float(assumptions.tax_rate)))
    steps.extend(
        explain_nwc(hist_ar, hist_inv, hist_ap, hist_deferred, hist_revenue)
    )
    steps.extend(
        explain_growth_and_margins(
            cogs_pct=float(assumptions.cogs_pct_revenue),
            sga_pct=float(assumptions.sga_pct_revenue),
            rd_pct=float(assumptions.rd_pct_revenue),
            da_pct=float(assumptions.da_pct_revenue),
            capex_path=list(assumptions.capex_pct_path or [assumptions.capex_pct_revenue]),
            sga_improv_bps=float(assumptions.sga_margin_improvement_bps),
        )
    )
    ebit = [float(income.loc[y, "operating_income"]) for y in proj_years]
    da = [float(cashflow.loc[y, "da"]) for y in proj_years]
    capex = [float(abs(cashflow.loc[y, "capex"])) for y in proj_years]
    dnwc = [float(cashflow.loc[y, "change_in_nwc"]) for y in proj_years]
    steps.extend(
        explain_fcff_and_dcf(
            ebit=ebit,
            da=da,
            capex=capex,
            dnwc=dnwc,
            tax_rate=float(assumptions.tax_rate),
            wacc=float(assumptions.wacc),
            g=float(assumptions.perpetual_growth),
            exit_multiple=float(exit_ev_ebitda),
            net_debt=float(net_debt),
            shares_000s=float(shares_000s),
            share_price=float(share_price),
            mid_year=True,
        )
    )
    return steps


def steps_to_rows(steps: Sequence[MathStep]) -> List[Dict[str, str]]:
    return [
        {
            "section": s.section,
            "formula": s.formula,
            "inputs": s.inputs,
            "result": s.result,
            "source": s.source,
        }
        for s in steps
    ]


def cover_blurb(wacc: float = MODEL3_WACC, exit_x: float = EXIT_EV_EBITDA) -> str:
    w = WaccInputs()
    return (
        f"{MODEL_NAME}: Ke=Rf+β·ERP={w.risk_free_rate:.2%}+{w.beta:.2f}×{w.equity_risk_premium:.2%}"
        f"={w.cost_of_equity:.2%}; WACC=80%×Ke+20%×Rd(1−t)={wacc:.2%}; "
        f"FCFF=EBIT(1−t)+DA−CapEx−ΔNWC; NWC=AR−AP−Deferred; "
        f"TV=EBITDA×{exit_x:.0f}x (Gordon cross-check); mid-year XNPV. "
        f"Open in Excel to recalculate."
    )
