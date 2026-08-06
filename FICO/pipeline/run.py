"""
CLI: fetch SEC XBRL → map to Pydantic → validate 3-statement → forecast → DCF → export.

Usage:
  python -m FICO.pipeline.run --ticker FICO
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import edgar, map_xbrl
from .dcf import dcf_annual_frame, run_dcf
from .export_csv import export_dcf, export_math_explained, export_three_statement
from .model3_assumptions import build_math_explained
from .three_statement import (
    ThreeStatementError,
    build_forecast,
    default_assumptions_from_history,
    validate_historical,
)

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUTPUT = ROOT / "output"

# Market bridge defaults — baked in model3_assumptions (10-Q / Yahoo; not LLM).
from .model13_assumptions import (
    CASH_10Q_000s,
    EXIT_EV_EBITDA,
    MKT_SECS_10Q_000s,
    MODEL_NAME,
    SHARE_PRICE,
    SHARES_OUTSTANDING_000s,
    TOTAL_DEBT_10Q_000s,
)

FICO_MARKET = {
    "share_price": SHARE_PRICE,
    "diluted_shares_000s": SHARES_OUTSTANDING_000s,
    "cash": CASH_10Q_000s,
    "marketable_securities": MKT_SECS_10Q_000s,
    "total_debt": TOTAL_DEBT_10Q_000s,
    "use_fy_net_debt": False,
}


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="XBRL → 3-statement → DCF pipeline")
    p.add_argument("--ticker", default="FICO")
    p.add_argument("--force-refresh", action="store_true")
    p.add_argument("--share-price", type=float, default=None)
    p.add_argument("--shares", type=float, default=None, help="Diluted shares in thousands")
    p.add_argument("--wacc", type=float, default=None)
    p.add_argument("--g", type=float, default=None, help="Perpetual growth")
    p.add_argument(
        "--exit-ev-ebitda",
        type=float,
        default=EXIT_EV_EBITDA,
        help=f"Exit EV/EBITDA multiple ({MODEL_NAME} default {EXIT_EV_EBITDA:.0f}x)",
    )
    p.add_argument(
        "--tv-method",
        choices=("exit", "gordon"),
        default="exit",
        help="Terminal value: exit EV/EBITDA (MODEL3 primary) or Gordon growth",
    )
    p.add_argument("--use-fy-net-debt", action="store_true", help="Net debt from FY BS instead of 10-Q bridge")
    p.add_argument(
        "--inject-excel",
        action="store_true",
        help="After CSV export, inject INPUTS only into CFI Named-Range template → FICO_Completed_Model.xlsx",
    )
    args = p.parse_args(argv)

    ticker = args.ticker.upper()
    print(f"[1/6] Fetching SEC companyfacts XBRL for {ticker}...")
    facts = edgar.fetch_companyfacts(
        ticker, cache_dir=DATA, force_refresh=args.force_refresh
    )

    print("[2/6] Mapping us-gaap tags → Pydantic 3-statement models...")
    fund = map_xbrl.build_fundamentals(facts, ticker=ticker)
    fund = map_xbrl.apply_audited_patches(fund)

    # Persist mapped JSON for auditability
    DATA.mkdir(parents=True, exist_ok=True)
    mapped_path = DATA / f"{ticker}_fundamentals.json"
    mapped_path.write_text(fund.model_dump_json(indent=2))
    print(f"      Wrote {mapped_path}")

    print("[3/6] Validating historical 3-statement linkages...")
    validation_notes: list[str] = []
    try:
        validate_historical(fund, tol=2.0)
        print("      PASS")
    except ThreeStatementError as e:
        # Soft-fail to notes for export; still raise for hard zeros
        validation_notes.append(str(e))
        print(f"      WARN: {e}")
        # Re-raise if revenue missing
        if "revenue is 0" in str(e):
            raise

    print(f"[4/6] Building {MODEL_NAME} forecast (Python math + Excel equations)...")
    from .wacc import WaccInputs
    from .model13_assumptions import MODEL13_WACC

    assumptions = default_assumptions_from_history(fund)
    assumptions.model_name = MODEL_NAME
    if args.wacc is not None:
        assumptions.wacc = args.wacc
    elif ticker == "FICO":
        assumptions.wacc = MODEL13_WACC
    if args.g is not None:
        assumptions.perpetual_growth = args.g
    wacc_notes = WaccInputs(tax_rate=assumptions.tax_rate).notes() if ticker == "FICO" else []

    # Market / share inputs
    mkt = dict(FICO_MARKET) if ticker == "FICO" else {
        "share_price": args.share_price or 0.0,
        "diluted_shares_000s": args.shares or 0.0,
        "cash": 0.0,
        "marketable_securities": 0.0,
        "total_debt": 0.0,
        "use_fy_net_debt": True,
    }
    if args.share_price is not None:
        mkt["share_price"] = args.share_price
    if args.shares is not None:
        mkt["diluted_shares_000s"] = args.shares
    if args.use_fy_net_debt:
        mkt["use_fy_net_debt"] = True

    assumptions.shares_thousands = float(mkt["diluted_shares_000s"])

    model = build_forecast(fund, assumptions)
    income = model["income_statement"]
    balance = model["balance_sheet"]
    cashflow = model["cash_flow"]
    proj_years = model["forecast_years"]

    # Net debt bridge
    last_hist = max(model["historical_years"])
    fy_cash = float(balance.loc[last_hist, "cash"])
    fy_debt = float(balance.loc[last_hist, "total_debt"])
    if mkt.get("use_fy_net_debt") or ticker != "FICO":
        cash = fy_cash
        mkt_secs = 0.0
        total_debt = fy_debt
    else:
        cash = float(mkt["cash"]) if mkt["cash"] else fy_cash
        # Prefer 10-Q cash+mkt when provided: split already in defaults
        mkt_secs = float(mkt.get("marketable_securities", 0.0))
        total_debt = float(mkt["total_debt"]) if mkt["total_debt"] else fy_debt
    net_debt = total_debt - cash - mkt_secs
    assumptions.net_debt_thousands = net_debt

    print("[5/6] Running DCF engine on 3-statement FCFF...")
    result = run_dcf(
        income,
        cashflow,
        assumptions,
        share_price=float(mkt["share_price"]),
        diluted_shares_000s=float(mkt["diluted_shares_000s"]),
        cash=cash,
        marketable_securities=mkt_secs,
        total_debt=total_debt,
        exit_ev_ebitda=args.exit_ev_ebitda,
        tv_method=args.tv_method,
        proj_years=proj_years,
    )
    result.ticker = ticker
    annual = dcf_annual_frame(result, income, cashflow, proj_years, assumptions.tax_rate)

    print("[6/6] Exporting CSV outputs (no Excel binaries)...")
    OUTPUT.mkdir(parents=True, exist_ok=True)
    # Remove legacy binary workbooks if present
    for legacy in (OUTPUT / f"3S_{ticker}.xlsx", OUTPUT / f"DCF_{ticker}.xlsx"):
        if legacy.exists():
            legacy.unlink()

    paths_3s = export_three_statement(
        OUTPUT,
        income,
        balance,
        cashflow,
        assumptions,
        ticker,
        validation_notes=validation_notes,
    )
    paths_dcf = export_dcf(
        OUTPUT,
        result,
        annual,
        ticker,
        share_price=float(mkt["share_price"]),
        diluted_shares_000s=float(mkt["diluted_shares_000s"]),
        cash=cash,
        marketable_securities=mkt_secs,
        total_debt=total_debt,
        net_debt=net_debt,
        exit_ev_ebitda=float(args.exit_ev_ebitda),
    )

    # Baked-in assumption math — every formula with inputs / result / source
    last_hist = max(model["historical_years"])
    math_steps = build_math_explained(
        assumptions=assumptions,
        income=income,
        cashflow=cashflow,
        proj_years=proj_years,
        net_debt=net_debt,
        share_price=float(mkt["share_price"]),
        shares_000s=float(mkt["diluted_shares_000s"]),
        exit_ev_ebitda=args.exit_ev_ebitda,
        hist_ar=float(balance.loc[last_hist, "accounts_receivable"]),
        hist_inv=float(balance.loc[last_hist, "inventory"]),
        hist_ap=float(balance.loc[last_hist, "accounts_payable"]),
        hist_deferred=float(balance.loc[last_hist, "deferred_revenue"])
        if "deferred_revenue" in balance.columns
        else 0.0,
        hist_revenue=float(income.loc[last_hist, "revenue"]),
    )
    math_path = export_math_explained(OUTPUT, math_steps, ticker=ticker)
    print(f"      Math explained → {math_path}")
    print("=== ASSUMPTION MATH (baked-in) ===")
    for step in math_steps:
        if step.section in ("WACC / CAPM", "Operating NWC", "Terminal value", "Enterprise → Equity"):
            print(f"  [{step.section}] {step.formula}")
            print(f"      → {step.result}")

    summary = {
        "model_name": MODEL_NAME,
        "ticker": ticker,
        "entity": fund.entity_name,
        "cik": fund.cik,
        "historical_years": model["historical_years"],
        "forecast_years": proj_years,
        "enterprise_value_000s": result.enterprise_value,
        "equity_value_000s": result.equity_value,
        "value_per_share": result.equity_value_per_share,
        "wacc": result.wacc,
        "wacc_stack": wacc_notes,
        "g": result.perpetual_growth,
        "exit_ev_ebitda": args.exit_ev_ebitda,
        "tv_method": args.tv_method,
        "net_debt_000s": net_debt,
        "nwc_pct_revenue": assumptions.nwc_pct_revenue,
        "capex_pct_path": assumptions.capex_pct_path,
        "outputs": {
            "three_statement": {k: str(v) for k, v in paths_3s.items()},
            "dcf": {k: str(v) for k, v in paths_dcf.items()},
            "math_explained": str(math_path),
        },
        "notes": result.notes,
        "math_explained": [
            {
                "section": s.section,
                "formula": s.formula,
                "result": s.result,
                "source": s.source,
            }
            for s in math_steps
        ],
    }
    summary_path = OUTPUT / f"{ticker}_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2))

    excel_out = None
    if args.inject_excel:
        print("[bonus] Injecting CSV INPUTS into CFI template (formulas untouched)…")
        from .inject_template import DEFAULT_OUT, inject_all
        from .prepare_template import OUT_TEMPLATE

        excel_out = inject_all(OUT_TEMPLATE, OUTPUT, DEFAULT_OUT, rebuild_template=True)
        summary["outputs"]["excel_completed_model"] = str(excel_out)
        summary_path.write_text(json.dumps(summary, indent=2))

    print()
    print("=== DONE ===")
    print("3-statement CSVs:")
    for pth in paths_3s.values():
        print(f"  {pth}")
    print("DCF CSVs:")
    for pth in paths_dcf.values():
        print(f"  {pth}")
    print(f"Summary:     {summary_path}")
    if excel_out:
        print(f"Excel model: {excel_out}  (open in Excel to recalculate formulas)")
    print(f"Intrinsic:   ${result.equity_value_per_share:,.2f} / share")
    print(f"EV:          ${result.enterprise_value:,.1f}k | Equity ${result.equity_value:,.1f}k")
    return 0


if __name__ == "__main__":
    sys.exit(main())
