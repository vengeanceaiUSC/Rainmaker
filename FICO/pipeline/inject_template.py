"""
Excel Template Injector — CSV inputs → Named Ranges (formulas preserved).

================================================================================
FORMULA PROTECTION RULE
================================================================================
We ONLY inject hardcoded historical numbers and assumption/driver inputs into
Excel Named Ranges that were deliberately attached to input cells.

We NEVER map or write calculated fields such as Gross Profit, Net Income,
Unlevered Free Cash Flow, or DCF Enterprise Value. Those cells keep their
original template formulas (e.g. Gross Profit = Revenue - COGS).

openpyxl.load_workbook(..., data_only=False)  # DEFAULT — required
  data_only=True would strip formulas down to last-cached values and break
  the model. We never use data_only=True when reading or saving the template.

After this script writes FICO/output/FICO_Completed_Model.xlsx, open the file
in Microsoft Excel / Google Sheets / LibreOffice. Excel will automatically
recalculate every untouched formula using the newly injected SEC inputs.
================================================================================
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from .named_range_map import (
    DCF_FORECAST_YEARS,
    FORECAST_N,
    HIST_N,
    HIST_YEARS,
    SCALAR_MAPS,
    SERIES_MAPS,
    FORMULA_OUTPUTS_DO_NOT_MAP,
    SHEET_3S,
)
from .bake_equations import bake_equations_into_dcf
from .assumption_explanations import (
    export_assumption_explanations_csv,
    write_assumption_explanations,
)
from .bake_forecast_equations import bake_forecast_equations
from .equation_explanations import export_all_equations_csv, write_equation_comments
from .write_source_index import export_source_index_csv, write_cover_source_index
from .export_model2 import export_model2_csvs
from .fix_schedules import fix_three_statement_schedules
from .model10_assumptions import EXIT_EV_EBITDA, MODEL_NAME, SGA_FLOOR_PCT
from .prepare_template import OUT_TEMPLATE, build_template
from .wire_dcf import wire_dcf_to_three_statement

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output"
DEFAULT_OUT = OUTPUT / "FICO_Completed_Model.xlsx"

BLUE = Font(name="Calibri", color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")

_CELL_REF_RE = re.compile(
    r"(?:'?([^']+)'?!)\$?([A-Za-z]+)\$?(\d+)"
)


class InjectWarning(UserWarning):
    pass


def resolve_named_range(wb, name: str) -> Optional[Tuple[str, str]]:
    """
    Resolve a Defined Name to (sheet_title, cell_coordinate), e.g. ('3 Statement Model', 'E24').
    Returns None if missing / unresolvable (caller should warn and continue).
    """
    try:
        defn = wb.defined_names[name]
    except KeyError:
        return None

    # Prefer destinations API when available
    try:
        destinations = list(defn.destinations)
        if destinations:
            sheet, coord = destinations[0]
            # coord may be '$E$24' or 'E24' or a range 'E24:I24' — take top-left
            coord = str(coord).replace("$", "").split(":")[0]
            return sheet, coord
    except Exception:
        pass

    attr = getattr(defn, "attr_text", None) or str(defn.attr_text)
    m = _CELL_REF_RE.search(attr)
    if not m:
        return None
    sheet, col, row = m.group(1), m.group(2), m.group(3)
    return sheet, f"{col}{row}"


def _offset_cell(coord: str, col_offset: int) -> str:
    col_letters, row = coordinate_from_string(coord)
    col_idx = column_index_from_string(col_letters) + col_offset
    return f"{get_column_letter(col_idx)}{row}"


def inject_value(wb, named_range: str, value: Any, *, style_input: bool = True) -> bool:
    """Write value into the cell pointed to by named_range. False if missing."""
    resolved = resolve_named_range(wb, named_range)
    if resolved is None:
        print(f"  WARNING: Named Range missing/unresolvable: {named_range} — skipped")
        return False
    sheet_name, coord = resolved
    if sheet_name not in wb.sheetnames:
        print(f"  WARNING: Sheet {sheet_name!r} for {named_range} not found — skipped")
        return False
    ws = wb[sheet_name]
    cell = ws[coord]
    # Guard: never overwrite a formula cell even if someone mis-named it
    if isinstance(cell.value, str) and cell.value.startswith("="):
        print(
            f"  WARNING: {named_range} → {sheet_name}!{coord} holds a formula; "
            f"refusing to overwrite (formula protection)"
        )
        return False
    cell.value = value
    if style_input and isinstance(value, (int, float)):
        cell.font = BLUE
        cell.fill = INPUT_FILL
    return True


def inject_series(wb, named_range_start: str, values: List[Any]) -> int:
    """Inject values horizontally starting at named_range_start. Returns # written."""
    resolved = resolve_named_range(wb, named_range_start)
    if resolved is None:
        print(f"  WARNING: Named Range missing/unresolvable: {named_range_start} — series skipped")
        return 0
    sheet_name, start_coord = resolved
    if sheet_name not in wb.sheetnames:
        print(f"  WARNING: Sheet {sheet_name!r} for {named_range_start} not found — series skipped")
        return 0
    ws = wb[sheet_name]
    written = 0
    for i, val in enumerate(values):
        coord = _offset_cell(start_coord, i)
        cell = ws[coord]
        if isinstance(cell.value, str) and cell.value.startswith("="):
            print(
                f"  WARNING: {named_range_start}+{i} → {sheet_name}!{coord} is a formula; "
                f"skipping (formula protection)"
            )
            continue
        cell.value = val
        if isinstance(val, (int, float)):
            cell.font = BLUE
            cell.fill = INPUT_FILL
        written += 1
    return written


def _load_wide_csv(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    if "line_item" not in df.columns:
        raise ValueError(f"{path} missing line_item column")
    df = df.set_index("line_item")
    # normalize year columns to int where possible
    new_cols = []
    for c in df.columns:
        s = str(c)
        if s.startswith("FY"):
            new_cols.append(int(s.replace("FY", "")))
        else:
            try:
                new_cols.append(int(float(s)))
            except ValueError:
                new_cols.append(s)
    df.columns = new_cols
    return df


def _load_kv_csv(path: Path, key_col: str, val_col: str) -> Dict[str, Any]:
    df = pd.read_csv(path)
    out: Dict[str, Any] = {}
    for _, row in df.iterrows():
        k = str(row[key_col])
        v = row[val_col]
        out[k] = v
    return out


def load_csv_bundle(output_dir: Path) -> Dict[str, Any]:
    """Load pipeline CSV outputs into structures the mapper understands."""
    bundle: Dict[str, Any] = {}
    bundle["income_statement"] = _load_wide_csv(output_dir / "3S_FICO_income_statement.csv")
    bundle["balance_sheet"] = _load_wide_csv(output_dir / "3S_FICO_balance_sheet.csv")
    bundle["cash_flow"] = _load_wide_csv(output_dir / "3S_FICO_cash_flow.csv")
    bundle["annual_fcff"] = _load_wide_csv(output_dir / "DCF_FICO_annual_fcff.csv")
    bundle["assumptions"] = _load_kv_csv(
        output_dir / "3S_FICO_assumptions.csv", "parameter", "value"
    )
    bundle["dcf_summary"] = _load_kv_csv(
        output_dir / "DCF_FICO_summary.csv", "metric", "value"
    )
    bundle["market_inputs"] = _load_kv_csv(
        output_dir / "DCF_FICO_market_inputs.csv", "input", "value"
    )
    # Derived scalars
    mkt = bundle["market_inputs"]
    cash = float(mkt.get("cash_000s", 0) or 0)
    mkt_secs = float(mkt.get("marketable_securities_000s", 0) or 0)
    mkt["cash_plus_mkt"] = cash + mkt_secs
    first_forecast = DCF_FORECAST_YEARS[0]
    last_hist = HIST_YEARS[-1]
    # Exit multiple: prefer DCF summary / assumptions; fall back to MODEL3 default
    dcf_sum = bundle["dcf_summary"]
    exit_x = dcf_sum.get("exit_ev_ebitda") or bundle["assumptions"].get("exit_ev_ebitda")
    try:
        exit_x = float(exit_x) if exit_x is not None else EXIT_EV_EBITDA
    except (TypeError, ValueError):
        exit_x = EXIT_EV_EBITDA
    bundle["meta"] = {
        "base_year": HIST_YEARS[0],
        "exit_ev_ebitda": exit_x,
        # DCF year headers = YEAR(DATE(YEAR(D10)+period,…)) → 2026–2030
        "dcf_transaction_date": date(last_hist, 9, 30),
        "dcf_fiscal_year_end": date(first_forecast, 9, 30),
    }
    return bundle


def _series_values(bundle: Dict[str, Any], csv_file: str, line_item: str, years) -> List[float]:
    df: pd.DataFrame = bundle[csv_file]
    if line_item not in df.index:
        # try alternate encodings
        alts = [line_item, line_item.replace("Δ", "d"), "change_in_nwc" if "NWC" in line_item else line_item]
        found = None
        for a in alts:
            if a in df.index:
                found = a
                break
        if found is None:
            raise KeyError(f"{line_item} not in {csv_file} CSV index={list(df.index)}")
        line_item = found
    vals = []
    for y in years:
        if y not in df.columns:
            raise KeyError(f"Year {y} not in {csv_file} columns={list(df.columns)}")
        vals.append(float(df.loc[line_item, y]))
    return vals


def _deferred(bs, y) -> float:
    if "deferred_revenue" in bs.index:
        return float(bs.loc["deferred_revenue", y])
    return 0.0


def _gross_ar(bs, y) -> float:
    """Gross trade receivables (MODEL6 — not net of deferred)."""
    return float(bs.loc["accounts_receivable", y])


def _gross_ar_series(bundle: Dict[str, Any]) -> List[float]:
    bs = bundle["balance_sheet"]
    return [_gross_ar(bs, y) for y in HIST_YEARS]


def _equity_capital_plug(bundle: Dict[str, Any]) -> List[float]:
    """Simplified CFI BS plug with GROSS AR and Deferred as a liability (MODEL6).

    Equity Capital = Cash+GrossAR+Inv+PPE − AP − Deferred − Debt − RE(0).
    """
    bs = bundle["balance_sheet"]
    plugs = []
    for y in HIST_YEARS:
        assets = (
            float(bs.loc["cash", y])
            + _gross_ar(bs, y)
            + float(bs.loc["inventory", y])
            + float(bs.loc["ppe_net", y])
        )
        liab = (
            float(bs.loc["accounts_payable", y])
            + _deferred(bs, y)
            + float(bs.loc["total_debt", y])
        )
        plugs.append(assets - liab)
    return plugs


def _debt_issuance(bundle: Dict[str, Any]) -> List[float]:
    bs = bundle["balance_sheet"]
    debts = [float(bs.loc["total_debt", y]) for y in HIST_YEARS]
    # Prior-year debt unknown for FY1 → 0 issuance assumption for first year delta vs itself
    out = [0.0]
    for i in range(1, len(debts)):
        out.append(debts[i] - debts[i - 1])
    return out


def _delta_nwc(bundle: Dict[str, Any]) -> List[float]:
    """ΔNWC from operating NWC = GrossAR + Inv − AP − Deferred (MODEL6)."""
    bs = bundle["balance_sheet"]
    nwc = []
    for y in HIST_YEARS:
        nwc.append(
            _gross_ar(bs, y)
            + float(bs.loc["inventory", y])
            - float(bs.loc["accounts_payable", y])
            - _deferred(bs, y)
        )
    # Need prior NWC for first year — approximate with same (delta 0) if unknown
    deltas = [0.0]
    for i in range(1, len(nwc)):
        deltas.append(nwc[i] - nwc[i - 1])
    return deltas


def _forecast_assumption_series(bundle: Dict[str, Any]) -> Dict[str, List[float]]:
    """Build J..N assumption input vectors from CSV assumptions + last hist year."""
    a = bundle["assumptions"]
    is_ = bundle["income_statement"]
    bs = bundle["balance_sheet"]
    cf = bundle["cash_flow"]
    last = HIST_YEARS[-1]

    growth_raw = str(a.get("revenue_growth_path", "0.1,0.1,0.1,0.1,0.1"))
    growths = [float(x) for x in growth_raw.replace('"', "").split(",")]
    while len(growths) < FORECAST_N:
        growths.append(growths[-1])
    growths = growths[:FORECAST_N]

    cogs_pct = float(a["cogs_pct_revenue"])
    tax = float(a["tax_rate"])
    sga0 = float(is_.loc["sga", last])
    # Normalize FY25 restructuring out of SGA base when present
    if last == 2025:
        from .model3_assumptions import RESTRUCTURING_NORMALIZE_000s

        sga0 = max(0.0, sga0 - RESTRUCTURING_NORMALIZE_000s)
    rd0 = float(is_.loc["rd", last])
    rev0 = float(is_.loc["revenue", last])
    ar0 = float(bs.loc["accounts_receivable", last])
    inv0 = float(bs.loc["inventory", last])
    ap0 = float(bs.loc["accounts_payable", last])
    deferred0 = (
        float(bs.loc["deferred_revenue", last])
        if "deferred_revenue" in bs.index
        else 0.0
    )
    cogs0 = float(is_.loc["cogs", last])
    da_pct_rev = float(a["da_pct_revenue"])
    # CFI uses D&A % of opening PPE — approximate from last year DA/PPE
    ppe0 = float(bs.loc["ppe_net", last])
    da0 = float(is_.loc["da", last])
    da_pct_ppe = (da0 / ppe0) if ppe0 else 0.25
    interest0 = float(is_.loc["interest_expense", last])
    debt0 = float(bs.loc["total_debt", last])
    int_pct = (interest0 / debt0) if debt0 else 0.05
    # MODEL6: GROSS AR days (standard DSO); deferred projected separately
    ar_days = round(ar0 / rev0 * 365) if rev0 else 0
    inv_days = round(inv0 / cogs0 * 365) if cogs0 else 0
    ap_days = round(ap0 / cogs0 * 365) if cogs0 else 0

    capex_path_raw = str(a.get("capex_pct_path", a.get("capex_pct_revenue", "0.02")))
    if "," in capex_path_raw:
        capex_pcts = [float(x) for x in capex_path_raw.replace('"', "").split(",")]
    else:
        capex_pcts = [float(capex_path_raw)] * FORECAST_N
    while len(capex_pcts) < FORECAST_N:
        capex_pcts.append(capex_pcts[-1])
    capex_pcts = capex_pcts[:FORECAST_N]

    sga_improv_bps = float(a.get("sga_margin_improvement_bps", 0) or 0)
    sga_pct0 = float(a.get("sga_pct_revenue", sga0 / rev0 if rev0 else 0.26))

    sga_levels = []
    rd_levels = []
    capex_levels = []
    rev = rev0
    for i, g in enumerate(growths):
        rev = rev * (1 + g)
        sga_pct = max(SGA_FLOOR_PCT, sga_pct0 - (sga_improv_bps / 10_000.0) * (i + 1))
        sga_levels.append(round(rev * sga_pct, 1))
        rd_levels.append(round(rd0 * (rev / rev0), 1))
        capex_levels.append(round(rev * capex_pcts[i], 1))

    return {
        "ASSUM_RevGrowth_Start": growths,
        "ASSUM_COGSPct_Start": [cogs_pct] * FORECAST_N,
        "ASSUM_SGA_Start": sga_levels,
        "ASSUM_RD_Start": rd_levels,
        "ASSUM_DAPctPPE_Start": [da_pct_ppe] * FORECAST_N,
        "ASSUM_IntPct_Start": [int_pct] * FORECAST_N,
        "ASSUM_TaxRate_Start": [tax] * FORECAST_N,
        "ASSUM_ARDays_Start": [ar_days] * FORECAST_N,
        "ASSUM_InvDays_Start": [inv_days] * FORECAST_N,
        "ASSUM_APDays_Start": [ap_days] * FORECAST_N,
        "ASSUM_Capex_Start": capex_levels,
        "ASSUM_DebtIssuance_Start": [0.0] * FORECAST_N,
        "ASSUM_EquityIssuance_Start": [0.0] * FORECAST_N,
    }


def inject_all(
    template_path: Path,
    output_dir: Path,
    out_path: Path,
    *,
    rebuild_template: bool = False,
) -> Path:
    # Guardrail: never accidentally map formula outputs
    for bad in FORMULA_OUTPUTS_DO_NOT_MAP:
        for sm in SERIES_MAPS:
            if sm.line_item == bad and sm.csv_file != "annual_fcff":
                # operating_income intentionally not in SERIES for 3S; EBIT is DCF input
                if sm.named_range_start.startswith("IS_") or sm.named_range_start.startswith("BS_"):
                    raise RuntimeError(f"Illegal formula-output mapping: {sm}")

    if rebuild_template or not template_path.exists():
        print(f"[prep] Building Named-Range template → {template_path}")
        build_template(template_path)

    print(f"[load] Reading CSVs from {output_dir} (agent-readable source of truth)")
    bundle = load_csv_bundle(output_dir)

    print(
        f"[open] load_workbook({template_path.name}, data_only=False) — "
        f"preserving all template formulas"
    )
    wb = load_workbook(template_path, data_only=False)

    # Wire DCF ↔ 3-statement links BEFORE any inject so formula-protection
    # will refuse to overwrite EBIT/D&A/CapEx/ΔNWC/TV with CSV hardcodes.
    print("[wire] Linking DCF drivers to 3-statement formulas (no CSV hardcodes)…")
    wire_dcf_to_three_statement(wb)

    written = 0
    skipped = 0

    # Scalars
    print("[inject] Scalar assumption / market inputs…")
    for sm in SCALAR_MAPS:
        src = bundle.get(sm.csv_file, {})
        if sm.key not in src:
            print(f"  WARNING: key {sm.key!r} missing in {sm.csv_file} — skipped")
            skipped += 1
            continue
        val = src[sm.key]
        if isinstance(val, (datetime, date)):
            if isinstance(val, datetime):
                val = val.date()
        else:
            try:
                if sm.named_range == "IS_BaseYear":
                    val = int(float(val))
                else:
                    val = float(val)
            except (TypeError, ValueError):
                pass
        if inject_value(wb, sm.named_range, val):
            written += 1
        else:
            skipped += 1

    # Historical series from mapping (3-statement inputs only — not DCF UFCF drivers)
    print("[inject] Historical INPUT series (horizontal from Start names)…")
    for sm in SERIES_MAPS:
        try:
            vals = _series_values(bundle, sm.csv_file, sm.line_item, sm.years)
        except Exception as e:
            print(f"  WARNING: {sm.named_range_start} / {sm.line_item}: {e}")
            skipped += 1
            continue
        n = inject_series(wb, sm.named_range_start, vals[: sm.n_cols])
        written += n
        if n == 0:
            skipped += 1

    # Derived historical inputs not stored as clean CSV lines
    print("[inject] Derived BS/CF inputs (gross AR, equity plug, debt issuance, ΔNWC)…")
    # MODEL6: BS AR is GROSS; Deferred is a separate WC liability (row 88).
    n = inject_series(wb, "BS_AR_Start", _gross_ar_series(bundle))
    written += n
    n = inject_series(wb, "BS_EquityCapital_Start", _equity_capital_plug(bundle))
    written += n
    # RE historical → 0 (CFI simplified; NI accumulates in forecast formulas)
    n = inject_series(wb, "BS_RE_Start", [0.0] * HIST_N)
    written += n
    n = inject_series(wb, "CF_DebtIssuance_Start", _debt_issuance(bundle))
    written += n
    n = inject_series(wb, "CF_EquityIssuance_Start", [0.0] * HIST_N)
    written += n
    # MODEL10 CF layout: 64=SBC, 65=ΔNWC, 66=CFO — clear template formulas on
    # hist E65:I65 (old CFO row) so ΔNWC values can be written.
    ws3 = wb[SHEET_3S]
    for col in ("E", "F", "G", "H", "I"):
        for row in (64, 65, 66):
            ws3[f"{col}{row}"].value = None
    n = inject_series(wb, "CF_DeltaNWC_Start", _delta_nwc(bundle))
    written += n
    # Opening cash FY1 ≈ prior-year cash; use FY2021 cash - net change if available, else FY21
    cash_fy1 = float(bundle["balance_sheet"].loc["cash", HIST_YEARS[0]])
    # From 10-K cheat sheet FY2020 cash was 157394; prefer that if we can't derive
    opening = 157394.0
    if inject_value(wb, "CF_OpeningCash_FY1", opening):
        written += 1

    # Forecast assumption drivers (leave J24+ formulas alone)
    print("[inject] Forecast assumption INPUTS only (not forecast IS formula cells)…")
    for name, vals in _forecast_assumption_series(bundle).items():
        n = inject_series(wb, name, vals)
        written += n
        if n == 0:
            skipped += 1

    # Sync WC/PPE schedules so ΔNWC and D&A are not polluted by CFI sample residue
    print("[fix] Syncing WC + PPE supporting schedules to FICO history…")
    fix_three_statement_schedules(wb, bundle)

    # MODEL10: DSO/DPO WC, DA% of sales, flat CapEx, SBC, cash tax, financing
    print("[bake] Writing MODEL10 forecast equations into 3-statement…")
    bake_forecast_equations(wb)

    # Bake LIVE CAPM / FCFF / TV equations into DCF columns Q–V; D6 ← WACC formula
    print("[bake] Writing live CAPM/FCFF/TV equations into DCF!Q:V…")
    from .model10_assumptions import CASH_TAX_RATE as _CASH_TAX

    tax = float(_CASH_TAX)
    bake_equations_into_dcf(wb, tax_rate=tax)

    # Re-assert cash tax + comps exit after CAPM bake (may touch nearby cells)
    from openpyxl.styles import Font as _Font, PatternFill as _Fill
    from .model10_assumptions import (
        CASH_TAX_RATE,
        EXIT_EV_EBITDA,
        EXIT_EV_EBITDA_BEAR,
        EXIT_EV_EBITDA_BULL,
        PEER_MEDIAN_EV_EBITDA,
        URL_FICO_EV_EBITDA,
        URL_PEER_COMPS,
    )

    dcf = wb["DCF Model"]
    # MODEL10: DCF uses cash tax rate (not book tax from 3S J13)
    dcf["D5"] = float(CASH_TAX_RATE)
    dcf["D5"].number_format = "0.00%"
    dcf["D5"].font = _Font(name="Calibri", color="0000FF")
    dcf["D5"].fill = _Fill("solid", fgColor="FFF2CC")
    dcf["B5"] = f"Cash Tax Rate (3yr avg IncomeTaxesPaid/EBT = {CASH_TAX_RATE:.2%})"

    dcf["D8"] = float(EXIT_EV_EBITDA)
    dcf["D8"].number_format = "0.0"
    dcf["C8"] = (
        f"BASE {EXIT_EV_EBITDA:.1f}x ≈ peer median {PEER_MEDIAN_EV_EBITDA:.1f}x. "
        f"Bull {EXIT_EV_EBITDA_BULL:.0f}x / Bear {EXIT_EV_EBITDA_BEAR:.1f}x."
    )
    dcf["C8"].font = _Font(name="Calibri", italic=True, size=8, color="595959")
    dcf["F8"] = "Peer EV/EBITDA comps (click)"
    dcf["F8"].hyperlink = URL_PEER_COMPS
    dcf["F8"].font = _Font(name="Calibri", size=8, color="0563C1", underline="single")
    dcf["G8"] = "FICO spot EV/EBITDA (click)"
    dcf["G8"].hyperlink = URL_FICO_EV_EBITDA
    dcf["G8"].font = _Font(name="Calibri", size=8, color="0563C1", underline="single")

    # After DCF bake: equation commentary on all math cells, then restore
    # richer HOW/WHY/SOURCE on assumption rows 7–20
    print("[bake] Attaching ~20-word commentary to every math equation…")
    n_comments = write_equation_comments(wb)
    write_assumption_explanations(wb)
    print(f"  equation comments written: {n_comments}")

    # Cover note + clickable Source Index + PROMINENT Assumptions PDF on Cover
    if "Cover Page" in wb.sheetnames:
        from openpyxl.styles import (
            Alignment as _CoverAlign,
            Font as _CoverFont,
            PatternFill as _CoverFill,
        )
        from .model10_assumptions import (
            ASSUMPTIONS_PDF_URL,
            ASSUMPTIONS_PDF_VIEW_URL,
            cover_blurb,
        )

        cover = wb["Cover Page"]
        cover["C12"] = f"FICO — {MODEL_NAME} (3-Statement + DCF)"

        # Giant orange PDF download banner — MUST be on Cover (directly under title)
        for merge in list(cover.merged_cells.ranges):
            m = str(merge)
            if m.startswith("C13:") or m.startswith("E13:") or m.startswith("C22:"):
                try:
                    cover.unmerge_cells(m)
                except Exception:
                    pass
        safe_url = ASSUMPTIONS_PDF_URL.replace('"', '""')
        cover["C13"] = (
            f'=HYPERLINK("{safe_url}",'
            f'"⬇ CLICK HERE — DOWNLOAD ASSUMPTIONS LIST PDF (NO CHARTS)")'
        )
        cover["C13"].hyperlink = ASSUMPTIONS_PDF_URL
        cover["C13"].font = _CoverFont(
            name="Calibri", bold=True, size=20, color="FFFFFF", underline="single"
        )
        cover["C13"].fill = _CoverFill("solid", fgColor="FF6B00")
        cover["C13"].alignment = _CoverAlign(
            horizontal="center", vertical="center", wrap_text=True
        )
        try:
            cover.merge_cells("C13:G13")
        except Exception:
            pass
        cover.row_dimensions[13].height = 44

        # Keep Table of Contents on C14 if template had it; add PDF URL on E12 as well
        cover["E12"] = f'=HYPERLINK("{safe_url}","⬇ ASSUMPTIONS PDF")'
        cover["E12"].hyperlink = ASSUMPTIONS_PDF_URL
        cover["E12"].font = _CoverFont(
            name="Calibri", bold=True, size=14, color="FFFFFF", underline="single"
        )
        cover["E12"].fill = _CoverFill("solid", fgColor="FF6B00")

        cover["C21"] = (
            cover_blurb()
            + " ASSUMPTIONS PDF DOWNLOAD is the orange banner on Cover row 13 (C13). "
            + "Also listed in Source Index (cols E–G)."
        )
        # Second Cover hit: Notes section — large orange again
        cover["C22"] = (
            f'=HYPERLINK("{safe_url}",'
            f'"⬇ DOWNLOAD ASSUMPTIONS LIST PDF — CLICK THIS LINK")'
        )
        cover["C22"].hyperlink = ASSUMPTIONS_PDF_URL
        cover["C22"].font = _CoverFont(
            name="Calibri", bold=True, size=16, color="FFFFFF", underline="single"
        )
        cover["C22"].fill = _CoverFill("solid", fgColor="FF6B00")
        cover["C22"].alignment = _CoverAlign(
            horizontal="left", vertical="center", wrap_text=True
        )
        cover.row_dimensions[22].height = 30
        cover["C23"] = f'=HYPERLINK("{safe_url}","{safe_url}")'
        cover["C23"].hyperlink = ASSUMPTIONS_PDF_URL
        cover["C23"].font = _CoverFont(
            name="Calibri", bold=True, size=10, color="FF6B00", underline="single"
        )
        n_src = write_cover_source_index(wb)
        print(f"[bake] Cover Source Index links: {n_src}")
        print(f"[bake] Cover C13 Assumptions PDF banner → {ASSUMPTIONS_PDF_URL}")

    _fix_hash_display(wb)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    print("[export] Writing MODEL10 CSV sheet dumps…")
    m2 = export_model2_csvs(out_path, out_path.parent)
    for sheet, pth in m2.items():
        print(f"  {sheet} → {pth}")
    expl = export_assumption_explanations_csv(out_path.parent, ticker="FICO")
    print(f"  Assumptions Explained → {expl}")
    eqs = export_all_equations_csv(out_path.parent, ticker="FICO")
    print(f"  All Equations Explained → {eqs}")
    src = export_source_index_csv(out_path.parent, ticker="FICO")
    print(f"  Source Links → {src}")

    print()
    print(f"=== INJECTION COMPLETE ({MODEL_NAME}) ===")
    print(f"Wrote:    {out_path}")
    print(f"Cells OK: {written}  |  warnings/skips: {skipped}")
    print(
        "DCF linked to 3-statement; formulas preserved (data_only=False). "
        "Open in Excel to recalculate."
    )
    return out_path


def _fix_hash_display(wb) -> None:
    """Widen columns / compact formats so Excel does not show ########."""
    num_fmt = "#,##0.0;(#,##0.0);-"
    pct_fmt = "0.0%"
    date_fmt = "yyyy-mm-dd"
    price_fmt = "$#,##0.00"

    if "3 Statement Model" in wb.sheetnames:
        ws = wb["3 Statement Model"]
        ws.column_dimensions["B"].width = 42
        for col in range(4, 15):
            ws.column_dimensions[get_column_letter(col)].width = 16
        for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row or 110, 110), min_col=4, max_col=14):
            for cell in row:
                if cell.value is None:
                    continue
                # Year headers (row 2) must stay as plain years, not 2,026.0
                if cell.row == 2:
                    cell.number_format = "0"
                    continue
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    if "%" in str(cell.number_format):
                        cell.number_format = pct_fmt
                    elif "YEAR(" in cell.value.upper():
                        cell.number_format = "0"
                    else:
                        cell.number_format = num_fmt
                elif isinstance(cell.value, float) and abs(cell.value) <= 2 and cell.row <= 20:
                    cell.number_format = pct_fmt
                elif isinstance(cell.value, (int, float)):
                    cell.number_format = num_fmt

    if "DCF Model" in wb.sheetnames:
        ws = wb["DCF Model"]
        ws.column_dimensions["B"].width = 36
        for col in range(3, 12):
            ws.column_dimensions[get_column_letter(col)].width = 16
        for addr in ("D9", "D10"):
            ws[addr].number_format = date_fmt
        ws["D11"].number_format = price_fmt
        for addr in ("D5", "D6", "D7"):
            ws[addr].number_format = pct_fmt
        ws["D8"].number_format = "0.0"
        for row in ws.iter_rows(min_row=15, max_row=40, min_col=4, max_col=13):
            for cell in row:
                if cell.value is None:
                    continue
                if isinstance(cell.value, (datetime, date)):
                    cell.number_format = date_fmt
                elif isinstance(cell.value, str) and cell.value.startswith("="):
                    u = cell.value.upper()
                    if "YEAR(" in u and "YEARFRAC" not in u and "DATE(" not in u:
                        cell.number_format = "0"
                    elif "DATE(" in u:
                        cell.number_format = date_fmt
                    else:
                        cell.number_format = num_fmt
                elif isinstance(cell.value, (int, float)):
                    cell.number_format = num_fmt


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Inject CSV inputs into CFI Named-Range template")
    p.add_argument("--template", type=Path, default=OUT_TEMPLATE)
    p.add_argument("--csv-dir", type=Path, default=OUTPUT)
    p.add_argument("--out", type=Path, default=DEFAULT_OUT)
    p.add_argument("--rebuild-template", action="store_true")
    args = p.parse_args(argv)

    # Sanity: required CSVs
    required = [
        "3S_FICO_income_statement.csv",
        "3S_FICO_balance_sheet.csv",
        "3S_FICO_cash_flow.csv",
        "3S_FICO_assumptions.csv",
        "DCF_FICO_summary.csv",
        "DCF_FICO_annual_fcff.csv",
        "DCF_FICO_market_inputs.csv",
    ]
    missing = [f for f in required if not (args.csv_dir / f).exists()]
    if missing:
        print("Missing CSV inputs — run `python -m FICO.pipeline.run` first:")
        for f in missing:
            print(f"  - {args.csv_dir / f}")
        return 1

    inject_all(
        args.template,
        args.csv_dir,
        args.out,
        rebuild_template=args.rebuild_template,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
