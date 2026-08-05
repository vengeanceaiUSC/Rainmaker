"""
Named-range mapping: CSV INPUTS only → Excel Defined Names.

STRICT RULE: Only map hardcoded driver/input cells (blue-font style in CFI).
Never map formula outputs (Gross Profit, Net Income, UFCF, Enterprise Value, etc.).
If a cell has no Named Range, the injector cannot touch it.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

# Sheet titles inside FICO/templates/CFI_Template.xlsx
SHEET_3S = "3 Statement Model"
SHEET_DCF = "DCF Model"

# Historical FY columns in the CFI 3-statement template: E..I = FY1..FY5
HIST_YEARS: List[int] = [2021, 2022, 2023, 2024, 2025]
HIST_N = len(HIST_YEARS)  # 5
# Forecast assumption columns: J..N = 5 years
FORECAST_N = 5


@dataclass(frozen=True)
class SeriesMap:
    """Map a CSV line_item across N years starting at a Named Range anchor."""

    named_range_start: str
    csv_file: str  # logical key: income_statement | balance_sheet | cash_flow | annual_fcff
    line_item: str
    years: Sequence[int]  # calendar years OR projection year labels handled by loader
    n_cols: int
    # optional transform applied to each value before write
    transform: Optional[str] = None  # e.g. "abs", "sga_as_salaries"


@dataclass(frozen=True)
class ScalarMap:
    """Map a single CSV scalar → one Named Range."""

    named_range: str
    csv_file: str  # assumptions | dcf_summary | market_inputs
    key: str
    transform: Optional[str] = None


# ---------------------------------------------------------------------------
# Named Range definitions to stamp onto the template (INPUT CELLS ONLY).
# attr_text uses Excel absolute refs; Start anchors support horizontal fill.
# ---------------------------------------------------------------------------
TEMPLATE_NAMED_RANGES: Dict[str, str] = {
    # --- 3-statement: base year driver ---
    "IS_BaseYear": f"'{SHEET_3S}'!$E$2",
    # --- IS historical inputs (Start = FY1 / col E); FY1..FY5 aliases ---
    "IS_Revenue_Start": f"'{SHEET_3S}'!$E$24",
    "IS_COGS_Start": f"'{SHEET_3S}'!$E$25",
    "IS_SGA_Start": f"'{SHEET_3S}'!$E$28",  # Salaries and Benefits ← SG&A
    "IS_RD_Start": f"'{SHEET_3S}'!$E$29",  # Rent/Overhead ← R&D residual
    "IS_DA_Start": f"'{SHEET_3S}'!$E$30",
    "IS_Interest_Start": f"'{SHEET_3S}'!$E$31",
    "IS_Tax_Start": f"'{SHEET_3S}'!$E$35",
    # Per-year aliases (same cells as Start+offset) — never create names on formula rows
    "IS_Revenue_FY1": f"'{SHEET_3S}'!$E$24",
    "IS_Revenue_FY2": f"'{SHEET_3S}'!$F$24",
    "IS_Revenue_FY3": f"'{SHEET_3S}'!$G$24",
    "IS_Revenue_FY4": f"'{SHEET_3S}'!$H$24",
    "IS_Revenue_FY5": f"'{SHEET_3S}'!$I$24",
    "IS_COGS_FY1": f"'{SHEET_3S}'!$E$25",
    "IS_COGS_FY2": f"'{SHEET_3S}'!$F$25",
    "IS_COGS_FY3": f"'{SHEET_3S}'!$G$25",
    "IS_COGS_FY4": f"'{SHEET_3S}'!$H$25",
    "IS_COGS_FY5": f"'{SHEET_3S}'!$I$25",
    # --- BS historical inputs ---
    "BS_Cash_Start": f"'{SHEET_3S}'!$E$41",
    "BS_AR_Start": f"'{SHEET_3S}'!$E$42",
    "BS_Inventory_Start": f"'{SHEET_3S}'!$E$43",
    "BS_PPE_Start": f"'{SHEET_3S}'!$E$44",
    "BS_AP_Start": f"'{SHEET_3S}'!$E$48",
    "BS_Debt_Start": f"'{SHEET_3S}'!$E$49",
    "BS_EquityCapital_Start": f"'{SHEET_3S}'!$E$52",
    "BS_RE_Start": f"'{SHEET_3S}'!$E$53",
    # --- CF historical inputs ---
    "CF_DeltaNWC_Start": f"'{SHEET_3S}'!$E$64",
    "CF_Capex_Start": f"'{SHEET_3S}'!$E$68",
    "CF_DebtIssuance_Start": f"'{SHEET_3S}'!$E$72",
    "CF_EquityIssuance_Start": f"'{SHEET_3S}'!$E$73",
    # E77 formula is =D78; inject prior-year opening cash into D78 so E77 formula stays intact
    "CF_OpeningCash_FY1": f"'{SHEET_3S}'!$D$78",
    # --- Forecast assumption INPUTS (J..N) — leave forecast IS formulas alone ---
    "ASSUM_RevGrowth_Start": f"'{SHEET_3S}'!$J$7",
    "ASSUM_COGSPct_Start": f"'{SHEET_3S}'!$J$8",
    "ASSUM_SGA_Start": f"'{SHEET_3S}'!$J$9",
    "ASSUM_RD_Start": f"'{SHEET_3S}'!$J$10",
    "ASSUM_DAPctPPE_Start": f"'{SHEET_3S}'!$J$11",
    "ASSUM_IntPct_Start": f"'{SHEET_3S}'!$J$12",
    "ASSUM_TaxRate_Start": f"'{SHEET_3S}'!$J$13",
    "ASSUM_ARDays_Start": f"'{SHEET_3S}'!$J$15",
    "ASSUM_InvDays_Start": f"'{SHEET_3S}'!$J$16",
    "ASSUM_APDays_Start": f"'{SHEET_3S}'!$J$17",
    "ASSUM_Capex_Start": f"'{SHEET_3S}'!$J$18",
    "ASSUM_DebtIssuance_Start": f"'{SHEET_3S}'!$J$19",
    "ASSUM_EquityIssuance_Start": f"'{SHEET_3S}'!$J$20",
    # --- DCF assumption INPUTS (never EV / Equity Value / UFCF formula cells) ---
    "DCF_TaxRate": f"'{SHEET_DCF}'!$D$5",
    "DCF_WACC": f"'{SHEET_DCF}'!$D$6",
    "DCF_PerpetualGrowth": f"'{SHEET_DCF}'!$D$7",
    "DCF_ExitMultiple": f"'{SHEET_DCF}'!$D$8",
    # Date drivers — year headers =YEAR(DATE(YEAR($D$10)+period,…))
    "DCF_TransactionDate": f"'{SHEET_DCF}'!$D$9",
    "DCF_FiscalYearEnd": f"'{SHEET_DCF}'!$D$10",
    "DCF_SharePrice": f"'{SHEET_DCF}'!$D$11",
    "DCF_Shares": f"'{SHEET_DCF}'!$D$12",
    "DCF_Debt": f"'{SHEET_DCF}'!$D$13",
    "DCF_Cash": f"'{SHEET_DCF}'!$D$14",
    # D15 / EBIT / D&A / CapEx / ΔNWC are FORMULAS linked to 3-statement — do not name as CSV inject targets
}

# Explicitly documented NON-targets (formula outputs) — do not add these as names.
FORMULA_OUTPUTS_DO_NOT_MAP = (
    "gross_profit",
    "net_income",
    "operating_income",
    "ebt",
    "total_assets",
    "total_liabilities",
    "total_equity",
    "cash_from_operations",
    "FCFF",
    "Unlevered FCF",
    "enterprise_value",
    "equity_value",
    "intrinsic_value_per_share",
    "PV of FCFF",
    "Terminal value",
    "EBIT",  # DCF EBIT is a live link to 3-statement
    "D&A",
    "CapEx",
    "ΔNWC",
)


# Series injections for historical 3-statement inputs
SERIES_MAPS: List[SeriesMap] = [
    SeriesMap("IS_Revenue_Start", "income_statement", "revenue", HIST_YEARS, HIST_N),
    SeriesMap("IS_COGS_Start", "income_statement", "cogs", HIST_YEARS, HIST_N),
    SeriesMap("IS_SGA_Start", "income_statement", "sga", HIST_YEARS, HIST_N),
    SeriesMap("IS_RD_Start", "income_statement", "rd", HIST_YEARS, HIST_N),
    SeriesMap("IS_DA_Start", "income_statement", "da", HIST_YEARS, HIST_N),
    SeriesMap("IS_Interest_Start", "income_statement", "interest_expense", HIST_YEARS, HIST_N),
    SeriesMap("IS_Tax_Start", "income_statement", "tax_expense", HIST_YEARS, HIST_N),
    SeriesMap("BS_Cash_Start", "balance_sheet", "cash", HIST_YEARS, HIST_N),
    SeriesMap("BS_AR_Start", "balance_sheet", "accounts_receivable", HIST_YEARS, HIST_N),
    SeriesMap("BS_Inventory_Start", "balance_sheet", "inventory", HIST_YEARS, HIST_N),
    SeriesMap("BS_PPE_Start", "balance_sheet", "ppe_net", HIST_YEARS, HIST_N),
    SeriesMap("BS_AP_Start", "balance_sheet", "accounts_payable", HIST_YEARS, HIST_N),
    SeriesMap("BS_Debt_Start", "balance_sheet", "total_debt", HIST_YEARS, HIST_N),
    SeriesMap("CF_Capex_Start", "cash_flow", "capex", HIST_YEARS, HIST_N),
    # NOTE: DCF EBIT / D&A / CapEx / ΔNWC are Excel formulas → 3-statement (see wire_dcf.py)
]

# Forecast calendar years for DCF headers (D10 = first forecast FYE)
DCF_FORECAST_YEARS: List[int] = [2026, 2027, 2028, 2029, 2030]

SCALAR_MAPS: List[ScalarMap] = [
    ScalarMap("IS_BaseYear", "meta", "base_year"),
    ScalarMap("DCF_WACC", "assumptions", "wacc"),
    ScalarMap("DCF_PerpetualGrowth", "assumptions", "perpetual_growth"),
    ScalarMap("DCF_TaxRate", "assumptions", "tax_rate"),
    ScalarMap("DCF_SharePrice", "dcf_summary", "share_price_market"),
    ScalarMap("DCF_Shares", "dcf_summary", "shares_outstanding_000s"),
    ScalarMap("DCF_Debt", "market_inputs", "total_debt_000s"),
    ScalarMap("DCF_Cash", "market_inputs", "cash_plus_mkt"),  # computed
    ScalarMap("DCF_ExitMultiple", "meta", "exit_ev_ebitda"),
    ScalarMap("DCF_TransactionDate", "meta", "dcf_transaction_date"),
    ScalarMap("DCF_FiscalYearEnd", "meta", "dcf_fiscal_year_end"),
]
