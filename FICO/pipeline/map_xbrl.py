"""Map us-gaap XBRL companyfacts tags -> strict Pydantic 3-statement models.

Tag resolution is deterministic: prefer an ordered list of candidate tags,
pick the latest FY 10-K fact for the company's fiscal year-end month.
"""

from __future__ import annotations

from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from .models import (
    BalanceSheet,
    CashFlowStatement,
    CompanyFundamentals,
    HistoricalThreeStatement,
    IncomeStatement,
    LineItem,
)

# Candidate XBRL tags per standard line (first hit with FY data wins)
TAG_MAP: Dict[str, Sequence[str]] = {
    "revenue": (
        "RevenueFromContractWithCustomerExcludingAssessedTax",
        "Revenues",
        "SalesRevenueNet",
        "RevenueFromContractWithCustomerIncludingAssessedTax",
    ),
    "cogs": (
        "CostOfRevenue",
        "CostOfGoodsAndServicesSold",
        "CostOfGoodsSold",
    ),
    "rd": ("ResearchAndDevelopmentExpense",),
    "sga": ("SellingGeneralAndAdministrativeExpense",),
    "da": (
        "DepreciationDepletionAndAmortization",
        "DepreciationAndAmortization",
    ),
    "operating_income": ("OperatingIncomeLoss",),
    "interest_expense": (
        "InterestExpense",
        "InterestExpenseNonoperating",
        "InterestIncomeExpenseNet",
    ),
    "other_income": ("OtherNonoperatingIncomeExpense", "NonoperatingIncomeExpense"),
    "tax_expense": ("IncomeTaxExpenseBenefit",),
    "net_income": ("NetIncomeLoss", "ProfitLoss"),
    "cash": (
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsAndShortTermInvestments",
        "Cash",
    ),
    "ar": (
        "AccountsReceivableNetCurrent",
        "AccountsReceivableNet",
    ),
    "inventory": ("InventoryNet", "InventoryFinishedGoods"),
    "current_assets": ("AssetsCurrent",),
    "ppe_net": ("PropertyPlantAndEquipmentNet",),
    "total_assets": ("Assets",),
    "ap": ("AccountsPayableCurrent", "AccountsPayableTradeCurrent"),
    "short_term_debt": (
        "LongTermDebtCurrent",
        "DebtCurrent",
        "ShortTermBorrowings",
    ),
    "current_liabilities": ("LiabilitiesCurrent",),
    "long_term_debt": (
        "LongTermDebtNoncurrent",
        "LongTermDebt",
        "LongTermDebtAndCapitalLeaseObligations",
    ),
    "total_liabilities": ("Liabilities",),
    "retained_earnings": ("RetainedEarningsAccumulatedDeficit", "RetainedEarnings"),
    "total_equity": (
        "StockholdersEquity",
        "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    ),
    "ocf": ("NetCashProvidedByUsedInOperatingActivities",),
    "capex": (
        "PaymentsToAcquirePropertyPlantAndEquipment",
        "PaymentsToAcquireProductiveAssets",
    ),
    "cfi": ("NetCashProvidedByUsedInInvestingActivities",),
    "cff": ("NetCashProvidedByUsedInFinancingActivities",),
    "cap_software": ("PaymentsToAcquireSoftware", "CapitalizedComputerSoftwareNet"),
}


def _fy_facts(
    facts: Dict[str, Any],
    tag: str,
    *,
    fye_month_day: str = "09-30",
    forms: Iterable[str] = ("10-K", "10-K/A"),
    ymin: int = 2018,
    ymax: int = 2030,
) -> Dict[int, Tuple[float, str]]:
    """Return {fiscal_year: (value_usd, filed)} for FY facts ending on fye_month_day."""
    usgaap = facts.get("facts", {}).get("us-gaap", {})
    if tag not in usgaap:
        return {}
    units = usgaap[tag].get("units", {})
    if "USD" not in units:
        return {}
    best: Dict[int, Tuple[float, str]] = {}
    for it in units["USD"]:
        if it.get("form") not in forms or it.get("fp") != "FY":
            continue
        end = it.get("end") or ""
        if not end.endswith(fye_month_day):
            continue
        y = int(end[:4])
        if y < ymin or y > ymax:
            continue
        filed = it.get("filed") or ""
        prev = best.get(y)
        if prev is None or filed >= prev[1]:
            best[y] = (float(it["val"]), filed)
    return best


def _pick_line(
    facts: Dict[str, Any],
    key: str,
    year: int,
    *,
    fye_month_day: str,
    scale: float = 1 / 1000.0,
) -> LineItem:
    for tag in TAG_MAP.get(key, ()):
        series = _fy_facts(facts, tag, fye_month_day=fye_month_day)
        if year in series:
            return LineItem(
                value=series[year][0] * scale,
                xbrl_tag=tag,
                source=f"xbrl:{tag}",
            )
    return LineItem(value=0.0, xbrl_tag=None, source="missing_default_zero")


def detect_fye_month_day(facts: Dict[str, Any]) -> str:
    """Infer dominant FYE month-day from NetIncomeLoss 10-K ends."""
    ends: Dict[str, int] = {}
    for tag in ("NetIncomeLoss", "Revenues", "Assets"):
        usgaap = facts.get("facts", {}).get("us-gaap", {})
        if tag not in usgaap or "USD" not in usgaap[tag]["units"]:
            continue
        for it in usgaap[tag]["units"]["USD"]:
            if it.get("form") not in ("10-K", "10-K/A") or it.get("fp") != "FY":
                continue
            end = it.get("end") or ""
            if len(end) >= 10:
                md = end[5:]
                ends[md] = ends.get(md, 0) + 1
    if not ends:
        return "12-31"
    return max(ends.items(), key=lambda kv: kv[1])[0]


def build_fundamentals(
    facts: Dict[str, Any],
    *,
    ticker: str,
    years: Optional[Sequence[int]] = None,
    fye_month_day: Optional[str] = None,
) -> CompanyFundamentals:
    fye = fye_month_day or detect_fye_month_day(facts)
    cik = str(facts.get("cik", "")).zfill(10)
    entity = facts.get("entityName") or ticker

    # Discover years from revenue or NI
    rev_years = set()
    for tag in TAG_MAP["revenue"]:
        rev_years |= set(_fy_facts(facts, tag, fye_month_day=fye))
    if not rev_years:
        for tag in TAG_MAP["net_income"]:
            rev_years |= set(_fy_facts(facts, tag, fye_month_day=fye))
    if years is None:
        years = sorted(y for y in rev_years if y >= 2021)[-5:]
    else:
        years = list(years)

    hist: List[HistoricalThreeStatement] = []
    for y in years:
        end = f"{y}-{fye}"
        revenue = _pick_line(facts, "revenue", y, fye_month_day=fye)
        cogs = _pick_line(facts, "cogs", y, fye_month_day=fye)
        rd = _pick_line(facts, "rd", y, fye_month_day=fye)
        sga = _pick_line(facts, "sga", y, fye_month_day=fye)
        da = _pick_line(facts, "da", y, fye_month_day=fye)
        opinc = _pick_line(facts, "operating_income", y, fye_month_day=fye)
        interest = _pick_line(facts, "interest_expense", y, fye_month_day=fye)
        # FICO prints interest expense, net — if InterestExpense missing, keep tag result
        other = _pick_line(facts, "other_income", y, fye_month_day=fye)
        tax = _pick_line(facts, "tax_expense", y, fye_month_day=fye)
        ni = _pick_line(facts, "net_income", y, fye_month_day=fye)

        is_stmt = IncomeStatement(
            fiscal_year=y,
            period_end=end,
            revenue=revenue,
            cogs=cogs,
            rd=rd,
            sga=sga,
            da=da,
            operating_income=opinc,
            interest_expense=interest,
            other_income=other,
            tax_expense=tax,
            net_income=ni,
        )

        cash = _pick_line(facts, "cash", y, fye_month_day=fye)
        ar = _pick_line(facts, "ar", y, fye_month_day=fye)
        inv = _pick_line(facts, "inventory", y, fye_month_day=fye)
        ca = _pick_line(facts, "current_assets", y, fye_month_day=fye)
        ppe = _pick_line(facts, "ppe_net", y, fye_month_day=fye)
        assets = _pick_line(facts, "total_assets", y, fye_month_day=fye)
        ap = _pick_line(facts, "ap", y, fye_month_day=fye)
        std = _pick_line(facts, "short_term_debt", y, fye_month_day=fye)
        cl = _pick_line(facts, "current_liabilities", y, fye_month_day=fye)
        ltd = _pick_line(facts, "long_term_debt", y, fye_month_day=fye)
        tl = _pick_line(facts, "total_liabilities", y, fye_month_day=fye)
        re = _pick_line(facts, "retained_earnings", y, fye_month_day=fye)
        equity = _pick_line(facts, "total_equity", y, fye_month_day=fye)

        # Infer liabilities if missing: Assets - Equity
        if tl.value == 0.0 and assets.value and equity.value:
            tl = LineItem(
                value=assets.value - equity.value,
                source="derived: assets - equity",
            )
        # other equity plug so BS can be checked on reported totals
        other_eq = LineItem(
            value=equity.value - re.value,
            source="derived: total_equity - retained_earnings",
        )
        other_liab = LineItem(
            value=max(0.0, tl.value - cl.value - ltd.value)
            if cl.value or ltd.value
            else 0.0,
            source="derived: total_liab - current - LT debt",
        )

        bs = BalanceSheet(
            fiscal_year=y,
            period_end=end,
            cash=cash,
            accounts_receivable=ar,
            inventory=inv,
            current_assets=ca,
            ppe_net=ppe,
            other_assets=LineItem(
                value=max(0.0, assets.value - ca.value - ppe.value) if ca.value else 0.0,
                source="derived: assets - CA - PPE",
            ),
            total_assets=assets,
            accounts_payable=ap,
            short_term_debt=std,
            current_liabilities=cl,
            long_term_debt=ltd,
            other_liabilities=other_liab,
            total_liabilities=tl,
            retained_earnings=re,
            other_equity=other_eq,
            total_equity=equity,
        )

        ocf = _pick_line(facts, "ocf", y, fye_month_day=fye)
        capex = _pick_line(facts, "capex", y, fye_month_day=fye)
        # Add capitalized software when present (FICO investing)
        cap_soft = _pick_line(facts, "cap_software", y, fye_month_day=fye)
        if cap_soft.value:
            capex = LineItem(
                value=abs(capex.value) + abs(cap_soft.value),
                xbrl_tag=f"{capex.xbrl_tag}+{cap_soft.xbrl_tag}",
                source="xbrl: PPE purchases + capitalized software",
            )
        else:
            capex = LineItem(value=abs(capex.value), xbrl_tag=capex.xbrl_tag, source=capex.source)
        cfi = _pick_line(facts, "cfi", y, fye_month_day=fye)
        cff = _pick_line(facts, "cff", y, fye_month_day=fye)
        cf_ni = LineItem(value=ni.value, xbrl_tag=ni.xbrl_tag, source="linked_from_IS")
        cf = CashFlowStatement(
            fiscal_year=y,
            period_end=end,
            net_income=cf_ni,
            da=da,
            cash_from_operations=ocf,
            capex=capex,
            cash_from_investing=cfi,
            cash_from_financing=cff,
            net_change_in_cash=LineItem(
                value=ocf.value + cfi.value + cff.value,
                source="derived: OCF+CFI+CFF",
            ),
        )
        hist.append(HistoricalThreeStatement(income_statement=is_stmt, balance_sheet=bs, cash_flow=cf))

    return CompanyFundamentals(
        ticker=ticker.upper(),
        cik=cik,
        entity_name=entity,
        years=hist,
    )


# Known 10-K note patches when XBRL taxonomy is incomplete (still not LLM).
# Values in $ thousands.
AUDITED_PATCHES: Dict[str, Dict[int, Dict[str, float]]] = {
    "FICO": {
        2023: {"interest_expense": 95_546},
        2024: {"interest_expense": 105_638, "total_debt": 2_209_021, "capex": 25_551},
        2025: {"interest_expense": 133_647, "total_debt": 3_055_691, "capex": 39_407},
    }
}


def apply_audited_patches(fund: CompanyFundamentals) -> CompanyFundamentals:
    patches = AUDITED_PATCHES.get(fund.ticker.upper(), {})
    for block in fund.years:
        y = block.income_statement.fiscal_year
        p = patches.get(y, {})
        if "interest_expense" in p:
            block.income_statement.interest_expense = LineItem(
                value=p["interest_expense"],
                source="10-K audited patch: Interest expense, net",
            )
        if "capex" in p:
            block.cash_flow.capex = LineItem(
                value=p["capex"],
                source="10-K audited patch: PPE + capitalized software",
            )
        if "total_debt" in p:
            # Split unknown; set LT debt to total and zero ST if both empty-ish
            td = p["total_debt"]
            st = block.balance_sheet.short_term_debt.value
            lt = block.balance_sheet.long_term_debt.value
            if abs(st + lt - td) > 1:
                block.balance_sheet.long_term_debt = LineItem(
                    value=td - st,
                    source="10-K audited patch: total debt footnote",
                )
    return fund
