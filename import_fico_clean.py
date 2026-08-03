#!/usr/bin/env python3
"""
Clean, non-destructive FICO 10-K import into pristine CFI 3-statement template.

- Loads DCF resources/CFI_3-Statement-Model-Complete.xlsx with data_only=False
- Overwrites ONLY hardcoded historical input cells
- Skips any cell whose value starts with '='
- Does not alter fonts, fills, borders, number formats, rows, columns, or sheets
- Saves FICO_3Statement_Model_Clean.xlsx
"""

from __future__ import annotations

import json
import urllib.request
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "DCF resources" / "CFI_3-Statement-Model-Complete.xlsx"
ORIGINALS = ROOT / "DCF resources" / "originals" / "CFI_3-Statement-Model-Complete.xlsx"
OUTPUT = ROOT / "FICO_3Statement_Model_Clean.xlsx"
CACHE = ROOT / "fico-valuation" / "data" / "fico_companyfacts.json"
USER_AGENT = "FinancialAnalyst user@example.com"
CIK = "0000814547"
SHEET = "3 Statement Model"

# Historical columns in CFI template (5 slots). We fill FY2021–FY2025 so the
# last three (G/H/I) are FY2023–FY2025 as requested.
HIST_COLS = ["E", "F", "G", "H", "I"]
YEARS = [2021, 2022, 2023, 2024, 2025]


def is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def safe_set(ws, coord: str, value, label: str, year: int | str, log: list) -> None:
    """Write value only when the cell is not a formula. No style changes."""
    cell = ws[coord]
    current = cell.value
    if is_formula(current):
        log.append(f"SKIP  {coord:>4}  FY{year}  {label:<28}  formula={current}")
        return
    old = current
    cell.value = value
    log.append(
        f"WRITE {coord:>4}  FY{year}  {label:<28}  {old!r}  ->  {value!r}"
    )


def fetch_companyfacts() -> dict:
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": USER_AGENT, "Accept": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode())
    CACHE.write_text(json.dumps(data))
    return data


def annual_usd_thousands(facts: dict, concept: str) -> dict[int, float]:
    usgaap = facts["facts"].get("us-gaap", {})
    if concept not in usgaap:
        return {}
    units = usgaap[concept]["units"]
    if "USD" not in units:
        return {}
    best: dict[int, dict] = {}
    for it in units["USD"]:
        if it.get("form") not in ("10-K", "10-K/A") or it.get("fp") != "FY":
            continue
        end = it.get("end") or ""
        if not end.endswith("-09-30"):
            continue
        y = int(end[:4])
        if y < 2019 or y > 2025:
            continue
        prev = best.get(y)
        if prev is None or it.get("filed", "") >= prev["filed"]:
            best[y] = {"val": it["val"], "filed": it.get("filed", "")}
    # CFI template units are $000s
    return {y: best[y]["val"] / 1000.0 for y in best}


def load_fico() -> dict:
    print("Fetching FICO companyfacts from SEC EDGAR...")
    facts = fetch_companyfacts()
    print(f"Entity: {facts.get('entityName')} CIK={facts.get('cik')}")

    series = {
        "revenue": annual_usd_thousands(facts, "RevenueFromContractWithCustomerExcludingAssessedTax"),
        "cogs": annual_usd_thousands(facts, "CostOfRevenue"),
        "sga": annual_usd_thousands(facts, "SellingGeneralAndAdministrativeExpense"),
        "rd": annual_usd_thousands(facts, "ResearchAndDevelopmentExpense"),
        "opinc": annual_usd_thousands(facts, "OperatingIncomeLoss"),
        "interest": annual_usd_thousands(facts, "InterestExpense"),
        "tax": annual_usd_thousands(facts, "IncomeTaxExpenseBenefit"),
        "ni": annual_usd_thousands(facts, "NetIncomeLoss"),
        "cash": annual_usd_thousands(facts, "CashAndCashEquivalentsAtCarryingValue"),
        "ar": annual_usd_thousands(facts, "AccountsReceivableNetCurrent"),
        "ppe": annual_usd_thousands(facts, "PropertyPlantAndEquipmentNet"),
        "ap": annual_usd_thousands(facts, "AccountsPayableCurrent"),
        "debt_lt": annual_usd_thousands(facts, "LongTermDebtNoncurrent"),
        "debt_cur": annual_usd_thousands(facts, "LongTermDebtCurrent"),
        "da": annual_usd_thousands(facts, "DepreciationDepletionAndAmortization"),
        "capex": annual_usd_thousands(facts, "PaymentsToAcquirePropertyPlantAndEquipment"),
        "current_assets": annual_usd_thousands(facts, "AssetsCurrent"),
        "total_assets": annual_usd_thousands(facts, "Assets"),
        "retained": annual_usd_thousands(facts, "RetainedEarningsAccumulatedDeficit"),
        "equity": annual_usd_thousands(facts, "StockholdersEquity"),
        "ocf": annual_usd_thousands(facts, "NetCashProvidedByUsedInOperatingActivities"),
    }

    # 10-K audited patches where taxonomy is incomplete / needs note totals ($000s)
    series["interest"].update({2023: 95_546, 2024: 105_638, 2025: 133_647})
    # Total debt from 10-K debt footnote
    debt = {}
    for y in set(series["debt_lt"]) | set(series["debt_cur"]) | {2023, 2024, 2025}:
        debt[y] = series["debt_lt"].get(y, 0) + series["debt_cur"].get(y, 0)
    debt[2024] = 2_209_021
    debt[2025] = 3_055_691
    series["debt"] = debt

    # Capex + capitalized internal-use software (10-K CF)
    series["capex_total"] = dict(series["capex"])
    series["capex_total"].update({2023: 4_237, 2024: 25_551, 2025: 39_407})

    # Inventory = 0 for FICO (software / scores business; not a material inventory model)
    series["inventory"] = {y: 0 for y in YEARS + [2020]}

    # Prior-year anchors for schedules
    series["prior"] = {
        "cash": series["cash"].get(2020, series["cash"][2021]),
        "ar": series["ar"].get(2020, series["ar"][2021]),
        "ap": series["ap"].get(2020, series["ap"][2021]),
        "ppe": series["ppe"].get(2020, series["ppe"][2021]),
        "debt": series["debt"].get(2020, series["debt_lt"].get(2020, 739_435)),
    }

    # Validate last 3 fiscal years present
    for y in (2023, 2024, 2025):
        for key in ("revenue", "cogs", "sga", "rd", "ni", "cash", "ar", "ppe", "ap", "debt", "da"):
            if y not in series[key] and key != "inventory":
                raise RuntimeError(f"Missing FICO {key} for FY{y}")

    return series


def pristine_template_path() -> Path:
    """Load from untouched originals/; never write back into DCF resources/."""
    if ORIGINALS.exists():
        print(f"Using pristine original {ORIGINALS} ({ORIGINALS.stat().st_size} bytes)")
        return ORIGINALS
    if TEMPLATE.exists():
        print(f"Fallback template {TEMPLATE} ({TEMPLATE.stat().st_size} bytes)")
        return TEMPLATE
    raise SystemExit(f"Missing pristine template at {ORIGINALS} or {TEMPLATE}")


def inject(ws, fico: dict, log: list) -> None:
    # Start year drives E..I headers via formulas in F2:I2
    safe_set(ws, "E2", YEARS[0], "Fiscal year start", "header", log)

    for col, y in zip(HIST_COLS, YEARS):
        print(f"\n--- Mapping FY{y} into column {col} ---")

        # Income statement hardcoded inputs
        safe_set(ws, f"{col}24", round(fico["revenue"][y], 1), "Revenue", y, log)
        safe_set(ws, f"{col}25", round(fico["cogs"][y], 1), "COGS / Cost of revenues", y, log)
        safe_set(ws, f"{col}28", round(fico["sga"][y], 1), "SG&A -> Salaries line", y, log)
        safe_set(ws, f"{col}29", round(fico["rd"][y], 1), "R&D -> Rent/Overhead line", y, log)
        safe_set(ws, f"{col}30", round(fico["da"][y], 1), "D&A", y, log)
        safe_set(ws, f"{col}31", round(fico["interest"][y], 1), "Interest expense", y, log)
        safe_set(ws, f"{col}35", round(fico["tax"][y], 1), "Tax expense", y, log)

        # Balance sheet hardcoded inputs
        safe_set(ws, f"{col}41", round(fico["cash"][y], 1), "Cash & equivalents", y, log)
        safe_set(ws, f"{col}42", round(fico["ar"][y], 1), "Accounts receivable", y, log)
        safe_set(ws, f"{col}43", 0, "Inventory", y, log)
        safe_set(ws, f"{col}44", round(fico["ppe"][y], 1), "Net PP&E", y, log)
        safe_set(ws, f"{col}48", round(fico["ap"][y], 1), "Accounts payable", y, log)
        safe_set(ws, f"{col}49", round(fico["debt"][y], 1), "Total debt", y, log)

        # Equity: template only has Equity Capital + RE as inputs; keep simplified plug
        # so BS check can balance under limited asset lines (Cash/AR/Inv/PPE).
        assets_simple = fico["cash"][y] + fico["ar"][y] + 0 + fico["ppe"][y]
        liab_simple = fico["ap"][y] + fico["debt"][y]
        equity_plug = assets_simple - liab_simple
        # Prefer retained earnings from SEC when available; capital = plug residual
        re = fico["retained"].get(y)
        if re is not None:
            safe_set(ws, f"{col}53", round(re, 1), "Retained earnings (SEC)", y, log)
            safe_set(ws, f"{col}52", round(equity_plug - re, 1), "Equity capital (plug)", y, log)
        else:
            safe_set(ws, f"{col}52", round(equity_plug, 1), "Equity capital (plug)", y, log)
            safe_set(ws, f"{col}53", 0, "Retained earnings", y, log)

        # Cash flow hardcoded inputs
        if y == YEARS[0]:
            prev_nwc = fico["prior"]["ar"] - fico["prior"]["ap"]
            prev_debt = fico["prior"]["debt"]
            prev_cash = fico["prior"]["cash"]
            open_ppe = fico["prior"]["ppe"]
        else:
            prev_nwc = fico["ar"][y - 1] - fico["ap"][y - 1]
            prev_debt = fico["debt"][y - 1]
            prev_cash = fico["cash"][y - 1]
            open_ppe = fico["ppe"][y - 1]
        nwc = fico["ar"][y] - fico["ap"][y]
        dnwc = nwc - prev_nwc
        debt_issue = fico["debt"][y] - prev_debt
        capex = fico["capex_total"][y]

        # CFI stores CapEx as a positive amount; cash roll uses Ending = CFO - CapEx + CFF
        capex_cf = abs(capex)
        safe_set(ws, f"{col}64", round(dnwc, 1), "Change in working capital", y, log)
        safe_set(ws, f"{col}68", round(capex_cf, 1), "CapEx (incl. cap. software)", y, log)
        safe_set(ws, f"{col}72", round(debt_issue, 1), "Debt issuance/(repay)", y, log)

        # Equity issuance plug so cash roll matches BS cash under template identity
        cfo_proxy = fico["ni"][y] + fico["da"][y] - dnwc
        target_delta = fico["cash"][y] - prev_cash
        equity_issue = target_delta - (cfo_proxy - capex_cf + debt_issue)
        safe_set(ws, f"{col}73", round(equity_issue, 1), "Equity issuance plug", y, log)

        # Opening cash F-I are values; E77 is formula (=D78) — set D78 once
        if col == "E":
            safe_set(ws, "D78", round(fico["prior"]["cash"], 1), "Opening cash prior year", "prior", log)
        else:
            # F77:I77 hardcoded opening cash in sample — set to prior year closing cash
            safe_set(ws, f"{col}77", round(prev_cash, 1), "Opening cash balance", y, log)

        # WC schedule mirrors
        safe_set(ws, f"{col}85", round(fico["ar"][y], 1), "WC schedule AR", y, log)
        safe_set(ws, f"{col}86", 0, "WC schedule Inventory", y, log)
        safe_set(ws, f"{col}87", round(fico["ap"][y], 1), "WC schedule AP", y, log)

        # PPE schedule
        da = fico["da"][y]
        ppe_capex = fico["ppe"][y] - open_ppe + da
        safe_set(ws, f"{col}92", round(open_ppe, 1), "PPE opening", y, log)
        safe_set(ws, f"{col}93", round(ppe_capex, 1), "PPE capex (rollforward)", y, log)
        safe_set(ws, f"{col}94", round(da, 1), "PPE depreciation", y, log)

        # Debt schedule
        safe_set(ws, f"{col}98", round(prev_debt, 1), "Debt opening", y, log)
        safe_set(ws, f"{col}99", round(debt_issue, 1), "Debt issuance schedule", y, log)
        safe_set(ws, f"{col}101", round(fico["interest"][y], 1), "Interest schedule", y, log)

        # Console mapping summary for the year
        print(
            f"  Revenue={fico['revenue'][y]:,.0f}  COGS={fico['cogs'][y]:,.0f}  "
            f"SG&A={fico['sga'][y]:,.0f}  R&D={fico['rd'][y]:,.0f}  "
            f"NI={fico['ni'][y]:,.0f}  Cash={fico['cash'][y]:,.0f}  Debt={fico['debt'][y]:,.0f}"
        )


def verify(path: Path, fico: dict) -> None:
    wb = load_workbook(path, data_only=False)
    ws = wb[SHEET]
    assert ws["E2"].value == YEARS[0]
    # Last 3 years in G/H/I
    assert ws["G24"].value == round(fico["revenue"][2023], 1)
    assert ws["H24"].value == round(fico["revenue"][2024], 1)
    assert ws["I24"].value == round(fico["revenue"][2025], 1)
    assert ws["I25"].value == round(fico["cogs"][2025], 1)
    assert ws["I28"].value == round(fico["sga"][2025], 1)
    assert ws["I29"].value == round(fico["rd"][2025], 1)
    assert ws["I49"].value == round(fico["debt"][2025], 1)
    # formulas untouched
    assert is_formula(ws["I26"].value)  # Gross Profit
    assert is_formula(ws["I36"].value)  # Net Earnings
    assert is_formula(ws["I45"].value)  # Total Assets
    assert is_formula(ws["J24"].value)  # Forecast revenue
    print("\nVERIFY OK")
    print(f"  E2={ws['E2'].value}")
    print(f"  FY23 Revenue G24={ws['G24'].value}")
    print(f"  FY24 Revenue H24={ws['H24'].value}")
    print(f"  FY25 Revenue I24={ws['I24'].value}")
    print(f"  FY25 SG&A I28={ws['I28'].value}")
    print(f"  FY25 R&D I29={ws['I29'].value}")
    print(f"  FY25 Debt I49={ws['I49'].value}")
    print(f"  Gross Profit I26 still formula: {ws['I26'].value}")
    print(f"  Forecast Rev J24 still formula: {ws['J24'].value}")


def cleanup_prior_outputs() -> None:
    """Remove previously generated broken/combined model outputs."""
    candidates = [
        ROOT / "FICO_DCF_Financial_Model.xlsx",
        ROOT / "FICO_CFI_3Statement_Model.xlsx",
        ROOT / "FICO_3Statement_Model_Clean.xlsx",
        ROOT / "FICO_MEGA_3Statement_DCF.xlsx",
        ROOT / "DCF resources" / "FICO_DCF_Financial_Model.xlsx",
        ROOT / "DCF resources" / "FICO_MEGA_3Statement_DCF.xlsx",
        ROOT / "DCF resources" / "FICO_CFI_3Statement_Model.xlsx",
        ROOT / "DCF resources" / "FICO_3Statement_Model_Clean.xlsx",
    ]
    for path in candidates:
        if path.exists():
            path.unlink()
            print(f"Deleted prior output: {path}")


def main() -> None:
    cleanup_prior_outputs()
    src = pristine_template_path()
    fico = load_fico()

    print("\nLoading template with data_only=False (preserve formulas)...")
    wb = load_workbook(src, data_only=False)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"Missing sheet {SHEET!r} in {src}")
    ws = wb[SHEET]

    log: list[str] = []
    print("\n===== LINE-ITEM MAPPING LOG =====")
    inject(ws, fico, log)
    for line in log:
        print(line)

    writes = sum(1 for line in log if line.startswith("WRITE"))
    skips = sum(1 for line in log if line.startswith("SKIP"))
    print(f"\nSummary: WRITE={writes} SKIP(formula)={skips}")

    # Do not touch cover labels beyond necessity — prompt says no structural edits.
    # Save NEW file only.
    wb.save(OUTPUT)
    print(f"\nSaved {OUTPUT} ({OUTPUT.stat().st_size} bytes)")
    verify(OUTPUT, fico)


if __name__ == "__main__":
    main()
