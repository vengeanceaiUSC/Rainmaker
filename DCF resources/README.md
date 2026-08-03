# DCF resources

CFI templates populated with **Fair Isaac Corp (FICO)**, plus an integrated mega model.

## Files

| File | What it is |
|---|---|
| **`FICO_MEGA_3Statement_DCF.xlsx`** | **Merged 3-statement + DCF (DCF linked to 3-statement forecast)** |
| `CFI_DCF-Model.xlsx` | Standalone CFI DCF with FICO inputs |
| `CFI_3-Statement-Model-Complete.xlsx` | Standalone CFI 3-statement with FICO inputs |

## MEGA model usage

1. Open `FICO_MEGA_3Statement_DCF.xlsx`
2. Edit yellow inputs on `01_Three_Statement` (growth, opex, tax, capex, etc.)
3. Read valuation on `02_DCF` — green cells flow from the 3-statement
4. See `03_Link_Map` for the formula map

## Units / sources

- Figures in **$ thousands** (except share price)
- Historicals: SEC EDGAR FY2025 10-K / companyfacts
- Yellow/blue = inputs; green = linked formulas

## Caveats

The CFI 3-statement template only has Cash / AR / Inventory / PP&E assets and AP / Debt liabilities, so FICO goodwill, deferred revenue, buybacks, and SBC are not fully modeled.
