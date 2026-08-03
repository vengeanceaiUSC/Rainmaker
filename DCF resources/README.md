# DCF resources

## Primary output

**`FICO_DCF_Financial_Model.xlsx`** (also at repo root) — CFI 3-statement + DCF templates filled with authentic FICO SEC EDGAR numbers. Native CFI formulas preserved.

### Download
https://github.com/vengeanceaiUSC/Rainmaker/raw/cursor/fico-dcf-valuation-model-d3ac/FICO_DCF_Financial_Model.xlsx

### Sheets
| Sheet | Contents |
|---|---|
| `02_FICO_SEC_Source_Data` | Raw EDGAR line items |
| `03_Three_Statement` | CFI 3-statement — hist FY2021–2025 in cols E–I (Revenue row 24, etc.) |
| `04_DCF` | CFI DCF — green cells linked to 3-statement forecast |

### Rebuild
```bash
python3 build_fico_model.py
```

## Blueprint templates
- `originals/` — pristine CFI downloads (untouched)
- `CFI_3-Statement-Model-Complete.xlsx` / `CFI_DCF-Model.xlsx` — standalone populated copies
