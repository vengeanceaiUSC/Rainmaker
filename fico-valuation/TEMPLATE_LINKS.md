# Downloadable valuation templates

Catalog refreshed via **Firecrawl search/scrape** (not Brave Search). Prefer direct `.xlsx` / `.xls` links when available.

## Direct Excel downloads

| Source | Template | URL |
|---|---|---|
| Damodaran | fcffsimpleginzu.xlsx (recommended FCFF) | https://pages.stern.nyu.edu/~adamodar/pc/fcffsimpleginzu.xlsx |
| Damodaran | fcffginzu.xlsx (full kitchen-sink FCFF) | https://pages.stern.nyu.edu/~adamodar/pc/fcffginzu.xlsx |
| Damodaran | fcff2st.xls (two-stage FCFF) | https://pages.stern.nyu.edu/~adamodar/pc/fcff2st.xls |
| Damodaran | fcff3st.xls (three-stage FCFF) | https://pages.stern.nyu.edu/~adamodar/pc/fcff3st.xls |
| Damodaran | firmmult.xls | https://pages.stern.nyu.edu/~adamodar/pc/firmmult.xls |
| CFI / Babson mirror | Bloomberg Practice / DCF Model xlsx | https://www.babson.edu/media/babson/assets/cutler-center/Bloomberg-Practice-Template.xlsx |
| CFI | Compact DCF xlsx | https://corporatefinanceinstitute.com/assets/DCF-Valuation-Compact-Complete.xlsx |
| Smartsheet | Discounted Cash Flow Valuation xlsx | https://www.smartsheet.com/sites/default/files/2020-07/IC-Discounted-Cash-Flow-Valuation-10840.xlsx |
| Wall Street Prep sample | LBO + DCF sample xlsx | https://s3.amazonaws.com/wsp_sample_file/excel-templates/lbo-w-dcf-model-sample.xlsx |
| Exinfm | LBO DCF Model xls | https://exinfm.com/excel%20files/LBO_DCF_Model.xls |

## Landing pages (free download / trial)

| Source | Template | URL | Notes |
|---|---|---|---|
| Damodaran | Spreadsheet hub | https://pages.stern.nyu.edu/~adamodar/New_Home_Page/spreadsh.htm | Full free suite |
| Damodaran | Equity spreadsheets | https://pages.stern.nyu.edu/~adamodar/New_Home_Page/eqspread.htm | 1/2/3-stage models |
| CFI | DCF model template | https://corporatefinanceinstitute.com/resources/financial-modeling/dcf-model-template/ | 3-statement + DCF |
| CFI | Comparable company analysis | https://corporatefinanceinstitute.com/resources/financial-modeling/comparable-company-analysis-template/ | Trading comps |
| CFI | 3-statement case study | https://corporatefinanceinstitute.com/resources/financial-modeling/case-study-3-statement-model-template/ | Integrated model |
| CT Acquisitions | 2026 free DCF | https://ctacquisitions.com/dcf-template-and-excel/ | 6-tab IB/PE style |
| Limelight | DCF calculator | https://www.golimelight.com/templates/dcf-calculator-template | WACC + scenarios |
| Excel Business Resource | Free DCF | https://excelbusinessresource.com/product/free-dcf-model-template-xls/ | FCFF dual TV |
| Keene Advisors | Free DCF | https://www.keeneadvisors.com/free-dcf-model-template | 3yr hist + 5yr proj |
| Wall Street Oasis | DCF template | https://www.wallstreetoasis.com/resources/templates/excel-financial-modeling/discounted-cash-flow-model-template | Community templates |
| Wall Street Oasis | LBO template | https://www.wallstreetoasis.com/resources/templates/excel-financial-modeling/leveraged-buyout-model-template | LBO |
| Wall Street Oasis | 3-statement | https://www.wallstreetoasis.com/resources/templates/excel-financial-modeling/3-statement-financial-model-template | 3-statement |
| Wall Street Prep | LBO course + template | https://www.wallstreetprep.com/knowledge/lbo-model/ | Free lesson + file |
| Wall Street Prep | Trading comps | https://www.wallstreetprep.com/knowledge/financial-modeling-quick-lesson-trading-comps/ | Comps lesson |
| FE Training | LBO model template | https://www.fe.training/free-resources/financial-modeling/lbo-model-template/ | Free resource |
| FE Training | Trading comps | https://www.fe.training/free-resources/investment-banking/trading-comparables-model-template/ | Free resource |
| Macabacus | DCF | https://macabacus.com/excel/templates/discounted-cash-flow | Trial to download |
| Macabacus | Long-form LBO | https://macabacus.com/excel/templates/lbo-model-long | Trial |
| Macabacus | Operating model | https://macabacus.com/excel/templates/operating-model | Trial |
| Cube Software | Free 3-statement | https://www.cubesoftware.com/finance-templates/3-statement-model | Free template |
| eFinancialModels | Free 3-statement | https://www.efinancialmodels.com/downloads/free-three-statement-model-38201/ | Free download |
| Vertex42 | Business valuation | https://www.vertex42.com/ExcelTemplates/business-valuation.html | Simple SME model |
| StableBread | Spreadsheet list | https://stablebread.com/spreadsheets/ | Damodaran mirrors / index |

## SEC / FICO source data

| Source | URL |
|---|---|
| FICO 10-K filings | https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000814547&type=10-K |
| FICO companyfacts XBRL | https://data.sec.gov/api/xbrl/companyfacts/CIK0000814547.json |

This repo also includes a ready-to-edit FICO model: `fico-valuation/FICO_DCF_Valuation_Model.xlsx`.

## Firecrawl usage in this repo

Web research for templates uses Firecrawl (`api.firecrawl.dev`), not Brave Search.

```bash
# put key in gitignored .env.local
export $(grep -v '^#' .env.local | xargs)
python3 fico-valuation/scripts/firecrawl_search.py "free DCF Excel template"
```
