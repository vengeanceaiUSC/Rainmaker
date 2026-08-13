# LBO — initial templates (with IRR + MoM)

Starter LBO Excel templates you can open and edit. **No FICO / company-specific injection.**

## Click to download (raw)

| File | Raw download | Notes |
|---|---|---|
| **Primary — complex LBO + IRR / cash-on-cash (MoM)** | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO_Complex_Template_IRR_MoM.xlsx | Wall Street Prep sample (BMC demo data) |
| Same under `LBO/output/` | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/output/LBO_Complex_Template_IRR_MoM.xlsx | Identical file |
| WSP named copy | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/WSP_LBO_with_DCF_Sample.xlsx | Same bytes as primary |
| Exinfm LBO + DCF (has `IRR_Returns` tab) | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/exinfm_LBO_DCF_Model.xlsx | Older multi-tab model |
| LBO Model ABC (no IRR sheet) | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/LBO_Model_ABC.xlsx | Complex ops/debt; exit returns incomplete |
| CFI Leveraged Finance | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/CFI_Leveraged_Finance_Template.xlsx | Simple IRR by leverage |
| CFI Debt Capacity | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/CFI_Debt_Capacity_Model_Template.xlsx | Credit metrics |

Use **raw** links. Open in Excel / Google Sheets to edit.

## What’s in the primary template

Original source: Wall Street Prep `lbo-w-dcf-model-sample.xlsx`  
https://s3.amazonaws.com/wsp_sample_file/excel-templates/lbo-w-dcf-model-sample.xlsx

Sheets: `Coversheet` · `LBO` · `DCF` · `Shares` · `52wkHL`

On the **`LBO`** sheet (scroll to ~row 235):

| Block | Contents |
|---|---|
| Sources & Uses | Debt tranches + sponsor equity |
| Operating model | Revenue → EBITDA → FCF |
| Debt schedule | Interest / amort / paydown |
| **RETURNS** | **Cash-on-cash (MoM)** + **IRR** by instrument and for sponsor equity |
| Sensitivity | IRR vs exit multiple, leverage, offer price, equity kickers |

Demo company in the file is **BMC** (sample only — replace inputs to model another target).

Educational / research use only — not investment advice.
