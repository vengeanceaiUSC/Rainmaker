# LBO — initial templates (with IRR + MoM)

Starter LBO Excel templates you can open and edit. **No FICO / company-specific injection.**

## Click to download (raw)

| File | Raw download | Notes |
|---|---|---|
| **Primary — LBO (IRR/MoM) + ZoomInfo S&M→AI tabs** | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO_Complex_Template_IRR_MoM.xlsx | Original WSP LBO sheets kept; ZI FY2025 10-K tabs added |
| Same under `LBO/output/` | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/output/LBO_Complex_Template_IRR_MoM.xlsx | Identical file |
| Pristine WSP original (no ZI tabs) | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/WSP_LBO_ORIGINAL_IRR_MoM.xlsx | Untouched original template |
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


## ZoomInfo S&M → AI (inside the primary LBO file)

The primary workbook above **keeps the original WSP LBO/DCF/IRR/MoM sheets** and adds:

- `ZI_01_10K_Source` … `ZI_05_Sensitivity` — FY2025 10-K S&M isolation + SDR→AI schedule (≥500 bps)

Standalone ZI-only copy (same ZI tabs): https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/ZI_SM_AI_EBITDA_Margin_Model.xlsx

Pristine WSP-only original (no ZI tabs): https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/lbo-44dc/LBO/templates/WSP_LBO_with_DCF_Sample.xlsx

Rebuild ZI tabs: `python3 LBO/zi_sm_ai/build_zi_sm_ai_model.py`


## vengeanceaiUSC-LBOMODEL1 (ZoomInfo numbers in LBO)

**Download:** https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx

LBO sheet inputs replaced with ZoomInfo (GTM) FY2025 / Q2’26 public figures; IRR & cash-on-cash refreshed. Notes: `LBO/output/LBOMODEL1_NOTES.md`.


### LBOMODEL1 assumptions (every tab) — downloadable
| Format | Raw URL |
|---|---|
| **PDF** | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.pdf |
| **Markdown** | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.md |
| **TXT** | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.txt |
| **CSV** | https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_ASSUMPTIONS_LIST.csv |

Covers **Coversheet, LBO, DCF, Shares, 52wkHL, ZI_01…ZI_05** with What / How / Why / Source per assumption (same style as FICO DCF model lists).
