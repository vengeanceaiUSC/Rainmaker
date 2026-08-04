#!/usr/bin/env python3
"""Populate pristine CFI templates with authentic FICO numbers and build MEGA linked workbook."""

from __future__ import annotations

import shutil
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
ORIG = ROOT / "originals"
BLUE = Font(name="Calibri", color="0000FF")
BLACK = Font(name="Calibri", color="000000")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
LINK_FILL = PatternFill("solid", fgColor="E2EFDA")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True)
TITLE = Font(name="Calibri", size=16, bold=True, color="1F4E79")
NOTE = Font(name="Calibri", size=10, italic=True, color="C00000")

# ---------------------------------------------------------------------------
# Authentic FICO figures ($ thousands) from SEC 10-K / companyfacts / 10-Q
# ---------------------------------------------------------------------------
YEARS = [2021, 2022, 2023, 2024, 2025]  # hist cols E-I
# Income statement
REV = {2021: 1_316_536, 2022: 1_377_270, 2023: 1_513_557, 2024: 1_717_526, 2025: 1_990_869}
COGS = {2021: 332_462, 2022: 302_174, 2023: 311_053, 2024: 348_206, 2025: 353_722}
RD = {2021: 171_231, 2022: 146_758, 2023: 159_950, 2024: 171_940, 2025: 188_347}
SGA = {2021: 396_281, 2022: 383_863, 2023: 400_565, 2024: 462_834, 2025: 513_028}
# Small IS items folded into overhead for template fit (amort/restructuring/gain)
OTHER_OPEX = {2021: 0, 2022: 0, 2023: -841, 2024: 917, 2025: 10_922}  # amort - gain + restructuring
DA = {2021: 25_592, 2022: 20_465, 2023: 14_638, 2024: 13_827, 2025: 14_952}
INTEREST = {2021: 40_092, 2022: 68_967, 2023: 95_546, 2024: 105_638, 2025: 133_647}
OTHER_INC = {2021: 0, 2022: 0, 2023: 6_340, 2024: 14_034, 2025: 11_392}
TAX = {2021: 81_058, 2022: 97_768, 2023: 124_249, 2024: 129_214, 2025: 150_649}
NI = {2021: 392_084, 2022: 373_541, 2023: 429_375, 2024: 512_811, 2025: 651_946}
OPINC = {2021: 505_489, 2022: 542_414, 2023: 642_830, 2024: 733_629, 2025: 924_850}

# Map into template expense lines while preserving reported operating income:
# GP - Salaries - Rent - DA = OPINC  => Salaries+Rent = GP - OPINC - DA
# Put SG&A into Salaries, residual (mostly R&D +/- other) into Rent.
SALARIES = {}
RENT = {}
for y in YEARS:
    gp = REV[y] - COGS[y]
    opex_ex_da = gp - OPINC[y] - DA[y]
    SALARIES[y] = SGA[y]
    RENT[y] = opex_ex_da - SGA[y]  # residual so EBIT/OpInc identity holds

# Balance sheet
CASH = {2020: 157_394, 2021: 195_354, 2022: 133_202, 2023: 136_778, 2024: 150_667, 2025: 134_136}
AR = {2020: 334_180, 2021: 312_107, 2022: 322_410, 2023: 387_947, 2024: 426_642, 2025: 529_148}
PPE = {2020: 46_419, 2021: 27_913, 2022: 17_580, 2023: 10_966, 2024: 38_465, 2025: 67_713}
AP = {2020: 23_033, 2021: 20_749, 2022: 17_273, 2023: 19_009, 2024: 22_473, 2025: 32_315}
DEBT = {2020: 739_435, 2021: 1_009_018, 2022: 1_823_669, 2023: 1_811_658, 2024: 2_209_021, 2025: 3_055_691}
CAPEX_PPE = {2021: 7_569, 2022: 6_029, 2023: 4_237, 2024: 8_884, 2025: 8_922}
# Total reinvestment used in DCF (PP&E + capitalized software)
CAPEX_TOTAL = {2021: 7_569, 2022: 6_029, 2023: 4_237, 2024: 25_551, 2025: 39_407}

# Market / valuation
PRICE = 1046.23
SHARES_K = 21_597.635
DEBT_10Q = 5_582_389
CASH_10Q = 248_444 + 56_093
WACC = 0.096
G = 0.03
EV_EBITDA = 25.0
TAX_RATE = 0.188

# Forecast FY26-30
GROWTHS = [0.14, 0.12, 0.10, 0.09, 0.08]
COGS_PCT = [0.178, 0.178, 0.175, 0.175, 0.170]
DA_PCT_PPE = 0.25
INT_PCT = 0.055
AR_DAYS = round(AR[2025] / REV[2025] * 365)
INV_DAYS = 0
AP_DAYS = round(AP[2025] / COGS[2025] * 365)


def style_input(cell):
    cell.font = BLUE
    cell.fill = INPUT_FILL


def style_link(cell):
    cell.font = BLACK
    cell.fill = LINK_FILL


def populate_three_statement(path_in: Path, path_out: Path) -> None:
    wb = load_workbook(path_in)
    cover = wb["Cover Page"]
    cover["C12"] = "FICO — 3 Statement Model"
    cover["C20"] = "Populated with Fair Isaac Corp (FICO) SEC 10-K numbers. Units: $ thousands. Educational only."
    cover["C21"] = "Yellow/blue cells = inputs. History FY2021–FY2025; forecast FY2026–FY2030."

    ws = wb["3 Statement Model"]
    ws["B2"] = "FICO FINANCIAL STATEMENTS"
    ws["E2"] = 2021  # drives all year headers

    # Clarify labels without breaking structure
    ws["B9"] = "SG&A ($000's)"
    ws["B10"] = "R&D + other opex residual ($000's)"
    ws["B28"] = "SG&A"
    ws["B29"] = "R&D + other opex residual"

    cols = ["E", "F", "G", "H", "I"]

    # --- Forecast assumptions J-N ---
    sal0, rent0 = SALARIES[2025], max(RENT[2025], 0)
    for i, col in enumerate(["J", "K", "L", "M", "N"]):
        g = GROWTHS[i]
        vals = {
            7: g,
            8: COGS_PCT[i],
            9: round(sal0 * (1 + g), 1),
            10: round(rent0 * (1 + g), 1),
            11: DA_PCT_PPE,
            12: INT_PCT,
            13: TAX_RATE,
            15: AR_DAYS,
            16: INV_DAYS,
            17: AP_DAYS,
            18: round(REV[2025] * (1 + g) * 0.02 if i == 0 else None, 1)
            if False
            else None,
            19: 0,
            20: 0,
        }
        # Capex path ~2% of projected revenue
        rev_proj = REV[2025]
        for gg in GROWTHS[: i + 1]:
            rev_proj *= 1 + gg
        vals[18] = round(rev_proj * 0.02, 1)

        for row, val in vals.items():
            cell = ws[f"{col}{row}"]
            cell.value = val
            style_input(cell)

    # --- Historical IS ---
    for col, y in zip(cols, YEARS):
        for row, val in {
            24: REV[y],
            25: COGS[y],
            28: round(SALARIES[y], 1),
            29: round(RENT[y], 1),
            30: DA[y],
            31: INTEREST[y],
            35: TAX[y],
        }.items():
            cell = ws[f"{col}{row}"]
            cell.value = val
            style_input(cell)

    # --- Historical BS (simplified template assets/liabilities) + equity plug ---
    for col, y in zip(cols, YEARS):
        assets = CASH[y] + AR[y] + 0 + PPE[y]
        liab = AP[y] + DEBT[y]
        equity = assets - liab
        for row, val in {
            41: CASH[y],
            42: AR[y],
            43: 0,
            44: PPE[y],
            48: AP[y],
            49: DEBT[y],
            52: equity,  # equity capital plug
            53: 0,  # RE starts at 0; forecast accumulates NI from I53
        }.items():
            cell = ws[f"{col}{row}"]
            cell.value = val
            style_input(cell)

    # --- Cash flow historical hardcodes ---
    debt_issue = {y: DEBT[y] - DEBT[y - 1] for y in YEARS}
    nwc = {y: AR[y] - AP[y] for y in [2020] + YEARS}
    dnwc = {y: nwc[y] - nwc[y - 1] for y in YEARS}

    for col, y in zip(cols, YEARS):
        ws[f"{col}64"] = dnwc[y]
        style_input(ws[f"{col}64"])
        ws[f"{col}68"] = CAPEX_TOTAL[y]  # total reinvestment for CF investing
        style_input(ws[f"{col}68"])
        ws[f"{col}72"] = debt_issue[y]
        style_input(ws[f"{col}72"])

    # Opening cash chain
    ws["E77"] = CASH[2020]
    style_input(ws["E77"])
    ws["F77"] = "=E78"
    ws["G77"] = "=F78"
    ws["H77"] = "=G78"
    ws["I77"] = "=H78"

    # Equity issuance plugs so closing cash matches BS cash each historical year.
    # closing = opening + CFO - CFI + CFF; CFF = debt_issue + equity_issue
    # CFO ≈ NI + DA - dnwc (template formulas); we set equity_issue after estimating.
    # Use direct plug: target delta cash = CASH[y]-CASH[y-1]
    # equity_issue = target_delta - (NI + DA - dnwc - capex + debt_issue)
    # But NI in model may differ slightly; compute plug from known cash bridge using reported NI.
    for col, y in zip(cols, YEARS):
        cfo_proxy = NI[y] + DA[y] - dnwc[y]
        cfi = CAPEX_TOTAL[y]
        target_delta = CASH[y] - CASH[y - 1]
        equity_issue = target_delta - (cfo_proxy - cfi + debt_issue[y])
        ws[f"{col}73"] = round(equity_issue, 1)
        style_input(ws[f"{col}73"])

    # Working capital schedule hist values
    for col, y in zip(cols, YEARS):
        ws[f"{col}85"] = AR[y]
        ws[f"{col}86"] = 0
        ws[f"{col}87"] = AP[y]
        for r in (85, 86, 87):
            style_input(ws[f"{col}{r}"])

    # PPE schedule: choose capex/da so closing PPE matches BS
    for col, y in zip(cols, YEARS):
        open_ppe = PPE[y - 1]
        close_ppe = PPE[y]
        da = DA[y]
        capex_sched = close_ppe - open_ppe + da
        ws[f"{col}92"] = open_ppe
        ws[f"{col}93"] = capex_sched
        ws[f"{col}94"] = da
        for r in (92, 93, 94):
            style_input(ws[f"{col}{r}"])

    # Debt schedule
    for col, y in zip(cols, YEARS):
        ws[f"{col}98"] = DEBT[y - 1]
        ws[f"{col}99"] = debt_issue[y]
        ws[f"{col}101"] = INTEREST[y]
        for r in (98, 99, 101):
            style_input(ws[f"{col}{r}"])

    # Notes
    ws["B131"] = (
        "FICO mapping: Revenue/COGS/Interest/Tax from 10-K; SG&A→Salaries; R&D residual→Rent so OpInc identity holds."
    )
    ws["B132"] = (
        "Template BS is simplified (Cash/AR/Inv/PPE vs AP/Debt). Equity is a plug. Yellow/blue = inputs."
    )
    ws["B133"] = (
        "Sources: SEC EDGAR FY2025 10-K / companyfacts; forecast drivers are base-case assumptions."
    )
    for r in (131, 132, 133):
        ws[f"B{r}"].font = NOTE

    wb.save(path_out)
    print("Wrote", path_out)


def populate_dcf(path_in: Path, path_out: Path) -> None:
    wb = load_workbook(path_in)
    cover = wb["Cover Page"]
    cover["C12"] = "FICO — DCF Model"
    cover["C20"] = "Populated with Fair Isaac Corp (FICO) projections from SEC fundamentals. Units: $ thousands."
    cover["C21"] = "Yellow/blue = inputs. Educational only — not investment advice."

    ws = wb["DCF Model"]
    ws["B2"] = "FICO DCF Model"

    # Assumptions
    ws["D5"] = TAX_RATE
    ws["D6"] = WACC
    ws["D7"] = G
    ws["D8"] = EV_EBITDA
    ws["D9"] = datetime(2025, 9, 30)
    ws["D10"] = datetime(2025, 9, 30)
    ws["D11"] = PRICE
    ws["D12"] = SHARES_K
    ws["D13"] = DEBT_10Q
    ws["D14"] = CASH_10Q
    ws["D15"] = CAPEX_TOTAL[2025]
    for addr in ["D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", "D13", "D14", "D15"]:
        style_input(ws[addr])

    # FICO FYE September
    for col in ["E", "F", "G", "H", "I"]:
        ws[f"{col}18"] = f"=DATE(YEAR($D$10)+{col}19,9,30)"

    # Build EBIT / DA / NWC from same forecast as 3-statement
    prev = REV[2025]
    # FY25 salaries/rent for opex ratio
    opex0 = SALARIES[2025] + max(RENT[2025], 0)
    for i, col in enumerate(["E", "F", "G", "H", "I"]):
        g = GROWTHS[i]
        rev = prev * (1 + g)
        cogs = rev * COGS_PCT[i]
        # Keep FY25 opex margin ex-DA roughly, grown
        opex = opex0 * (1 + g)
        da = rev * 0.008
        ebit = rev - cogs - opex - da
        dnwc = max(0, rev - prev) * 0.05
        ws[f"{col}21"] = round(ebit, 1)
        ws[f"{col}23"] = round(da, 1)
        ws[f"{col}25"] = round(dnwc, 1)
        for r in (21, 23, 25):
            style_input(ws[f"{col}{r}"])
        prev = rev

    ws["B50"] = "FICO notes"
    ws["B51"] = "Debt/Cash from 10-Q 2026-06-30. Shares/price from market. Tax 18.8% FY25 ETR."
    ws["B52"] = "EBIT path uses same growth/COGS/opex logic as the FICO 3-statement forecast."
    for r in (50, 51, 52):
        ws[f"B{r}"].font = NOTE

    wb.save(path_out)
    print("Wrote", path_out)


def build_mega(three_path: Path, dcf_path: Path, out_path: Path, banner: Path | None) -> None:
    # Start from populated 3-statement
    wb = load_workbook(three_path)
    wb["Cover Page"].title = "00_Cover"
    wb["3 Statement Model"].title = "01_Three_Statement"

    # Copy DCF sheet
    src_wb = load_workbook(dcf_path)
    src = src_wb["DCF Model"]
    dst = wb.create_sheet("02_DCF")
    for row in src.iter_rows():
        for cell in row:
            new = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new.font = copy(cell.font)
                new.fill = copy(cell.fill)
                new.border = copy(cell.border)
                new.number_format = cell.number_format
                new.alignment = copy(cell.alignment)
    for letter, dim in src.column_dimensions.items():
        dst.column_dimensions[letter].width = dim.width
    for idx, dim in src.row_dimensions.items():
        dst.row_dimensions[idx].height = dim.height
    for merged in src.merged_cells.ranges:
        dst.merge_cells(str(merged))

    # Link DCF forecast to 3-statement
    dst["B2"] = "FICO MEGA DCF (linked to 01_Three_Statement)"
    dst["D5"] = "='01_Three_Statement'!J13"
    style_link(dst["D5"])
    # Keep market debt/cash as 10-Q overrides (more current than FY25 BS)
    # but document; leave D13/D14 as hardcoded 10Q values already in populated DCF
    dst["D15"] = "='01_Three_Statement'!I68"
    style_link(dst["D15"])

    for dcol, scol in zip(["E", "F", "G", "H", "I"], ["J", "K", "L", "M", "N"]):
        dst[f"{dcol}21"] = (
            f"='01_Three_Statement'!{scol}26-'01_Three_Statement'!{scol}28-"
            f"'01_Three_Statement'!{scol}29-'01_Three_Statement'!{scol}30"
        )
        dst[f"{dcol}23"] = f"='01_Three_Statement'!{scol}30"
        dst[f"{dcol}24"] = f"='01_Three_Statement'!{scol}18"
        dst[f"{dcol}25"] = f"='01_Three_Statement'!{scol}89"
        for r in (21, 23, 24, 25):
            style_link(dst[f"{dcol}{r}"])

    dst["B21"] = "EBIT (from 3-stmt)"
    dst["B23"] = "Plus: D&A (from 3-stmt)"
    dst["B24"] = "Less: Capex (from 3-stmt)"
    dst["B25"] = "Less: Changes in NWC (from 3-stmt)"
    dst["B55"] = "GREEN cells are linked from 01_Three_Statement. Edit yellow drivers there; DCF updates."
    dst["B55"].font = NOTE
    dst["B56"] = "YELLOW on this sheet: WACC D6, g D7, EV/EBITDA D8, Price D11, Shares D12, Debt D13, Cash D14 (10-Q)."
    dst["B56"].font = NOTE

    # Instructions
    ins = wb.create_sheet("00_Instructions", 0)
    ins["A1"] = "FICO MEGA 3-Statement + DCF"
    ins["A1"].font = TITLE
    lines = [
        "",
        "This file uses the original CFI template layouts, filled with authentic FICO SEC numbers.",
        "",
        "Sheets",
        "01_Three_Statement — CFI 3-statement template with FICO FY2021–FY2025 history + FY2026–FY2030 forecast",
        "02_DCF — CFI DCF template; green forecast cells pull from 01_Three_Statement",
        "04_FICO_Source_Data — raw 10-K line items used for the inputs",
        "",
        "How to use",
        "1) Confirm historical yellow inputs on 01_Three_Statement match 04_FICO_Source_Data",
        "2) Edit forecast assumptions J7:N20 on 01_Three_Statement",
        "3) Read intrinsic value on 02_DCF",
        "",
        "Color key: Yellow/blue = inputs | Green = linked from 3-statement | Black = formulas",
    ]
    for i, t in enumerate(lines, start=2):
        ins[f"A{i}"] = t
    ins.column_dimensions["A"].width = 110

    if banner and banner.exists():
        try:
            img = XLImage(str(banner))
            img.width = 900
            img.height = 506
            ins.add_image(img, "A16")
            ins["A15"] = "PROOF BANNER (edit marker)"
            ins["A15"].font = Font(name="Impact", size=18, color="FFFFFF")
            ins["A15"].fill = PatternFill("solid", fgColor="111111")
        except Exception as e:
            print("banner skip", e)

    # Source data sheet with authentic numbers
    src = wb.create_sheet("04_FICO_Source_Data")
    src["A1"] = "Authentic FICO source inputs ($ thousands)"
    src["A1"].font = TITLE
    src["A2"] = "From SEC EDGAR FY2025 10-K / companyfacts / 10-Q 2026-06-30"
    headers = ["Line item", 2021, 2022, 2023, 2024, 2025, "Source"]
    for c, h in enumerate(headers, 1):
        cell = src.cell(4, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    rows = [
        ("Revenue", REV, "10-K / XBRL"),
        ("Cost of revenues", COGS, "10-K / XBRL"),
        ("R&D", RD, "10-K / XBRL"),
        ("SG&A", SGA, "10-K / XBRL"),
        ("Operating income", OPINC, "10-K / XBRL"),
        ("Interest expense, net", INTEREST, "10-K / XBRL"),
        ("Tax provision", TAX, "10-K / XBRL"),
        ("Net income", NI, "10-K / XBRL"),
        ("D&A (CF)", DA, "10-K CF"),
        ("Cash", CASH, "10-K BS"),
        ("Accounts receivable", AR, "10-K BS"),
        ("PP&E net", PPE, "10-K BS"),
        ("Accounts payable", AP, "10-K BS"),
        ("Total debt", DEBT, "10-K debt note / BS"),
        ("Capex + capitalized software", CAPEX_TOTAL, "10-K CF"),
    ]
    for i, (name, series, source) in enumerate(rows, start=5):
        src.cell(i, 1, name)
        for j, y in enumerate(YEARS, start=2):
            src.cell(i, j, series.get(y))
        src.cell(i, 7, source)

    src["A22"] = "Market bridge used in DCF"
    src["A23"] = "Share price"
    src["B23"] = PRICE
    src["A24"] = "Shares outstanding (000s)"
    src["B24"] = SHARES_K
    src["A25"] = "Debt (10-Q 2026-06-30)"
    src["B25"] = DEBT_10Q
    src["A26"] = "Cash + marketable securities (10-Q)"
    src["B26"] = CASH_10Q

    for col in range(1, 8):
        src.column_dimensions[get_column_letter(col)].width = 18
    src.column_dimensions["A"].width = 36

    # Link map
    link = wb.create_sheet("03_Link_Map")
    link["A1"] = "DCF ← 3-Statement links"
    link["A1"].font = TITLE
    for i, h in enumerate(["DCF cell", "Meaning", "Source"], 1):
        cell = link.cell(3, i, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for r, row in enumerate(
        [
            ("02_DCF!D5", "Tax rate", "01_Three_Statement!J13"),
            ("02_DCF!E21:I21", "EBIT FY26-30", "3-stmt GP - SG&A - R&D residual - D&A (J-N)"),
            ("02_DCF!E23:I23", "D&A", "01_Three_Statement!J30:N30"),
            ("02_DCF!E24:I24", "Capex", "01_Three_Statement!J18:N18"),
            ("02_DCF!E25:I25", "ΔNWC", "01_Three_Statement!J89:N89"),
            ("02_DCF!D13/D14", "Debt / Cash", "10-Q bridge (manual yellow inputs)"),
        ],
        start=4,
    ):
        for c, v in enumerate(row, 1):
            link.cell(r, c, v)
    for col in ["A", "B", "C"]:
        link.column_dimensions[col].width = 48

    # Cover tweak
    cover = wb["00_Cover"]
    cover["C12"] = "FICO MEGA — 3-Statement + DCF"
    cover["C15"] = "01_Three_Statement"
    cover["C16"] = "02_DCF"
    cover["C20"] = "Original CFI templates filled with authentic FICO SEC numbers; DCF linked to 3-statement forecast."

    order = ["00_Instructions", "00_Cover", "01_Three_Statement", "02_DCF", "03_Link_Map", "04_FICO_Source_Data"]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.save(out_path)
    print("Wrote", out_path, "sheets", wb.sheetnames)


def verify(three_path: Path, dcf_path: Path, mega_path: Path) -> None:
    wb = load_workbook(three_path)
    ws = wb["3 Statement Model"]
    assert ws["E2"].value == 2021
    assert ws["I24"].value == REV[2025], ws["I24"].value
    assert ws["I25"].value == COGS[2025]
    assert ws["I28"].value == round(SALARIES[2025], 1)
    assert ws["I41"].value == CASH[2025]
    assert ws["I49"].value == DEBT[2025]
    assert ws["J7"].value == GROWTHS[0]
    print("3-statement OK: FY25 rev", ws["I24"].value, "debt", ws["I49"].value)

    wb = load_workbook(dcf_path)
    ws = wb["DCF Model"]
    assert ws["D11"].value == PRICE
    assert ws["D13"].value == DEBT_10Q
    assert ws["E21"].value is not None
    print("DCF OK: price", ws["D11"].value, "EBIT Y1", ws["E21"].value)

    wb = load_workbook(mega_path)
    assert "01_Three_Statement" in wb.sheetnames
    assert "02_DCF" in wb.sheetnames
    assert "04_FICO_Source_Data" in wb.sheetnames
    ws = wb["01_Three_Statement"]
    assert ws["I24"].value == REV[2025]
    dcf = wb["02_DCF"]
    assert isinstance(dcf["E21"].value, str) and "01_Three_Statement" in dcf["E21"].value
    print("MEGA OK: linked EBIT", dcf["E21"].value)
    print("MEGA sheets:", wb.sheetnames)


def main() -> None:
    assert (ORIG / "CFI_3-Statement-Model-Complete.xlsx").exists()
    assert (ORIG / "CFI_DCF-Model.xlsx").exists()

    three_out = ROOT / "CFI_3-Statement-Model-Complete.xlsx"
    dcf_out = ROOT / "CFI_DCF-Model.xlsx"
    mega_out = ROOT / "FICO_MEGA_3Statement_DCF.xlsx"
    banner = ROOT / "assets" / "fico_eminem_banner.png"

    populate_three_statement(ORIG / "CFI_3-Statement-Model-Complete.xlsx", three_out)
    populate_dcf(ORIG / "CFI_DCF-Model.xlsx", dcf_out)
    build_mega(three_out, dcf_out, mega_out, banner if banner.exists() else None)
    verify(three_out, dcf_out, mega_out)

    (ROOT / "README.md").write_text(
        """# DCF resources

CFI original templates filled with **authentic Fair Isaac (FICO)** SEC numbers.

## Files

| File | Description |
|---|---|
| `FICO_MEGA_3Statement_DCF.xlsx` | **Use this** — merged CFI 3-statement + DCF, DCF linked to 3-statement |
| `CFI_3-Statement-Model-Complete.xlsx` | Standalone CFI 3-statement with FICO FY21–25 + forecast |
| `CFI_DCF-Model.xlsx` | Standalone CFI DCF with FICO assumptions/projections |
| `originals/` | Untouched pristine CFI downloads |
| `04_FICO_Source_Data` sheet (inside MEGA) | Raw 10-K line items |

## Where the FICO numbers live

On `01_Three_Statement` / standalone 3-statement:
- Years start at **2021** (cell E2)
- Historical inputs in columns **E–I** (Revenue row 24, COGS 25, SG&A 28, R&D residual 29, D&A 30, Interest 31, Tax 35, BS 41–53, CF/schedules below)
- Forecast drivers in **J–N** rows 7–20

On `02_DCF`:
- Assumptions D5–D15 (tax/WACC/g/multiple/price/shares/debt/cash/capex)
- Green EBIT/D&A/Capex/NWC cells link from the 3-statement in the MEGA file

## Download

https://github.com/vengeanceaiUSC/Rainmaker/raw/cursor/fico-dcf-valuation-model-d3ac/DCF%20resources/FICO_MEGA_3Statement_DCF.xlsx
""",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
