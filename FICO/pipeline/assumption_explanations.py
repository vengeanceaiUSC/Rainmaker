"""Plain-English explanations for every 3-statement assumption row (MODEL4)."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import List, Tuple

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .named_range_map import SHEET_3S

_FORECAST_COLS = ("J", "K", "L", "M", "N")
_COMMENT_AUTHOR = "vengeanceaiUSCMODEL4"

# (row, short_name, what_it_is, how_set, why, source)
ASSUMPTION_EXPLANATIONS: List[Tuple[int, str, str, str, str, str]] = [
    (
        7,
        "Revenue Growth",
        "Year-over-year % increase in revenue for each forecast year.",
        "Yellow POLICY input: 27% / 16% / 13% / 10% / 7%.",
        "Y1 ≈ company FY2026 revenue guidance (~$2.53B / FY25 $1.991B − 1 ≈ 27%). "
        "Then fade toward terminal g=3% (not straight-lined). This is the main "
        "forward-looking judgment; it drives Rev → GP → EBT → Net Earnings → CF.",
        "FICO FY2026 guidance + fade policy",
    ),
    (
        8,
        "COGS % of Revenue",
        "Cost of revenues as a % of sales (gross margin = 1 − this).",
        "Excel equation: I25/I24 (FY25 COGS ÷ FY25 Revenue) held flat.",
        "Locks in latest reported cost structure (~17.8%). Scores mix is high-margin; "
        "holding FY25 is conservative vs further mix shift.",
        "SEC 10-K FY2025 — Cost of revenues / Total revenues",
    ),
    (
        9,
        "SGA % of Revenue",
        "Operating opex (Salaries & Benefits / SG&A) as % of sales.",
        "Equation: (I28 − 10,922)/I24 − 50bps × year. "
        "Strips FY25 restructuring; then −0.5% of sales each year.",
        "Normalize one-time restructuring ($10.9M). Mild efficiency grind reflects "
        "operating leverage as revenue scales; floored at 5%.",
        "SEC 10-K FY2025 — SG&A + restructuring note",
    ),
    (
        10,
        "R&D % of Revenue",
        "Research & development (template label may say Rent and Overhead) as % of sales.",
        "Excel equation: I29/I24 held flat (~9.46%).",
        "FICO reinvests steadily in analytics/software. Flat % grows dollars with revenue.",
        "SEC 10-K FY2025 — Research and development",
    ),
    (
        11,
        "D&A % of PP&E Open",
        "Depreciation & amortization as % of opening net PP&E.",
        "Excel equation: I30/I44 (FY25 DA ÷ FY25 PPE). Forecast DA = OpenPPE × this %.",
        "Ties DA to the asset base the schedule rolls forward (opens at I44).",
        "SEC 10-K FY2025 — D&A and PP&E",
    ),
    (
        12,
        "Interest % of Debt Open",
        "Interest expense as % of opening debt balance.",
        "Excel equation: I101/I98 (FY25 interest ÷ FY25 opening debt in schedule).",
        "Approximates average coupon / cost of debt on the book. Debt held flat "
        "(issuance = 0), so interest stays linked to that stock.",
        "SEC 10-K FY2025 — Interest expense, net + debt footnote",
    ),
    (
        13,
        "Tax Rate (% of EBT)",
        "Effective book tax rate applied to forecast EBT.",
        "Excel equation: I35/I33 (FY25 tax ÷ simplified FY25 EBT in this template).",
        "Uses reported effective rate (~19% on template EBT; ~18.8% on 10-K EBT). "
        "UFCF in DCF also uses unlevered EBIT × t.",
        "SEC 10-K FY2025 — Provision for income taxes / Income before taxes",
    ),
    (
        15,
        "Accounts Receivable (Days)",
        "AR days sales outstanding used to project AR = Rev × days/365.",
        "Excel equation: ROUND(I42/I24×365, 0). I42 is net AR (AR − deferred).",
        "Operating working-capital view: contract liabilities (deferred revenue) "
        "reduce net AR so ΔNWC is not overstated.",
        "SEC 10-K — AR; DeferredRevenueCurrent (XBRL)",
    ),
    (
        16,
        "Inventory (Days)",
        "Inventory days (Inv = COGS × days/365).",
        "Hard zero — FICO is software / scores; inventory is immaterial.",
        "No inventory cycle to fund.",
        "SEC 10-K — Inventory ≈ $0",
    ),
    (
        17,
        "Accounts Payable (Days)",
        "AP days used to project AP = COGS × days/365.",
        "Excel equation: ROUND(I48/I25×365, 0) from FY25.",
        "Payable timing offsets AR in operating NWC.",
        "SEC 10-K FY2025 — Accounts payable",
    ),
    (
        18,
        "CapEx % of Revenue",
        "Capital investment (PPE + capitalized software) as % of sales.",
        "Equation: fade from I68/I24 (FY25 CapEx/Sales ~1.98%) toward 1.0% steady. "
        "Weights 85%→65%→45%→30%→0% on the peak.",
        "FY25 CapEx was elevated by capitalized internal-use software. Fading avoids "
        "locking a peak reinvestment rate forever.",
        "SEC 10-K — PP&E purchases + capitalized software",
    ),
    (
        19,
        "Debt Issuance (Repayment)",
        "New borrowing (+) or repayment (−) in the forecast.",
        "POLICY input = 0 every year.",
        "Hold debt stock flat at FY25 level for the explicit period; interest follows. "
        "Net debt for DCF equity bridge uses the later 10-Q amount separately.",
        "Modeling choice (financing not in FCFF)",
    ),
    (
        20,
        "Equity Issued (Repaid)",
        "New equity issued (+) or buybacks (−) in the forecast BS/CF.",
        "POLICY input = 0 every year.",
        "Buybacks are real historically but are financing — excluded from FCFF. "
        "Share count for $/share is the spot shares outstanding assumption.",
        "Modeling choice (financing not in FCFF)",
    ),
]

def _cell_commentary(name: str, how: str, why: str, source: str) -> str:
    """Text embedded in Excel cell comments (hover on assumption cells)."""
    return (
        f"{name}\n"
        f"HOW: {how}\n"
        f"WHY: {why}\n"
        f"SOURCE: {source}"
    )


def _column_c_note(how: str, why: str, source: str) -> str:
    """Visible note beside the label — must NOT start with '=' (#NAME?)."""
    # Keep readable in-sheet; full detail also lives in the cell comment.
    why_short = why if len(why) <= 160 else why[:157] + "…"
    return f"WHY: {why_short} | SOURCE: {source} | HOW: {how}"

HEADER_FILL = PatternFill("solid", fgColor="C65911")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
SUB_FILL = PatternFill("solid", fgColor="FCE4D6")
SUB_FONT = Font(name="Calibri", bold=True, size=9, color="833C0C")
BODY_FONT = Font(name="Calibri", size=9)
TITLE_FONT = Font(name="Calibri", bold=True, size=12, color="C65911")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def explanation_rows() -> List[dict]:
    rows = []
    for row, name, what, how, why, source in ASSUMPTION_EXPLANATIONS:
        rows.append(
            {
                "row": row,
                "assumption": name,
                "what_it_is": what,
                "how_set_in_model": how,
                "why_this_choice": why,
                "source": source,
            }
        )
    return rows


def write_assumption_explanations(wb) -> None:
    """Put WHY + SOURCE commentary in each assumption cell (Excel comments) + col C + legend."""
    if SHEET_3S not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {SHEET_3S}")
    ws = wb[SHEET_3S]

    ws["B4"] = (
        "vengeanceaiUSCMODEL4: hover any yellow/green assumption cell (J–N) "
        "for WHY + SOURCE commentary; col C also shows WHY | SOURCE | HOW"
    )
    ws["B4"].font = Font(name="Calibri", bold=True, color="1F4E79")
    ws["B4"].fill = PatternFill("solid", fgColor="D6EAF8")

    for row, name, what, how, why, source in ASSUMPTION_EXPLANATIONS:
        # Visible note next to the assumption label
        ws[f"C{row}"] = _column_c_note(how, why, source)
        ws[f"C{row}"].font = Font(name="Calibri", italic=True, size=8, color="595959")
        ws[f"C{row}"].alignment = WRAP

        # Commentary INSIDE each forecast assumption cell (red-triangle Excel comment)
        text = _cell_commentary(name, how, why, source)
        for col in _FORECAST_COLS:
            cell = ws[f"{col}{row}"]
            comment = Comment(text, _COMMENT_AUTHOR)
            comment.width = 320
            comment.height = 140
            cell.comment = comment

        # Also comment the row label so auditors see it without opening J–N
        label_cell = ws[f"B{row}"]
        label_comment = Comment(text, _COMMENT_AUTHOR)
        label_comment.width = 320
        label_comment.height = 140
        label_cell.comment = label_comment

        ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 15, 36)

    # Legend block to the right of the assumptions (starting col P)
    ws["P4"] = "ASSUMPTIONS EXPLAINED — every driver (also in each cell comment)"
    ws["P4"].font = TITLE_FONT
    ws.merge_cells("P4:U4")

    ws["P5"] = (
        "Yellow = policy judgment. Green = Excel equations linked to FY25 (col I). "
        "Hover J–N (or B label) for WHY + SOURCE. Column C is text only (not a formula)."
    )
    ws["P5"].font = Font(name="Calibri", italic=True, size=9, color="595959")
    ws.merge_cells("P5:U5")

    headers = [
        ("P6", "Row"),
        ("Q6", "Assumption"),
        ("R6", "What it is"),
        ("S6", "How set in model"),
        ("T6", "Why this choice"),
        ("U6", "Source"),
    ]
    for coord, label in headers:
        cell = ws[coord]
        cell.value = label
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN
        cell.alignment = Alignment(vertical="center")

    # Place each explanation on the same sheet row as the assumption (side-by-side)
    for row, name, what, how, why, source in ASSUMPTION_EXPLANATIONS:
        values = [row, name, what, how, why, source]
        for col_idx, val in enumerate(values, start=16):  # P=16
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.font = BODY_FONT if col_idx > 17 else SUB_FONT
            cell.alignment = WRAP
            cell.border = THIN
            if col_idx <= 17:
                cell.fill = SUB_FILL
            # Commentary also on the legend "Why" / "Source" cells
            if col_idx in (20, 21):  # T, U
                cell.comment = Comment(
                    _cell_commentary(name, how, why, source),
                    _COMMENT_AUTHOR,
                )

    # Column widths for readability
    ws.column_dimensions["C"].width = 72
    ws.column_dimensions["P"].width = 6
    ws.column_dimensions["Q"].width = 22
    ws.column_dimensions["R"].width = 36
    ws.column_dimensions["S"].width = 42
    ws.column_dimensions["T"].width = 48
    ws.column_dimensions["U"].width = 36


def export_assumption_explanations_csv(out_dir: Path, *, ticker: str = "FICO") -> Path:
    """Write MODEL4_*_ASSUMPTIONS_EXPLAINED.csv."""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"MODEL4_{ticker}_ASSUMPTIONS_EXPLAINED.csv"
    rows = explanation_rows()
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "row",
                "assumption",
                "what_it_is",
                "how_set_in_model",
                "why_this_choice",
                "source",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)
    return path
