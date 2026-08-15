#!/usr/bin/env python3
"""
Add Baseline Yes/No on Assumptions_Drivers (col M — no column shift).
Sync 00_Assumptions_List numbers to current Drivers.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

REPO = Path(__file__).resolve().parents[2]
XLSX = REPO / "vengeanceaiUSC_LBOMODEL2.xlsx"
XLSX2 = REPO / "LBO" / "output" / "vengeanceaiUSC_LBOMODEL2.xlsx"

HDR = PatternFill("solid", fgColor="1F4E79")
HDR_F = Font(name="Calibri", color="FFFFFF", bold=True, size=9)
YES_F = PatternFill("solid", fgColor="C6EFCE")
NO_F = PatternFill("solid", fgColor="FCE4D6")
BOLD = Font(name="Calibri", bold=True, size=10)
BLACK = Font(name="Calibri", size=10)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(wrap_text=True, vertical="top")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
LINK = Font(name="Calibri", color="0563C1", underline="single", size=10)

# YES = derived from company past performance / FY24 numbers
# NO  = external source, market rate, MODEL CONST, or AI thesis lever
BASELINE = {
    5: "NO",
    6: "NO",
    7: "NO",
    8: "NO",
    9: "NO",
    10: "NO",
    11: "NO",
    12: "NO",
    13: "NO",
    14: "NO",
    15: "NO",
    16: "NO",
    17: "NO",
    18: "NO",
    19: "NO",
    20: "NO",
    21: "YES",  # ZI FY24 S&M-derived
    22: "YES",  # carve from S&M baseline
    23: "YES",  # carve from S&M baseline
    24: "NO",
    25: "NO",
    27: "NO",
    28: "NO",
}


def patch_drivers(ws):
    # Col A — leftmost, next to Driver — Baseline Yes/No
    ws["A4"] = "Baseline Yes/No"
    ws["A4"].fill = HDR
    ws["A4"].font = HDR_F
    ws["A4"].alignment = WRAP
    ws.column_dimensions["A"].width = 16

    ws["B2"] = (
        "Assumptions Drivers — Col A = Baseline Yes/No. "
        "YES = derived from company past performance (ZI FY24 pools). "
        "NO = external source / market rate / MODEL CONST / AI thesis lever. "
        "Verbatim click-links: cols I/J/K."
    )
    ws["B3"] = (
        "Baseline YES = rows 21–23 only (S&M $425 / payroll $318.8 / commission $106.2 from FY24). "
        "All forward rates and AI levers = Baseline NO."
    )

    for r, yn in BASELINE.items():
        cell = ws[f"A{r}"]
        cell.value = yn
        cell.font = BOLD
        cell.alignment = WRAP
        cell.border = THIN
        cell.fill = YES_F if yn == "YES" else NO_F

    # Fix leftover 15% in CapEx why
    if isinstance(ws["D12"].value, str) and "15%" in ws["D12"].value:
        ws["D12"] = ws["D12"].value.replace("15%", "18%")

    ws["B35"] = "Baseline legend (col A):"
    ws["C35"] = (
        "YES = yellow value derived from ZI past performance / FY24 S&M pools. "
        "NO = external publication, market quote, MODEL CONST, or AI operating lever."
    )
    ws["C35"].alignment = LEFT


def patch_assumptions_list(ws):
    # Unmerge before edits
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    # Sync values that drifted from Drivers
    ws.cell(9, 4).value = "−18% in Year 1"
    ws.cell(9, 12).value = (
        "Payroll_Y1 = Assumptions_Drivers!C22 × (1 − C7). C7=18%. Baseline payroll = C22."
    )
    ws.cell(9, 13).value = (
        "Sales payroll cut 18% in Y1 — matches published US B2B SaaS net SDR headcount "
        "decline of 18% YoY (Assumptions_Drivers!C7). Synced from Drivers (was −15%)."
    )

    ws.cell(13, 4).value = "18% of revenue"
    ws.cell(13, 13).value = (
        "G&A = 18% of revenue (Assumptions_Drivers!C11 = Blossom 2025 SaaS median). "
        "Same % in AI and Base; not ZI filing ~24.3%."
    )

    ws.cell(16, 4).value = "SOFR 4.30% + 4.00% (= 8.30%)"
    ws.cell(17, 4).value = "SOFR 4.30% + 5.00% (= 9.30%)"

    ws.cell(21, 4).value = "12.9x EV / EBITDA"
    ws.cell(21, 12).value = (
        "Exit_EV = Y5_EBITDA × Assumptions_Drivers!C20 (12.9x); "
        "Exit_Equity = Exit_EV − Debt_end; MOIC / IRR from that."
    )
    ws.cell(21, 13).value = (
        "Exit multiple 12.9x EV/EBITDA = Aventis private SaaS 1st quartile "
        "(Assumptions_Drivers!C20). Synced from Drivers (was 13.0x)."
    )

    # Baseline Yes/No on list
    ws.cell(6, 14).value = "Baseline Yes/No"
    ws.cell(6, 14).fill = HDR
    ws.cell(6, 14).font = HDR_F
    ws.column_dimensions["N"].width = 14

    list_bl = {r: "NO" for r in range(7, 22)}
    for r, yn in list_bl.items():
        cell = ws.cell(r, 14, yn)
        cell.fill = NO_F
        cell.font = BOLD
        cell.alignment = WRAP
        cell.border = THIN

    # Clear old footer rows 22+ (now unmerged)
    for r in range(22, max(ws.max_row, 45) + 1):
        for c in range(2, 15):
            ws.cell(r, c).value = None

    ws.cell(22, 2).value = "BASELINE POOLS (must match Assumptions_Drivers)"
    ws.cell(22, 2).font = BOLD

    baselines = [
        (23, 16, "S&M baseline ($m)", "$425.0m", "C21", "YES", "ZI FY24-derived (~35%×$1,214.3m near filing $414.1m)"),
        (24, 17, "Payroll baseline ($m)", "$318.8m", "C22", "YES", "75% × $425m FY24 S&M carve"),
        (25, 18, "Commission baseline ($m)", "$106.2m", "C23", "YES", "25% × $425m FY24 S&M carve"),
        (26, 19, "CapEx base % of rev", "2.0%", "C24", "NO", "External SaaS CapEx 1–3% midpoint"),
        (27, 20, "WC base % of Δrev", "2.0%", "C25", "NO", "MODEL CONST plug"),
    ]
    for r, num, name, val, cref, yn, note in baselines:
        ws.cell(r, 2).value = num
        ws.cell(r, 3).value = name
        ws.cell(r, 4).value = val
        ws.cell(r, 5).value = "LBO (not DCF)"
        ws.cell(r, 6).value = "Assumptions_Drivers"
        ws.cell(r, 7).value = cref
        ws.cell(r, 8).value = f"Go to Assumptions_Drivers!{cref}"
        ws.cell(r, 13).value = note
        cell = ws.cell(r, 14, yn)
        cell.fill = YES_F if yn == "YES" else NO_F
        cell.font = BOLD
        cell.border = THIN

    ws.cell(29, 2).value = "NOTE"
    ws.cell(29, 2).font = BOLD
    ws.cell(30, 2).value = (
        "This list must match Assumptions_Drivers: payroll cut −18%, exit 12.9x, AI infra $34m, "
        "G&A 18%, SOFR 4.30%. Col N = Baseline Yes/No (same rule as Drivers col A)."
    )
    ws.cell(30, 2).alignment = LEFT

    ws.cell(32, 2).value = "SOURCES / VERBATIM"
    ws.cell(32, 2).font = BOLD
    ws.cell(33, 2).value = (
        "Assumptions_Drivers: Col A Baseline Yes/No | Col E source | Cols I/J/K clickable Verbatim. "
        "PDF: https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_DRIVER_SOURCES.pdf"
    )
    ws.cell(33, 2).alignment = LEFT

    if ws["B4"].value and "MODEL1" in str(ws["B4"].value):
        ws["B4"] = (
            "PDF: https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/"
            "cursor/vengeanceaiusclbomodel2-44dc/LBO/output/LBOMODEL2_DRIVER_SOURCES.pdf"
        )


def main():
    wb = load_workbook(XLSX)
    patch_drivers(wb["Assumptions_Drivers"])
    patch_assumptions_list(wb["00_Assumptions_List"])

    d = wb["Assumptions_Drivers"]
    assert d["C7"].value == 0.18
    assert d["C20"].value == 12.9
    assert d["A21"].value == "YES"
    assert d["A7"].value == "NO"
    assert d["I9"].value and "HYPERLINK" in str(d["I9"].value)

    al = wb["00_Assumptions_List"]
    assert "18%" in str(al.cell(9, 4).value)
    assert "12.9" in str(al.cell(21, 4).value)
    assert "18%" in str(al.cell(13, 4).value)

    wb.save(XLSX)
    shutil.copy2(XLSX, XLSX2)
    print("Synced Assumptions_List; added Drivers col A Baseline Yes/No")
    for r in (7, 9, 21, 22, 23):
        print(f"  Drivers A{r}={d[f'A{r}'].value} | {d[f'B{r}'].value}={d[f'C{r}'].value}")
    print("List:", al.cell(9, 4).value, "|", al.cell(13, 4).value, "|", al.cell(21, 4).value)


if __name__ == "__main__":
    main()
