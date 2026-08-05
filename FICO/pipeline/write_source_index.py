"""Write a clickable Source Index on the Cover Page (MODEL5)."""

from __future__ import annotations

import csv
from pathlib import Path

from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .model6_assumptions import MODEL_NAME, SOURCE_LINKS

LINK_FONT = Font(name="Calibri", size=10, color="0563C1", underline="single")
HDR_FONT = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
LABEL_FONT = Font(name="Calibri", bold=True, size=10)
BODY_FONT = Font(name="Calibri", size=9)
HDR_FILL = PatternFill("solid", fgColor="C65911")
NOTE_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def write_cover_source_index(wb) -> int:
    """Place a Source Index with clickable URLs on Cover Page (cols E–G)."""
    if "Cover Page" not in wb.sheetnames:
        return 0
    ws = wb["Cover Page"]

    start = 24  # below typical cover content
    ws.cell(row=start, column=5).value = f"{MODEL_NAME} — SOURCE INDEX (click any link)"
    ws.cell(row=start, column=5).font = HDR_FONT
    ws.cell(row=start, column=5).fill = HDR_FILL
    ws.merge_cells(start_row=start, start_column=5, end_row=start, end_column=7)

    ws.cell(row=start + 1, column=5).value = (
        "Every assumption / WACC / filing source used in this workbook. "
        "Same links appear on 3-statement col U/V and DCF col U/V."
    )
    ws.cell(row=start + 1, column=5).font = Font(name="Calibri", italic=True, size=9)
    ws.cell(row=start + 1, column=5).fill = NOTE_FILL
    ws.merge_cells(start_row=start + 1, start_column=5, end_row=start + 1, end_column=7)

    headers = [("E", "Source"), ("F", "Clickable link"), ("G", "URL")]
    for col, title in headers:
        cell = ws[f"{col}{start + 2}"]
        cell.value = title
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="833C0C")
        cell.border = THIN

    n = 0
    for i, (name, url) in enumerate(sorted(SOURCE_LINKS.items())):
        r = start + 3 + i
        label = ws.cell(row=r, column=5, value=name)
        label.font = LABEL_FONT
        label.border = THIN
        label.alignment = Alignment(wrap_text=True, vertical="center")

        link = ws.cell(row=r, column=6, value=name)
        link.hyperlink = url
        link.font = LINK_FONT
        link.border = THIN
        link.alignment = Alignment(wrap_text=True, vertical="center")

        raw = ws.cell(row=r, column=7, value=url)
        raw.hyperlink = url
        raw.font = LINK_FONT
        raw.border = THIN
        raw.alignment = Alignment(wrap_text=True, vertical="center")
        n += 1

    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 36
    ws.column_dimensions["G"].width = 70
    return n


def export_source_index_csv(out_dir: Path, *, ticker: str = "FICO") -> Path:
    """Write MODEL6_*_SOURCE_LINKS.csv — every named source + URL."""
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"MODEL6_{ticker}_SOURCE_LINKS.csv"
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["source", "url"])
        w.writeheader()
        for name, url in sorted(SOURCE_LINKS.items()):
            w.writerow({"source": name, "url": url})
    return path
