"""Export vengeanceaiUSCMODEL6 sheet snapshots as agent-readable CSVs."""

from __future__ import annotations

from pathlib import Path
from typing import Dict

import pandas as pd
from openpyxl import load_workbook


def export_model2_csvs(xlsx_path: Path, out_dir: Path) -> Dict[str, Path]:
    """Write MODEL6_*.csv copies of Cover / 3-Statement / DCF sheets (formulas as text)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(xlsx_path, data_only=False)
    mapping = {
        "Cover Page": "MODEL6_Cover_Page.csv",
        "3 Statement Model": "MODEL6_3_Statement_Model.csv",
        "DCF Model": "MODEL6_DCF_Model.csv",
    }
    written: Dict[str, Path] = {}
    for sheet, fname in mapping.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        rows = []
        for row in ws.iter_rows(values_only=True):
            # trim trailing empties
            vals = list(row)
            while vals and vals[-1] is None:
                vals.pop()
            if any(v is not None and v != "" for v in vals):
                rows.append(vals)
        path = out_dir / fname
        # ragged rows → pad
        width = max((len(r) for r in rows), default=0)
        norm = [list(r) + [None] * (width - len(r)) for r in rows]
        pd.DataFrame(norm).to_csv(path, index=False, header=False)
        written[sheet] = path
    return written
