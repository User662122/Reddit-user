#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analyse_format.py - work out the per-row-type formatting of a formatted sample.

For a hand-formatted "GSTR1 SYSTEM DATA" file it prints, for every row class
(title / header / section / data / total / blank), the style applied to each
column, plus the sheet geometry.  Used to derive the client profiles in
format_gstr1.py.

    python3 analyse_format.py "BDB GSTR1 SYSTEM DATA JULY-2026.xlsx"
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

DATA_HINTS = ("TOTAL", "GST", "TAX FREE")


def classify(ws, row):
    values = [ws.cell(row, c).value for c in range(1, ws.max_column + 1)]
    if all(v is None for v in values):
        return "blank"
    texts = [v for v in values if isinstance(v, str)]
    if any(t.strip().upper().startswith(DATA_HINTS) for t in texts):
        return "total"
    if any(isinstance(v, (int, float)) for v in values if v is not None) and any(
            isinstance(v, str) and v.count("-") >= 2 and len(v) > 8 for v in texts):
        return "data"
    if texts:
        return "section/title"
    return "?"


def style_of(cell):
    bits = []
    if cell.font.bold:
        bits.append("B")
    if cell.font.underline:
        bits.append("U")
    if cell.alignment.horizontal:
        bits.append(cell.alignment.horizontal[:1].upper())
    if cell.alignment.wrap_text:
        bits.append("W")
    if cell.number_format != "General":
        bits.append(f"[{cell.number_format}]")
    return "".join(bits) or "."


def main():
    for name in sys.argv[1:]:
        wb = load_workbook(Path(name))
        ws = wb.active
        print("=" * 90)
        print(f"{Path(name).name}   ({ws.max_row} rows)")
        header = None
        for row in range(1, min(ws.max_row, 12) + 1):
            first = ws.cell(row, 1).value
            if isinstance(first, str) and first.strip().lower() == "gst no":
                header = row
                break
        print(f"header row: {header}")
        for row in range(1, ws.max_row + 1):
            kind = classify(ws, row)
            styles = [f"{get_column_letter(c)}:{style_of(ws.cell(row, c))}"
                      for c in range(1, ws.max_column + 1)
                      if style_of(ws.cell(row, c)) != "."]
            label = ""
            for c in range(1, ws.max_column + 1):
                v = ws.cell(row, c).value
                if isinstance(v, str) and v.strip():
                    label = v.strip()[:38]
                    break
            print(f"{row:>3} {kind:<12} {label:<40} {' '.join(styles)}")
        widths = {}
        for dim in ws.column_dimensions.values():
            if dim.width:
                for i in range(dim.min or 1, (dim.max or dim.min or 1) + 1):
                    widths[get_column_letter(i)] = round(dim.width, 3)
        print(f"widths : {widths}")
        print(f"heights: { {k: v.height for k, v in ws.row_dimensions.items() if v.height} }")
        print(f"default row height {ws.sheet_format.defaultRowHeight}; "
              f"page {ws.page_setup.orientation} paper {ws.page_setup.paperSize} "
              f"scale {ws.page_setup.scale}; margins L{ws.page_margins.left:.3f} "
              f"R{ws.page_margins.right:.3f} T{ws.page_margins.top:.3f} "
              f"B{ws.page_margins.bottom:.3f} H{ws.page_margins.header:.3f} "
              f"F{ws.page_margins.footer:.3f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
