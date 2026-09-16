#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compare_workbooks.py - check a generated file against a hand-formatted sample.

    python3 compare_workbooks.py "BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" generated.xlsx

Reports every difference in values, fonts (bold / underline), alignment,
number formats, column widths, row heights and page setup.

Purely invisible differences are marked "(cosmetic)" and do not make the
comparison fail:

  * a number format on a cell that is empty      -> nothing is displayed
  * "right" vs "General" on a number             -> Excel right-aligns either way

Exit code: 0 = same formatting, 1 = real differences found.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required.  Install it with:  pip install openpyxl")

NUMERIC = (int, float, dt.datetime, dt.date)


def is_empty(value):
    return value is None or (isinstance(value, str) and not value.strip())


def effective_align(cell):
    """Excel's default alignment when nothing is set explicitly."""
    horizontal = cell.alignment.horizontal
    if horizontal:
        return horizontal
    return "right" if isinstance(cell.value, NUMERIC) else "left"


def compare_cells(ref_ws, got_ws, max_row, max_col, show_cosmetic):
    diffs, cosmetic = [], []

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            ref, got = ref_ws.cell(row, col), got_ws.cell(row, col)
            where = f"{get_column_letter(col)}{row}"

            ref_value = None if is_empty(ref.value) else ref.value
            got_value = None if is_empty(got.value) else got.value
            if ref_value != got_value:
                diffs.append(f"{where}: value  expected {ref_value!r}, got {got_value!r}")

            for label, a, b in (
                ("bold", ref.font.bold, got.font.bold),
                ("underline", ref.font.underline, got.font.underline),
                ("italic", ref.font.italic, got.font.italic),
                ("wrap", bool(ref.alignment.wrap_text), bool(got.alignment.wrap_text)),
            ):
                if bool(a) == bool(b):
                    continue
                if is_empty(ref.value) and is_empty(got.value):
                    # a font on a cell with nothing in it cannot be seen
                    cosmetic.append(f"{where}: {label} on an empty cell  "
                                    f"expected {bool(a)}, got {bool(b)}")
                else:
                    diffs.append(f"{where}: {label}  expected {bool(a)}, got {bool(b)}")

            if effective_align(ref) != effective_align(got):
                note = (f"{where}: align  expected {effective_align(ref)}"
                        f" (set: {ref.alignment.horizontal}), "
                        f"got {effective_align(got)} (set: {got.alignment.horizontal})")
                if is_empty(ref.value) and is_empty(got.value):
                    cosmetic.append(note.replace("align ", "align on an empty cell "))
                else:
                    diffs.append(note)

            if ref.number_format != got.number_format:
                if is_empty(ref.value) and is_empty(got.value):
                    cosmetic.append(f"{where}: number format on an empty cell  "
                                    f"expected {ref.number_format!r}, "
                                    f"got {got.number_format!r}")
                else:
                    diffs.append(f"{where}: number format  expected "
                                 f"{ref.number_format!r}, got {got.number_format!r}")
    return diffs, cosmetic


def column_widths(ws):
    """{column index -> width}, resolving column ranges like min=4 max=5."""
    widths = {}
    for dim in ws.column_dimensions.values():
        if dim.width is None:
            continue
        for index in range(dim.min or 1, (dim.max or dim.min or 1) + 1):
            widths[index] = dim.width
    return widths


def compare_sheet_layout(ref_ws, got_ws):
    diffs = []

    ref_widths, got_widths = column_widths(ref_ws), column_widths(got_ws)
    for index in sorted(set(ref_widths) | set(got_widths)):
        letter = get_column_letter(index)
        a, b = ref_widths.get(index), got_widths.get(index)
        if a is None or b is None or abs(a - b) > 0.011:
            diffs.append(f"column {letter}: width  expected {a}, got {b}")

    ref_heights = {k: v.height for k, v in ref_ws.row_dimensions.items() if v.height}
    got_heights = {k: v.height for k, v in got_ws.row_dimensions.items() if v.height}
    for row in sorted(set(ref_heights) | set(got_heights)):
        a, b = ref_heights.get(row), got_heights.get(row)
        if a != b:
            diffs.append(f"row {row}: height  expected {a}, got {b}")

    for label, a, b in (
        ("page orientation", ref_ws.page_setup.orientation, got_ws.page_setup.orientation),
        ("paper size", ref_ws.page_setup.paperSize, got_ws.page_setup.paperSize),
        ("print scale", ref_ws.page_setup.scale, got_ws.page_setup.scale),
        ("margin left", ref_ws.page_margins.left, got_ws.page_margins.left),
        ("margin right", ref_ws.page_margins.right, got_ws.page_margins.right),
        ("margin top", ref_ws.page_margins.top, got_ws.page_margins.top),
        ("margin bottom", ref_ws.page_margins.bottom, got_ws.page_margins.bottom),
        ("default row height", ref_ws.sheet_format.defaultRowHeight,
         got_ws.sheet_format.defaultRowHeight),
    ):
        if a != b and not (isinstance(a, float) and isinstance(b, float) and abs(a - b) < 1e-9):
            diffs.append(f"{label}: expected {a}, got {b}")

    return diffs


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    show_cosmetic = "--all" in argv
    argv = [a for a in argv if a != "--all"]
    if len(argv) != 2:
        print(__doc__)
        return 2

    ref_path, got_path = Path(argv[0]), Path(argv[1])
    for path in (ref_path, got_path):
        if not path.is_file():
            print(f"ERROR: file not found: {path}")
            return 2

    ref_wb, got_wb = load_workbook(ref_path), load_workbook(got_path)
    print(f"reference : {ref_path.name}")
    print(f"generated : {got_path.name}")

    total_diffs, total_cosmetic = 0, 0
    for ws_ref in ref_wb.worksheets:
        if ws_ref.title not in got_wb.sheetnames:
            print(f"\nsheet '{ws_ref.title}': MISSING in the generated file")
            total_diffs += 1
            continue
        ws_got = got_wb[ws_ref.title]
        max_row = max(ws_ref.max_row, ws_got.max_row)
        max_col = max(ws_ref.max_column, ws_got.max_column)

        diffs, cosmetic = compare_cells(ws_ref, ws_got, max_row, max_col, show_cosmetic)
        diffs += compare_sheet_layout(ws_ref, ws_got)
        total_diffs += len(diffs)
        total_cosmetic += len(cosmetic)

        print(f"\nsheet '{ws_ref.title}': {len(diffs)} difference(s), "
              f"{len(cosmetic)} cosmetic note(s)")
        for line in diffs:
            print(f"  DIFF     {line}")
        if show_cosmetic:
            for line in cosmetic:
                print(f"  cosmetic {line}")
        elif cosmetic:
            print(f"  ({len(cosmetic)} cosmetic note(s) hidden - use --all to see them)")

    print()
    if total_diffs:
        print(f"RESULT: {total_diffs} real difference(s) - formatting does NOT match.")
        return 1
    print("RESULT: identical formatting "
          f"({total_cosmetic} invisible cosmetic note(s)).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
