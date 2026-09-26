#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
compare_workbooks.py - check a generated file against a hand-formatted sample.

    python3 compare_workbooks.py "BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" out/"BDB ... JULY-2026.xlsx"
    python3 compare_workbooks.py --folder out/         # every generated file against
                                                       # the samples in this folder
    python3 compare_workbooks.py REF.xlsx GOT.xlsx --all

Reports every difference in values, bold/underline, alignment, number formats,
column widths, row heights and page setup.

A reference cell that holds a formula (the samples' combined
"TOTAL FOR GST 18 %" rows use =E11+E27) is evaluated and compared with the
value in the generated file, so a formula and its sum count as equal.

Differences that cannot be seen are reported as "cosmetic" and do not fail the
comparison:

  * a font / number format / alignment on an empty cell
  * a cell that is right aligned only because it holds a number

Exit code: 0 = same formatting, 1 = real differences found.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required.  Install it with:  pip install openpyxl")

NUMERIC = (int, float, dt.datetime, dt.date)
CELL_REF = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def is_empty(value):
    return value is None or (isinstance(value, str) and not value.strip())


def effective_align(cell):
    """Excel's default alignment when nothing is set explicitly."""
    horizontal = cell.alignment.horizontal
    if horizontal:
        return horizontal
    return "right" if isinstance(cell.value, NUMERIC) else "left"


def column_widths(ws):
    """{column index -> width}, resolving column ranges like min=4 max=5."""
    widths = {}
    for dim in ws.column_dimensions.values():
        if dim.width is None:
            continue
        for index in range(dim.min or 1, (dim.max or dim.min or 1) + 1):
            widths[index] = dim.width
    return widths


def eval_formula(formula, ws, depth=0):
    """Value of a simple '=A1+B2+...' formula (None when not that simple)."""
    if depth > 5 or not isinstance(formula, str) or not formula.startswith("="):
        return None
    body = formula[1:].strip()
    if not body or any(ch in body for ch in "*/()"):
        return None
    total = 0.0
    for part in body.split("+"):
        match = CELL_REF.match(part.strip())
        if not match:
            return None
        value = ws[f"{match.group(1)}{match.group(2)}"].value
        if isinstance(value, str) and value.startswith("="):
            value = eval_formula(value, ws, depth + 1)
        if not isinstance(value, (int, float)):
            return None
        total += value
    return total


def compare_sheets(ref_ws, got_ws, show_cosmetic=False):
    diffs, cosmetic = [], []
    max_row = max(ref_ws.max_row, got_ws.max_row)
    max_col = max(ref_ws.max_column, got_ws.max_column)

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            ref, got = ref_ws.cell(row, col), got_ws.cell(row, col)
            where = f"{get_column_letter(col)}{row}"

            ref_value = None if is_empty(ref.value) else ref.value
            got_value = None if is_empty(got.value) else got.value

            # a "=E11+E27" formula and its sum are the same thing
            if isinstance(ref_value, str) and ref_value.startswith("="):
                ref_value = eval_formula(ref_value, ref_ws)
            if isinstance(got_value, str) and got_value.startswith("="):
                got_value = eval_formula(got_value, got_ws)
            if isinstance(ref_value, (int, float)) and isinstance(got_value, (int, float)) \
                    and abs(ref_value - got_value) < 1e-6:
                ref_value = got_value
            if isinstance(ref_value, str) and isinstance(got_value, str) \
                    and ref_value.strip() == got_value.strip():
                ref_value = got_value
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


def compare_layout(ref_ws, got_ws):
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
        if a != b and not (isinstance(a, float) and isinstance(b, float)
                           and abs(a - b) < 1e-9):
            diffs.append(f"{label}: expected {a}, got {b}")

    return diffs


def compare_files(ref_path, got_path, show_cosmetic=False):
    """Compare two workbooks; returns a dict with the differences found."""
    ref_wb, got_wb = load_workbook(ref_path), load_workbook(got_path)
    differences, cosmetic = [], []

    for ws_ref in ref_wb.worksheets:
        if ws_ref.title not in got_wb.sheetnames:
            differences.append(f"sheet '{ws_ref.title}' is missing in the generated file")
            continue
        diffs, notes = compare_sheets(ws_ref, got_wb[ws_ref.title], show_cosmetic)
        differences += diffs
        cosmetic += notes
        differences += compare_layout(ws_ref, got_wb[ws_ref.title])

    return dict(reference=Path(ref_path), generated=Path(got_path),
                differences=differences, cosmetic=len(cosmetic), notes=cosmetic)


def register_period(path):
    """('JULY', 2026) read from the 'Sales Register From ... TO ...' line."""
    try:
        ws = load_workbook(path, read_only=True).active
    except Exception:
        return None
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 12, 12)):
        for cell in row:
            if isinstance(cell.value, str):
                match = re.search(
                    r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})\s*(?:TO|\u2013|-)\s*"
                    r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", cell.value, re.I)
                if match:
                    month, year = int(match.group(2)), int(match.group(3))
                    if 1 <= month <= 12:
                        return month, year + (2000 if year < 100 else 0)
    return None


def client_code(name):
    match = re.match(r"^[A-Za-z]+", name)
    return match.group(0).upper() if match else ""


def find_sample(generated, samples_dir):
    """The client's own file for the same month, in the samples folder."""
    code = client_code(generated.stem)
    period = register_period(generated)
    for candidate in sorted(Path(samples_dir).glob("*GSTR1*.xlsx")):
        if candidate.resolve() == Path(generated).resolve():
            continue
        other = client_code(candidate.stem)
        if not other or not (other.startswith(code) or code.startswith(other)):
            continue
        if period is None or register_period(candidate) == period:
            return candidate
    return None


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    show_cosmetic = "--all" in argv
    argv = [a for a in argv if a != "--all"]

    if "--folder" in argv:
        index = argv.index("--folder")
        folder = argv[index + 1] if len(argv) > index + 1 else "."
        samples_dir = "."
        if "--samples" in argv:
            samples_dir = argv[argv.index("--samples") + 1]
        generated_files = [p for p in sorted(Path(folder).glob("*GSTR1*.xlsx"))
                           if p.resolve() != Path(samples_dir).resolve()]
        pairs = []
        for generated in generated_files:
            sample = find_sample(generated, samples_dir)
            if sample is None:
                print(f"{generated.name}: no sample for this client/month in "
                      f"{samples_dir} - skipped")
                continue
            pairs.append((sample, generated))
        if not pairs:
            print(f"nothing to compare in {folder}")
            return 2
    else:
        if len(argv) < 2:
            print(__doc__)
            return 2
        pairs = [(argv[0], argv[1])]

    total = 0
    for ref_path, got_path in pairs:
        if got_path is None:
            print(f"{Path(ref_path).name}: no formatted counterpart found")
            continue
        ref_path, got_path = Path(ref_path), Path(got_path)
        for path in (ref_path, got_path):
            if not path.is_file():
                print(f"ERROR: file not found: {path}")
                return 2
        result = compare_files(ref_path, got_path, show_cosmetic)
        print(f"reference : {ref_path.name}")
        print(f"generated : {got_path.name}")
        print(f"  {len(result['differences'])} difference(s), "
              f"{result['cosmetic']} cosmetic note(s)")
        for line in result["differences"][:40]:
            print(f"    DIFF     {line}")
        if show_cosmetic:
            for line in result["notes"][:80]:
                print(f"    cosmetic {line}")
        if result["differences"]:
            print("  RESULT: formatting does NOT match.")
            total += 1
        else:
            print("  RESULT: identical formatting "
                  f"({result['cosmetic']} invisible cosmetic note(s)).")
        print()

    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
