#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
format_gstr1.py - monthly formatter for the Tally "Sales Register" export.

It turns the raw export  (BDBJUL.xlsx)  into the hand-formatted file
"BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" so the cosmetics never have to be done
by hand again.

What it does (identical to the July sample)
-------------------------------------------
  * company name line (first title line)                  -> bold
  * column heading row                                    -> bold + centred + wrapped, row height 30
  * section heading line, e.g. "GST - GST 18 %",
    "TXF - TAX FREE", "GST - GST 5 %", ...                -> bold + single underline
  * every invoice line
        Bill.No, Date, State                              -> centred
        SGST %, CGST %, IGST %                            -> centred in the
                sections that carry a rate; a section where every rate is 0
                (TAX FREE / RES bills) is left exactly as Tally exported it
        Date                                              -> d-mmm-yy
        Bill Amt, SGST Amt, CGST Amt, IGST Amt, Taxable   -> 0.00
  * every "TOTAL FOR ..." / "GRAND TOTAL" row
        whole row bold
        label moved from the Bill.No column to the Party Name column,
        right aligned, suffixed " REG :" for GST-rate totals and " : " for
        the other totals (exactly as in the sample)
        money columns -> 0.00
  * column widths, landscape A4 print setup, narrow margins

Nothing is hard-coded to row numbers or to the number of invoices: the header
row, the sections, the invoice lines and the total rows are all found with
regular expressions, so a month with 5 invoices or 500 invoices, one GST rate
or five, formats just the same.

Usage
-----
    python3 format_gstr1.py BDBJUL.xlsx
        -> writes "BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" next to the input
           (month/year are read from the "Sales Register From .. TO .." line)

    python3 format_gstr1.py BDBAUG.xlsx -o "aug done.xlsx"
    python3 format_gstr1.py BDBJUL.xlsx --dry-run          # show what would happen
    python3 format_gstr1.py BDBJUL.xlsx --inplace          # edit the file itself

Requires: openpyxl   (pip install openpyxl)
"""

from __future__ import annotations

import argparse
import calendar
import re
import shutil
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required.  Install it with:  pip install openpyxl")


# ===========================================================================
#  SETTINGS  -  everything you may ever want to tweak lives in this block
# ===========================================================================

# ---- total-row labels -----------------------------------------------------
# Suffix added to the label of a "TOTAL FOR GST <rate> %" row.
GST_TOTAL_SUFFIX = " REG :"
# Suffix added to every other "TOTAL FOR ..." row (e.g. TOTAL FOR TAX FREE).
OTHER_TOTAL_SUFFIX = " : "
# Suffix added to the "GRAND TOTAL" row.
GRAND_TOTAL_SUFFIX = " : "

# Cell the total labels are written into - "party" = the Party Name column (B).
TOTAL_LABEL_COLUMN = "party"

# ---- fonts / alignment ----------------------------------------------------
BOLD_TITLE_ROWS = 1          # how many title lines above the table get bolded
HEADER_ROW_HEIGHT = 30       # height of the heading row
MARK_HEADER_WRAP = True      # wrap the heading row text
RIGHT_ALIGN_TOTAL_NUMBERS = True  # right align the money cells of a total row

# Alignment of the SGST% / CGST% / IGST% columns:
#   "auto"   (default, as in the July sample) -> centre them for the sections
#            that really carry a tax rate; a section where every rate is 0
#            (TAX FREE / RES bills) is left exactly as Tally exported it
#   "always" -> centre every percent cell, whatever the section
#   "never"  -> never touch them, leave Tally's own alignment
PERCENT_ALIGNMENT = "auto"

# ---- numbers --------------------------------------------------------------
MONEY_FORMAT = "0.00"        # amount columns
DATE_FORMAT = "d-mmm-yy"     # date column

# ---- sheet geometry -------------------------------------------------------
# Column widths taken from the formatted July file (Excel character units).
COLUMN_WIDTHS = {
    "A": 17.85546875, "B": 30.7109375, "C": 8.7109375, "D": 10.7109375,
    "E": 10.7109375, "F": 10.5703125, "G": 5.7109375, "H": 10.7109375,
    "I": 5.7109375, "J": 9.5703125, "K": 5.7109375, "L": 7.7109375,
    "M": 10.7109375,
}
DEFAULT_ROW_HEIGHT = 15
PAGE_ORIENTATION = "landscape"
PAGE_PAPER_SIZE = 9          # 9 = A4
PAGE_SCALE = 95              # print scale %
PAGE_MARGINS = dict(         # "narrow" margins, as in the sample file
    left=0.5905511811023623, right=0.1968503937007874,
    top=0.5905511811023623, bottom=0.3937007874015748,
    header=0.3149606299212598, footer=0.3149606299212598,
)

# ---- output file name -----------------------------------------------------
# {prefix} comes from the input file name (BDBJUL.xlsx -> BDB)
OUTPUT_NAME_TEMPLATE = "{prefix} GSTR1 SYSTEM DATA {month}-{year}.xlsx"
DEFAULT_PREFIX = "BDB"

# ===========================================================================
#  REGEX PATTERNS  -  this is what makes the script month-independent
# ===========================================================================

# Column headings of the Tally export, matched loosely (case / dots / spaces).
COLUMN_PATTERNS = {
    "gst_no":   r"^gst\s*no\.?$",
    "party":    r"^party\s*name$",
    "bill_no":  r"^bill\s*\.?\s*no\.?$",
    "date":     r"^date$",
    "bill_amt": r"^bill\s*a\s*m?t\.?$",
    "state":    r"^state$",
    "sgst_pct": r"^sgst\s*%$",
    "sgst_amt": r"^sgst\s*amt\.?$",
    "cgst_pct": r"^cgst\s*%$",
    "cgst_amt": r"^cgst\s*amt\.?$",
    "igst_pct": r"^igst\s*%$",
    "igst_amt": r"^igst\s*amt\.?$",
    "taxable":  r"^taxable\s*amt\.?$",
}
COLUMN_PATTERNS = {k: re.compile(v, re.I) for k, v in COLUMN_PATTERNS.items()}

# which of those columns get which treatment
MONEY_COLUMNS = ("bill_amt", "sgst_amt", "cgst_amt", "igst_amt", "taxable")
PERCENT_COLUMNS = ("sgst_pct", "cgst_pct", "igst_pct")

# a section heading line of the register, e.g. "GST - GST 18 %" / "TXF - TAX FREE"
RE_SECTION_LABEL = re.compile(r"^\s*[A-Za-z][A-Za-z0-9&/.]{0,10}\s*-\s*\S.*$")

# a "TOTAL FOR ..." or "GRAND TOTAL" row (works on already-suffixed labels too)
RE_TOTAL_LABEL = re.compile(r"^\s*(?:TOTAL\s+FOR\b|GRAND\s+TOTAL\b)", re.I)
RE_GST_TOTAL = re.compile(r"^\s*TOTAL\s+FOR\s+GST\b.*$", re.I)
RE_GRAND_TOTAL = re.compile(r"^\s*GRAND\s+TOTAL\b", re.I)

# "Sales Register From 01-07-26 TO 31-07-26"
RE_REGISTER_PERIOD = re.compile(
    r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})\s*(?:TO|\u2013|-)\s*"
    r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})",
    re.I,
)

# trailing month in a file name: BDBJUL / BDB_JULY_2026 / XYZJAN26 ...
RE_MONTH_TAIL = re.compile(
    r"^(?P<prefix>.+?)[\s_\-]*(?P<mon>JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|"
    r"SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER|JAN|FEB|MAR|APR|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)$",
    re.I,
)


# ===========================================================================
#  small styling helpers
# ===========================================================================

def set_bold(cell, underline=None):
    """Bold a cell (keeping its font family/size/colour) and set underline."""
    font = copy(cell.font)
    font.bold = True
    if underline is not None:
        font.underline = underline
    cell.font = font


def set_alignment(cell, horizontal=None, wrap_text=None):
    alignment = copy(cell.alignment)
    if horizontal is not None:
        alignment.horizontal = horizontal
    if wrap_text is not None:
        alignment.wrap_text = wrap_text
    cell.alignment = alignment


def cell_text(value):
    return value.strip() if isinstance(value, str) else None


# ===========================================================================
#  finding the table inside a sheet
# ===========================================================================

def find_header_row(ws, max_scan=40):
    """Row that holds the column headings ('Gst No', 'Party Name', ...)."""
    best, best_hits = None, 0
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or max_scan, max_scan)):
        hits = 0
        for cell in row:
            text = cell_text(cell.value)
            if text and any(p.match(text) for p in COLUMN_PATTERNS.values()):
                hits += 1
        # a real heading row matches several known headings
        if hits >= 3 and hits > best_hits:
            best, best_hits = row[0].row, hits
    return best


def map_columns(ws, header_row):
    """{logical name -> column number} built from the heading text."""
    columns, unknown = {}, []
    for cell in ws[header_row]:
        text = cell_text(cell.value)
        if text is None:
            continue
        for name, pattern in COLUMN_PATTERNS.items():
            if name not in columns and pattern.match(text):
                columns[name] = cell.column
                break
        else:
            unknown.append(text)
    if unknown:
        print(f"    note: unknown heading(s) ignored: {', '.join(unknown)}")
    for needed in ("bill_no", "date", "party"):
        if needed not in columns:
            print(f"    WARNING: no '{needed}' column found - that part is skipped")
    return columns


def is_data_row(ws, row, columns):
    """True when the row looks like an invoice line (has a date or an amount)."""
    date_col = columns.get("date")
    if date_col and isinstance(ws.cell(row, date_col).value, datetime):
        return True
    for name in MONEY_COLUMNS:
        col = columns.get(name)
        if col and isinstance(ws.cell(row, col).value, (int, float)):
            return True
    return False


def classify(ws, row, columns, last_col):
    """'title' | 'section' | 'data' | 'total' | 'blank' | 'unknown'."""
    values = [ws.cell(row, c).value for c in range(1, last_col + 1)]
    if all(v is None or (isinstance(v, str) and not v.strip()) for v in values):
        return "blank"
    for value in values:
        text = cell_text(value)
        if text and RE_TOTAL_LABEL.match(text):
            return "total"
    if is_data_row(ws, row, columns):
        return "data"
    for value in values:
        text = cell_text(value)
        if text and RE_SECTION_LABEL.match(text):
            return "section"
    return "unknown"


def find_register_period(ws):
    """('JULY', 2026) from the 'Sales Register From 01-07-26 TO 31-07-26' line."""
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 12, 12)):
        for cell in row:
            text = cell_text(cell.value)
            if not text:
                continue
            match = RE_REGISTER_PERIOD.search(text)
            if not match:
                continue
            _, month, year = match.group(1), match.group(2), match.group(3)
            month_no = int(month)
            year_no = int(year) + (2000 if len(year) == 2 else 0)
            if 1 <= month_no <= 12:
                return calendar.month_name[month_no].upper(), year_no
    return None, None


# ===========================================================================
#  the formatting itself
# ===========================================================================

def new_total_label(text):
    """'TOTAL FOR GST 18 %' -> 'TOTAL FOR GST 18 % REG :'   (idempotent)."""
    base = text.strip()
    base = re.sub(r"[\s:]+$", "", base)                 # drop trailing ' :'
    base = re.sub(r"\s*\bREG\b\s*$", "", base, re.I)    # drop a trailing 'REG'
    base = re.sub(r"[\s:]+$", "", base).strip()

    if RE_GRAND_TOTAL.match(base):
        return base + GRAND_TOTAL_SUFFIX
    if RE_GST_TOTAL.match(base):
        return base + GST_TOTAL_SUFFIX
    return base + OTHER_TOTAL_SUFFIX


def style_data_row(ws, row, columns, center_percents):
    for name in MONEY_COLUMNS:
        col = columns.get(name)
        if col:
            ws.cell(row, col).number_format = MONEY_FORMAT
    for name in ("bill_no", "date", "state"):
        col = columns.get(name)
        if col:
            set_alignment(ws.cell(row, col), horizontal="center")
    if center_percents:
        for name in PERCENT_COLUMNS:
            col = columns.get(name)
            if col:
                set_alignment(ws.cell(row, col), horizontal="center")
    date_col = columns.get("date")
    if date_col:
        ws.cell(row, date_col).number_format = DATE_FORMAT


def carries_tax(ws, rows, columns):
    """True when one of these rows holds a non-zero SGST/CGST/IGST percentage."""
    for row in rows:
        for name in PERCENT_COLUMNS:
            col = columns.get(name)
            if col is None:
                continue
            value = ws.cell(row, col).value
            if isinstance(value, (int, float)) and value != 0:
                return True
    return False


def percent_decision(ws, rows, columns):
    """(centre the percent cells?, note for the log) for a group of invoice rows."""
    if PERCENT_ALIGNMENT == "always":
        return True, "percent columns centred"
    if PERCENT_ALIGNMENT == "never":
        return False, "percent columns left untouched"
    if carries_tax(ws, rows, columns):
        return True, "percent columns centred"
    return False, "percent columns left as exported (all 0 %)"


def style_total_row(ws, row, columns, last_col, stats):
    # 1. bold the whole row (empty cells included - that is what the sample has)
    for col in range(1, last_col + 1):
        set_bold(ws.cell(row, col))

    # 2. money cells keep 2 decimals (and are right aligned)
    for name in MONEY_COLUMNS:
        col = columns.get(name)
        if col is None:
            continue
        cell = ws.cell(row, col)
        cell.number_format = MONEY_FORMAT
        if RIGHT_ALIGN_TOTAL_NUMBERS:
            set_alignment(cell, horizontal="right")

    # 3. move the label into the Party Name column, right aligned, with suffix
    label_cell = None
    for col in range(1, last_col + 1):
        text = cell_text(ws.cell(row, col).value)
        if text and RE_TOTAL_LABEL.match(text):
            label_cell = ws.cell(row, col)
            break
    if label_cell is None:
        return

    new_text = new_total_label(label_cell.value)
    target_col = columns.get(TOTAL_LABEL_COLUMN)
    if target_col is None:
        target_col = label_cell.column
    target = ws.cell(row, target_col)

    if label_cell.coordinate != target.coordinate:
        label_cell.value = None                 # clear the old position
        stats["relabelled"].append((label_cell.coordinate, target.coordinate, new_text))
    target.value = new_text
    set_bold(target)
    set_alignment(target, horizontal="right")


def style_header_row(ws, header_row, last_col):
    for col in range(1, last_col + 1):
        cell = ws.cell(header_row, col)
        set_bold(cell)
        set_alignment(cell, horizontal="center",
                      wrap_text=True if MARK_HEADER_WRAP else None)
    ws.row_dimensions[header_row].height = HEADER_ROW_HEIGHT


def style_titles(ws, header_row):
    """Bold the company-name line(s) above the table."""
    if BOLD_TITLE_ROWS <= 0:
        return
    done = 0
    for row in range(1, header_row):
        cells = [ws.cell(row, c) for c in range(1, ws.max_column + 1)]
        if not any(c.value not in (None, "") for c in cells):
            continue
        done += 1
        if done > BOLD_TITLE_ROWS:
            break
        for cell in cells:
            if cell.value not in (None, ""):
                set_bold(cell)


def apply_sheet_geometry(ws):
    for letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[letter].width = width
    ws.sheet_format.defaultRowHeight = DEFAULT_ROW_HEIGHT
    ws.page_setup.orientation = PAGE_ORIENTATION
    ws.page_setup.paperSize = PAGE_PAPER_SIZE
    ws.page_setup.scale = PAGE_SCALE
    for key, value in PAGE_MARGINS.items():
        setattr(ws.page_margins, key, value)


def format_worksheet(ws, verbose=True):
    """Format one sheet.  Returns a stats dict (counts + notes)."""
    stats = {"invoice_rows": 0, "total_rows": 0, "sections": [], "relabelled": [],
             "unknown_rows": [], "ok": False}

    header_row = find_header_row(ws)
    if header_row is None:
        if verbose:
            print(f"    sheet '{ws.title}': no column heading row found - skipped")
        return stats

    columns = map_columns(ws, header_row)
    last_col = header_row and max(
        (c.column for c in ws[header_row] if c.value not in (None, "")),
        default=ws.max_column,
    )

    style_titles(ws, header_row)
    style_header_row(ws, header_row, last_col)

    # ---- pass 1: classify every row and group the invoice lines by section ----
    plan, loose_rows, current = [], [], None
    for row in range(header_row + 1, ws.max_row + 1):
        kind = classify(ws, row, columns, last_col)
        plan.append((row, kind))
        if kind == "section":
            for col in range(1, last_col + 1):
                text = cell_text(ws.cell(row, col).value)
                if text and RE_SECTION_LABEL.match(text):
                    current = {"label": text, "row": row, "col": col,
                               "invoices": 0, "rows": []}
                    stats["sections"].append(current)
                    break
        elif kind == "data":
            stats["invoice_rows"] += 1
            if current is not None:
                current["invoices"] += 1
                current["rows"].append(row)
            else:
                loose_rows.append(row)          # invoice line before any section
        elif kind == "unknown":
            stats["unknown_rows"].append(row)

    # which sections get their SGST% / CGST% / IGST% cells centred
    for section in stats["sections"]:
        section["center_percents"], section["percent_note"] = percent_decision(
            ws, section["rows"], columns)
    loose_center, _ = percent_decision(ws, loose_rows, columns)

    # ---- pass 2: apply the styles ----
    section_of_row = {row: s for s in stats["sections"] for row in s["rows"]}
    for row, kind in plan:
        if kind == "data":
            if row in section_of_row:
                center = section_of_row[row]["center_percents"]
            else:
                center = loose_center
            style_data_row(ws, row, columns, center)
        elif kind == "total":
            style_total_row(ws, row, columns, last_col, stats)
            stats["total_rows"] += 1
        elif kind == "section":
            for section in stats["sections"]:
                if section["row"] == row:
                    set_bold(ws.cell(row, section["col"]), underline="single")
                    break

    apply_sheet_geometry(ws)
    stats["ok"] = True
    stats["header_row"] = header_row
    stats["columns"] = columns
    return stats


# ===========================================================================
#  output file name
# ===========================================================================

def prefix_from_filename(path):
    """BDBJUL.xlsx -> 'BDB', BDB_JULY_2026.xlsx -> 'BDB', XYZJAN26.xlsx -> 'XYZ'."""
    stem = path.stem.strip()
    base = re.sub(r"[\s_\-]*\d{2,4}$", "", stem).strip(" _-")     # drop a trailing year
    match = RE_MONTH_TAIL.match(base)
    prefix = match.group("prefix") if match else stem             # else: whole file name
    prefix = re.sub(r"[_\s]+", " ", prefix).strip(" _-")
    if "GSTR1" in prefix.upper():        # already an output name -> keep 'BDB' only
        prefix = prefix.split()[0]
    return prefix or DEFAULT_PREFIX


def build_output_path(src, ws, outdir=None, template=OUTPUT_NAME_TEMPLATE):
    month, year = find_register_period(ws)
    prefix = prefix_from_filename(src)
    if month is None:
        print("    WARNING: could not read the 'Sales Register From ... TO ...' line;"
              " using the file name for the output name")
        name = f"{prefix} GSTR1 SYSTEM DATA {src.stem}.xlsx"
    else:
        name = template.format(prefix=prefix, month=month, year=year, ext=src.suffix)
    folder = Path(outdir) if outdir else src.parent
    return folder / name


# ===========================================================================
#  main
# ===========================================================================

def format_workbook(wb, out_path, sheet_name=None, dry_run=False):
    """Format the sheets of an already loaded workbook and save it."""
    sheets = [wb[sheet_name]] if sheet_name else wb.worksheets

    all_stats = []
    for ws in sheets:
        print(f"    sheet '{ws.title}':")
        stats = format_worksheet(ws)
        all_stats.append(stats)
        if stats["ok"]:
            print(f"      heading row .......... {stats['header_row']}")
            print(f"      invoice lines ........ {stats['invoice_rows']}")
            print(f"      total rows ........... {stats['total_rows']}")
            for section in stats["sections"]:
                print(f"      section '{section['label']}' -> {section['invoices']} "
                      f"invoice(s), {section['percent_note']}")
            for old, new, text in stats["relabelled"]:
                print(f"      label {old} -> {new}: {text!r}")
            if stats["unknown_rows"]:
                print(f"      WARNING: unrecognised row(s) left untouched: "
                      f"{stats['unknown_rows']}")
        else:
            print("      nothing formatted")

    if not any(s["ok"] for s in all_stats):
        raise SystemExit("ERROR: no sheet with a recognisable Sales Register heading row. "
                         "Is this the right file?")

    if dry_run:
        print(f"[dry run] would write: {out_path}")
        return None

    out_path.parent.mkdir(parents=True, exist_ok=True)
    backup = out_path.with_suffix(out_path.suffix + ".bak")
    if out_path.exists() and not backup.exists():
        shutil.copy2(out_path, backup)
        print(f"    existing {out_path.name} backed up as {backup.name}")

    wb.save(out_path)
    print(f"    saved -> {out_path}")
    return out_path


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Apply the monthly GST-R1 sales-register formatting to a raw "
                    "Tally export (openpyxl, regex driven).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="examples:\n"
               "  python3 format_gstr1.py BDBJUL.xlsx\n"
               "  python3 format_gstr1.py BDBAUG.xlsx -o out/AUG.xlsx\n"
               "  python3 format_gstr1.py BDBJUL.xlsx --dry-run\n",
    )
    parser.add_argument("input", help="raw Tally sales register, e.g. BDBJUL.xlsx")
    parser.add_argument("-o", "--output", help="output file (default: name built from "
                                               "the register month)")
    parser.add_argument("--outdir", help="folder for the auto-named output file")
    parser.add_argument("--sheet", help="only format this sheet")
    parser.add_argument("--inplace", action="store_true",
                        help="format the input file itself (a .bak is kept)")
    parser.add_argument("--dry-run", action="store_true",
                        help="report what would be done, write nothing")
    args = parser.parse_args(argv)

    src = Path(args.input).expanduser()
    if not src.is_file():
        raise SystemExit(f"ERROR: file not found: {src}")
    if src.suffix.lower() not in (".xlsx", ".xlsm"):
        raise SystemExit("ERROR: this needs an .xlsx/.xlsm file (Tally 'Excel' export). "
                         f"Got: {src.name}")

    print(f"formatting {src.name}")
    wb = load_workbook(src)                      # keeps every existing style

    if args.output:
        out_path = Path(args.output).expanduser()
        if out_path.parent == Path(".") and not out_path.is_absolute():
            out_path = src.parent / out_path     # bare -o name -> next to the input
    elif args.inplace:
        out_path = src
    else:
        first = wb[args.sheet] if args.sheet else wb.worksheets[0]
        out_path = build_output_path(src, first, outdir=args.outdir)

    return format_workbook(wb, out_path, sheet_name=args.sheet,
                           dry_run=args.dry_run)


if __name__ == "__main__":
    main()
