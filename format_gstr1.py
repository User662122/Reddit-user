#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
format_gstr1.py - monthly formatter for the Tally "Sales Register" exports.

One command turns every raw export into the hand-formatted file of that client:

    python3 format_gstr1.py BDBJUL.xlsx DPBJUL.xlsx HKIJUL.xlsx HTEJUL.xlsx

     BDBJUL.xlsx -> BDB GSTR1 SYSTEM DATA JULY-2026.xlsx
     DPBJUL.xlsx -> DPB GSTR1 SYSTEM DATA JULY 2026.xlsx
     HKIJUL.xlsx -> HKI GSTR1 SYSTEM DATA JULY-2026.xlsx
     HTEJUL.xlsx -> HTEI GSTR1 SYSTEM DATA JULY 2026.xlsx

Each client is described by a profile (the CLIENT PROFILES block).  The profile
is picked automatically from the input file name (BDBJUL -> BDB), or forced with
--profile.  Everything the profile does not mention falls back to the default
profile, so a brand-new client usually needs no code at all - just a file name
that starts with its code, or an explicit --profile.

What is formatted (all driven by the profile, per client)
---------------------------------------------------------
  * company-name line                     -> bold
  * column-heading row                    -> bold, centred, row height 30,
                                             "wrap" on the columns that need it
  * section line ("GST - GST 18 %",
    "TXF - TAX FREE", "G18 - GST 18%", ..) -> bold + underlined
  * every invoice line                    -> chosen columns centred
                                             (Bill No / Date / State / % cols),
                                             chosen columns wrapped,
                                             date as d-mmm-yy,
                                             chosen money columns as 0.00
  * every "TOTAL FOR ..." row             -> whole row bold, label moved to the
                                             Party Name column, right aligned and
                                             suffixed (" REG :", " : ", " REG : ")
                                             exactly like that client's sample,
                                             money columns as 0.00
  * "GRAND TOTAL" row                     -> same treatment, " : " suffix
  * column widths, landscape A4 print setup and margins per client

Two extra steps some clients need:
  * combine_same_rate - when a month is split into several sections with the same
    GST rate (e.g. registered + un-registered), the samples add a combined
    "TOTAL FOR GST 18 % : " row = sum of those sections (a live formula).
  * drop_blank_between_title_and_header - one client's sample has no empty row
    between "Sales Register From ..." and the heading row.

Nothing is bound to row numbers or to the number of invoices: the heading row,
the columns, the sections, the invoice lines and the total rows are all found
with regular expressions, so a month with 3 invoices or 300, one GST rate or
five, formats just the same.

Usage
-----
    python3 format_gstr1.py FILE [FILE ...] [options]

      -o FILE            write to this file (single input only)
      --outdir DIR       write the auto-named files into DIR
      --profile NAME     force a client profile (BDB, DPB, HKI, HTEI)
      --dry-run          report what would be done, write nothing
      --verify [REF]     after saving, compare with the hand-formatted file
                         (auto-found next to the input unless REF is given)
      --inplace          format the files themselves (a .bak is kept)
      --list-profiles    show the client profiles and stop

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
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required.  Install it with:  pip install openpyxl")


# ===========================================================================
#  REGEX PATTERNS  -  this block is what makes the script month-independent
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

MONEY_COLUMNS = ("bill_amt", "sgst_amt", "cgst_amt", "igst_amt", "taxable")
PERCENT_COLUMNS = ("sgst_pct", "cgst_pct", "igst_pct")

# a section heading line of the register, e.g. "GST - GST 18 %" / "TXF - TAX FREE"
RE_SECTION_LABEL = re.compile(r"^\s*[A-Za-z][A-Za-z0-9&/.]{0,10}\s*-\s*\S.*$")

# a "TOTAL FOR ..." / "GRAND TOTAL" row (works on already-suffixed labels too)
RE_TOTAL_LABEL = re.compile(r"^\s*(?:TOTAL\s+FOR\b|GRAND\s+TOTAL\b)", re.I)
RE_GST_TOTAL = re.compile(r"^\s*TOTAL\s+FOR\s+GST\b", re.I)
RE_GRAND_TOTAL = re.compile(r"^\s*GRAND\s+TOTAL\b", re.I)

# section kinds
RE_TAX_FREE = re.compile(r"\bTAX\s*FREE\b", re.I)
RE_UNREGISTERED = re.compile(r"\bUN\s*REG", re.I)

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
#  SETTINGS  -  values shared by every client profile
# ===========================================================================

MONEY_FORMAT = "0.00"        # amount columns
DATE_FORMAT = "d-mmm-yy"     # date column
HEADER_ROW_HEIGHT = 30       # height of the heading row
TITLE_BOLD_ROWS = 1          # how many title lines above the table get bolded
DEFAULT_ROW_HEIGHT = 15
PAGE_ORIENTATION = "landscape"
PAGE_PAPER_SIZE = 9          # 9 = A4
DEFAULT_PREFIX = "BDB"
RIGHT_ALIGN_TOTAL_NUMBERS = True
RIGHT_ALIGN_TOTAL_LABEL = True

# ===========================================================================
#  CLIENT PROFILES  -  one entry per client, taken from their formatted sample
# ===========================================================================
#
# Fields used by the formatter:
#   output_name_template       how the saved file is named ({prefix}/{month}/{year})
#   output_prefix              fixed prefix for the output name (else from file name)
#   drop_blank_between_title_and_header
#                              remove the empty spacer row between the
#                              "Sales Register From ..." line and the heading row
#   header_text_fixes          heading text to correct on the heading row
#   header_wrap                "all", or the columns that get wrap on the heading row
#   center_data                invoice-line columns that are centred
#   wrap_data                  invoice-line columns that get wrap
#   money_data                 invoice-line columns formatted 0.00
#   money_total                total-row columns formatted 0.00
#   center_zero_percent        centre the SGST%/CGST%/IGST% cells even in a section
#                              whose rates are all 0 (TAX FREE / RES bills)
#   mark_registered_section    write " REG" into the heading of a registered
#                              GST-rate section
#   reg_suffix / other_suffix / grand_suffix
#                              label suffixes for a registered section, an
#                              un-registered or TAX FREE section, and GRAND TOTAL
#   combine_same_rate          add a combined "TOTAL FOR GST <rate> % : " row that
#                              adds up the GST sections that share a rate
#   combine_columns            the columns those rows add up
#   combine_position           "right_after_total" or "before_next_section"
#   party_name_fixes           {name in Tally : corrected name} - corrections that
#                              were made by hand in the sample (e.g. adding "- RENT")
#   sections_without_underline section headings the sample left plain bold"
#   column_widths              Excel column widths
#   page_scale                 print scale %
#   margins                    page margins (inches)

NARROW_MARGINS = dict(left=0.5905511811023623, right=0.1968503937007874,
                      top=0.5905511811023623, bottom=0.3937007874015748,
                      header=0.31496062992125984, footer=0.31496062992125984)


def _profile(**overrides):
    """A client profile = the defaults plus the client's own settings."""
    profile = dict(DEFAULT_PROFILE)
    profile.update(overrides)
    return profile


DEFAULT_PROFILE = dict(
    name="default",
    output_name_template="{prefix} GSTR1 SYSTEM DATA {month}-{year}.xlsx",
    output_prefix=None,
    drop_blank_between_title_and_header=False,
    header_text_fixes={},
    header_wrap="all",
    center_data=("bill_no", "date", "state", "sgst_pct", "cgst_pct", "igst_pct"),
    wrap_data=(),
    money_data=MONEY_COLUMNS,
    money_total=MONEY_COLUMNS,
    center_zero_percent=False,
    mark_registered_section=False,
    reg_suffix=" REG :",
    other_suffix=" : ",
    grand_suffix=" : ",
    combine_same_rate=False,
    combine_columns=MONEY_COLUMNS,
    combine_position="right_after_total",
    party_name_fixes={},
    sections_without_underline=(),
    column_widths={
        "A": 17.85546875, "B": 30.7109375, "C": 8.7109375, "D": 10.7109375,
        "E": 10.7109375, "F": 10.5703125, "G": 5.7109375, "H": 10.7109375,
        "I": 5.7109375, "J": 9.5703125, "K": 5.7109375, "L": 7.7109375,
        "M": 10.7109375,
    },
    page_scale=95,
    margins=dict(NARROW_MARGINS),
)

CLIENT_PROFILES = {
    # 06- BHAVNA DEVENDRA BACHKANIWALA  (BDB GSTR1 SYSTEM DATA JULY-2026.xlsx)
    "BDB": _profile(
        name="BDB",
        output_prefix="BDB",
        output_name_template="{prefix} GSTR1 SYSTEM DATA {month}-{year}.xlsx",
        reg_suffix=" REG :",
        other_suffix=" : ",
        grand_suffix=" : ",
        column_widths={
            "A": 17.85546875, "B": 30.7109375, "C": 8.7109375, "D": 10.7109375,
            "E": 10.7109375, "F": 10.5703125, "G": 5.7109375, "H": 10.7109375,
            "I": 5.7109375, "J": 9.5703125, "K": 5.7109375, "L": 7.7109375,
            "M": 10.7109375,
        },
        page_scale=95,
    ),

    # 01- DEVENDRA PANNALAL BACHKANIWALA  (DPB GSTR1 SYSTEM DATA JULY 2026.xlsx)
    "DPB": _profile(
        name="DPB",
        output_prefix="DPB",
        output_name_template="{prefix} GSTR1 SYSTEM DATA {month} {year}.xlsx",
        header_text_fixes={"bill_no": "Bill No", "bill_amt": "Bill Amt"},
        header_wrap=("sgst_pct", "sgst_amt", "cgst_pct", "cgst_amt",
                     "igst_pct", "igst_amt"),
        center_data=("bill_no", "date", "state", "sgst_pct", "cgst_pct", "igst_pct"),
        money_data=MONEY_COLUMNS,
        money_total=MONEY_COLUMNS,
        center_zero_percent=True,
        reg_suffix=" REG : ",
        other_suffix=" : ",
        grand_suffix=" : ",
        combine_same_rate=True,
        combine_columns=("bill_amt", "sgst_amt", "cgst_amt", "taxable"),
        combine_position="before_next_section",
        party_name_fixes={           # " - RENT" was written in by hand on this row
            "PIYUSH JAGDISHBHAI ZALAWADIYA": "PIYUSH JAGDISHBHAI ZALAWADIYA - RENT",
        },
        column_widths={
            "A": 18.28515625, "B": 37.140625, "C": 9.140625, "D": 9.85546875,
            "E": 10.7109375, "F": 10.140625, "G": 6.0, "H": 9.42578125,
            "I": 6.28515625, "J": 9.5703125, "K": 5.5703125, "L": 7.140625,
            "M": 12.0,
        },
        page_scale=90,
    ),

    # 05-HIMSON KNITTING IND. P.LTD.  (HKI GSTR1 SYSTEM DATA JULY-2026.xlsx)
    "HKI": _profile(
        name="HKI",
        output_prefix="HKI",
        output_name_template="{prefix} GSTR1 SYSTEM DATA {month}-{year}.xlsx",
        header_text_fixes={"bill_no": "Bill No", "bill_amt": "Bill Amt"},
        header_wrap=("sgst_pct", "cgst_pct", "igst_pct", "taxable"),
        center_data=("bill_no", "date", "sgst_pct", "cgst_pct", "igst_pct"),
        wrap_data=("sgst_pct", "cgst_pct"),
        money_data=(),                      # amounts left as Tally exported them
        money_total=MONEY_COLUMNS,
        center_zero_percent=True,
        reg_suffix=" : ",                   # this client does not use "REG"
        other_suffix=" : ",
        grand_suffix=" : ",
        column_widths={
            "A": 17.7109375, "B": 37.7109375, "C": 8.7109375, "D": 8.85546875,
            "E": 10.7109375, "F": 10.5703125, "G": 5.7109375, "H": 9.42578125,
            "I": 5.7109375, "J": 9.5703125, "K": 5.7109375, "L": 9.0,
            "M": 10.7109375,
        },
        page_scale=90,
        margins=dict(NARROW_MARGINS, bottom=0.1968503937007874),
    ),

    # 01-HIMSON TEXTILES ENGG IND PVT LTD  (HTEI ... JULY 2026.xlsx)
    "HTE": _profile(
        name="HTE",
        output_prefix="HTEI",
        output_name_template="{prefix} GSTR1 SYSTEM DATA {month} {year}.xlsx",
        drop_blank_between_title_and_header=True,
        header_text_fixes={"bill_no": "Bill No.", "bill_amt": "Bill Amt"},
        header_wrap=("sgst_pct", "sgst_amt", "cgst_pct", "cgst_amt",
                     "igst_pct", "igst_amt"),
        center_data=("bill_no", "date", "sgst_pct", "cgst_pct", "igst_pct"),
        money_data=("bill_amt", "sgst_amt", "cgst_amt", "taxable"),
        money_total=("bill_amt", "sgst_amt", "cgst_amt", "taxable"),
        center_zero_percent=True,
        mark_registered_section=True,
        reg_suffix=" REG : ",
        other_suffix=" : ",
        grand_suffix=" : ",
        combine_same_rate=True,
        combine_columns=("bill_amt", "sgst_amt", "cgst_amt", "taxable"),
        combine_position="right_after_total",
        sections_without_underline=(    # this heading is bold but not underlined
            "GUR - GST 18% UN REGI",
        ),
        column_widths={
            "A": 18.28515625, "B": 36.5703125, "C": 9.85546875, "D": 9.7109375,
            "E": 10.42578125, "F": 10.42578125, "G": 5.7109375, "H": 9.7109375,
            "I": 5.7109375, "J": 9.7109375, "K": 5.7109375, "L": 6.28515625,
            "M": 11.42578125,
        },
        page_scale=90,
        margins=dict(NARROW_MARGINS, top=0.3937007874015748,
                     bottom=0.1968503937007874),
    ),
}


def profile_for(path, forced=None):
    """Pick the client profile: --profile wins, else the longest name match."""
    if forced:
        key = forced.strip().upper()
        if key in CLIENT_PROFILES:
            return CLIENT_PROFILES[key]
        matches = [k for k in CLIENT_PROFILES if k.startswith(key)]
        if len(matches) == 1:
            return CLIENT_PROFILES[matches[0]]
        raise SystemExit(f"ERROR: unknown profile '{forced}'. "
                         f"Known: {', '.join(sorted(CLIENT_PROFILES))}")
    stem = path.stem.upper()
    matches = [k for k in CLIENT_PROFILES if stem.startswith(k)]
    if matches:
        best = max(matches, key=len)
        return CLIENT_PROFILES[best]
    print(f"    note: no client profile matches '{path.name}' "
          f"(known: {', '.join(sorted(CLIENT_PROFILES))}) - using the default")
    return DEFAULT_PROFILE


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


def number(value):
    return value if isinstance(value, (int, float)) else 0


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
        if hits >= 3 and hits > best_hits:      # a real heading row matches several
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
    """'blank' | 'total' | 'data' | 'section' | 'unknown'."""
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
            month_no = int(match.group(2))
            year_no = int(match.group(3)) + (2000 if len(match.group(3)) == 2 else 0)
            if 1 <= month_no <= 12:
                return calendar.month_name[month_no].upper(), year_no
    return None, None


# ===========================================================================
#  the layout of one sheet: sections, their invoices and their totals
# ===========================================================================

def section_text(label):
    """'GST - GST 18 %' -> 'GST 18 %'   (drop the leading section code)."""
    parts = re.split(r"\s*-\s*", label.strip(), maxsplit=1)
    return (parts[1] if len(parts) == 2 else parts[0]).strip()


def build_layout(ws, columns, header_row, last_col):
    """Read the sheet into (rows, sections, totals, grand_total, loose rows)."""
    sections, totals, loose = [], [], []
    current = None
    rows = []
    for row in range(header_row + 1, ws.max_row + 1):
        kind = classify(ws, row, columns, last_col)
        rows.append((row, kind))
        if kind == "section":
            for col in range(1, last_col + 1):
                text = cell_text(ws.cell(row, col).value)
                if text and RE_SECTION_LABEL.match(text):
                    current = {
                        "row": row, "col": col, "label": text,
                        "raw_label": text, "data_rows": [], "total_row": None,
                        "tax_free": bool(RE_TAX_FREE.search(text)),
                        "unregistered": bool(RE_UNREGISTERED.search(text)),
                    }
                    sections.append(current)
                    break
        elif kind == "data":
            if current is not None:
                current["data_rows"].append(row)
            else:
                loose.append(row)
        elif kind == "total":
            totals.append(row)
            if current is not None and current["total_row"] is None:
                current["total_row"] = row

    grand_total = None
    for row in totals:
        for col in range(1, last_col + 1):
            text = cell_text(ws.cell(row, col).value)
            if text and RE_GRAND_TOTAL.match(text):
                grand_total = row
                break

    return dict(rows=rows, sections=sections, totals=totals,
                grand_total=grand_total, loose=loose)


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


# ===========================================================================
#  optional restructuring steps
# ===========================================================================

def drop_blank_rows_above_header(ws, header_row, stats):
    """Remove the empty spacer row(s) between the titles and the heading row."""
    to_delete = []
    row = header_row - 1
    while row >= 1:
        values = [ws.cell(row, c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None and not (isinstance(v, str) and not v.strip())
               for v in values):
            break                        # reached the last title line
        to_delete.append(row)
        row -= 1
    for row in reversed(to_delete):
        ws.delete_rows(row, 1)
    if to_delete:
        stats["dropped_rows"] = to_delete
        header_row -= len(to_delete)
    return header_row


RE_GST_RATE = re.compile(r"(\d+(?:\.\d+)?)\s*%")


def rate_of(label):
    """'GST - GST 18 %' -> '18';  None when the heading has no rate."""
    match = RE_GST_RATE.search(label)
    return match.group(1) if match else None


def combined_rows_to_add(layout, profile):
    """Which combined 'TOTAL FOR GST <rate> %' rows this sheet needs.

    Whenever one GST rate is used by more than one section (for example a
    registered and an un-registered section at 18 %), the clients' samples add a
    row that adds those sections up.  Each rate group gets its own row, so a
    month with 5 % and 18 % sections stays correct.
    """
    groups = {}
    for section in layout["sections"]:
        if section["tax_free"] or not section["data_rows"] or not section["total_row"]:
            continue
        rate = rate_of(section["label"])
        if rate is None:
            continue
        groups.setdefault(rate, []).append(section)

    todo = []
    for rate, sections in groups.items():
        if len(sections) < 2:
            continue
        last = sections[-1]
        if profile["combine_position"] == "before_next_section":
            following = [s for s in layout["sections"] if s["row"] > last["total_row"]]
            insert_at = min(s["row"] for s in following) if following \
                else last["total_row"] + 1
        else:
            insert_at = last["total_row"] + 1
        todo.append(dict(rate=rate, sections=sections, insert_at=insert_at))
    return todo


def write_combined_totals(ws, layout, columns, profile, stats):
    """Insert the combined rows (bottom-up, so the row numbers stay valid)."""
    todo = combined_rows_to_add(layout, profile)
    written = []
    for item in sorted(todo, key=lambda i: -i["insert_at"]):
        sections, insert_at = item["sections"], item["insert_at"]
        ws.insert_rows(insert_at, 1)

        registered = [s for s in sections if not s["unregistered"]]
        label_text = "TOTAL FOR " + total_text((registered or sections)[0]) \
                     + profile["other_suffix"]
        label_col = columns.get("party") or columns.get("bill_no") or 2
        ws.cell(insert_at, label_col).value = label_text

        formulas = []
        for name in profile["combine_columns"]:
            col = columns.get(name)
            if col is None:
                continue
            addresses = [f"{get_column_letter(col)}{s['total_row']}"
                         for s in sections if s["total_row"]]
            if addresses:
                formula = "=" + "+".join(addresses)
                ws.cell(insert_at, col).value = formula
                formulas.append(f"{get_column_letter(col)}{insert_at}{formula}")
        written.append(dict(row=insert_at, label=label_text, formulas=formulas,
                            rate=item["rate"],
                            sections=[s["label"] for s in sections]))
    stats["combined"] = written
    return written


# ===========================================================================
#  styling
# ===========================================================================

def style_titles(ws, header_row):
    """Bold the company-name line(s) above the table."""
    if TITLE_BOLD_ROWS <= 0:
        return
    done = 0
    for row in range(1, header_row):
        cells = [ws.cell(row, c) for c in range(1, ws.max_column + 1)]
        if not any(c.value not in (None, "") for c in cells):
            continue
        done += 1
        if done > TITLE_BOLD_ROWS:
            break
        for cell in cells:
            if cell.value not in (None, ""):
                set_bold(cell)


def apply_header_fixes(ws, header_row, columns, profile):
    """Correct the heading text the way the client's sample has it."""
    for name, text in profile["header_text_fixes"].items():
        col = columns.get(name)
        if col:
            ws.cell(header_row, col).value = text


def style_header_row(ws, header_row, columns, last_col, profile):
    wrap = profile["header_wrap"]
    wrap_columns = None if wrap == "all" else {
        columns.get(name) for name in wrap if columns.get(name)}
    for col in range(1, last_col + 1):
        cell = ws.cell(header_row, col)
        set_bold(cell)
        set_alignment(cell, horizontal="center",
                      wrap_text=True if (wrap_columns is None
                                         or col in wrap_columns) else None)
    ws.row_dimensions[header_row].height = HEADER_ROW_HEIGHT


def style_data_row(ws, row, columns, profile, center_percents):
    for name in profile["money_data"]:
        col = columns.get(name)
        if col:
            ws.cell(row, col).number_format = MONEY_FORMAT
    for name in profile["center_data"]:
        if name in PERCENT_COLUMNS and not center_percents:
            continue
        col = columns.get(name)
        if col:
            set_alignment(ws.cell(row, col), horizontal="center")
    for name in profile["wrap_data"]:
        col = columns.get(name)
        if col:
            set_alignment(ws.cell(row, col), wrap_text=True)
    date_col = columns.get("date")
    if date_col:
        ws.cell(row, date_col).number_format = DATE_FORMAT


def style_total_row(ws, row, columns, last_col, profile, label=None):
    # 1. bold the whole row (empty cells included - that is what the samples have)
    for col in range(1, last_col + 1):
        set_bold(ws.cell(row, col))

    # 2. money cells keep 2 decimals and are right aligned
    for name in profile["money_total"]:
        col = columns.get(name)
        if col is None:
            continue
        cell = ws.cell(row, col)
        cell.number_format = MONEY_FORMAT
        if RIGHT_ALIGN_TOTAL_NUMBERS and not isinstance(cell.value, str):
            set_alignment(cell, horizontal="right")

    # 3. the label goes into the Party Name column, right aligned
    target_col = columns.get("party") or columns.get("bill_no") or 2
    if label is not None:
        for col in range(1, last_col + 1):
            text = cell_text(ws.cell(row, col).value)
            if text and RE_TOTAL_LABEL.match(text):
                if col != target_col:
                    ws.cell(row, col).value = None
                break
        ws.cell(row, target_col).value = label
    target = ws.cell(row, target_col)
    if target.value not in (None, ""):
        set_bold(target)
        if RIGHT_ALIGN_TOTAL_LABEL:
            set_alignment(target, horizontal="right")


def total_text(section):
    """The wording inside 'TOTAL FOR ...' for a section, without any suffix."""
    text = section_text(section.get("raw_label") or section["label"])
    if not (section["tax_free"] or section["unregistered"]):
        text = re.sub(r"\s*\bREG\b\s*$", "", text, flags=re.I)   # idempotent
    return text.strip()


def total_label(section, profile):
    """The 'TOTAL FOR ...' text of a section, the way that client writes it."""
    text = total_text(section)
    if section["tax_free"] or section["unregistered"]:
        suffix = profile["other_suffix"]
    else:
        suffix = profile["reg_suffix"]
    return "TOTAL FOR " + text + suffix


def apply_sheet_geometry(ws, profile):
    for letter, width in profile["column_widths"].items():
        ws.column_dimensions[letter].width = width
    ws.sheet_format.defaultRowHeight = DEFAULT_ROW_HEIGHT
    ws.page_setup.orientation = PAGE_ORIENTATION
    ws.page_setup.paperSize = PAGE_PAPER_SIZE
    ws.page_setup.scale = profile["page_scale"]
    for key, value in profile["margins"].items():
        setattr(ws.page_margins, key, value)


# ===========================================================================
#  one sheet, end to end
# ===========================================================================

def format_worksheet(ws, profile, verbose=True):
    stats = {"invoice_rows": 0, "total_rows": 0, "sections": [], "relabelled": [],
             "unknown_rows": [], "combined": [], "dropped_rows": [],
             "party_fixed": [], "ok": False}

    header_row = find_header_row(ws)
    if header_row is None:
        if verbose:
            print(f"    sheet '{ws.title}': no column heading row found - skipped")
        return stats

    if profile["drop_blank_between_title_and_header"]:
        header_row = drop_blank_rows_above_header(ws, header_row, stats)

    columns = map_columns(ws, header_row)
    last_col = max((c.column for c in ws[header_row] if c.value not in (None, "")),
                   default=ws.max_column)
    apply_header_fixes(ws, header_row, columns, profile)

    # ---- read the layout, add the combined GST row when the client needs it ----
    layout = build_layout(ws, columns, header_row, last_col)
    combine_allowed = not any(isinstance(ws.cell(r, c).value, str) and
                              str(ws.cell(r, c).value).startswith("=")
                              for r in range(1, ws.max_row + 1)
                              for c in range(1, last_col + 1))
    if profile["combine_same_rate"] and combine_allowed:
        write_combined_totals(ws, layout, columns, profile, stats)
        layout = build_layout(ws, columns, header_row, last_col)

    # ---- titles and heading row ----
    style_titles(ws, header_row)
    style_header_row(ws, header_row, columns, last_col, profile)

    # ---- section headings (and the " REG" mark one client writes) ----
    for section in layout["sections"]:
        cell = ws.cell(section["row"], section["col"])
        plain = section["label"] in profile["sections_without_underline"]
        set_bold(cell, underline=None if plain else "single")
        if (profile["mark_registered_section"] and not section["tax_free"]
                and not section["unregistered"]
                and not re.search(r"\bREG\b", section["label"], re.I)):
            section["label"] = section["label"].rstrip() + " REG"
            cell.value = section["label"]
        section["stats"] = dict(invoices=len(section["data_rows"]),
                                label=section["label"],
                                total_label=total_label(section, profile))
        stats["sections"].append(section["stats"])

    # ---- corrections that were made by hand in the client's file ----
    fixes = profile.get("party_name_fixes") or {}
    party_col = columns.get("party")
    fixed = []
    if fixes and party_col:
        for row, kind in layout["rows"]:
            if kind != "data":
                continue
            text = cell_text(ws.cell(row, party_col).value)
            if text in fixes:
                ws.cell(row, party_col).value = fixes[text]
                fixed.append((row, text, fixes[text]))
    stats["party_fixed"] = fixed

    # ---- invoice lines ----
    section_of_row = {row: s for s in layout["sections"] for row in s["data_rows"]}
    loose_center = (profile["center_zero_percent"]
                    or carries_tax(ws, layout["loose"], columns))
    for row, kind in layout["rows"]:
        if kind != "data":
            continue
        stats["invoice_rows"] += 1
        section = section_of_row.get(row)
        if section is None:
            center = loose_center
        elif profile["center_zero_percent"]:
            center = True
        else:
            center = carries_tax(ws, section["data_rows"], columns)
        style_data_row(ws, row, columns, profile, center)

    # ---- total rows ----
    combined_rows = {c["row"] for c in stats["combined"]}
    for row, kind in layout["rows"]:
        if kind != "total":
            continue
        stats["total_rows"] += 1
        if row in combined_rows:
            style_total_row(ws, row, columns, last_col, profile)      # keep its label
            continue
        label = None
        if row == layout["grand_total"]:
            label = "GRAND TOTAL" + profile["grand_suffix"]
        else:
            for section in layout["sections"]:
                if section["total_row"] == row:
                    label = total_label(section, profile)
                    break
        if label:
            before = [ws.cell(row, c).value for c in range(1, last_col + 1)]
            style_total_row(ws, row, columns, last_col, profile, label=label)
            stats["relabelled"].append((row, label, before))

    # ---- sheet geometry ----
    apply_sheet_geometry(ws, profile)
    stats["ok"] = True
    stats["header_row"] = header_row
    stats["columns"] = columns
    stats["layout"] = layout
    return stats


# ===========================================================================
#  output file name
# ===========================================================================

def prefix_from_filename(path):
    """BDBJUL.xlsx -> 'BDB', BDB_JULY_2026.xlsx -> 'BDB', XYZJAN26.xlsx -> 'XYZ'."""
    stem = path.stem.strip()
    base = re.sub(r"[\s_\-]*\d{2,4}$", "", stem).strip(" _-")     # drop a trailing year
    match = RE_MONTH_TAIL.match(base)
    prefix = match.group("prefix") if match else stem
    prefix = re.sub(r"[_\s]+", " ", prefix).strip(" _-")
    if "GSTR1" in prefix.upper():        # already an output name -> keep the first word
        prefix = prefix.split()[0]
    return prefix or DEFAULT_PREFIX


def build_output_path(src, ws, profile, outdir=None):
    month, year = find_register_period(ws)
    prefix = profile["output_prefix"] or prefix_from_filename(src)
    if month is None:
        print("    WARNING: could not read the 'Sales Register From ... TO ...' line;"
              " using the file name for the output name")
        name = f"{prefix} GSTR1 SYSTEM DATA {src.stem}.xlsx"
    else:
        name = profile["output_name_template"].format(prefix=prefix, month=month,
                                                      year=year)
    folder = Path(outdir) if outdir else src.parent
    return folder / name


# ===========================================================================
#  reporting / verification
# ===========================================================================

def report(stats, profile):
    print(f"    profile ............... {profile['name']}")
    print(f"    heading row ........... {stats['header_row']}")
    print(f"    invoice lines ......... {stats['invoice_rows']}")
    print(f"    total rows ............ {stats['total_rows']}")
    if stats["dropped_rows"]:
        print(f"    removed blank row(s) .. {stats['dropped_rows']}")
    for section in stats["sections"]:
        print(f"    section {section['label']!r} -> {section['invoices']} invoice(s),"
              f" total {section['total_label']!r}")
    for info in stats["combined"]:
        print(f"    combined {info['rate']}% GST row .. row {info['row']}: "
              f"{info['label']!r} ({'; '.join(info['formulas'])})")
        print(f"                            sums up {info['sections']}")
    for row, label, before in stats["relabelled"]:
        old = next((v for v in before if isinstance(v, str)), None)
        print(f"    label row {row}: {old!r} -> {label!r}")
    for row, old, new in stats.get("party_fixed", []):
        print(f"    party name row {row}: {old!r} -> {new!r}")
    if stats["unknown_rows"]:
        print(f"    WARNING: unrecognised row(s) left untouched: {stats['unknown_rows']}")


def find_reference(src, out_path, ws, profile):
    """Look for this client's hand-formatted file next to the input."""
    month, year = find_register_period(ws)
    prefix = (profile["output_prefix"] or prefix_from_filename(src)).upper()
    for candidate in sorted(src.parent.glob("*GSTR1*.xlsx")):
        if candidate.resolve() in (out_path.resolve(), src.resolve()):
            continue
        if not candidate.stem.upper().startswith(prefix):
            continue
        try:
            ref_ws = load_workbook(candidate, read_only=True).active
            ref_month, ref_year = find_register_period(ref_ws)
        except Exception:
            continue
        if (ref_month, ref_year) == (month, year):
            return candidate
    return None


def verify_against(reference, generated):
    """Compare a generated file with the hand-formatted one (100% check)."""
    try:
        from compare_workbooks import compare_files
    except ImportError:
        print("    (compare_workbooks.py not found next to this script - "
              "verification skipped)")
        return None
    result = compare_files(reference, generated)
    print(f"    verify vs {reference.name}: ", end="")
    if result["differences"]:
        print(f"FAIL - {len(result['differences'])} difference(s)")
        for line in result["differences"][:20]:
            print(f"        {line}")
    else:
        print(f"IDENTICAL ({result['cosmetic']} invisible cosmetic note(s))")
    return result


# ===========================================================================
#  main
# ===========================================================================

def format_file(src, profile, out_path, sheet_name=None, dry_run=False,
                reference=None, verify=False):
    wb = load_workbook(src)
    sheets = [wb[sheet_name]] if sheet_name else wb.worksheets

    ok = False
    for ws in sheets:
        print(f"    sheet '{ws.title}':")
        stats = format_worksheet(ws, profile)
        if stats["ok"]:
            report(stats, profile)
            ok = True
        else:
            print("      nothing formatted")

    if not ok:
        print("    ERROR: no sheet with a recognisable Sales Register heading row "
              "- is this the right file?")
        return None, None

    if dry_run:
        print(f"    [dry run] would write {out_path}")
        return None, None

    out_path.parent.mkdir(parents=True, exist_ok=True)
    backup = out_path.with_suffix(out_path.suffix + ".bak")
    if out_path.exists() and not backup.exists():
        shutil.copy2(out_path, backup)
        print(f"    existing {out_path.name} backed up as {backup.name}")

    wb.save(out_path)
    print(f"    saved -> {out_path}")

    result = None
    if verify:
        if reference is None:
            reference = find_reference(src, out_path, wb.active, profile)
        if reference is None:
            print("    verify: no hand-formatted file with the same month found - "
                  "pass it explicitly with --verify FILE")
        else:
            result = verify_against(reference, out_path)
    return out_path, result


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Apply every client's monthly GST-R1 sales-register formatting "
                    "to raw Tally exports (openpyxl, regex driven).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="examples:\n"
               "  python3 format_gstr1.py BDBJUL.xlsx\n"
               "  python3 format_gstr1.py BDBJUL.xlsx DPBJUL.xlsx HKIJUL.xlsx "
               "HTEJUL.xlsx\n"
               "  python3 format_gstr1.py BDBAUG.xlsx --verify\n"
               "  python3 format_gstr1.py FOOJUL.xlsx --profile DPB\n",
    )
    parser.add_argument("inputs", nargs="*", help="raw Tally sales registers")
    parser.add_argument("-o", "--output", help="output file (single input only)")
    parser.add_argument("--outdir", help="folder for the auto-named output files")
    parser.add_argument("--profile", help="force a client profile (BDB, DPB, HKI, HTE)")
    parser.add_argument("--sheet", help="only format this sheet")
    parser.add_argument("--dry-run", action="store_true",
                        help="report what would be done, write nothing")
    parser.add_argument("--verify", nargs="?", const="", default=None,
                        metavar="REFERENCE",
                        help="compare the result with the hand-formatted file "
                             "(auto-found next to the input when no file is given)")
    parser.add_argument("--inplace", action="store_true",
                        help="format the inputs themselves (a .bak is kept)")
    parser.add_argument("--list-profiles", action="store_true",
                        help="show the client profiles and stop")
    args = parser.parse_args(argv)

    if args.list_profiles:
        for key, profile in sorted(CLIENT_PROFILES.items()):
            print(f"{key:<5} -> output '{profile['output_name_template']}' "
                  f"(prefix {profile['output_prefix']}), scale {profile['page_scale']}%")
        return 0

    if not args.inputs:
        parser.error("give at least one input file (or --list-profiles)")

    failures = 0
    for name in args.inputs:
        src = Path(name).expanduser()
        if not src.is_file():
            print(f"ERROR: file not found: {src}")
            failures += 1
            continue
        if src.suffix.lower() not in (".xlsx", ".xlsm"):
            print(f"ERROR: {src.name}: this needs an .xlsx/.xlsm file "
                  "(Tally 'Excel' export)")
            failures += 1
            continue

        profile = profile_for(src, args.profile)
        print(f"formatting {src.name}")

        wb = load_workbook(src)
        first = wb[args.sheet] if args.sheet else wb.worksheets[0]

        if args.output:
            out_path = Path(args.output).expanduser()
            if out_path.parent == Path(".") and not out_path.is_absolute():
                out_path = src.parent / out_path
        elif args.inplace:
            out_path = src
        else:
            out_path = build_output_path(src, first, profile, outdir=args.outdir)

        if out_path.resolve() == src.resolve() and not args.inplace:
            print(f"    skipped: {src.name} is itself a formatted file - the result "
                  f"would overwrite it.\n             (use --inplace if that is what "
                  f"you want, or give the raw export instead)")
            continue

        reference = args.verify if isinstance(args.verify, str) and args.verify else None
        _, result = format_file(src, profile, out_path, sheet_name=args.sheet,
                                dry_run=args.dry_run, reference=reference,
                                verify=args.verify is not None)
        if result and result["differences"]:
            failures += 1

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
