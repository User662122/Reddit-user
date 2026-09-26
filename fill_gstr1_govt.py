#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_gstr1_govt.py - fill the government GSTR-1 workbook with a month's data.

Third step of the monthly routine:

    1. Tally "Sales Register" export        (BDBJUL.xlsx, HTEJUL.xlsx, ...)
    2. format_gstr1.py      -> ".. GSTR1 SYSTEM DATA JULY-2026.xlsx"
    3. fill_gstr1_govt.py   -> ".. GSTR1 JULY-2026.xlsx"      <- this script

Step 3 does exactly what is done by hand every month: take the previous month's
GSTR-1 file (the one that was uploaded to the GST portal), delete the previous
month's data out of it and enter the new month's data in every worksheet.
Everything else is kept as it is - the 23 worksheets of the portal's offline
utility, the frozen title / summary / heading rows, every number format, every
data validation and every column width.

Usage
-----
    python3 fill_gstr1_govt.py "out/BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" \
            --template "BDB GSTR1 JUNE-2026.xlsx"

        -> writes "govt_out/BDB GSTR1 JULY-2026.xlsx" (name from the template)
           and compares it with the hand-made file next to it

    python3 fill_gstr1_govt.py --all                  # BDB, DPB, HKI and HTEI
    python3 fill_gstr1_govt.py --all --outdir .       # write next to the files

Options
-------
      --template FILE     previous month's government file (what the new month
                          is built from)
      -o, --output FILE   exact output path (single input only)
      --outdir DIR        where to write (default: govt_out).  The name of the
                          generated file is the template's name with the new
                          month, so it can collide with a hand-made file that is
                          already there - hence the separate folder; an existing
                          file is always backed up as .bak first.
      --all               do all four clients in one run (uses --arranged-dir
                          and --govt-dir to find the files)
      --verify [FILE]     compare with the hand-made file (default: on; the
                          file is looked up when no name is given)
      --no-verify         skip the comparison
      --dry-run           report what would be written, change nothing
      --sst-as-is         never touch sharedStrings.xml (text that is new is
                          written as an inline string instead)

Where the data goes (nothing is bound to a row number or to the number of
invoices - sections, invoice lines and series are found with regular
expressions, so a month with 3 invoices or 300 fills the same)
----------------------------------------------------------------------------
The arranged file is read invoice by invoice and every invoice is put in one of
three groups, exactly the way the columns are filled in by hand:

    registered     column A (Gst No) filled            -> b2b,sez,de, hsn(b2b)
    un-registered  no Gst No, but tax amounts          -> b2cs, hsn(b2c)
    tax free       no Gst No and no tax amounts        -> exemp, hsn(b2b)

    b2b,sez,de   one row per registered invoice
    b2cs         one row per un-registered invoice
    exemp        the tax free total, under the description already in the sheet
    hsn(b2b)     registered HSN summary, plus the tax free row when there is one
    hsn(b2c)     un-registered HSN summary
    docs         one row per invoice-number series ('' for 29-37, 'AKR/', 'RES/')

The other 17 worksheets carry no data for these clients and stay untouched.
The summary formulas of row 3 (no. of recipients, no. of invoices, taxable
value, ...) keep their formulas; only the cached result is refreshed, so the
totals are correct even before Excel recalculates.

Rows that held data last month but have none this month are cleared the way
they are cleared by hand: the values are deleted, and in the docs worksheet the
series rows stay behind with the digits stripped ('AST/01' -> 'AST/') and a
zero count.  Lines that never held a number - HTEI's 'WIND MILL' HSN line,
HKI's empty second docs line, BDB's zeroed hsn(b2c) line - are deliberately
left alone, because that is what the hand-made files do.
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import shutil
import sys
import zipfile
from collections import Counter, OrderedDict
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

# ---------------------------------------------------------------------------
#  what this workbook is made of
# ---------------------------------------------------------------------------

B2B, B2CS, EXEMP = "b2b,sez,de", "b2cs", "exemp"
HSN_B2B, HSN_B2C, DOCS = "hsn(b2b)", "hsn(b2c)", "docs"

#: worksheets that carry data, in the order they are reported
DATA_SHEETS = [B2B, B2CS, EXEMP, HSN_B2B, HSN_B2C, DOCS]

#: columns that hold this month's data
SHEET_COLUMNS = {
    B2B: "ABCDEFGHIJKLM",
    B2CS: "ABCDEFGH",
    EXEMP: "ABCD",
    HSN_B2B: "ABCDEFGHIJK",
    HSN_B2C: "ABCDEFGHIJK",
    DOCS: "ABCDE",
}

#: columns emptied on a row that is not used any more ("old data deleted")
CLEAR_COLUMNS = {
    B2B: "ABCDEFGHIJKLM",
    B2CS: "ABCDEFGH",
    EXEMP: "ABCD",
    HSN_B2B: "ABCDEFGHIJ",           # the Cess column is left as it is
    HSN_B2C: "ABCDEFGHIJ",
    DOCS: "E",                       # B, C and D are handled apart
}

#: labels typed into a worksheet when the row is brand new
SHEET_LABELS = {
    B2B: {"G": "N", "I": "Regular B2B"},
    B2CS: {"A": "OE"},
    EXEMP: {"A": "RESIDENTIAL RENT"},
    HSN_B2C: {"A": "997212", "B": "RENT INCOME"},
    DOCS: {"A": "Invoices for outward supply"},
}
HSN_B2B_LABELS = [{"A": "997212", "B": "RENT INCOME"},
                  {"A": "997212", "B": "RENT INCOME RESIDENTIAL"}]

HEADER_PATTERNS = OrderedDict([
    ("gst", r"gst\s*(no|in)"),
    ("party", r"party|particular"),
    ("bill", r"bill\s*\.?\s*(no|number)"),
    ("date", r"date"),
    ("amount", r"bill\s*amt|invoice\s*value|bill\s*amount"),
    ("state", r"state|place\s*of\s*supply"),
    ("sgst_pct", r"sgst\s*%"),
    ("sgst", r"sgst\s*amt"),
    ("cgst_pct", r"cgst\s*%"),
    ("cgst", r"cgst\s*amt"),
    ("igst_pct", r"igst\s*%"),
    ("igst", r"igst\s*amt"),
    ("taxable", r"taxable"),
])

MONTHS = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
          "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]

MONTH_RE = re.compile(r"(JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|"
                      r"JUN(?:E)?|JUL(?:Y)?|AUG(?:UST)?|SEP(?:TEMBER)?|"
                      r"OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?)", re.I)

ELEMENT_REF = re.compile(r'r="([A-Z]{1,3})(\d+)"')
ROW_RE = re.compile(r'<row\b[^>]*?r="(\d+)"[^>]*?>.*?</row>', re.S)
CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
VALUE_RE = re.compile(r"<v>(.*?)</v>", re.S)
FORMULA_RE = re.compile(r"<f\b[^>]*?(?:/>|>(.*?)</f>)", re.S)

XL_EPOCH = dt.date(1899, 12, 30)          # Excel day 1, 1900 date system


class Fill:
    """Value typed in only when the cell is still empty (a column label)."""

    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


# ---------------------------------------------------------------------------
#  helpers
# ---------------------------------------------------------------------------

def is_empty(value):
    return value is None or (isinstance(value, str) and not value.strip())


def as_number(value):
    """int/float for anything numeric, else None."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if re.fullmatch(r"-?\d+", text):
            return int(text)
        if re.fullmatch(r"-?\d*\.\d+", text):
            return float(text)
    return None


def num_text(value):
    """A number written the way Excel writes it into <v>...</v>."""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    return str(value)


def unescape(text):
    return (text.replace("&lt;", "<").replace("&gt;", ">")
            .replace("&quot;", '"').replace("&apos;", "'")
            .replace("&amp;", "&"))


def to_serial(value):
    """datetime / date / 'dd-mm-yy' text -> Excel serial number."""
    if isinstance(value, dt.datetime):
        value = value.date()
    if isinstance(value, dt.date):
        return (value - XL_EPOCH).days
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        text = value.strip()
        match = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
        if match:
            year, month, day = (int(p) for p in match.groups())
        else:
            match = re.fullmatch(r"(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", text)
            if not match:
                return None
            day, month, year = (int(p) for p in match.groups())
            year += 2000 if year < 100 else 0
        try:
            return (dt.date(year, month, day) - XL_EPOCH).days
        except ValueError:
            return None
    return None


def text_key(value):
    """COUNTIF compares text without case, numbers as numbers."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    return str(value).strip().lower()


def distinct_fp(values):
    """Excel's SUMPRODUCT((rng<>"")/COUNTIF(rng,rng&"")) - distinct count.

    The formula adds 1/count for every cell in row order, so the result can
    carry the same tiny floating point remainder Excel shows (three recipients
    with seven invoices on one of them -> 2.9999999999999996).
    """
    keys = [text_key(v) for v in values if not is_empty(v)]
    counts = Counter(keys)
    total = 0.0
    for key in keys:
        total += 1.0 / counts[key]
    return total


def leading_prefix(bill):
    """Invoice-number series of an invoice: 29 -> '', 'AKR/68' -> 'AKR/'."""
    return re.match(r"^\D*", str(bill)).group(0)


def series_prefix(bill):
    """Series of an invoice number already in a worksheet (None when empty)."""
    if is_empty(bill):
        return None
    text = str(bill).strip()
    if re.fullmatch(r"\d+(\.0+)?", text):
        return ""                     # the plain 29, 30, 31 ... series
    return leading_prefix(text)       # 'AKR/', 'RES/', 'AST/', ...


def series_sort_key(bill):
    text = str(bill).strip()
    if re.fullmatch(r"\d+(\.0+)?", text):
        return (0, float(text), "")
    return (1, 0.0, text)


def month_name(index):
    return MONTHS[index - 1]


def shift_month(month, year, delta):
    index = year * 12 + (month - 1) + delta
    return index % 12 + 1, index // 12


# ---------------------------------------------------------------------------
#  reading the arranged "GSTR1 SYSTEM DATA" file
# ---------------------------------------------------------------------------

def read_arranged(path, verbose=True):
    """<- registered / un-registered / tax free invoice lists + the period."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header_row, columns = None, {}
    for row in range(1, min(ws.max_row, 60) + 1):
        texts = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row, col).value
            if isinstance(value, str) and value.strip():
                texts[col] = re.sub(r"\s+", " ", value.strip().lower())
        if not texts:
            continue
        found, used = {}, set()
        for key, pattern in HEADER_PATTERNS.items():
            for col, text in texts.items():
                if col in used:
                    continue
                if re.search(pattern, text):
                    found[key] = col
                    used.add(col)
                    break
        if len(found) >= 8 and "bill" in found and "amount" in found:
            header_row, columns = row, found
            break
    if header_row is None:
        raise SystemExit(f"ERROR: no 'Gst No / Party / Bill' heading row in {path}")

    def value(row, key):
        col = columns.get(key)
        return ws.cell(row, col).value if col else None

    registered, unregistered, taxfree = [], [], []
    for row in range(header_row + 1, ws.max_row + 1):
        bill = value(row, "bill")
        amount = as_number(value(row, "amount"))
        party = value(row, "party")
        gst = value(row, "gst")
        if amount is None or (is_empty(bill) and is_empty(party)):
            continue
        label = f"{party if isinstance(party, str) else ''} " \
                f"{bill if isinstance(bill, str) else ''}"
        if re.search(r"total", label, re.I):
            continue
        invoice = dict(
            row=row,
            gst=str(gst).strip() if not is_empty(gst) else None,
            party=party,
            bill=bill,
            date=value(row, "date"),
            amount=amount,
            state=value(row, "state"),
            sgst_pct=as_number(value(row, "sgst_pct")) or 0,
            sgst=as_number(value(row, "sgst")) or 0,
            cgst_pct=as_number(value(row, "cgst_pct")) or 0,
            cgst=as_number(value(row, "cgst")) or 0,
            igst_pct=as_number(value(row, "igst_pct")) or 0,
            igst=as_number(value(row, "igst")) or 0,
            taxable=as_number(value(row, "taxable")),
        )
        invoice["rate"] = int(round(invoice["sgst_pct"] + invoice["cgst_pct"]
                                    + invoice["igst_pct"]))
        if invoice["taxable"] is None:
            invoice["taxable"] = invoice["amount"]
        if invoice["gst"]:
            registered.append(invoice)
        elif any(invoice[key] for key in ("sgst", "cgst", "igst",
                                          "sgst_pct", "cgst_pct", "igst_pct")):
            unregistered.append(invoice)
        else:
            taxfree.append(invoice)

    states = [inv["state"] for inv in registered if not is_empty(inv["state"])]
    if not states:
        states = [inv["state"] for inv in registered + unregistered + taxfree
                  if not is_empty(inv["state"])]
    default_state = Counter(states).most_common(1)[0][0] if states else None

    period = None
    for row in range(1, min(header_row + 2, 12) + 1):
        for col in range(1, ws.max_column + 1):
            text = ws.cell(row, col).value
            if not isinstance(text, str):
                continue
            match = re.search(r"from\s+(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})",
                              text, re.I)
            if match:
                day, month, year = (int(p) for p in match.groups())
                period = (month, year + 2000 if year < 100 else year)
    wb.close()

    if verbose:
        print(f"    arranged file ........ {Path(path).name}")
        print(f"    heading row .......... {header_row}")
        print("    columns .............. " + ", ".join(
            f"{get_column_letter(c)}={k}" for k, c in columns.items()))
        print(f"    registered invoices .. {len(registered)}")
        print(f"    un-registered ........ {len(unregistered)}")
        print(f"    tax free ............. {len(taxfree)}")
        print(f"    place of supply ...... {default_state!r}")
        if period:
            print(f"    period ............... {month_name(period[0])}-{period[1]}")

    return dict(path=Path(path), header_row=header_row, columns=columns,
                registered=registered, unregistered=unregistered,
                taxfree=taxfree, default_state=default_state, period=period)


# ---------------------------------------------------------------------------
#  the six worksheets that are filled
# ---------------------------------------------------------------------------

def group_totals(invoices):
    """Numbers shared by the HSN lines and the summary of a group."""
    def total(key):
        return sum(inv[key] for inv in invoices) or 0

    return dict(count=len(invoices), value=total("amount"),
                rate=int(round(sum(inv["rate"] for inv in invoices)
                               / len(invoices))) if invoices else 0,
                taxable=total("taxable"), igst=total("igst"),
                cgst=total("cgst"), sgst=total("sgst"))


def plan_b2b(data):
    """One row per registered invoice."""
    labels = {col: Fill(text) for col, text in SHEET_LABELS[B2B].items()}
    rows = []
    for inv in data["registered"]:
        row = {"A": inv["gst"], "B": inv["party"], "C": inv["bill"],
               "D": to_serial(inv["date"]), "E": inv["amount"],
               "F": inv["state"], "K": inv["rate"], "L": inv["taxable"]}
        row.update(labels)
        rows.append(row)
    return rows


def plan_b2cs(data):
    """One row per un-registered invoice."""
    labels = {col: Fill(text) for col, text in SHEET_LABELS[B2CS].items()}
    rows = []
    for inv in data["unregistered"]:
        row = {"B": inv["state"] or data["default_state"],
               "D": inv["rate"], "E": inv["taxable"]}
        row.update(labels)
        rows.append(row)
    return rows


def plan_exemp(data):
    """The tax free total, in the row that is already in the sheet."""
    if not data["taxfree"]:
        return []
    row = dict((col, Fill(text)) for col, text in SHEET_LABELS[EXEMP].items())
    row["C"] = sum(inv["taxable"] for inv in data["taxfree"]) or 0
    return [row]


def plan_hsn_b2b(data):
    """Row 5 = registered, row 6 = tax free (the slots of the worksheet)."""
    rows = []
    for index, invoices in enumerate((data["registered"], data["taxfree"])):
        if not invoices:
            rows.append({})
            continue
        totals = group_totals(invoices)
        row = {"D": totals["count"], "E": totals["value"],
               "F": 0 if index == 1 else totals["rate"], "G": totals["taxable"],
               "H": totals["igst"], "I": totals["cgst"], "J": totals["sgst"]}
        row.update({col: Fill(text)
                    for col, text in HSN_B2B_LABELS[index].items()})
        rows.append(row)
    while rows and not rows[-1]:
        rows.pop()
    return rows


def plan_hsn_b2c(data):
    """The un-registered HSN summary."""
    if not data["unregistered"]:
        return []
    totals = group_totals(data["unregistered"])
    row = {"D": totals["count"], "E": totals["value"], "F": totals["rate"],
           "G": totals["taxable"], "H": totals["igst"], "I": totals["cgst"],
           "J": totals["sgst"]}
    row.update({col: Fill(text) for col, text in SHEET_LABELS[HSN_B2C].items()})
    return [row]


def plan_docs(data):
    """One row per invoice-number series, in the order the series appear."""
    series = OrderedDict()
    for inv in data["registered"] + data["unregistered"] + data["taxfree"]:
        series.setdefault(leading_prefix(inv["bill"]), []).append(inv["bill"])
    rows = []
    for bills in series.values():
        ordered = sorted(bills, key=series_sort_key)
        row = dict((col, Fill(text)) for col, text in SHEET_LABELS[DOCS].items())
        row.update({"B": ordered[0], "C": ordered[-1], "D": len(ordered)})
        rows.append(row)
    return rows


SHEET_PLANNERS = {
    B2B: plan_b2b,
    B2CS: plan_b2cs,
    EXEMP: plan_exemp,
    HSN_B2B: plan_hsn_b2b,
    HSN_B2C: plan_hsn_b2c,
    DOCS: plan_docs,
}


def row3_summary(sheet, values):
    """Row 3 of `sheet`, recalculated from every value now in the sheet."""
    def column(letter):
        return [row.get(letter) for row in values.values()]

    def total(letter):
        numbers = [as_number(v) for v in column(letter)]
        return sum(n for n in numbers if n is not None) or 0

    if sheet == B2B:
        return {"A3": distinct_fp(column("A")), "C3": distinct_fp(column("C")),
                "E3": total("E"), "L3": total("L"), "M3": total("M")}
    if sheet == B2CS:
        return {"E3": total("E"), "F3": total("F")}
    if sheet == EXEMP:
        return {"B3": total("B"), "C3": total("C"), "D3": total("D")}
    if sheet in (HSN_B2B, HSN_B2C):
        return {"A3": distinct_fp(column("A")), "E3": total("E"),
                "G3": total("G"), "H3": total("H"), "I3": total("I"),
                "J3": total("J"), "K3": total("K")}
    if sheet == DOCS:
        return {"D3": total("D"), "E3": total("E")}
    raise KeyError(sheet)


# ---------------------------------------------------------------------------
#  shared strings
# ---------------------------------------------------------------------------

class SharedStrings:
    """Reads <sst> and appends the strings that are not in it yet."""

    def __init__(self, raw):
        self.raw = raw.decode("utf-8") if raw is not None else None
        self.index, self.values = {}, []
        self.count = self.unique = self.added = 0
        if self.raw is None:
            return
        head = re.search(r"<sst\b([^>]*)>", self.raw)
        if head:
            count = re.search(r'count="(\d+)"', head.group(1))
            unique = re.search(r'uniqueCount="(\d+)"', head.group(1))
            self.count = int(count.group(1)) if count else 0
            self.unique = int(unique.group(1)) if unique else 0
        for position, match in enumerate(re.finditer(r"<si>(.*?)</si>",
                                                     self.raw, re.S)):
            text = self.text_of(match.group(1))
            self.values.append(text)
            self.index.setdefault(text, position)

    @staticmethod
    def text_of(body):
        parts = re.findall(r"<t\b[^>]*>(.*?)</t>", body, re.S)
        return "".join(unescape(part) for part in parts)

    def index_of(self, text):
        """Index of `text`, appending it to the table when it is new."""
        if text in self.index:
            return self.index[text]
        if self.raw is None:
            return None
        body = f'<si><t xml:space="preserve">{escape(str(text))}</t></si>'
        self.raw = self.raw.replace("</sst>", body + "</sst>")
        self.index[text] = self.unique + self.added
        self.added += 1
        return self.index[text]

    def to_bytes(self, count=None):
        if self.raw is None:
            return None
        head = re.search(r"<sst\b([^>]*)>", self.raw)
        if head:
            attrs = re.sub(r'count="\d+"',
                           f'count="{self.count if count is None else count}"',
                           head.group(1))
            attrs = re.sub(r'uniqueCount="\d+"',
                           f'uniqueCount="{self.unique + self.added}"', attrs)
            self.raw = self.raw[:head.start(1)] + attrs + self.raw[head.end(1):]
        return self.raw.encode("utf-8")


# ---------------------------------------------------------------------------
#  editing one worksheet, value by value
# ---------------------------------------------------------------------------

class SheetEditor:
    """Changes cell values of a worksheet and keeps every other byte."""

    def __init__(self, xml, sheet, strings):
        self.xml = xml
        self.sheet = sheet
        self.strings = strings
        self.rows = {}
        for match in ROW_RE.finditer(xml):
            self.rows[int(match.group(1))] = (match.start(), match.end(),
                                              match.group(0))
        self.styles = {}
        for row in sorted(self.rows):
            for col, attrs, _inner, _raw in self.cells(row):
                if col in self.styles:
                    continue
                style = re.search(r'\ss="[^"]*"', attrs)
                if style:
                    self.styles[col] = style.group(0)
            if len(self.styles) >= len(SHEET_COLUMNS[sheet]):
                break

    # -- reading ---------------------------------------------------------
    def cells(self, row):
        """[(column, attributes, inner, raw)] of one worksheet row."""
        entry = self.rows.get(row)
        if not entry:
            return []
        found = []
        for match in CELL_RE.finditer(entry[2]):
            attrs, inner = match.group(1), match.group(2) or ""
            ref = ELEMENT_REF.search(attrs)
            if ref:
                found.append((ref.group(1), attrs, inner, match.group(0)))
        return found

    def values(self, row):
        """{column: value} of one worksheet row, shared strings resolved."""
        out = {}
        for col, attrs, inner, _raw in self.cells(row):
            value = self.cell_value(attrs, inner)
            if value is not None:
                out[col] = value
        return out

    def cell_value(self, attrs, inner):
        kind = re.search(r't="([^"]+)"', attrs)
        kind = kind.group(1) if kind else "n"
        if "<is>" in inner:
            text = re.findall(r"<t\b[^>]*>(.*?)</t>", inner, re.S)
            return "".join(unescape(part) for part in text) or None
        match = VALUE_RE.search(inner)
        if not match:
            return None
        raw = unescape(match.group(1))
        if kind == "s":
            try:
                return self.strings.values[int(raw)]
            except (IndexError, ValueError):
                return raw
        if kind in ("str", "inlineStr"):
            return raw
        number = as_number(raw)
        return number if number is not None else raw

    def content_rows(self, from_row=5):
        """{row: {column: value}} for the rows that hold something."""
        out = {}
        for row in sorted(self.rows):
            if row < from_row:
                continue
            text = self.rows[row][2]
            if "<v>" not in text and "<is>" not in text:
                continue
            values = self.values(row)
            if values:
                out[row] = values
        return out

    # -- writing ---------------------------------------------------------
    def write_row(self, row, edits):
        """Set or clear cells of one row.  edits = {column: value or None}."""
        entry = self.rows.get(row)
        if not entry:
            raise KeyError(f"{self.sheet}: row {row} is missing")
        items = []
        for match in CELL_RE.finditer(entry[2]):
            attrs = match.group(1)
            ref = ELEMENT_REF.search(attrs)
            if ref:
                items.append([column_index_from_string(ref.group(1)),
                              ref.group(1), attrs, match.group(0)])
        present = {item[1] for item in items}
        for col in edits:
            if col not in present:
                items.append([column_index_from_string(col), col,
                              f' r="{col}{row}"' + self.styles.get(col, ""),
                              None])
        items.sort(key=lambda item: item[0])

        parts = []
        for _index, col, attrs, raw in items:
            parts.append(raw if col not in edits
                         else self.cell_xml(col, row, attrs, edits[col]))
        body = entry[2]
        start = body.index("<c") if "<c" in body else len(body)
        end = body.rindex("</c>") + 4 if "</c>" in body else start
        self._replace(row, body[:start] + "".join(parts) + body[end:])

    def cell_xml(self, col, row, attrs, value):
        """The <c> element for `value`; attrs is what the cell already had."""
        if isinstance(value, Fill):
            value = value.value
        keep = re.sub(r'\st="[^"]*"', "", attrs)
        if not keep.strip() or 'r="' not in keep:
            keep = f' r="{col}{row}"' + self.styles.get(col, "")
        if value is None:
            return f"<c{keep}/>"
        if isinstance(value, str):
            index = self.strings.index_of(value)
            if index is None:
                space = ' xml:space="preserve"' if value != value.strip() else ""
                return (f'<c{keep} t="inlineStr"><is><t{space}>{escape(value)}'
                        "</t></is></c>")
            return f'<c{keep} t="s"><v>{index}</v></c>'
        return f"<c{keep}><v>{num_text(value)}</v></c>"

    def set_cached(self, column, row, value):
        """Refresh the cached result of a row-3 summary formula."""
        entry = self.rows.get(row)
        if not entry:
            return False
        for match in CELL_RE.finditer(entry[2]):
            attrs, inner = match.group(1), match.group(2) or ""
            ref = ELEMENT_REF.search(attrs)
            if not ref or ref.group(1) != column:
                continue
            if "<v>" in inner:
                new_inner = VALUE_RE.sub(f"<v>{num_text(value)}</v>", inner, 1)
            else:
                formula = FORMULA_RE.search(inner)
                insert = formula.end() if formula else 0
                if not formula:
                    new_inner = f"<v>{num_text(value)}</v>"
                else:
                    new_inner = inner[:insert] + f"<v>{num_text(value)}</v>" \
                        + inner[insert:]
            new_cell = f"<c{attrs}>{new_inner}</c>"
            body = entry[2]
            self._replace(row, body[:match.start()] + new_cell
                          + body[match.end():])
            return True
        return False

    def _replace(self, row, new_text):
        start, end, _old = self.rows[row]
        delta = len(new_text) - (end - start)
        self.xml = self.xml[:start] + new_text + self.xml[end:]
        moved = {}
        for number, (s, e, text) in self.rows.items():
            if s > start:
                moved[number] = (s + delta, e + delta, text)
            else:
                moved[number] = (s, e, text)
        moved[row] = (start, start + len(new_text), new_text)
        self.rows = moved

    def to_bytes(self):
        return self.xml.encode("utf-8")


# ---------------------------------------------------------------------------
#  plan -> cell edits, including the rows the new month does not use
# ---------------------------------------------------------------------------

def place_docs_rows(plans, source_rows, first_row=5):
    """{row: plan} - a document series keeps the row it already occupies.

    'AKR/' stays in the AKR row, 'RES/' in the RES row and so on, wherever
    those rows are; a series the previous month did not have goes into the
    first row that is free.
    """
    placed, used, waiting = {}, set(), []
    occupied = {}
    for row in sorted(source_rows):
        prefix = series_prefix(source_rows[row].get("B"))
        if prefix is not None:
            occupied.setdefault(prefix, []).append(row)
    for plan in plans:
        prefix = leading_prefix(plan.get("B"))
        row = next((r for r in occupied.get(prefix, ()) if r not in used), None)
        if row is None:
            waiting.append(plan)
            continue
        used.add(row)
        placed[row] = plan
    free = [r for r in sorted(source_rows) if r not in used]
    row = first_row
    for plan in waiting:
        if free:
            target = free.pop(0)
        else:
            while row in used or row in source_rows:
                row += 1
            target = row
        used.add(target)
        placed[target] = plan
    return placed


def build_edits(editor, sheet, plans):
    """Work out what has to change, compared with the previous month."""
    columns = SHEET_COLUMNS[sheet]
    source_rows = editor.content_rows(5)

    placement = place_docs_rows(plans, source_rows) if sheet == DOCS else None
    last = max([5 + len(plans) - 1] + list(source_rows)
               + (list(placement) if placement else []) + [4])
    edits, final = {}, {}
    cleared = kept = 0

    for row in range(5, last + 1):
        index = row - 5
        source = source_rows.get(row, {})
        if placement is not None:
            plan = placement.get(row)
        else:
            plan = plans[index] if index < len(plans) else None

        if plan:
            row_edits = {}
            for col, value in plan.items():
                if not isinstance(value, Fill):
                    row_edits[col] = value
                    final.setdefault(row, {})[col] = value
                elif is_empty(source.get(col)):
                    row_edits[col] = value.value
                    final.setdefault(row, {})[col] = value.value
            for col, value in source.items():
                final.setdefault(row, {}).setdefault(col, value)
            if row_edits:
                edits[row] = row_edits
            continue

        # a row that is not used by the new month any more
        stale = [col for col in columns
                 if isinstance(source.get(col), (int, float))
                 and not isinstance(source.get(col), bool)
                 and source[col] != 0]
        if stale:
            if sheet == DOCS:
                row_edits = {}
                for col in ("B", "C"):
                    text = None if is_empty(source.get(col)) else str(source[col])
                    stripped = None if text is None else \
                        (re.sub(r"\d+$", "", text).strip() or None)
                    row_edits[col] = stripped
                    final.setdefault(row, {})[col] = stripped
                row_edits["D"] = 0
                final.setdefault(row, {})["D"] = 0
                for col in CLEAR_COLUMNS[DOCS]:
                    row_edits[col] = None
                    final.setdefault(row, {})[col] = None
            else:
                row_edits = {col: None for col in CLEAR_COLUMNS[sheet]
                             if col in source}
                final.setdefault(row, {}).update(
                    {col: None for col in row_edits})
                for col, value in source.items():
                    if col not in row_edits:
                        final.setdefault(row, {}).setdefault(col, value)
            edits[row] = row_edits
            cleared += 1
        else:
            if source:
                kept += 1
                final.setdefault(row, {}).update(source)

    return edits, final, dict(cleared=cleared, kept=kept)


# ---------------------------------------------------------------------------
#  the workbook
# ---------------------------------------------------------------------------

def sheet_paths(zf):
    """{worksheet name: zip entry} from workbook.xml and its relationships."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    out = OrderedDict()
    for sheet in workbook.find("m:sheets", ns):
        target = targets[sheet.get(rid)]
        out[sheet.get("name")] = target if target.startswith("xl/") \
            else "xl/" + target.lstrip("/")
    return out


def fill_workbook(arranged, template, out_path, dry_run=False, verbose=True,
                  sst_as_is=False):
    """Write `out_path`: the `template` workbook with `arranged`'s data in it."""
    data = read_arranged(arranged, verbose=verbose)

    zf = zipfile.ZipFile(template)
    paths = sheet_paths(zf)
    missing = [sheet for sheet in DATA_SHEETS if sheet not in paths]
    if missing:
        raise SystemExit(f"ERROR: {Path(template).name} has no worksheet(s) "
                         f"{', '.join(missing)} - is this the GSTR-1 utility "
                         "file?")
    strings = SharedStrings(zf.read("xl/sharedStrings.xml")
                            if "xl/sharedStrings.xml" in zf.namelist() else None)

    editors, report = {}, OrderedDict()
    for sheet in DATA_SHEETS:
        editor = SheetEditor(zf.read(paths[sheet]).decode("utf-8"), sheet,
                             strings)
        plans = SHEET_PLANNERS[sheet](data)
        edits, final, stats = build_edits(editor, sheet, plans)
        for row in sorted(edits):
            editor.write_row(row, edits[row])
        row3 = row3_summary(sheet, final)
        for coord, value in row3.items():
            editor.set_cached(coord[0], int(coord[1:]), value)
        editors[sheet] = editor
        report[sheet] = dict(rows=len(plans), row3=row3, **stats)

    if verbose:
        for sheet in DATA_SHEETS:
            info = report[sheet]
            if info["rows"]:
                totals = "  ".join(f"{cell}={_show(value)}"
                                   for cell, value in info["row3"].items())
            else:
                totals = "no new data - worksheet left as it is"
            print(f"    {sheet:<12} {info['rows']:>3} row(s)   {totals}")
            if info["cleared"] or info["kept"]:
                print(f"                 cleared {info['cleared']} old row(s), "
                      f"kept {info['kept']} leftover row(s)")

    if dry_run:
        print(f"    [dry run] {Path(out_path).name} was not written")
        zf.close()
        return report, None

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        backup = out_path.with_suffix(out_path.suffix + ".bak")
        if not backup.exists():
            shutil.copy2(out_path, backup)
            print(f"    existing {out_path.name} backed up as {backup.name}")

    count = None
    if strings.raw is not None and not sst_as_is:
        count = 0
        for sheet, path in paths.items():
            body = editors[sheet].xml if sheet in editors \
                else zf.read(path).decode("utf-8")
            count += body.count('t="s"')

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as out:
        for info in zf.infolist():
            name = info.filename
            if name == "xl/sharedStrings.xml" and strings.raw is not None \
                    and not sst_as_is:
                out.writestr(info, strings.to_bytes(count))
                continue
            modified = next((sheet for sheet, path in paths.items()
                             if path == name and sheet in editors), None)
            if modified:
                out.writestr(info, editors[modified].to_bytes())
            else:
                out.writestr(info, zf.read(name))
    zf.close()

    print(f"    saved -> {out_path}  "
          f"({out_path.stat().st_size / 1048576:.1f} MB, "
          f"{strings.added} new shared string(s))")
    return report, out_path


def _show(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


# ---------------------------------------------------------------------------
#  verification against the hand-made workbook
# ---------------------------------------------------------------------------

def scan_workbook(path):
    """{sheet: {coordinate: (kind, payload, style)}} of every filled cell."""
    zf = zipfile.ZipFile(path)
    paths = sheet_paths(zf)
    strings = SharedStrings(zf.read("xl/sharedStrings.xml")
                            if "xl/sharedStrings.xml" in zf.namelist() else None)
    out = OrderedDict()
    for sheet, path in paths.items():
        body = zf.read(path).decode("utf-8")
        cells = {}
        for row in ROW_RE.finditer(body):
            for match in CELL_RE.finditer(row.group(0)):
                attrs, inner = match.group(1), match.group(2) or ""
                ref = ELEMENT_REF.search(attrs)
                if not ref:
                    continue
                style = re.search(r'\ss="(\d+)"', attrs)
                style = style.group(1) if style else None
                coord = f"{ref.group(1)}{ref.group(2)}"
                kind = re.search(r't="([^"]+)"', attrs)
                kind = kind.group(1) if kind else "n"
                if "<is>" in inner:
                    text = re.findall(r"<t\b[^>]*>(.*?)</t>", inner, re.S)
                    cells[coord] = ("text",
                                    "".join(unescape(t) for t in text), style)
                    continue
                value = VALUE_RE.search(inner)
                if value is None:
                    continue
                raw = unescape(value.group(1))
                formula = FORMULA_RE.search(inner)
                if formula is not None:
                    cells[coord] = ("formula", formula.group(1) or "", style,
                                    raw)
                elif kind == "s":
                    try:
                        cells[coord] = ("text", strings.values[int(raw)], style)
                    except (IndexError, ValueError):
                        cells[coord] = ("text", raw, style)
                else:
                    cells[coord] = ("number", raw, style)
        out[sheet] = cells
    zf.close()
    return out


def same_payload(a, b):
    """Do two scanned cells carry the same value (style aside)?"""
    if not a or not b or a[0] != b[0]:
        return False
    if a[0] == "number":
        left, right = as_number(a[1]), as_number(b[1])
        if left is None or right is None:
            return a[1] == b[1]
        return abs(left - right) < 1e-9
    if a[0] == "text":
        return a[1].strip() == b[1].strip()
    return a[1] == b[1] and a[3] == b[3]          # formula text and result


def compare_govt(reference, generated):
    """Cell by cell comparison of a generated workbook with the hand-made one.

    A different value is a difference.  A cell that holds the same value but
    another style id is only reported as a note: the hand-made file can have
    been re-saved by another Excel build, which renumbers the styles of the
    untouched template cells.
    """
    ref, got = scan_workbook(reference), scan_workbook(generated)
    differences, notes, compared = [], [], 0
    for sheet in ref:
        if sheet not in got:
            differences.append(f"worksheet {sheet!r} is missing")
            continue
        order = sorted(set(ref[sheet]) | set(got[sheet]),
                       key=lambda c: (int(re.sub(r"\D", "", c) or 0),
                                      re.sub(r"\d", "", c)))
        for coord in order:
            a, b = ref[sheet].get(coord), got[sheet].get(coord)
            compared += 1
            if a == b:
                continue
            if same_payload(a, b):
                notes.append(f"{sheet}!{coord}: same value, style "
                             f"{a[2]} -> {b[2]}")
                continue
            differences.append(f"{sheet}!{coord}: expected {a}, got {b}")
    return dict(reference=Path(reference), generated=Path(generated),
                differences=differences, compared=compared, notes=notes)


# ---------------------------------------------------------------------------
#  finding the files of a month
# ---------------------------------------------------------------------------

def government_period(path):
    """(month, year) of a government workbook - from its data, else its name."""
    try:
        zf = zipfile.ZipFile(path)
        paths = sheet_paths(zf)
        for sheet in (B2B, B2CS):
            if sheet not in paths:
                continue
            body = zf.read(paths[sheet]).decode("utf-8")
            for row in ROW_RE.finditer(body):
                number = int(row.group(1))
                if number < 5:
                    continue
                for match in CELL_RE.finditer(row.group(0)):
                    attrs, inner = match.group(1), match.group(2) or ""
                    if f'r="D{number}"' not in attrs:
                        continue
                    value = VALUE_RE.search(inner)
                    if not value:
                        continue
                    serial = as_number(unescape(value.group(1)))
                    if serial is None or not 40000 < serial < 80000:
                        continue
                    date = XL_EPOCH + dt.timedelta(days=int(serial))
                    zf.close()
                    return date.month, date.year
        zf.close()
    except Exception:
        pass
    match = MONTH_RE.search(path.stem)
    if match:
        token = match.group(1).upper()
        month = next((i + 1 for i, name in enumerate(MONTHS)
                      if name.startswith(token[:3])), None)
        year = re.search(r"(20\d{2})", path.stem)
        if month:
            return month, int(year.group(1)) if year else None
    return None


def is_government_file(path):
    try:
        zf = zipfile.ZipFile(path)
        has = B2B in sheet_paths(zf)
        zf.close()
        return has
    except Exception:
        return False


def template_for(arranged, govt_dir):
    """The previous month's government workbook of the same client."""
    data = read_arranged(arranged, verbose=False)
    if not data["period"]:
        return None
    prefix = Path(arranged).stem.split()[0].upper()
    wanted = shift_month(data["period"][0], data["period"][1], -1)
    fallback, fallback_period = None, None
    for candidate in sorted(Path(govt_dir).glob("*GSTR1*.xlsx")):
        if "SYSTEM DATA" in candidate.stem.upper():
            continue
        if not candidate.stem.upper().startswith(prefix):
            continue
        if candidate.resolve() == Path(arranged).resolve():
            continue
        if not is_government_file(candidate):
            continue
        period = government_period(candidate)
        if period == wanted:
            return candidate
        if period and data["period"] and period < data["period"] \
                and (fallback_period is None or period > fallback_period):
            fallback, fallback_period = candidate, period
    return fallback


def find_reference(arranged, template, out_path):
    """The hand-made government file of the same month, if there is one."""
    data = read_arranged(arranged, verbose=False)
    if not data["period"]:
        return None
    prefix = Path(arranged).stem.split()[0].upper()
    skip = {Path(p).resolve() for p in (arranged, template, out_path)
            if p is not None}
    for folder in {Path(arranged).parent, Path(template).parent, Path(".")}:
        for candidate in sorted(folder.glob("*GSTR1*.xlsx")):
            if candidate.resolve() in skip:
                continue
            if not candidate.stem.upper().startswith(prefix):
                continue
            if "SYSTEM DATA" in candidate.stem.upper():
                continue
            if not is_government_file(candidate):
                continue
            if government_period(candidate) == data["period"]:
                return candidate
    return None


def output_name(template, period):
    """'.. GSTR1 JUNE-2026.xlsx' + (7, 2026) -> '.. GSTR1 JULY-2026.xlsx'."""
    month, year = period
    path = Path(template)
    match = MONTH_RE.search(path.stem)
    if not match:
        return path.with_name(f"{path.stem} {month_name(month)}-{year}"
                              f"{path.suffix}")
    token = match.group(1)
    new_month = month_name(month) if len(token) > 3 else month_name(month)[:3]
    tail = path.stem[match.end():]
    year_in_tail = re.search(r"\d{4}", tail)
    if year_in_tail:
        tail = tail[:year_in_tail.start()] + str(year) + tail[year_in_tail.end():]
    return path.with_name(path.stem[:match.start()] + new_month + tail
                          + path.suffix)


# ---------------------------------------------------------------------------
#  command line
# ---------------------------------------------------------------------------

def run_one(arranged, template, out_path, verify=None, dry_run=False,
            verbose=True):
    print(f"{Path(out_path).name}")
    print(f"    template ............. {Path(template).name}")
    report, written = fill_workbook(arranged, template, out_path,
                                    dry_run=dry_run, verbose=verbose)
    result = None
    if written is not None and verify is not False:
        reference = Path(verify) if verify else \
            find_reference(arranged, template, out_path)
        if reference is None:
            print("    verify ............... no hand-made file to compare with")
        else:
            result = compare_govt(reference, written)
            print(f"    verify vs {Path(reference).name}: ", end="")
            if result["differences"]:
                print(f"FAIL - {len(result['differences'])} difference(s)")
                for line in result["differences"][:20]:
                    print(f"        {line}")
            else:
                note = f", {len(result['notes'])} formatting note(s)" \
                    if result["notes"] else ""
                print(f"IDENTICAL ({result['compared']} cell(s) compared{note})")
                for line in result["notes"][:5]:
                    print(f"        note: {line}")
    return report, result


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Fill the government GSTR-1 workbook with a month's data.")
    parser.add_argument("arranged", nargs="*",
                        help="formatted '.. GSTR1 SYSTEM DATA ..' file(s)")
    parser.add_argument("--template", help="previous month's government file")
    parser.add_argument("-o", "--output", help="output file (single input only)")
    parser.add_argument("--outdir", default="govt_out",
                        help="folder for the generated file(s) "
                             "(default: govt_out, use . for this folder)")
    parser.add_argument("--all", action="store_true",
                        help="do every client in one run")
    parser.add_argument("--arranged-dir", default="out",
                        help="folder with the SYSTEM DATA files (default: out)")
    parser.add_argument("--govt-dir", default=".",
                        help="folder with the government files (default: .)")
    parser.add_argument("--verify", nargs="?", const="", default="",
                        help="compare with the hand-made file (default: on)")
    parser.add_argument("--no-verify", dest="verify", action="store_const",
                        const=None, help="skip the comparison")
    parser.add_argument("--dry-run", action="store_true",
                        help="report what would be written, change nothing")
    parser.add_argument("--sst-as-is", action="store_true",
                        help="never touch sharedStrings.xml")
    args = parser.parse_args(argv)

    jobs = []
    if args.all:
        for arranged in sorted(Path(args.arranged_dir)
                               .glob("*GSTR1*SYSTEM DATA*.xlsx")):
            template = template_for(arranged, args.govt_dir)
            if template is None:
                print(f"    no previous month's government file for "
                      f"{arranged.name} in {args.govt_dir} - skipped")
                continue
            jobs.append((arranged, template))
        if not jobs:
            parser.error(f"no '.. GSTR1 SYSTEM DATA ..' file in "
                         f"{args.arranged_dir}")
    else:
        if not args.arranged:
            parser.error("give a SYSTEM DATA file, or use --all")
        for name in args.arranged:
            arranged = Path(name)
            template = Path(args.template) if args.template else \
                template_for(arranged, args.govt_dir)
            if template is None:
                parser.error(f"no government file found for {name}; "
                             "use --template")
            jobs.append((arranged, template))

    exit_code = 0
    for position, (arranged, template) in enumerate(jobs):
        if position:
            print()
        period = (read_arranged(arranged, verbose=False)["period"]
                  or government_period(template))
        if args.output and len(jobs) == 1:
            out_path = Path(args.output)
        else:
            name = output_name(template, period) if period else \
                template.with_name(template.stem + " filled.xlsx")
            out_path = Path(args.outdir) / name.name if args.outdir \
                else template.parent / name.name
        _report, result = run_one(arranged, template, out_path, args.verify,
                                  dry_run=args.dry_run)
        if result and result["differences"]:
            exit_code = 1
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
