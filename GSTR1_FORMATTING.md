# Monthly GST-R1 sales-register formatting

Two small tools that remove the manual Excel work every month:

| file | what it is |
|---|---|
| `format_gstr1.py` | formats a raw Tally export (`BDBJUL.xlsx`) exactly like the hand-made `BDB GSTR1 SYSTEM DATA JULY-2026.xlsx` |
| `compare_workbooks.py` | checks a generated file against any sample, cell by cell |

## Install

```bash
pip install openpyxl
```

## Use

```bash
# raw export in, formatted file out (name built from the register month)
python3 format_gstr1.py BDBJUL.xlsx
#   -> BDB GSTR1 SYSTEM DATA JULY-2026.xlsx

python3 format_gstr1.py BDBAUG.xlsx                 # -> ... AUGUST-2026.xlsx
python3 format_gstr1.py BDBSEP.xlsx -o "sep done.xlsx"
python3 format_gstr1.py BDBJUL.xlsx --outdir out/    # keep outputs in a folder
python3 format_gstr1.py BDBJUL.xlsx --dry-run        # show what would happen
python3 format_gstr1.py BDBJUL.xlsx --inplace        # edit the file itself (keeps a .bak)
```

The source file is never touched unless you pass `--inplace`. If the output file
already exists, the old one is kept as `<name>.xlsx.bak` before it is replaced.

## What it formats

* company-name line → **bold**
* heading row → **bold, centred, wrapped**, row height 30
* section line (`GST - GST 18 %`, `TXF - TAX FREE`, `GST - GST 5 %`, …) → **bold + underlined**
* invoice line → `Bill.No`, `Date`, `State` centred; date as `d-mmm-yy`;
  `Bill Amt`, `SGST Amt`, `CGST Amt`, `IGST Amt`, `Taxable Amt` as `0.00`;
  `SGST %`, `CGST %`, `IGST %` centred in sections that carry a rate (a section
  whose rates are all 0, e.g. TAX FREE / RES bills, is left as Tally exported it)
* `TOTAL FOR …` / `GRAND TOTAL` row → whole row **bold**, label moved to the
  Party-Name column and right aligned, suffixed `REG :` for GST totals and ` : `
  for the rest, money columns as `0.00`
* column widths, `d-mmm-yy`, landscape A4 @ 95 %, narrow margins — as in the sample

## Why it works for any month

Nothing depends on row numbers or on how many invoices a section has; everything
is found with regular expressions (`RE_…` block in `format_gstr1.py`):

* the heading row = the row that matches several known headings (`Gst No`,
  `Party Name`, `Bill.No`, `Date`, …)
* the columns are looked up by heading name, not by letter
* a section line = `XXX - something` in a row that has no date/amount
* a total row = text starting with `TOTAL FOR` or `GRAND TOTAL`
* an invoice line = anything else that has a date or an amount
* the month/year of the output name comes from the
  `Sales Register From 01-07-26 TO 31-07-26` line, and the client prefix from
  the input file name (`BDBJUL.xlsx` → `BDB`)

So five GST rates, one section or six, 3 invoices or 300 — no code change needed.
Re-running on an already formatted file changes nothing (the `REG :` suffix is not
added twice).

## Settings

Everything you may want to change is in the `SETTINGS` block at the top of
`format_gstr1.py`, e.g.

```python
GST_TOTAL_SUFFIX   = " REG :"    # suffix for 'TOTAL FOR GST 18 %'
OTHER_TOTAL_SUFFIX = " : "       # suffix for 'TOTAL FOR TAX FREE'
GRAND_TOTAL_SUFFIX = " : "       # suffix for 'GRAND TOTAL'
MONEY_FORMAT       = "0.00"
DATE_FORMAT        = "d-mmm-yy"
PERCENT_ALIGNMENT  = "auto"      # "auto" | "always" | "never"
HEADER_ROW_HEIGHT  = 30
COLUMN_WIDTHS      = {...}
```

## Checking the output

```bash
python3 compare_workbooks.py "BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" "out/BDB GSTR1 SYSTEM DATA JULY-2026.xlsx"
```

It reports every difference in values, bold/underline, alignment, number formats,
column widths, row heights and page setup, and ignores differences that cannot be
seen (a number format or font on an empty cell). Add `--all` to see those too.
Exit code is 0 when the formatting matches.

The July pair in this repository is a good smoke test:

```bash
mkdir -p out
python3 format_gstr1.py BDBJUL.xlsx --outdir out
python3 compare_workbooks.py "BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" "out/BDB GSTR1 SYSTEM DATA JULY-2026.xlsx"
# -> RESULT: identical formatting
```
