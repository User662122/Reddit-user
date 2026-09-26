# Monthly GST-R1 sales-register formatting

Two tools that remove the manual Excel work every month:

| file | what it is |
|---|---|
| `format_gstr1.py` | formats the raw Tally exports of all four clients exactly like their hand-made "GSTR1 SYSTEM DATA" files |
| `compare_workbooks.py` | checks a generated file against a sample, cell by cell |

## Install

```bash
pip install openpyxl
```

## Use

```bash
# one client
python3 format_gstr1.py BDBJUL.xlsx
#   -> BDB GSTR1 SYSTEM DATA JULY-2026.xlsx

# every client in one go, and check the result against their own files
python3 format_gstr1.py BDBJUL.xlsx DPBJUL.xlsx HKIJUL.xlsx HTEJUL.xlsx --verify

# a whole month's folder (raw exports + the clients' files side by side)
python3 format_gstr1.py *.xlsx --outdir out/ --verify
```

The client is recognised from the file name (`BDBJUL.xlsx` → BDB, `DPBJUL.xlsx` →
DPB, `HKIJUL.xlsx` → HKI, `HTEJUL.xlsx` → HTEI). Everything else is read from the
file itself, so the output name is built from the "Sales Register From ... TO ..."
line:

| input | output |
|---|---|
| `BDBJUL.xlsx` | `BDB GSTR1 SYSTEM DATA JULY-2026.xlsx` |
| `DPBJUL.xlsx` | `DPB GSTR1 SYSTEM DATA JULY 2026.xlsx` |
| `HKIJUL.xlsx` | `HKI GSTR1 SYSTEM DATA JULY-2026.xlsx` |
| `HTEJUL.xlsx` | `HTEI GSTR1 SYSTEM DATA JULY 2026.xlsx` |

### Options

| option | meaning |
|---|---|
| `-o FILE` | write to this file (one input only) |
| `--outdir DIR` | write the auto-named files into DIR |
| `--profile NAME` | force a client profile (for a new client, or a file name that does not match) |
| `--verify [REF]` | after saving, compare with the client's own file (found automatically next to the input, or pass REF) |
| `--dry-run` | report what would be done, write nothing |
| `--inplace` | format the input file itself (keeps a `<name>.xlsx.bak`) |
| `--list-profiles` | show the client profiles |

The input is never modified unless you pass `--inplace`. If an output file already
exists it is kept as `<name>.xlsx.bak` before being replaced. A file that is
already in "GSTR1 SYSTEM DATA" form is skipped automatically, so `*.xlsx` is safe.

## What is formatted

Common to all four clients:

* company-name line → **bold**
* heading row → **bold, centred**, row height 30, "wrap" on the columns that need it
* section line (`GST - GST 18 %`, `TXF - TAX FREE`, `G18 - GST 18%`, …) → **bold + underlined**
* invoice line → `Bill No` / `Date` / `State` centred, date as `d-mmm-yy`,
  the amount columns as `0.00`, the `%` columns centred
* `TOTAL FOR …` row → whole row **bold**, label moved to the Party-Name column,
  right aligned and suffixed exactly like that client writes it
  (`TOTAL FOR GST 18 % REG : ` for BDB/DPB/HTEI, `TOTAL FOR GST 18% : ` for HKI),
  `TOTAL FOR TAX FREE : `, `GRAND TOTAL : `
* column widths, landscape A4 print setup, and the client's own margins/scale

Per client (these differ, which is why each has a profile in `CLIENT PROFILES`):

| | BDB | DPB | HKI | HTEI |
|---|---|---|---|---|
| heading text corrected | – | `Bill No`, `Bill Amt` | `Bill No`, `Bill Amt` | `Bill No.`, `Bill Amt` |
| amount format on invoice lines | all 5 | all 5 | – | Bill/SGST/CGST/Taxable |
| `REG` in the total label | yes | yes | no | yes |
| `REG` added to the section heading | – | – | – | yes |
| combined GST row when a rate appears twice | – | yes | – | yes |
| empty row above the heading row removed | – | – | – | yes |
| print scale | 95 % | 90 % | 90 % | 90 % |

"Combined GST row": when a month splits one GST rate over more than one section
(registered + un-registered, or the same rate again later), DPB and HTEI add a row
like `TOTAL FOR GST 18 % : ` that sums those sections — written as a live formula
(`=E11+E27`) exactly like in their files. Each rate group gets its own row, so a
month with 5 % and 18 % sections stays correct.

## Why it works for any month

Nothing depends on row numbers or on how many invoices a section has; everything
is found with regular expressions (`REGEX PATTERNS` block in `format_gstr1.py`):

* the heading row = the row that matches several known headings
  (`Gst No`, `Party Name`, `Bill.No`, `Date`, …)
* columns are looked up by heading text, not by letter
* a section line = `XXX - something` in a row that has no date/amount
* a total row = text starting with `TOTAL FOR` or `GRAND TOTAL`
* an invoice line = anything else that has a date or an amount
* the month/year of the output name comes from the
  `Sales Register From 01-07-26 TO 31-07-26` line

So five GST rates, one section or six, 3 invoices or 300 — no code change needed.
Re-running on an already formatted file changes nothing (no double `REG :`, no
second combined row).

## Adding a new client

1. If the file name starts with the client's code, a profile is found automatically;
   otherwise run with `--profile DPB` (any existing profile as a starting point).
2. To make it permanent, copy an entry in `CLIENT PROFILES` in `format_gstr1.py` and
   adjust: output name template, heading fixes, which columns are centred/wrapped/
   formatted, the total suffixes, `combine_same_rate`, column widths, scale, margins.
3. Check it against that client's own file: `python3 format_gstr1.py RAW.xlsx --verify`.

## Checking the output

```bash
python3 compare_workbooks.py "BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" "out/BDB GSTR1 SYSTEM DATA JULY-2026.xlsx"
python3 compare_workbooks.py --folder out/          # every file in out/ vs the samples here
python3 compare_workbooks.py --folder out/ --all    # also show the invisible notes
```

It reports every difference in values, bold/underline, alignment, number formats,
column widths, row heights and page setup, and evaluates the samples' formulas
(`=E11+E27`) before comparing. Differences that cannot be seen — a font, number
format or alignment sitting on an *empty* cell — are listed as "cosmetic" notes and
do not fail the check. Exit code is 0 when the formatting matches.

The four July pairs in this repository are the test:

```bash
mkdir -p out
python3 format_gstr1.py BDBJUL.xlsx DPBJUL.xlsx HKIJUL.xlsx HTEJUL.xlsx --outdir out --verify
#   -> all four: IDENTICAL (only invisible cosmetic notes)
```

## Two quirks copied from the samples

These exist in the clients' own files and can be switched off in one line:

* `DPB`: the party name on one invoice was corrected by hand to
  `PIYUSH JAGDISHBHAI ZALAWADIYA - RENT` → `party_name_fixes` in the DPB profile.
* `HTEI`: the heading `GUR - GST 18% UN REGI` is bold but **not** underlined in
  their July file (every other heading is) → `sections_without_underline`.
