# Filling the government GSTR-1 sheet

`fill_gstr1_govt.py` is the third step of the monthly routine:

| step | command | result |
| --- | --- | --- |
| 1 | Tally export | `BDBJUL.xlsx`, `DPBJUL.xlsx`, `HKIJUL.xlsx`, `HTEJUL.xlsx` |
| 2 | `python3 format_gstr1.py BDBJUL.xlsx DPBJUL.xlsx HKIJUL.xlsx HTEJUL.xlsx --outdir out` | `out/BDB GSTR1 SYSTEM DATA JULY-2026.xlsx` … |
| 3 | `python3 fill_gstr1_govt.py --all` | `BDB GSTR1 JULY-2026.xlsx` … |

Step 3 is the manual work it replaces: take the previous month's government file
(the one that was uploaded to the portal), delete the old month's data and type
the new month's data in every worksheet.

```bash
# one client
python3 fill_gstr1_govt.py "out/BDB GSTR1 SYSTEM DATA JULY-2026.xlsx" \
        --template "BDB GSTR1 JUNE-2026.xlsx"

# all four clients (SYSTEM DATA files are looked up in out/, the previous
# month's government files next to the script)
python3 fill_gstr1_govt.py --all
python3 fill_gstr1_govt.py --all --outdir .         # write next to the files
```

The output is named after the template (`BDB GSTR1 JUNE-2026.xlsx` ->
`BDB GSTR1 JULY-2026.xlsx`), the month comes from the "Sales Register From … TO
…" line of the arranged file.  It is written to `govt_out/` unless `--outdir`
says otherwise — the name is the same as the hand-made file of that month, and
that file must not be overwritten by accident.  (If the output does exist it is
copied to `.bak` before anything is written.)

## What is written where

Every invoice of the arranged file falls into one of three groups:

| group | how it is recognised | goes to |
| --- | --- | --- |
| **registered** (B2B) | column A `Gst No` is filled | `b2b,sez,de`, `hsn(b2b)` |
| **un-registered** | no Gst No, but tax amounts | `b2cs`, `hsn(b2c)` |
| **tax free / exempt** | no Gst No and no tax amounts | `exemp`, `hsn(b2b)` |

| worksheet | one row per … | columns taken from the arranged file |
| --- | --- | --- |
| `b2b,sez,de` | registered invoice | A Gst No, B Party, C Bill No, D Date, E Bill Amt, F State, K rate, L Taxable |
| `b2cs` | un-registered invoice | B place of supply, D rate, E Taxable (`Type` = `OE`) |
| `exemp` | month (the exempt total) | C = sum of the tax free taxable value |
| `hsn(b2b)` | registered group, + tax free group | D quantity (no. of invoices), E total value, F rate, G taxable, H IGST, I CGST, J SGST |
| `hsn(b2c)` | un-registered group | same columns as `hsn(b2b)` |
| `docs` | invoice-number series (`AKR/`, `RES/`, plain numbers) | B Sr. No. From, C Sr. No. To, D total number |

The other 17 worksheets carry no data for these clients and are left as they
are.  The summary formulas of row 3 (no. of recipients, no. of invoices, taxable
value, …) keep their formulas; only the cached result is refreshed so the totals
are right even before Excel recalculates.

Rows, sections and invoice-number series are found with regular expressions —
nothing is bound to a row number or to the number of invoices, so a month with
3 invoices or 300, with or without un-registered / tax free sales, is filled the
same way.

## Deleting the old month

A row that held data last month but has none this month is cleared the way it is
cleared by hand:

* the values are deleted (all data cells of the row),
* in `docs` the series row stays behind with the digits stripped (`AST/01` ->
  `AST/`) and a zero count,
* a line that never held a number is deliberately left alone — that is what the
  hand-made files do: HTEI's `WIND MILL` HSN line, HKI's empty second docs line,
  BDB's zeroed `hsn(b2c)` line, HTEI's `WIND MILL INCOME` exemp line.

In `docs` a series keeps the row it already occupies (`AKR/` stays in the AKR
row, `RES/` in the RES row); a series the previous month did not have yet goes
into the first free row.

## Verifying

After writing, the file is compared cell by cell with the hand-made file of the
same month, when it can be found next to the arranged or the template file
(`--verify FILE` to name it, `--no-verify` to skip):

```
    saved -> BDB GSTR1 JULY-2026.xlsx  (5.9 MB, 1 new shared string(s))
    verify vs BDB GSTR1 JULY-2026.xlsx: IDENTICAL (1520 cell(s) compared)
```

The comparison looks at every filled cell of all 23 worksheets — numbers, text,
dates, the cached values of the row-3 formulas — and at the style of each cell.
A cell that holds the same value with another style number is reported as a
*formatting note* instead of a difference, because a file that was re-saved by
another Excel build gets its style numbers renumbered; HKI's hand-made July file
has 416 of those (399 of them in the "Help Instruction" sheet, the rest on
empty sheet titles and five receiver-name cells).  The values are identical
everywhere in all four files.

## Notes and assumptions

* The template must be the previous month's **filled** file (the habit this
  script replaces). A freshly downloaded portal template also works, but then
  the leftover lines listed above do not exist yet.
* One HSN row is written per group. If un-registered sales are ever reported in
  two categories (HTEI had a second `ASSETS` HSN line in June), the extra line
  is cleared — the arranged file has nothing that identifies the category.
* `WIND MILL INCOME` in `exemp` and `WIND MILL` in `hsn(b2b)` are kept as they
  are; the arranged file only carries the tax free total, which is written to
  the `RESIDENTIAL RENT` line.
* The company name, GSTIN and other title cells are not touched.
