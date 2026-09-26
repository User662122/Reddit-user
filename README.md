Code is not created now as I have no real Api to test what will works perfectly when i will create. 
A personal, non-commercial AI assistant that monitors user-selected subreddits, generates AI-powered draft replies for review, and posts replies only after explicit user approval. The application follows Reddit's Developer Terms, Data API Terms, and subreddit rules.

---

### Excel / GST helpers

* `format_gstr1.py` — formats the raw monthly Tally sales registers of all four
  clients (BDB, DPB, HKI, HTEI) exactly like their hand-made
  "GSTR1 SYSTEM DATA" files, so the monthly formatting is a single command.
  Regex driven, so the number of invoices, sections and GST rates can change
  every month.
* `fill_gstr1_govt.py` — fills the government GSTR-1 offline-utility workbook of
  a month: it takes the previous month's government file, deletes the old data
  and writes the new month's data from the arranged file into the applicable
  worksheets (`b2b,sez,de`, `b2cs`, `exemp`, `hsn(b2b)`, `hsn(b2c)`, `docs`),
  then compares the result with the hand-made file.
  Regex driven, so a month with more or fewer invoices needs no change.
* `compare_workbooks.py` — cell-by-cell check of a generated file against a sample.

See [GSTR1_FORMATTING.md](GSTR1_FORMATTING.md) for step 2 and
[GSTR1_GOVT_FILLING.md](GSTR1_GOVT_FILLING.md) for step 3.
