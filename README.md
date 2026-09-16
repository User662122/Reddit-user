Code is not created now as I have no real Api to test what will works perfectly when i will create. 
A personal, non-commercial AI assistant that monitors user-selected subreddits, generates AI-powered draft replies for review, and posts replies only after explicit user approval. The application follows Reddit's Developer Terms, Data API Terms, and subreddit rules.

---

### Excel / GST helpers

* `format_gstr1.py` — formats the raw monthly Tally sales register (`BDBJUL.xlsx`)
  exactly like the hand-made `BDB GSTR1 SYSTEM DATA JULY-2026.xlsx`, so the
  monthly formatting is a single command. Regex driven, so the number of invoices
  and sections can change every month.
* `compare_workbooks.py` — cell-by-cell check of a generated file against a sample.

See [GSTR1_FORMATTING.md](GSTR1_FORMATTING.md) for usage.
