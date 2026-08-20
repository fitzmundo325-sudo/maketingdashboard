# HYG Marketing Hub

A Flask app built from `HYG_MARKETING_MASTER_DATABASE.xlsx`. It turns the six brand
tabs (GOLDILOCKS, SAVORY, ICEBERGS, TATERS, CHATIME, ELEVATE) into a searchable,
editable web database instead of a spreadsheet.

## What it does

- **Dashboard** — a card per brand with resource/category counts.
- **Brand pages** — resources grouped into the same categories that existed as
  section headers in each sheet (COINBANK, NOVELTIES, SPONSORSHIP DATABASE,
  RUNNING PROMO, MANCOMM REPORTS, GIFT CERTS, etc.), with in-brand filtering by
  keyword and status.
- **Global search** — search name / link / instructions / created-by across every brand at once.
- **Full CRUD** — add, edit, or delete resources and categories from the UI (no more
  hunting for the right row in Excel).
- **JSON API** — `/api/brands` and `/api/brand/<id>/resources` for hooking this up
  to other tools (e.g. your GAS/PHP systems) later.

## Data model

- `Brand` (GOLDILOCKS, SAVORY, ...)
- `Category` (the ALL-CAPS section headers from each sheet, scoped per brand)
- `Resource` (a row: name, link/reference, instructions, created_by, access, status)

The original "No." column was kept as free text (`ref_no`) since the sheet mixed
numbers, dates, and `#REF!` errors — it's not used for anything functional, just
carried over for reference.

## Run it

```bash
pip install -r requirements.txt
python app.py
```

Then open http://localhost:5000

On first run the app auto-creates `marketing_hub.db` (SQLite) and seeds it from
`parsed_data.json` (the cleaned Excel export). Delete `marketing_hub.db` and restart
to re-seed from scratch — any edits made in the UI will be lost when you do that.

## Notes on the source data

- `ELEVATE` was an empty sheet in the original workbook — it's created as a brand
  with zero resources so it's ready to fill in.
- Many "Spreadsheet Link" cells in the original file were plain sheet/file names
  rather than URLs (e.g. `COINBANK DATABASE`), not live hyperlinks. Those show as
  plain text with a file icon in the UI; anything starting with `http(s)://` renders
  as a clickable link.
- A few rows had category headers with no rows under them, or rows before the
  first category header (treated as "General").

## Next steps you might want

- Swap SQLite for a shared DB (Postgres/MySQL) if more than one person needs to edit at once.
- Add login/auth if this goes beyond your local machine.
- Wire the `link` field to actually open Google Drive files via the Drive API instead of pasting names.
