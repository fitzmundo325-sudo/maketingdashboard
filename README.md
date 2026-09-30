# HYG Marketing Department Database

A Flask app that catalogs marketing resources by brand, imported from Excel workbooks.

## Project structure

```
marketing-hub/
├── app.py                  # thin entry point: from website import create_app
├── run.cmd                 # windows launcher
├── .env                    # MH_SECRET_KEY, MH_APP_DB_NAME (not committed)
├── requirements.txt
├── parsed_data.json        # seed data (auto-loads on first run if DB empty)
└── website/                # application package
    ├── __init__.py         # create_app() factory, db, seeding, CLI commands
    ├── models.py           # User, Brand, Category, Resource, GoogleSheet(+Rows)
    ├── views.py            # main blueprint (dashboard, brands, categories, CRUD, Google links)
    ├── auth.py             # login/logout blueprint (Flask-Login)
    ├── api_handles.py      # /apis JSON blueprint
    ├── excel_mappings.py   # per-sheet Excel column mappings
    ├── excel_import.py     # shared Excel importer
    ├── import_*.py         # one-off importer entry scripts
    ├── templates/
    └── static/
```

Structure mirrors the cficountsystem project (application-factory + blueprints).

## Run it

```bash
pip install -r requirements.txt
python app.py
```

Then open http://localhost:5000 and log in.

### Create the first user

```bash
flask --app app create-admin
```

(Username + password are prompted; passwords are hashed with werkzeug.)

## Seed data

On first run the app auto-creates `marketing_hub.db` (SQLite) and seeds it from
`parsed_data.json` (the cleaned Excel export) if the DB is empty. Delete
`marketing_hub.db` and restart to re-seed from scratch — any edits made in the
UI will be lost when you do that.

## Notes on the source data

- `ELEVATE` was an empty sheet in the original workbook — it's created as a brand
  with zero resources so it's ready to fill in.
- Many "Spreadsheet Link" cells in the original file were plain sheet/file names
  rather than URLs (e.g. `COINBANK DATABASE`), not live hyperlinks. Those show as
  plain text with a file icon in the UI; anything starting with `http(s)://` renders
  as a clickable link.

## Excel imports

The one-off importers run as modules from the project root:

```bash
python -m website.import_coinbank
python -m website.excel_import "path/to/file.xlsx" [--update | --backfill]
```

### Master database import

`HYG MARKETING MASTER DATABASE.xlsx` (one sheet per brand) is supported by a
dedicated importer that reads the workbook's own layout:

- each sheet becomes a **brand** (GOLDILOCKS, SAVORY, ICEBERGS, TATERS,
  CHATIME, ELEVATE)
- the header row is located dynamically (column order differs per sheet)
- ALL-CAPS section rows (e.g. `NOVELTIES`, `MANCOMM REPORTS`, `GIFT CERTS`)
  become **categories**; the resources under them are grouped accordingly
- cleanup is automatic: `#REF!` refs, `None`/`0.0` links, float `1.0`-style
  numbers, datetime cells, and status values stuck in the Name column

```bash
python -m website.import_master "C:\path\to\HYG MARKETING MASTER DATABASE.xlsx" [--replace]
# or
flask --app app import-master
```

`--replace` removes each matched brand's existing resources first (brands and
categories are kept). Without it, resources are appended to the matching
categories.

Sheets that aren't shared publicly show as "failed" in the registry with the
HTTP error — they keep their links but their data can't be mirrored without
Google OAuth.

## Google Links registry + mirrored sheet data

Every Google URL on a resource (Sheets, Docs, Slides, Forms, Drive files,
Apps Script) is tracked in its own `google_sheets` database table, deduplicated
by document ID. Google **Sheets** additionally get their contents mirrored into
the `google_sheet_rows` table via the public CSV export endpoint — no Google
login required.

- **Registry page:** `/sheets` ("Google Links" in the top nav) lists all
  entries with sync status, type pills, search, and add/edit/remove/sync.
- **Data viewer:** clicking a sheet link (or the "View Data" pill on resources)
  opens `/sheets/<id>` inside the app, showing the mirrored rows with search.
  "Open in Google" is still available on that page.
- **Auto-registration:** adding/editing a resource with a Google link upserts
  the registry row and pulls the sheet's data on first sight.
- **CLI:** `flask --app app backfill-google-sheets` (register new links),
  `flask --app app sync-google-sheets` (re-pull every sheet's data).

Registry rows keep: title, canonical URL, document ID, type, owner, status,
notes, sync timestamp/status/error, and the resources using them.
