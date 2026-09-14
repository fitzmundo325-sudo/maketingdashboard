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
    ├── models.py           # User, Brand, Category, Resource
    ├── views.py            # main blueprint (dashboard, brands, categories, CRUD)
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
