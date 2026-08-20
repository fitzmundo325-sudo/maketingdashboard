"""Shared Excel importer used by import_coinbank.py and import_qr.py.

Stores the full Excel row as Resource.data (JSON) so the edit form can render
one labeled field per column. Use --backfill to fill `data` on rows imported
before this feature existed (only rows with data IS NULL are touched).
"""
import os
import sys
from datetime import datetime

import openpyxl

from app import app, db, Brand, Category, Resource, get_or_create_category
from excel_mappings import FILES, PARENT_CATEGORIES


def fmt(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def sheet_rows(ws, header_row=0):
    rows = list(ws.iter_rows(values_only=True))
    hdr_vals = rows[header_row]
    last_hdr = -1
    for i, v in enumerate(hdr_vals):
        if fmt(v):
            last_hdr = i
    data_rows = rows[header_row + 1:]
    max_cols = max([len(r) for r in data_rows] + [0])
    extra_cols = [
        col for col in range(last_hdr + 1, max_cols)
        if any(col < len(r) and r[col] not in (None, "") for r in data_rows)
    ]
    header = []
    for i in range(max_cols):
        if i <= last_hdr:
            key = fmt(hdr_vals[i]) if i < len(hdr_vals) else None
            key = key or f"col{i + 1}"
        elif i in extra_cols:
            key = f"col{i + 1}"
        else:
            continue
        if isinstance(key, str) and "\n" in key:
            key = key.split("\n")[0].strip()
        base, n = key, 2
        while key in header:
            key = f"{base}_{n}"
            n += 1
        header.append(key)
    out = []
    for r in data_rows:
        if not any(v not in (None, "") for v in r):
            continue
        vals = [fmt(v) for v in r]
        vals.extend([None] * (len(header) - len(vals)))
        out.append(dict(zip(header, vals[:len(header)])))
    return out


def clean_row(row, mapping):
    exclude = set(mapping.get("exclude", []))
    return {k: v for k, v in row.items() if k not in exclude}


def _make_resource(row, mapping, tab, i):
    row = clean_row(row, mapping)
    name = row.get(mapping["name"]) or "(Untitled)"
    labels = mapping.get("labels", {})
    parts = []
    if mapping.get("all_columns"):
        skip = {mapping.get("name"), mapping.get("created_by"), mapping.get("status")}
        for key, val in row.items():
            if val in (None, "") or key in skip:
                continue
            parts.append(f"{labels.get(key, key)}: {val}")
    else:
        for key in mapping["instructions"]:
            val = row.get(key)
            if val not in (None, ""):
                parts.append(f"{labels.get(key, key)}: {val}")
        notes_key = mapping.get("notes")
        if notes_key:
            note = row.get(notes_key)
            if note not in (None, ""):
                parts.append(f"{labels.get(notes_key, 'NOTE')}: {note}")
    return Resource(
        ref_no=row.get(mapping.get("ref")),
        name=name,
        instructions="\n".join(parts) or None,
        created_by=row.get(mapping.get("created_by")) if mapping.get("created_by") else None,
        status=row.get(mapping.get("status")) if mapping.get("status") else None,
        data=row,
        sort_order=i,
        category=tab,
    )


def _sync_from_data(res, mapping):
    from app import rebuild_from_data
    rebuild_from_data(res)


def run_import(xlsx_path, backfill=False, update=False):
    if not os.path.exists(xlsx_path):
        print(f"File not found: {xlsx_path}")
        return False
    filename = os.path.basename(xlsx_path)
    file_cfg = FILES.get(filename)
    if not file_cfg:
        print(f"No mapping for file: {filename}")
        return False
    parent_name = PARENT_CATEGORIES.get(filename)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    with app.app_context():
        parent = Category.query.filter_by(name=parent_name).first()
        if not parent:
            brand = Brand.query.order_by(Brand.id).first()
            if brand is None:
                print("No brand found; cannot create parent category.")
                return False
            top = get_or_create_category(brand, "Coin Bank")
            parent = get_or_create_category(brand, parent_name, parent=top)
            print(f"Created parent category '{parent_name}' under '{top.name}'.")

        for sheet_name, mapping in file_cfg["sheets"].items():
            if sheet_name not in wb.sheetnames:
                print(f"Sheet '{sheet_name}' missing, skipping.")
                continue
            tab = get_or_create_category(parent.brand, sheet_name, parent=parent)
            rows = sheet_rows(wb[sheet_name], header_row=mapping.get("header_row", 0))

            if update:
                updated = added = 0
                order_map = {r.sort_order: r for r in tab.resources}
                name_map = {}
                for r in tab.resources:
                    key = r.name
                    name_map.setdefault(key, []).append(r)
                for i, row in enumerate(rows, start=1):
                    res = order_map.get(i)
                    if res is None and mapping.get("name"):
                        cands = name_map.get(row.get(mapping["name"]), [])
                        res = cands[0] if cands else None
                    if res is None:
                        res = _make_resource(row, mapping, tab, i)
                        db.session.add(res)
                        added += 1
                    else:
                        res.data = clean_row(row, mapping)
                        _sync_from_data(res, mapping)
                        updated += 1
                db.session.commit()
                print(f"'{sheet_name}': updated {updated} resources, added {added} "
                      f"(file has {len(rows)} data rows, tab has {len(tab.resources)}).")
                continue

            if backfill:
                updated = 0
                cleaned = 0
                order_map = {r.sort_order: r for r in tab.resources}
                for i, row in enumerate(rows, start=1):
                    res = order_map.get(i)
                    if res is None:
                        continue
                    if res.data is None:
                        res.data = clean_row(row, mapping)
                        _sync_from_data(res, mapping)
                        updated += 1
                    else:
                        cleaned_data = clean_row(res.data, mapping)
                        if cleaned_data != res.data:
                            res.data = cleaned_data
                            cleaned += 1
                db.session.commit()
                print(f"'{sheet_name}': backfilled {updated} resources, cleaned {cleaned}.")
            else:
                if tab.resources:
                    print(f"'{sheet_name}': already has {len(tab.resources)} resources, skipping (use --backfill).")
                    continue
                for i, row in enumerate(rows, start=1):
                    db.session.add(_make_resource(row, mapping, tab, i))
                db.session.commit()
                print(f"'{sheet_name}': imported {len(rows)} resources into tab '{sheet_name}'.")
        print("Done.")
        return True


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else ""
    if not path:
        print("Usage: python excel_import.py <xlsx path> [--backfill] [--update]")
        sys.exit(1)
    run_import(path, backfill="--backfill" in sys.argv, update="--update" in sys.argv)