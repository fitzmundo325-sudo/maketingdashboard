from datetime import datetime, date
import csv
import io
import re
import urllib.parse
import urllib.request

from flask_login import UserMixin
from . import db


class User(db.Model, UserMixin):
    __tablename__ = "user"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True)
    full_name = db.Column(db.String(100))
    role = db.Column(db.String(100), default="Admin")
    password = db.Column(db.String(200))
    date_added = db.Column(db.DateTime(timezone=True), default=datetime.utcnow)


class Brand(db.Model):
    __tablename__ = "brands"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), unique=True, nullable=False)
    color = db.Column(db.String(20), default="#333333")

    categories = db.relationship("Category", backref="brand", cascade="all, delete-orphan", order_by="Category.name")

    @property
    def resource_count(self):
        return sum(len(c.resources) for c in self.categories)


class Category(db.Model):
    __tablename__ = "categories"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(150), nullable=False)
    brand_id = db.Column(db.Integer, db.ForeignKey("brands.id"), nullable=False)
    parent_id = db.Column(db.Integer, db.ForeignKey("categories.id"), nullable=True)

    parent = db.relationship("Category", remote_side=[id], backref="children")

    resources = db.relationship("Resource", backref="category", cascade="all, delete-orphan", order_by="Resource.sort_order")

    __table_args__ = (db.UniqueConstraint("name", "brand_id", name="uq_category_brand"),)

    @property
    def resource_count(self):
        return len(self.resources) + sum(c.resource_count for c in self.children)


class Resource(db.Model):
    __tablename__ = "resources"
    id = db.Column(db.Integer, primary_key=True)
    ref_no = db.Column(db.String(20))          # original "No." column, kept as text (had #REF!, dates, floats)
    name = db.Column(db.String(300), nullable=False)
    link = db.Column(db.Text)                  # URL or free-text reference to a file/sheet
    instructions = db.Column(db.Text)
    created_by = db.Column(db.String(120))
    access = db.Column(db.String(120))
    status = db.Column(db.String(60))
    sort_order = db.Column(db.Integer, default=0)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    data = db.Column(db.JSON)  # original Excel row as {"column": value, ...}

    category_id = db.Column(db.Integer, db.ForeignKey("categories.id"), nullable=False)
    google_sheet_id = db.Column(db.Integer, db.ForeignKey("google_sheets.id"), nullable=True)

    @property
    def is_url(self):
        return bool(self.link) and self.link.strip().lower().startswith(("http://", "https://"))

    @property
    def data_keys(self):
        return list(self.data.keys()) if isinstance(self.data, dict) else []


GENERIC_FIELDS = [
    ("Program / Project Name", "name"),
    ("Link / Reference", "link"),
    ("Specific Instructions", "instructions"),
    ("Created By", "created_by"),
    ("Access", "access"),
    ("Status", "status"),
]


# ---------------------------------------------------------------------------
# Google Sheets registry + mirrored data
# ---------------------------------------------------------------------------

GOOGLE_DOC_DOMAINS = ("docs.google.com", "drive.google.com", "script.google.com")


def _extract_google_id(url):
    """Google document ID from a sharing URL, else None."""
    m = (
        re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
        or re.search(r"/document/d/([a-zA-Z0-9_-]+)", url)
        or re.search(r"/presentation/d/([a-zA-Z0-9_-]+)", url)
        or re.search(r"/forms/d/([a-zA-Z0-9_-]+)", url)
        or re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
        or re.search(r"/macros/s/([a-zA-Z0-9_-]+)", url)
        or re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    )
    return m.group(1) if m else None


def google_doc_kind(url):
    """Human label for a Google docs URL: Sheet, Doc, Slides, Form, Drive file."""
    u = (url or "").lower()
    if "/spreadsheets/" in u:
        return "Sheet"
    if "/document/" in u:
        return "Doc"
    if "/presentation/" in u:
        return "Slides"
    if "/forms/" in u:
        return "Form"
    if "script.google.com" in u:
        return "Apps Script"
    if "drive.google.com" in u:
        return "Drive file"
    return "Google"


def _resolve_image_formula(formula, row_idx, col_idx, ws):
    """Resolve a Sheets =IMAGE(...) formula into a real URL.

    Handles string literals and ENCODEURL(A1-style cell references) that point
    at text cells; anything else (operators, other functions) is evaluated
    via Excel's grid context when possible, else the ref is dropped. Returns
    None when no usable URL can be built."""
    from openpyxl.formula.translate import Translator  # noqa: F401  (ensures openpyxl formula support)
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

    body = formula[formula.upper().find("(") + 1:formula.rfind(")")]

    # Split top-level commas (respecting quotes and nested parens)
    parts, depth, in_str, cur = [], 0, False, ""
    for ch in body:
        if ch == '"':
            in_str = not in_str
            cur += ch
        elif not in_str and ch in "(\":":
            depth += 1
            cur += ch
        elif not in_str and ch == ")":
            depth -= 1
            cur += ch
        elif not in_str and ch == "," and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    if cur:
        parts.append(cur)
    if not parts:
        return None

    expr = parts[0].strip()
    out = []
    i = 0
    while i < len(expr):
        m = re.match(r'"([^"]*)"', expr[i:])
        if m:
            out.append(m.group(1))
            i += m.end()
            continue
        m = re.match(r"(?i)ENCODEURL\s*\(\s*([A-Z]+[0-9]+)\s*\)", expr[i:])
        ref = m.group(1) if m else None
        if m:
            col_letters = re.match(r"[A-Z]+", ref, re.IGNORECASE).group(0).upper()
            row_no = int(re.search(r"[0-9]+", ref).group(0))
            val = ws.cell(row=row_no, column=column_index_from_string(col_letters)).value
            if val is None:
                val = ""
            out.append(urllib.parse.quote(str(val), safe=""))
            i += m.end()
            continue
        m = re.match(r"&\s*", expr[i:])
        if m:
            i += m.end()
            continue
        # Unknown token (function call, number, whitespace...) -> skip it
        m = re.match(r"[A-Za-z_][A-Za-z0-9_.]*\s*\([^)]*\)", expr[i:])
        if m:
            i += m.end()
            continue
        if expr[i].isspace():
            i += 1
            continue
        m = re.match(r"[^\"]+?", expr[i:])
        if m and expr[i] in "+/?:=&.-_":
            out.append(expr[i])
            i += 1
        else:
            i += 1
    url = "".join(out).strip()
    return url or None


def _grids_from_xlsx(payload):
    """Every worksheet of a workbook as [(tab_title, [[cell, ...], ...]), ...].

    Cell values come from the data layer; cells whose formula is =IMAGE(...)
    become 'img:<resolved url>' so the UI can render them as images."""
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(payload), data_only=True)
    wb_f = openpyxl.load_workbook(io.BytesIO(payload), data_only=False)

    grids = []
    for ws, ws_f in zip(wb.worksheets, wb_f.worksheets):
        def cell_out(c, cf):
            if isinstance(cf.value, str) and cf.value.startswith("=") and "IMAGE(" in cf.value.upper():
                # referenced cells must be read from the DATA workbook (computed values),
                # not the formula workbook, or we'd encode the formula text itself
                url = _resolve_image_formula(cf.value, c.row, c.column, ws)
                return f"img:{url}" if url else None
            if c.value is None:
                return None
            if isinstance(c.value, datetime):
                return c.value.strftime("%b %d, %Y")
            if isinstance(c.value, date):
                return c.value.strftime("%b %d, %Y")
            if isinstance(c.value, float) and c.value.is_integer():
                return str(int(c.value))
            text = str(c.value).strip()
            return text or None

        grid = []
        for row in ws.iter_rows():
            out = []
            for c in row:
                cf = ws_f.cell(row=c.row, column=c.column)
                out.append(cell_out(c, cf))
            # trim trailing empties but keep at least one cell
            while out and out[-1] is None:
                out.pop()
            if not out:
                continue
            grid.append(out)
        # trim trailing empty rows
        while grid and not any(v is not None for v in grid[-1]):
            grid.pop()
        grids.append((ws.title, grid))
    return grids


class GoogleSheet(db.Model):
    """Registry of every Google link (Sheets/Docs/Slides/Forms/Drive) in the
    system, deduplicated by document ID so one file shared in N resources is a
    single row linked to all of them.

    Google **Sheets** additionally get every one of their tabs mirrored into
    the google_sheet_tabs/google_sheet_rows tables via the public XLSX export
    endpoint."""
    __tablename__ = "google_sheets"
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(300), nullable=False)
    url = db.Column(db.Text, nullable=False)
    doc_id = db.Column(db.String(128), unique=True, index=True)  # dedupe key
    doc_type = db.Column(db.String(20), default="Sheet")        # Sheet/Doc/Slides/Form/Drive file/Apps Script
    owner = db.Column(db.String(120))
    status = db.Column(db.String(60))
    notes = db.Column(db.Text)
    last_synced_at = db.Column(db.DateTime)
    last_sync_ok = db.Column(db.Boolean)
    last_sync_error = db.Column(db.Text)
    added_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    tabs = db.relationship("GoogleSheetTab", backref="sheet", cascade="all, delete-orphan",
                           order_by="GoogleSheetTab.tab_index")
    resources = db.relationship("Resource",
                                backref="google_sheet",
                                foreign_keys="Resource.google_sheet_id")

    @property
    def is_syncable(self):
        """Only Sheets can be pulled via the public XLSX export endpoint."""
        return self.doc_type == "Sheet" and bool(self.doc_id)

    @property
    def total_rows(self):
        """Mirrored data rows across all tabs (excl. one header row per tab)."""
        return sum(t.row_count for t in self.tabs)

    @property
    def edit_url(self):
        """Stable Google-side /edit link built from the document ID."""
        if not self.doc_id:
            return self.url
        return {
            "Sheet": f"https://docs.google.com/spreadsheets/d/{self.doc_id}/edit",
            "Doc": f"https://docs.google.com/document/d/{self.doc_id}/edit",
            "Slides": f"https://docs.google.com/presentation/d/{self.doc_id}/edit",
            "Form": f"https://docs.google.com/forms/d/{self.doc_id}/edit",
            "Drive file": f"https://drive.google.com/file/d/{self.doc_id}/view",
            "Apps Script": f"https://script.google.com/macros/s/{self.doc_id}/dev",
        }.get(self.doc_type, self.url)

    @property
    def resource_count(self):
        return len(self.resources)

    def sync_data(self):
        """Fetch ALL tabs of the workbook via the public XLSX export endpoint
        and store them in google_sheet_tabs + google_sheet_rows.

        Uses the XLSX (not CSV) so =IMAGE(...) formulas come along: the URL is
        resolved from the formula (literals + ENCODEURL(cellref) refs) and the
        mirrored cell stores 'img:<url>', which the data viewer renders as an
        actual image. Returns True on success."""
        if not self.is_syncable:
            return False
        from flask import current_app
        from . import db as _db

        xlsx_url = f"https://docs.google.com/spreadsheets/d/{self.doc_id}/export?format=xlsx"
        try:
            req = urllib.request.Request(xlsx_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as resp:
                payload = resp.read()
            grids = _grids_from_xlsx(payload)
        except Exception as e:
            self.last_synced_at = datetime.utcnow()
            self.last_sync_ok = False
            self.last_sync_error = str(e)[:500]
            _db.session.commit()
            return False

        # Refresh from a clean slate (tabs + rows are fully derived from the sheet)
        GoogleSheetRow.query.filter_by(sheet_id=self.id).delete()
        GoogleSheetTab.query.filter_by(sheet_id=self.id).delete()
        _db.session.flush()
        total = 0
        for tab_index, (tab_title, grid) in enumerate(grids):
            tab = GoogleSheetTab(
                sheet_id=self.id,
                title=tab_title or f"Tab {tab_index + 1}",
                tab_index=tab_index,
                row_count=max(len(grid) - 1, 0),  # exclude the header row
            )
            _db.session.add(tab)
            _db.session.flush()
            for i, row in enumerate(grid):
                _db.session.add(GoogleSheetRow(
                    sheet_id=self.id,
                    tab_id=tab.id,
                    sheet_index=i,
                    row_json={str(j): c for j, c in enumerate(row)},
                    col_count=max(len(row), 1),
                ))
            total += max(len(grid) - 1, 0)
        self.last_synced_at = datetime.utcnow()
        self.last_sync_ok = True
        self.last_sync_error = None
        _db.session.commit()
        current_app.logger.info("Synced %d tabs / %d rows from sheet '%s'",
                                len(grids), total, self.title)
        return True

    @classmethod
    def upsert_from_url(cls, url, title=None, **extra):
        """Find-or-create by document ID (falling back to the raw URL),
        merging in optional fields (owner/status/notes). Returns (row, created)."""
        url = (url or "").strip()
        if not url.lower().startswith(("http://", "https://")):
            return None, False
        if not any(domain in url for domain in GOOGLE_DOC_DOMAINS):
            return None, False
        doc_id = _extract_google_id(url)
        kind = google_doc_kind(url)
        row = None
        if doc_id:
            row = cls.query.filter_by(doc_id=doc_id).first()
        if row is None:
            row = cls.query.filter_by(url=url).first()
        created = row is None
        if created:
            row = cls(
                title=(title or "Untitled Google file").strip(),
                url=url,
                doc_id=doc_id,
                doc_type=kind,
            )
            db.session.add(row)
            db.session.flush()  # populate row.id so callers can link to it right away
        elif title:
            row.title = title.strip()
        for key, val in extra.items():
            if val not in (None, ""):
                setattr(row, key, val)
        if row.doc_type != kind and kind != "Google":
            row.doc_type = kind
        return row, created

    @classmethod
    def backfill_from_resources(cls):
        """Register every Google link found on existing resources and pull
        each sheet's data on first sight. Returns the number of rows created."""
        created = 0
        for r in Resource.query.all():
            if not r.is_url or "google." not in (r.link or ""):
                continue
            row, was_created = cls.upsert_from_url(r.link, title=r.name)
            if was_created:
                created += 1
            if row and r.google_sheet_id != row.id:
                r.google_sheet_id = row.id
            if row and row.is_syncable and row.last_synced_at is None:
                row.sync_data()
        db.session.commit()
        return created


class GoogleSheetTab(db.Model):
    """One tab (worksheet) of a mirrored Google Sheet."""
    __tablename__ = "google_sheet_tabs"
    id = db.Column(db.Integer, primary_key=True)
    sheet_id = db.Column(db.Integer, db.ForeignKey("google_sheets.id"), nullable=False)
    title = db.Column(db.String(300), nullable=False)
    tab_index = db.Column(db.Integer, default=0)          # position in the workbook
    row_count = db.Column(db.Integer, default=0)          # mirrored data rows (excl. header)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    rows = db.relationship("GoogleSheetRow", backref="tab", cascade="all, delete-orphan",
                           order_by="GoogleSheetRow.sheet_index")


class GoogleSheetRow(db.Model):
    """One mirrored row of one tab of a Google Sheet: row_json = {"0": "A-cell", ...}."""
    __tablename__ = "google_sheet_rows"
    id = db.Column(db.Integer, primary_key=True)
    sheet_id = db.Column(db.Integer, db.ForeignKey("google_sheets.id"), nullable=False)
    tab_id = db.Column(db.Integer, db.ForeignKey("google_sheet_tabs.id"), nullable=False)
    sheet_index = db.Column(db.Integer, nullable=False)   # 0-based position within the tab
    col_count = db.Column(db.Integer, default=0)
    row_json = db.Column(db.JSON)                          # {"0": value, "1": value, ...}

    @property
    def values(self):
        return [self.row_json.get(str(i), "") for i in range(self.col_count)] if self.row_json else []
