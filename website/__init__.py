import os
import json
import secrets
from pathlib import Path

from flask import Flask, request
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_caching import Cache

db = SQLAlchemy()
cache = Cache()

DB_NAME = "marketing_hub.db"

BASE_DIR = Path(__file__).resolve().parent.parent
PACKAGE_DIR = Path(__file__).resolve().parent

BRAND_COLORS = {
    "GOLDILOCKS": "#D4AF37",
    "SAVORY": "#000080",
    "ICEBERGS": "#00A0DF",
    "TATERS": "#E4002B",
    "CHATIME": "#6E2C8C",
    "ELEVATE": "#2C3E50",
}

# Brand logos (files under static/img/brands/). Brands not listed fall back to a letter avatar.
BRAND_LOGO_FILES = {
    "GOLDILOCKS": "img/brands/goldilocks.png",
    "ICEBERGS": "img/brands/icebergs.png",
    "TATERS": "img/brands/taters.png",
    "CHATIME": "img/brands/chatime.png",
}


def brand_logo(brand_name):
    """Versioned static URL for a brand's logo, or None if it has no logo file."""
    filename = BRAND_LOGO_FILES.get((brand_name or "").strip().upper())
    if not filename:
        return None
    return static_url(filename)


def static_url(filename):
    """Static URL with an mtime-based version (?v=...) so browser caches
    pick up replaced images/CSS immediately - no Ctrl+F5 needed."""
    from flask import current_app, url_for

    try:
        url = url_for("static", filename=filename)
    except RuntimeError:
        return None
    file_path = os.path.join(current_app.static_folder, *filename.split("/"))
    try:
        version = str(int(os.path.getmtime(file_path)))
    except OSError:
        version = "0"
    return f"{url}?v={version}"

STATUS_CHOICES = ["In Progress", "Completed", "For Review", "On Hold", "Not Started", "N/A"]

# Brands pinned to the top of menus, in this order; all others follow alphabetically
PINNED_BRAND_ORDER = ["GOLDILOCKS"]


def sort_brands(brands):
    """Sort brands with pinned ones first (GOLDILOCKS on top), rest alphabetical."""
    def key(brand):
        name = (brand.name or "").upper()
        pinned_index = PINNED_BRAND_ORDER.index(name) if name in PINNED_BRAND_ORDER else len(PINNED_BRAND_ORDER)
        return (pinned_index, name)
    return sorted(brands, key=key)


def _load_local_env():
    """Load key=value pairs from a root .env file (same convention as cficountsystem)."""
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        if not key:
            continue

        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def _ensure_resource_data_column():
    """Add the `data` JSON column to an existing resources table (create_all won't alter)."""
    from sqlalchemy import text

    with db.engine.connect() as conn:
        existing_columns = {
            row[1]
            for row in conn.execute(text("PRAGMA table_info('resources')")).fetchall()
        }
        if "data" not in existing_columns:
            conn.execute(text("ALTER TABLE resources ADD COLUMN data TEXT"))
        conn.commit()


def _ensure_user_role_column():
    from sqlalchemy import text

    with db.engine.connect() as conn:
        existing_columns = {
            row[1]
            for row in conn.execute(text("PRAGMA table_info('user')")).fetchall()
        }
        if "role" not in existing_columns:
            conn.execute(text("ALTER TABLE user ADD COLUMN role VARCHAR(100) DEFAULT 'Admin'"))
        conn.commit()


def _ensure_google_sheets_tables():
    """Create the google_sheets/google_sheet_rows tables + resources.google_sheet_id
    FK on existing DBs (db.create_all() adds new tables but won't alter existing ones)."""
    from sqlalchemy import text

    with db.engine.connect() as conn:
        existing_columns = {
            row[1]
            for row in conn.execute(text("PRAGMA table_info('resources')")).fetchall()
        }
        if "google_sheet_id" not in existing_columns:
            conn.execute(text("ALTER TABLE resources ADD COLUMN google_sheet_id INTEGER"))
        # multi-tab support: rows move under google_sheet_tabs
        row_cols = {
            row[1]
            for row in conn.execute(text("PRAGMA table_info('google_sheet_rows')")).fetchall()
        }
        if row_cols and "tab_id" not in row_cols:
            conn.execute(text("ALTER TABLE google_sheet_rows ADD COLUMN tab_id INTEGER"))
        conn.commit()


def _seed_from_json(path):
    """Populate the database from the parsed Excel JSON, only if empty."""
    from .models import Brand, Category, Resource

    if Brand.query.first():
        return  # already seeded

    if not os.path.exists(path):
        return

    with open(path) as f:
        data = json.load(f)

    for brand_name, entries in data.items():
        brand = Brand.query.filter_by(name=brand_name).first()
        if not brand:
            brand = Brand(name=brand_name, color=BRAND_COLORS.get(brand_name, "#333333"))
            db.session.add(brand)
            db.session.flush()

        order_counter = {}
        for entry in entries:
            cat_name = entry.get("category") or "General"
            parent_id = None
            cat = Category.query.filter_by(brand_id=brand.id, name=cat_name, parent_id=parent_id).first()
            if not cat:
                cat = Category(name=cat_name, brand_id=brand.id, parent_id=parent_id)
                db.session.add(cat)
                db.session.flush()

            order_counter[cat.id] = order_counter.get(cat.id, 0) + 1
            resource = Resource(
                ref_no=entry.get("no"),
                name=entry.get("name") or "(Untitled)",
                link=entry.get("link"),
                instructions=entry.get("instructions"),
                created_by=entry.get("created_by"),
                access=entry.get("access"),
                status=entry.get("status"),
                sort_order=order_counter[cat.id],
                category=cat,
            )
            db.session.add(resource)

    db.session.commit()
    print("Seeded database from parsed_data.json.")


def create_app():
    _load_local_env()
    app = Flask(__name__)
    db_name = os.environ.get("MH_APP_DB_NAME", DB_NAME)
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{BASE_DIR / db_name}"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0

    configured_secret_key = os.environ.get("MH_SECRET_KEY")
    app.config["SECRET_KEY"] = configured_secret_key or secrets.token_hex(32)

    db.init_app(app)
    cache.init_app(app, config={
        "CACHE_TYPE": "SimpleCache",
        "CACHE_DEFAULT_TIMEOUT": 300,
    })

    from .views import views
    from .auth import auth
    from .api_handles import api_handles

    app.register_blueprint(auth, url_prefix="/")
    app.register_blueprint(views, url_prefix="/")
    app.register_blueprint(api_handles, url_prefix="/apis")

    login_manager = LoginManager()
    login_manager.login_view = "auth.login"
    login_manager.login_message = "Please login to access this page."
    login_manager.login_message_category = "info"
    login_manager.init_app(app)

    from .models import User

    @login_manager.user_loader
    def load_user(id):
        return User.query.get(int(id))

    from .models import (Brand, Category, Resource, GoogleSheet, GoogleSheetTab,
                         GoogleSheetRow, User)  # noqa: F401

    with app.app_context():
        db.create_all()
        _ensure_resource_data_column()
        _ensure_user_role_column()
        _ensure_google_sheets_tables()
        _seed_from_json(str(BASE_DIR / "parsed_data.json"))
        # First run after this feature: register every Google link already on resources
        if GoogleSheet.query.first() is None:
            created = GoogleSheet.backfill_from_resources()
            if created:
                print(f"Seeded google_sheets registry with {created} links from existing resources.")
        # One-time migration: rows synced before multi-tab support carry no tab.
        # Re-pull any sheets whose tab structure is missing.
        stale = [s for s in GoogleSheet.query.filter_by(doc_type="Sheet").all()
                 if s.last_sync_ok and not s.tabs]
        if stale:
            print(f"Re-syncing {len(stale)} sheet(s) for multi-tab support...")
            for s in stale:
                s.sync_data()
        print("Created database!")

    @app.context_processor
    def inject_sidebar():
        from .models import Brand

        brands = sort_brands(Brand.query.all())
        active_brand_id = request.view_args.get("brand_id") if request.view_args else None
        return {
            "sidebar_brands": brands,
            "active_brand_id": active_brand_id,
            "field_value": field_value,
            "brand_logo": brand_logo,
            "static_url": static_url,
            "sheet_data_url": sheet_data_url,
        }

    @app.after_request
    def prevent_css_staleness(resp):
        if resp.content_type and resp.content_type.startswith("text/html"):
            resp.headers["Cache-Control"] = "no-store"
        return resp

    @app.cli.command("init-db")
    def init_db_command():
        """Initialize and seed the database."""
        with app.app_context():
            db.create_all()
            _seed_from_json(str(BASE_DIR / "parsed_data.json"))
        print("Database initialized and seeded.")

    @app.cli.command("create-admin")
    def create_admin_command():
        """Create an admin user (prompts for username/password)."""
        import getpass
        from werkzeug.security import generate_password_hash
        from .models import User

        with app.app_context():
            username = input("Username: ").strip()
            if not username:
                print("Username is required.")
                return
            if User.query.filter_by(username=username).first():
                print(f"User '{username}' already exists.")
                return
            password = getpass.getpass("Password: ")
            user = User(
                username=username,
                role="Admin",
                password=generate_password_hash(password),
            )
            db.session.add(user)
            db.session.commit()
            print(f"Admin user '{username}' created.")

    @app.cli.command("backfill-google-sheets")
    def backfill_google_sheets_command():
        """Scan all resources, register new Google links, and pull sheet data."""
        with app.app_context():
            created = GoogleSheet.backfill_from_resources()
        print(f"Backfill complete: {created} new Google link(s) registered.")

    @app.cli.command("sync-google-sheets")
    def sync_google_sheets_command():
        """Re-pull data (all tabs) for every registered Google Sheet."""
        with app.app_context():
            sheets = GoogleSheet.query.filter_by(doc_type="Sheet").all()
            ok = sum(1 for s in sheets if s.sync_data())
        print(f"Synced {ok}/{len(sheets)} sheets (all tabs).")
    @app.cli.command("import-master")
    def import_master_command():
        """Import HYG MARKETING MASTER DATABASE.xlsx (per-sheet brand sections)."""
        import click
        from .import_master import import_master as run_master_import

        xlsx_path = click.prompt(
            "Path to the master workbook",
            default=str(BASE_DIR / "HYG MARKETING MASTER DATABASE.xlsx"),
        )
        replace = click.confirm("Remove existing resources of these brands first?", default=False)
        run_master_import(xlsx_path, replace=replace)

    return app


def field_value(resource, key):
    """Cell value for a column key: from resource.data, else the generic column."""
    from .models import GENERIC_FIELDS

    if resource.data:
        return resource.data.get(key, "")
    attr_map = {label: attr for label, attr in GENERIC_FIELDS}
    # Labels used by the master workbook import
    attr_map.update({"Spreadsheet Link": "link", "No.": "ref_no"})
    attr = attr_map.get(key)
    if attr:
        return getattr(resource, attr) or ""
    return ""


def sheet_data_url(resource):
    """Internal URL of a resource's Google Sheet data page, or None when the
    link isn't a Google URL that has a registry entry."""
    if not resource.is_url or "google." not in (resource.link or ""):
        return None
    from .models import GoogleSheet, _extract_google_id
    doc_id = _extract_google_id(resource.link)
    row = GoogleSheet.query.filter_by(doc_id=doc_id).first() if doc_id else None
    return f"/sheets/{row.id}" if row else None
