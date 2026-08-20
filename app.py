import os
import json
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from flask_sqlalchemy import SQLAlchemy

from excel_mappings import mapping_for

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__)
app.config['SECRET_KEY'] = 'hyg-marketing-hub-dev-key'
app.config['SQLALCHEMY_DATABASE_URI'] = f"sqlite:///{os.path.join(BASE_DIR, 'marketing_hub.db')}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

db = SQLAlchemy(app)

BRAND_COLORS = {
    "GOLDILOCKS": "#D4AF37",
    "SAVORY": "#000080",
    "ICEBERGS": "#00A0DF",
    "TATERS": "#E4002B",
    "CHATIME": "#6E2C8C",
    "ELEVATE": "#2C3E50",
}

STATUS_CHOICES = ["In Progress", "Completed", "For Review", "On Hold", "Not Started", "N/A"]


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

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

    __table_args__ = (db.UniqueConstraint('name', 'brand_id', name='uq_category_brand'),)

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

    @property
    def is_url(self):
        return bool(self.link) and self.link.strip().lower().startswith(("http://", "https://"))

    @property
    def data_keys(self):
        return list(self.data.keys()) if isinstance(self.data, dict) else []


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_or_create_brand(name):
    brand = Brand.query.filter_by(name=name).first()
    if not brand:
        brand = Brand(name=name, color=BRAND_COLORS.get(name, "#333333"))
        db.session.add(brand)
        db.session.flush()
    return brand


def get_or_create_category(brand, name, parent=None):
    parent_id = parent.id if parent else None
    cat = Category.query.filter_by(brand_id=brand.id, name=name, parent_id=parent_id).first()
    if not cat:
        cat = Category(name=name, brand_id=brand.id, parent_id=parent_id)
        db.session.add(cat)
        db.session.flush()
    return cat


def seed_from_json(path):
    """Populate the database from the parsed Excel JSON, only if empty."""
    if Brand.query.first():
        return  # already seeded

    with open(path) as f:
        data = json.load(f)

    for brand_name, entries in data.items():
        brand = get_or_create_brand(brand_name)
        order_counter = {}
        for entry in entries:
            cat_name = entry.get("category") or "General"
            cat = get_or_create_category(brand, cat_name)
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


# ---------------------------------------------------------------------------
# Template context (sidebar data on every page)
# ---------------------------------------------------------------------------

@app.context_processor
def inject_sidebar():
    brands = Brand.query.order_by(Brand.name).all()
    active_brand_id = request.view_args.get("brand_id") if request.view_args else None
    css_path = os.path.join(app.static_folder, "css", "style.css")
    try:
        css_version = str(int(os.path.getmtime(css_path)))
    except OSError:
        css_version = "0"
    return {
        "sidebar_brands": brands,
        "active_brand_id": active_brand_id,
        "css_version": css_version,
        "field_value": field_value,
    }


def field_value(resource, key):
    """Cell value for a column key: from resource.data, else the generic column."""
    if resource.data:
        return resource.data.get(key, "")
    attr_map = {label: attr for label, attr in GENERIC_FIELDS}
    attr = attr_map.get(key)
    if attr:
        return getattr(resource, attr) or ""
    return ""


@app.after_request
def prevent_css_staleness(resp):
    if resp.content_type and resp.content_type.startswith("text/html"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


# ---------------------------------------------------------------------------
# Routes - Dashboard
# ---------------------------------------------------------------------------

@app.route("/")
def dashboard():
    brands = Brand.query.order_by(Brand.name).all()
    total_resources = Resource.query.count()
    total_categories = Category.query.filter(Category.parent_id.is_(None)).count()
    return render_template(
        "dashboard.html",
        brands=brands,
        total_resources=total_resources,
        total_categories=total_categories,
    )


@app.route("/brand/<int:brand_id>")
def brand_view(brand_id):
    brand = Brand.query.get_or_404(brand_id)
    q = request.args.get("q", "").strip()
    status_filter = request.args.get("status", "").strip()

    categories = [c for c in brand.categories if c.parent_id is None]
    if q or status_filter:
        categories = []
        for cat in brand.categories:
            filtered = []
            for r in cat.resources:
                if q and q.lower() not in " ".join(filter(None, [r.name, r.link, r.instructions, r.created_by])).lower():
                    continue
                if status_filter and (r.status or "").lower() != status_filter.lower():
                    continue
                filtered.append(r)
            if filtered:
                # build a lightweight object clone with filtered resources for template use
                cat.filtered_resources = filtered
                categories.append(cat)
    else:
        for cat in categories:
            cat.filtered_resources = cat.resources

    return render_template(
        "brand.html",
        brand=brand,
        categories=categories,
        q=q,
        status_filter=status_filter,
        status_choices=STATUS_CHOICES,
    )


@app.route("/search")
def search():
    q = request.args.get("q", "").strip()
    results = []
    if q:
        like = f"%{q}%"
        results = (
            Resource.query.filter(
                db.or_(
                    Resource.name.ilike(like),
                    Resource.link.ilike(like),
                    Resource.instructions.ilike(like),
                    Resource.created_by.ilike(like),
                )
            )
            .order_by(Resource.name)
            .all()
        )
    return render_template("search.html", q=q, results=results)


# ---------------------------------------------------------------------------
# Routes - Resource CRUD
# ---------------------------------------------------------------------------

@app.route("/category/<int:category_id>")
def category_view(category_id):
    category = Category.query.get_or_404(category_id)
    children = category.children or []
    tabs = children if (children and category.parent_id is not None) else None
    tab_cols = {t.id: (template_keys_for(t), field_labels_for(t)) for t in (tabs or [])}
    return render_template(
        "category.html",
        category=category,
        categories=children if tabs is None else None,
        tabs=tabs,
        tab_cols=tab_cols,
        col_keys=template_keys_for(category) if tabs is None else [],
        field_labels=field_labels_for(category) if tabs is None else {},
        status_choices=STATUS_CHOICES,
    )


@app.route("/category/<int:category_id>/resource/new", methods=["GET", "POST"])
def resource_new(category_id):
    category = Category.query.get_or_404(category_id)
    fallback = url_for("brand_view", brand_id=category.brand_id)
    if request.method == "POST":
        submitted = collect_data_fields(request.form)
        if any(k.startswith("data_") for k in request.form.keys()):
            resource = Resource(data=submitted, category=category, sort_order=len(category.resources) + 1)
            db.session.add(resource)
            rebuild_from_data(resource)
        else:
            resource = Resource(
                name=request.form.get("name", "").strip() or "(Untitled)",
                link=request.form.get("link", "").strip(),
                instructions=request.form.get("instructions", "").strip(),
                created_by=request.form.get("created_by", "").strip(),
                access=request.form.get("access", "").strip(),
                status=request.form.get("status", "").strip(),
                category_id=category.id,
                sort_order=len(category.resources) + 1,
            )
            db.session.add(resource)
        db.session.commit()
        flash(f"Added resource '{resource.name}' to {category.name}.", "success")
        return redirect(submitted_back_url(fallback))
    return render_template(
        "resource_form.html",
        category=category,
        resource=None,
        status_choices=STATUS_CHOICES,
        template_keys=template_keys_for(category),
        field_labels=field_labels_for(category),
        back_url=back_url(fallback),
    )


@app.route("/resource/<int:resource_id>/edit", methods=["GET", "POST"])
def resource_edit(resource_id):
    resource = Resource.query.get_or_404(resource_id)
    fallback = url_for("brand_view", brand_id=resource.category.brand_id)
    if request.method == "POST":
        if any(k.startswith("data_") for k in request.form.keys()):
            resource.data = collect_data_fields(request.form)
            rebuild_from_data(resource)
        else:
            resource.name = request.form.get("name", "").strip() or "(Untitled)"
            resource.link = request.form.get("link", "").strip()
            resource.instructions = request.form.get("instructions", "").strip()
            resource.created_by = request.form.get("created_by", "").strip()
            resource.access = request.form.get("access", "").strip()
            resource.status = request.form.get("status", "").strip()
        db.session.commit()
        flash(f"Updated resource '{resource.name}'.", "success")
        return redirect(submitted_back_url(fallback))
    template = "_resource_form.html" if request.args.get("fragment") else "resource_form.html"
    return render_template(
        template,
        category=resource.category,
        resource=resource,
        status_choices=STATUS_CHOICES,
        template_keys=template_keys_for(resource.category),
        field_labels=field_labels_for(resource.category),
        back_url=back_url(fallback),
    )


def back_url(fallback):
    """URL of the page the user came from, if internal; else the fallback."""
    ref = (request.referrer or "").split("#", 1)[0]
    if ref.startswith(request.host_url):
        return ref
    return fallback


def submitted_back_url(fallback):
    """Return-to URL after a form POST: hidden 'next' field, else the referrer, else fallback."""
    nxt = (request.form.get("next") or "").split("#", 1)[0]
    if nxt.startswith(request.host_url):
        return nxt
    return back_url(fallback)


def collect_data_fields(form):
    """Dynamic Excel fields submitted as data_<column> -> value (empty -> None, key kept)."""
    out = {}
    for key, value in form.items():
        if key.startswith("data_"):
            field_key = key[len("data_"):]
            if isinstance(value, str):
                value = value.strip() or None
            out[field_key] = value
    return out


GENERIC_FIELDS = [
    ("Program / Project Name", "name"),
    ("Link / Reference", "link"),
    ("Specific Instructions", "instructions"),
    ("Created By", "created_by"),
    ("Access", "access"),
    ("Status", "status"),
]


def generic_data_for(resource):
    """Build a modular data dict from a resource's legacy columns (no Excel mapping)."""
    out = {}
    for label, attr in GENERIC_FIELDS:
        val = getattr(resource, attr)
        if val not in (None, ""):
            out[label] = val
    return out


def template_keys_for(category):
    """Column keys for the add-resource form: the most complete sibling's keys, else generic keys."""
    best = []
    for sibling in category.resources:
        if sibling.data and len(sibling.data_keys) > len(best):
            best = sibling.data_keys
    return best or [label for label, _ in GENERIC_FIELDS]


def field_labels_for(category):
    mapping = mapping_for(category.name)
    return mapping.get("labels", {}) if mapping else {}


def rebuild_from_data(resource):
    """Sync resource columns from data using the sheet mapping (or generic fields)."""
    if not resource.data:
        return
    mapping = mapping_for(resource.category.name)
    if not mapping:
        for label, attr in GENERIC_FIELDS:
            val = resource.data.get(label)
            if val is not None:
                setattr(resource, attr, val)
        if not resource.name:
            resource.name = "(Untitled)"
        return
    name_key = mapping.get("name")
    if name_key and resource.data.get(name_key):
        resource.name = resource.data[name_key]
    ref_key = mapping.get("ref")
    if ref_key and resource.data.get(ref_key):
        resource.ref_no = resource.data[ref_key]
    labels = mapping.get("labels", {})
    parts = []
    if mapping.get("all_columns"):
        skip = {name_key, mapping.get("created_by"), mapping.get("status")}
        for key, val in resource.data.items():
            if val in (None, "") or key in skip:
                continue
            parts.append(f"{labels.get(key, key)}: {val}")
    else:
        for key in mapping["instructions"]:
            val = resource.data.get(key)
            if val not in (None, ""):
                parts.append(f"{labels.get(key, key)}: {val}")
        notes_key = mapping.get("notes")
        if notes_key:
            note = resource.data.get(notes_key)
            if note not in (None, ""):
                parts.append(f"{labels.get(notes_key, 'NOTE')}: {note}")
    resource.instructions = "\n".join(parts) or None
    created_by_key = mapping.get("created_by")
    if created_by_key:
        resource.created_by = resource.data.get(created_by_key)
    status_key = mapping.get("status")
    if status_key:
        resource.status = resource.data.get(status_key)
    if not resource.name or resource.name == "(Untitled)":
        resource.name = resource.data.get(mapping.get("ref")) or "(Untitled)"


@app.route("/resource/<int:resource_id>/delete", methods=["POST"])
def resource_delete(resource_id):
    resource = Resource.query.get_or_404(resource_id)
    brand_id = resource.category.brand_id
    name = resource.name
    db.session.delete(resource)
    db.session.commit()
    flash(f"Deleted resource '{name}'.", "info")
    return redirect(url_for("brand_view", brand_id=brand_id))


@app.route("/brand/<int:brand_id>/category/new", methods=["POST"])
def category_new(brand_id):
    brand = Brand.query.get_or_404(brand_id)
    name = request.form.get("name", "").strip()
    if name:
        get_or_create_category(brand, name)
        db.session.commit()
        flash(f"Added category '{name}'.", "success")
    return redirect(url_for("brand_view", brand_id=brand.id))


@app.route("/category/<int:category_id>/project/new", methods=["POST"])
def category_project_new(category_id):
    category = Category.query.get_or_404(category_id)
    name = request.form.get("name", "").strip()
    if name:
        get_or_create_category(category.brand, name, parent=category)
        db.session.commit()
        flash(f"Added project '{name}'.", "success")
    return redirect(url_for("category_view", category_id=category.id))


# ---------------------------------------------------------------------------
# JSON API (lightweight, useful for future integrations e.g. GAS/PHP tools)
# ---------------------------------------------------------------------------

@app.route("/api/brands")
def api_brands():
    brands = Brand.query.order_by(Brand.name).all()
    return jsonify([{"id": b.id, "name": b.name, "resource_count": b.resource_count} for b in brands])


@app.route("/api/brand/<int:brand_id>/resources")
def api_brand_resources(brand_id):
    brand = Brand.query.get_or_404(brand_id)
    out = []
    for cat in brand.categories:
        for r in cat.resources:
            out.append({
                "id": r.id,
                "category": cat.name,
                "name": r.name,
                "link": r.link,
                "status": r.status,
                "created_by": r.created_by,
                "access": r.access,
            })
    return jsonify(out)


# ---------------------------------------------------------------------------
# CLI helper: init + seed DB
# ---------------------------------------------------------------------------

@app.cli.command("init-db")
def init_db_command():
    db.create_all()
    seed_from_json(os.path.join(BASE_DIR, "parsed_data.json"))
    print("Database initialized and seeded.")


def ensure_db():
    with app.app_context():
        db.create_all()
        migrate_add_data_column()
        seed_path = os.path.join(BASE_DIR, "parsed_data.json")
        if os.path.exists(seed_path):
            seed_from_json(seed_path)


def migrate_add_data_column():
    """Add the `data` JSON column to an existing database (create_all won't alter tables)."""
    from sqlalchemy import text
    cols = {row[1] for row in db.session.execute(text("PRAGMA table_info(resources)")).fetchall()}
    if "data" not in cols:
        db.session.execute(text("ALTER TABLE resources ADD COLUMN data TEXT"))
        db.session.commit()
        print("Migrated: added resources.data column.")


if __name__ == "__main__":
    ensure_db()
    app.run(debug=True, host="0.0.0.0", port=5000)
