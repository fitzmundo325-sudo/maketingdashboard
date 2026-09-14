from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required

from . import db, STATUS_CHOICES
from . import sort_brands
from .models import Brand, Category, Resource, GENERIC_FIELDS
from .excel_mappings import mapping_for

views = Blueprint("views", __name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_or_create_brand(name):
    brand = Brand.query.filter_by(name=name).first()
    if not brand:
        from . import BRAND_COLORS
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


# Labels used by resources imported from HYG MARKETING MASTER DATABASE.xlsx
MASTER_DATA_LABELS = {
    "Program / Project Name": "name",
    "Spreadsheet Link": "link",
    "Specific Instructions": "instructions",
    "Date Created": None,          # informational only
    "Created By": "created_by",
    "Access": "access",
    "Status": "status",
    "No.": None,                   # mapped to ref_no below
}


def rebuild_from_data(resource):
    """Sync resource columns from data using the sheet mapping (or generic fields)."""
    if not resource.data:
        return
    mapping = mapping_for(resource.category.name)
    if not mapping:
        # Master-imported resources carry the workbook's own labels
        if any(label in resource.data for label in MASTER_DATA_LABELS):
            for label, attr in MASTER_DATA_LABELS.items():
                val = resource.data.get(label)
                if attr and val is not None:
                    setattr(resource, attr, val)
            ref_val = resource.data.get("No.")
            if ref_val is not None:
                resource.ref_no = ref_val
            if not resource.name:
                resource.name = "(Untitled)"
            return
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


# ---------------------------------------------------------------------------
# Routes - Dashboard
# ---------------------------------------------------------------------------

@views.route("/")
@login_required
def dashboard():
    brands = sort_brands(Brand.query.all())
    total_resources = Resource.query.count()
    total_categories = Category.query.filter(Category.parent_id.is_(None)).count()
    return render_template(
        "dashboard.html",
        brands=brands,
        total_resources=total_resources,
        total_categories=total_categories,
    )


@views.route("/brand/<int:brand_id>")
@login_required
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


@views.route("/search")
@login_required
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

@views.route("/category/<int:category_id>")
@login_required
def category_view(category_id):
    category = Category.query.get_or_404(category_id)
    q = request.args.get("q", "").strip()
    status_filter = request.args.get("status", "").strip()
    filtering = bool(q or status_filter)

    def _match(r):
        if q and q.lower() not in " ".join(
            filter(None, [r.name, r.link, r.instructions, r.created_by, r.ref_no])
        ).lower():
            return False
        if status_filter and (r.status or "").lower() != status_filter.lower():
            return False
        return True

    children = category.children or []
    tabs = children if (children and category.parent_id is not None) else None
    tab_cols = {t.id: (template_keys_for(t), field_labels_for(t)) for t in (tabs or [])}

    if filtering:
        for t in tabs or []:
            t.filtered_resources = [r for r in t.resources if _match(r)]
        if tabs is None:
            for c in children:
                c.filtered_resources = [r for r in c.resources if _match(r)]
        resources = [r for r in category.resources if _match(r)]
    else:
        resources = category.resources

    return render_template(
        "category.html",
        category=category,
        categories=children if tabs is None else None,
        tabs=tabs,
        tab_cols=tab_cols,
        col_keys=template_keys_for(category) if tabs is None else [],
        field_labels=field_labels_for(category) if tabs is None else {},
        resources=resources,
        q=q,
        status_filter=status_filter,
        filtering=filtering,
        status_choices=STATUS_CHOICES,
    )


@views.route("/category/<int:category_id>/resource/new", methods=["GET", "POST"])
@login_required
def resource_new(category_id):
    category = Category.query.get_or_404(category_id)
    fallback = url_for("views.brand_view", brand_id=category.brand_id)
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


@views.route("/resource/<int:resource_id>/edit", methods=["GET", "POST"])
@login_required
def resource_edit(resource_id):
    resource = Resource.query.get_or_404(resource_id)
    fallback = url_for("views.brand_view", brand_id=resource.category.brand_id)
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


@views.route("/resource/<int:resource_id>/delete", methods=["POST"])
@login_required
def resource_delete(resource_id):
    resource = Resource.query.get_or_404(resource_id)
    brand_id = resource.category.brand_id
    name = resource.name
    db.session.delete(resource)
    db.session.commit()
    flash(f"Deleted resource '{name}'.", "info")
    return redirect(url_for("views.brand_view", brand_id=brand_id))


@views.route("/brand/<int:brand_id>/category/new", methods=["POST"])
@login_required
def category_new(brand_id):
    brand = Brand.query.get_or_404(brand_id)
    name = request.form.get("name", "").strip()
    if name:
        get_or_create_category(brand, name)
        db.session.commit()
        flash(f"Added category '{name}'.", "success")
    return redirect(url_for("views.brand_view", brand_id=brand.id))


@views.route("/category/<int:category_id>/project/new", methods=["POST"])
@login_required
def category_project_new(category_id):
    category = Category.query.get_or_404(category_id)
    name = request.form.get("name", "").strip()
    if name:
        get_or_create_category(category.brand, name, parent=category)
        db.session.commit()
        flash(f"Added project '{name}'.", "success")
    return redirect(url_for("views.category_view", category_id=category.id))
