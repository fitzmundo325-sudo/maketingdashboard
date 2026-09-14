from flask import Blueprint, jsonify
from flask_login import login_required

from . import sort_brands
from .models import Brand, Resource

api_handles = Blueprint("api_handles", __name__)


@api_handles.route("/brands")
@login_required
def api_brands():
    brands = sort_brands(Brand.query.all())
    return jsonify([{"id": b.id, "name": b.name, "resource_count": b.resource_count} for b in brands])


@api_handles.route("/brand/<int:brand_id>/resources")
@login_required
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
