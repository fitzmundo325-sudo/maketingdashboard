from datetime import datetime

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
