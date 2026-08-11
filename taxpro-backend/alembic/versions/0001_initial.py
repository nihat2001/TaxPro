"""initial: profiles, portals, vector_documents + 31 portal seed

Revision ID: 0001
Revises:
Create Date: 2026-08-11
"""

import uuid

import sqlalchemy as sa
from pgvector.sqlalchemy import Vector
from sqlalchemy.dialects import postgresql

from alembic import op

revision = "0001"
down_revision = None
branch_labels = None
depends_on = None


def _portal_id(name: str) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f"https://taxpro.portal/{name}"))


_STANDARD_MAPPING = {
    "ad_soyad": "input[name='fullname']",
    "voen": "input[name='voen']",
    "nomre": "input[name='phone']",
    "asan_id": "input[name='asanid']",
    "pin1": "input[name='pin1']",
    "pin2": "input[name='pin2']",
    "puk": "input[name='puk']",
}


def _portals() -> list[dict]:
    state = [
        ("e-Vergi (yeni)", "https://evergi.gov.az"),
        ("e-Sosial", "https://esosial.az"),
        ("e-customs", "https://e.customs.gov.az"),
        ("Statistika", "https://www.stat.gov.az"),
        ("Vergi xidməti", "https://www.taxes.gov.az"),
        ("e-Vergi (köhnə)", "https://e-taxes.gov.az"),
        ("AYNA", "https://ayna.gov.az"),
        ("myGov", "https://my.gov.az"),
        ("MyGovID", "https://mygovid.gov.az"),
        ("Asan İmza Portalı", "https://asan-imza.az"),
    ]
    banks = [
        ("Access Bank", "https://www.accessbank.az"),
        ("AFB Bank", "https://www.afbbank.az"),
        ("ABB", "https://www.abb-bank.az"),
        ("Azər-Türk Bank", "https://www.azer-turkbank.com"),
        ("Bank Avrasiya", "https://www.avrasiya.com"),
        ("Bank of Baku", "https://www.bankofbaku.com"),
        ("Bank Respublika", "https://www.bankrespublika.az"),
        ("Birbank Biznes", "https://birbank.az"),
        ("BMI Baku", "https://www.bmibaku.az"),
        ("BTB Bank", "https://www.btbbank.az"),
        ("Expressbank", "https://www.expressbank.az"),
        ("PashaBank Digital", "https://www.pashabank.digital"),
        ("Premium Bank", "https://www.premiumbank.az"),
        ("Rabitabank", "https://www.rabitabank.com"),
        ("Turanbank", "https://www.turanbank.az"),
        ("Unibank", "https://www.unibank.az"),
        ("VTB Bank", "https://www.vtb.az"),
        ("Xalq Bank", "https://www.xalqbank.az"),
        ("Yapı Kredi Bank", "https://www.yapikredi.com.az"),
        ("Yelo Bank", "https://www.yelo.az"),
        ("Ziraat Bank", "https://www.ziraatbank.com.az"),
    ]

    rows: list[dict] = []
    for sort_order, (name, url) in enumerate(state + banks, start=1):
        category = "dövlət" if (name, url) in state else "bank"
        rows.append(
            {
                "id": _portal_id(name),
                "name": name,
                "url": url,
                "category": category,
                "logo_url": None,
                "selector_mapping": dict(_STANDARD_MAPPING),
                "sort_order": sort_order,
            }
        )
    return rows


def upgrade() -> None:
    op.execute("CREATE EXTENSION IF NOT EXISTS vector")

    op.create_table(
        "profiles",
        sa.Column("id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("ad_soyad", sa.String(length=255), nullable=False),
        sa.Column("voen", sa.String(length=50), nullable=False),
        sa.Column("nomre", sa.String(length=50), nullable=False),
        sa.Column("asan_id", sa.String(length=50), nullable=False),
        sa.Column("bitme_tarixi", sa.Date(), nullable=True),
        sa.Column("qeyd", sa.Text(), nullable=True),
        sa.Column("pin1", sa.Text(), nullable=False),
        sa.Column("pin2", sa.Text(), nullable=False),
        sa.Column("puk", sa.Text(), nullable=False),
        sa.Column("reng", sa.String(length=20), nullable=False),
        sa.Column("ikon", sa.String(length=50), nullable=False),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            server_default=sa.text("now()"),
            nullable=False,
        ),
        sa.Column(
            "updated_at",
            sa.DateTime(timezone=True),
            server_default=sa.text("now()"),
            nullable=False,
        ),
        sa.PrimaryKeyConstraint("id"),
    )

    op.create_table(
        "portals",
        sa.Column("id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("name", sa.String(length=255), nullable=False),
        sa.Column("url", sa.String(length=500), nullable=False),
        sa.Column("category", sa.String(length=20), nullable=False),
        sa.Column("logo_url", sa.String(length=500), nullable=True),
        sa.Column("selector_mapping", postgresql.JSONB(), nullable=False),
        sa.Column("sort_order", sa.Integer(), nullable=False),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("name"),
    )

    op.create_table(
        "vector_documents",
        sa.Column("id", postgresql.UUID(as_uuid=True), nullable=False),
        sa.Column("profile_id", postgresql.UUID(as_uuid=True), nullable=True),
        sa.Column("content", sa.Text(), nullable=False),
        sa.Column("embedding", Vector(dim=1536), nullable=True),
        sa.Column("source_type", sa.String(length=50), nullable=False),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            server_default=sa.text("now()"),
            nullable=False,
        ),
        sa.ForeignKeyConstraint(
            ["profile_id"],
            ["profiles.id"],
            ondelete="SET NULL",
        ),
        sa.PrimaryKeyConstraint("id"),
    )

    portals_table = sa.table(
        "portals",
        sa.column("id", postgresql.UUID(as_uuid=True)),
        sa.column("name", sa.String()),
        sa.column("url", sa.String()),
        sa.column("category", sa.String()),
        sa.column("logo_url", sa.String()),
        sa.column("selector_mapping", postgresql.JSONB()),
        sa.column("sort_order", sa.Integer()),
    )
    op.bulk_insert(portals_table, _portals())


def downgrade() -> None:
    op.drop_table("vector_documents")
    op.drop_table("portals")
    op.drop_table("profiles")
