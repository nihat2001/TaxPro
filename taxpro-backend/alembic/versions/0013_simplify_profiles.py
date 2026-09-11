"""simplify profiles to ad_soyad, nomre, asan_id (drop voen/pin/... columns)

Revision ID: 0013
Revises: 0012
Create Date: 2026-09-11

Profil modeli 3 sahəyə endirildi (Ad Soyad, Telefon, Asan ID). PIN/PUK həssas
data idi və artıq heç bir kod yolu istifadə etmir; VÖEN/bitmə tarixi/qeyd/rəng/ikon
da çıxarıldı. Sütunlar nullable edilmir — birbaşa silinir.
"""

import sqlalchemy as sa

from alembic import op

revision = "0013"
down_revision = "0012"
branch_labels = None
depends_on = None

_DROPPED = (
    "voen",
    "bitme_tarixi",
    "qeyd",
    "pin1",
    "pin2",
    "puk",
    "reng",
    "ikon",
)


def upgrade() -> None:
    for column in _DROPPED:
        op.drop_column("profiles", column)


def downgrade() -> None:
    # Geri qaytarma: sütunlar nullable olaraq bərpa olunur (data itirildiyi üçün boş).
    op.add_column("profiles", sa.Column("voen", sa.String(length=50), nullable=True))
    op.add_column("profiles", sa.Column("bitme_tarixi", sa.Date(), nullable=True))
    op.add_column("profiles", sa.Column("qeyd", sa.Text(), nullable=True))
    op.add_column("profiles", sa.Column("pin1", sa.Text(), nullable=True))
    op.add_column("profiles", sa.Column("pin2", sa.Text(), nullable=True))
    op.add_column("profiles", sa.Column("puk", sa.Text(), nullable=True))
    op.add_column("profiles", sa.Column("reng", sa.String(length=20), nullable=True))
    op.add_column("profiles", sa.Column("ikon", sa.String(length=50), nullable=True))
