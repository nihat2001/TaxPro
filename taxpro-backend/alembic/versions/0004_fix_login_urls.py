"""e-Sosial ve e-Customs birbasa login sehifeleri

Revision ID: 0004
Revises: 0003
Create Date: 2026-08-11
"""

from alembic import op

revision = "0004"
down_revision = "0003"
branch_labels = None
depends_on = None

_FIXES = {
    "e-Sosial": ("https://e-sosial.az/login", "https://e-sosial.az"),
    "e-customs": (
        "https://e.customs.gov.az/auth/login",
        "https://e.customs.gov.az",
    ),
}


def upgrade() -> None:
    for name, (new_url, _old_url) in _FIXES.items():
        op.execute(f"UPDATE portals SET url = '{new_url}' WHERE name = '{name}'")


def downgrade() -> None:
    for name, (_new_url, old_url) in _FIXES.items():
        op.execute(f"UPDATE portals SET url = '{old_url}' WHERE name = '{name}'")
