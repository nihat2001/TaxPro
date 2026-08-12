"""fix portal urls: AYNA, e-Sosial, e-Vergi (koene)

Revision ID: 0008
Revises: 0007
Create Date: 2026-08-12
"""

from alembic import op

revision = "0008"
down_revision = "0007"
branch_labels = None
depends_on = None

_UPDATES = {
    "AYNA": ("https://ayna.gov.az/az", "https://mygovid.gov.az/auth"),
    "e-Sosial": ("https://e-social.gov.az/az", "https://mygovid.gov.az/auth"),
    "e-Vergi (köhnə)": (
        "https://www.e-taxes.gov.az",
        "https://new.e-taxes.gov.az/eportal/login",
    ),
}


def upgrade() -> None:
    for name, (new_url, _old_url) in _UPDATES.items():
        op.execute(f"UPDATE portals SET url = '{new_url}' WHERE name = '{name}'")


def downgrade() -> None:
    for name, (_new_url, old_url) in _UPDATES.items():
        op.execute(f"UPDATE portals SET url = '{old_url}' WHERE name = '{name}'")
