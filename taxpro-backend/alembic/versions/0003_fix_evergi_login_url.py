"""e-Vergi (yeni) birbasa login sehifesine kecid

Revision ID: 0003
Revises: 0002
Create Date: 2026-08-11
"""

from alembic import op

revision = "0003"
down_revision = "0002"
branch_labels = None
depends_on = None

_NAME = "e-Vergi (yeni)"
_NEW_URL = "https://new.e-taxes.gov.az/eportal/login"
_OLD_URL = "https://www.taxes.gov.az/az"


def upgrade() -> None:
    op.execute(
        f"UPDATE portals SET url = '{_NEW_URL}' WHERE name = '{_NAME}'"
    )


def downgrade() -> None:
    op.execute(
        f"UPDATE portals SET url = '{_OLD_URL}' WHERE name = '{_NAME}'"
    )
