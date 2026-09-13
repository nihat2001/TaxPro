"""Access Bank portalına rəsmi logo_url əlavəsi

Revision ID: 0016
Revises: 0015
Create Date: 2026-09-13

Sxem dəyişmir — yalnız data. Portalın `logo_url`-i NULL olduğu üçün extension
Google favicon fallback-inə düşürdü, `myaccess.az` domeni üçün isə favicon
tapılmadığından hərf-avatar görünürdü. myAccess-in manifest ikonu (kvadrat,
192x192) təyin olunur.
"""

import sqlalchemy as sa

from alembic import op

revision = "0016"
down_revision = "0015"
branch_labels = None
depends_on = None

_LOGO_URL = "https://myaccess.az/login/images/logo_myAccess-192.png"


def upgrade() -> None:
    op.execute(
        sa.text(
            "UPDATE portals SET logo_url = :logo WHERE name = 'Access Bank'"
        ).bindparams(logo=_LOGO_URL)
    )


def downgrade() -> None:
    op.execute("UPDATE portals SET logo_url = NULL WHERE name = 'Access Bank'")
