"""Azər-Türk Bank portalına rəsmi logo_url əlavəsi

Revision ID: 0017
Revises: 0016
Create Date: 2026-09-13

Sxem dəyişmir — yalnız data. Digər banklarla (0009/0016) eyni struktur:
`portals.logo_url` üçün xarici şəkil URL-i. Business app-in öz favicon-u WAF
tərəfindən bot sorğularına bağlı olduğu üçün Google favicon xidmətindən
(rəsmi `atb.az` favicon-u) istifadə olunur.
"""

import sqlalchemy as sa

from alembic import op

revision = "0017"
down_revision = "0016"
branch_labels = None
depends_on = None

_LOGO_URL = "https://www.google.com/s2/favicons?domain=atb.az&sz=64"


def upgrade() -> None:
    op.execute(
        sa.text(
            "UPDATE portals SET logo_url = :logo WHERE name = 'Azər-Türk Bank'"
        ).bindparams(logo=_LOGO_URL)
    )


def downgrade() -> None:
    op.execute(
        "UPDATE portals SET logo_url = NULL WHERE name = 'Azər-Türk Bank'"
    )
