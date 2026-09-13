"""Rabitabank: köhnə qeyri-sabit UUID selector-unun təmizlənməsi

Revision ID: 0018
Revises: 0017
Create Date: 2026-09-13

Sxem dəyişmir — yalnız data.

Rabitəbank-ın Vue tətbiqi telefon input-una hər səhifə yüklənişində yeni UUID
(id="<uuid>") verir. Qeydiyyat rejimi əvvəllər həmin UUID-ni selector kimi
saxlamışdı (`#<uuid>`), ona görə selector növbəti yükləmədə işləmirdi. Bu
miqrasiya yalnız UUID formalı köhnə selectoru silir; stabil selector sonra
Record Mode ilə yenidən tutulur.
"""

import sqlalchemy as sa

from alembic import op

revision = "0018"
down_revision = "0017"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.execute(
        sa.text(
            "UPDATE portals SET selector_mapping = selector_mapping - 'nomre' "
            "WHERE name = 'Rabitabank' "
            "AND (selector_mapping->>'nomre') ~ :pattern"
        ).bindparams(pattern=r"^#?[0-9a-fA-F-]{36}$")
    )


def downgrade() -> None:
    # Qeyri-sabit UUID geri qaytarılmır (mənasızdır).
    pass
