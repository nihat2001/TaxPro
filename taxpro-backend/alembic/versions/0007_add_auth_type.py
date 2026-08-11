"""portals.auth_type column

Revision ID: 0007
Revises: 0006
Create Date: 2026-08-11
"""

import sqlalchemy as sa

from alembic import op

revision = "0007"
down_revision = "0006"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column(
        "portals",
        sa.Column(
            "auth_type",
            sa.String(length=20),
            nullable=False,
            server_default="unknown",
        ),
    )


def downgrade() -> None:
    op.drop_column("portals", "auth_type")
