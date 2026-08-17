"""normalize existing profile phone numbers to 9-digit national format

Revision ID: 0011
Revises: 0010
Create Date: 2026-08-13
"""

from alembic import op

revision = "0011"
down_revision = "0010"
branch_labels = None
depends_on = None

_UPDATE_SQL = r"""
UPDATE profiles SET nomre = CASE
  WHEN nomre ~ '^\+?994[0-9]{9}$' THEN substr(regexp_replace(nomre, '\D', '', 'g'), 4)
  WHEN nomre ~ '^0[0-9]{9}$'      THEN substr(regexp_replace(nomre, '\D', '', 'g'), 2)
  WHEN nomre ~ '^[0-9]{9}$'       THEN regexp_replace(nomre, '\D', '', 'g')
  ELSE nomre
END
"""


def upgrade() -> None:
    op.execute(_UPDATE_SQL)


def downgrade() -> None:
    # Ilkin format itirildiyi ucun data transformasiyasi geri qaytarila bilmez.
    pass
