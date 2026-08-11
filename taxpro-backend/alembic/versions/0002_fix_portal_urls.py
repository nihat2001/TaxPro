"""fix portal urls: e-Vergi, e-Sosial, AFB, Azer-Turk, BTB, Ziraat

Revision ID: 0002
Revises: 0001
Create Date: 2026-08-11
"""

from alembic import op

revision = "0002"
down_revision = "0001"
branch_labels = None
depends_on = None

_FIXES = {
    "e-Vergi (yeni)": "https://new.e-taxes.gov.az/eportal/login",
    "e-Sosial": "https://mygovid.gov.az/auth",
    "e-customs": "https://e.customs.gov.az/auth/login",
    "AFB Bank": "https://business.afb.az/login",
    "Azər-Türk Bank": "https://atb.az",
    "BTB Bank": "https://btb.az",
    "Ziraat Bank": "https://ziraatbank.az",
}


def upgrade() -> None:
    for name, url in _FIXES.items():
        safe_name = name.replace("'", "''")
        safe_url = url.replace("'", "''")
        op.execute(
            f"UPDATE portals SET url = '{safe_url}' WHERE name = '{safe_name}'"
        )


def downgrade() -> None:
    pass
