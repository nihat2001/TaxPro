"""portal login urls: birbasa giris sehifeleri

Revision ID: 0005
Revises: 0004
Create Date: 2026-08-11
"""

from alembic import op

revision = "0005"
down_revision = "0004"
branch_labels = None
depends_on = None

_UPDATES = {
    "e-Vergi (yeni)": (
        "https://new.e-taxes.gov.az/eportal/login",
        "https://new.e-taxes.gov.az/etaxes/",
    ),
    "e-Vergi (köhnə)": (
        "https://new.e-taxes.gov.az/eportal/login",
        "https://e-taxes.gov.az",
    ),
    "e-Sosial": ("https://mygovid.gov.az/auth", "https://e-sosial.az/login"),
    "Asan İmza Portalı": (
        "https://portal.asxm.gov.az/login",
        "https://asan-imza.az",
    ),
    "myGov": ("https://mygovid.gov.az/auth", "https://my.gov.az"),
    "AYNA": ("https://mygovid.gov.az/auth", "https://ayna.gov.az"),
    "Statistika": ("https://www.stat.gov.az/", "https://www.stat.gov.az"),
    "Vergi xidməti": ("https://www.taxes.gov.az/az", "https://www.taxes.gov.az"),
    "Access Bank": (
        "https://www.accessbank.az/az/requests/onlayn-kredit?utm_source=Website_slider",
        "https://www.accessbank.az",
    ),
    "AFB Bank": ("https://business.afb.az/login", "https://afb.az"),
    "ABB": ("https://abb-bank.az/", "https://www.abb-bank.az"),
    "Bank Avrasiya": ("https://bankavrasiya.az/", "https://www.avrasiya.com"),
}


def upgrade() -> None:
    for name, (new_url, _old_url) in _UPDATES.items():
        op.execute(f"UPDATE portals SET url = '{new_url}' WHERE name = '{name}'")


def downgrade() -> None:
    for name, (_new_url, old_url) in _UPDATES.items():
        op.execute(f"UPDATE portals SET url = '{old_url}' WHERE name = '{name}'")
