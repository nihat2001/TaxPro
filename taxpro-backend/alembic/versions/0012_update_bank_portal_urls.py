"""update bank portal urls to direct internet-banking login pages

Revision ID: 0012
Revises: 0011
Create Date: 2026-08-25
"""

from alembic import op

revision = "0012"
down_revision = "0011"
branch_labels = None
depends_on = None

# name -> (yeni url, köhnə url)
_UPDATES = {
    "ABB": (
        "https://cb.abb-bank.az/login",
        "https://abb-bank.az/",
    ),
    "Access Bank": (
        "https://myaccess.az/login",
        "https://www.accessbank.az/az/requests/onlayn-kredit?utm_source=Website_slider",
    ),
    "AFB Bank": (
        "https://business.afb.az/login",
        "https://business.afb.az/login",
    ),
    "Azər-Türk Bank": (
        "https://business.atb.az/corporate/login",
        "https://atb.az",
    ),
    "Bank Avrasiya": (
        "https://online.bankavrasiya.az/login",
        "https://bankavrasiya.az/",
    ),
    "Bank of Baku": (
        "https://www.e-bankofbaku.com/InterBankNew/login",
        "https://www.bankofbaku.com",
    ),
    "Bank Respublika": (
        "https://iofis.bankrespublika.az/login",
        "https://www.bankrespublika.az",
    ),
    "Birbank Biznes": (
        "https://my.birbank.business/login",
        "https://birbank.az",
    ),
    "BMI Baku": (
        "https://ib.bmibaku.az/login.aspx",
        "https://www.bmibaku.az",
    ),
    "BTB Bank": (
        "https://ib.btb.az/",
        "https://btb.az",
    ),
    "Expressbank": (
        "https://cb.expressbank.az/login",
        "https://www.expressbank.az",
    ),
    "PashaBank Digital": (
        "https://pashabank.digital/login",
        "https://www.pashabank.digital",
    ),
    "Premium Bank": (
        "https://ibank.premiumbank.az/auth",
        "https://www.premiumbank.az",
    ),
    "Rabitabank": (
        "https://ib.rabitabank.com/InterBank/",
        "https://www.rabitabank.com",
    ),
    "Turanbank": (
        "https://online.turanbank.az/login/individual/asan",
        "https://www.turanbank.az",
    ),
    "Unibank": (
        "https://online.unibank.az/corporate/",
        "https://www.unibank.az",
    ),
    "VTB Bank": (
        "https://online.vtb.az/ru/auth/signin",
        "https://www.vtb.az",
    ),
    "Xalq Bank": (
        "https://xalqonline.az/az/auth/login",
        "https://www.xalqbank.az",
    ),
    "Yapı Kredi Bank": (
        "https://intbank.yapikredi.com.az/WebApplication.UI/Default.aspx?CustomerType=Retail",
        "https://www.yapikredi.com.az",
    ),
    "Yelo Bank": (
        "https://ibank.yelo.az/az/login/",
        "https://www.yelo.az",
    ),
    "Ziraat Bank": (
        "https://online.ziraatbank.az/WebApplication.UI/Default.aspx?CustomerType=Retail",
        "https://ziraatbank.az",
    ),
}


def upgrade() -> None:
    for name, (new_url, _old_url) in _UPDATES.items():
        safe = name.replace("'", "''")
        op.execute(f"UPDATE portals SET url = '{new_url}' WHERE name = '{safe}'")


def downgrade() -> None:
    for name, (_new_url, old_url) in _UPDATES.items():
        safe = name.replace("'", "''")
        op.execute(f"UPDATE portals SET url = '{old_url}' WHERE name = '{safe}'")