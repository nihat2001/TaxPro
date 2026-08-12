"""merge e-Vergi portals + static logo_url overrides

Revision ID: 0009
Revises: 0008
Create Date: 2026-08-12
"""

import json
import uuid

from alembic import op

revision = "0009"
down_revision = "0008"
branch_labels = None
depends_on = None

_STANDARD_MAPPING = {
    "ad_soyad": "input[name='fullname']",
    "voen": "input[name='voen']",
    "nomre": "input[name='phone']",
    "asan_id": "input[name='asanid']",
    "pin1": "input[name='pin1']",
    "pin2": "input[name='pin2']",
    "puk": "input[name='puk']",
}

_LOGO_UPDATES = {
    "e-Sosial": "https://cdn.aptoide.com/imgs/0/d/6/0d642f4f67f51c8e92860f04ec1e86d2_icon.png",
    "myGov": (
        "https://media.licdn.com/dms/image/v2/D4E0BAQENUtjdN6JaUA/"
        "company-logo_200_200/company-logo_200_200/0/1710534896931/mygovaz_logo"
        "?e=2147483647&v=beta&t=gKmZP6Do-w2FIsM6-3EnXjHmEXd6f8R0SCPNQL80ebQ"
    ),
    "MyGovID": (
        "https://media.licdn.com/dms/image/v2/D4E0BAQENUtjdN6JaUA/"
        "company-logo_200_200/company-logo_200_200/0/1710534896931/mygovaz_logo"
        "?e=2147483647&v=beta&t=gKmZP6Do-w2FIsM6-3EnXjHmEXd6f8R0SCPNQL80ebQ"
    ),
    "AFB Bank": (
        "https://media.licdn.com/dms/image/v2/C4E0BAQGFwjoiLgYVTg/"
        "company-logo_200_200/company-logo_200_200/0/1663138554265/afbbank_logo"
        "?e=2147483647&v=beta&t=oKh1uMsBjsQ-bV0ejjpUlbmlWSe0VkfwzAk30zH_eUk"
    ),
    "Bank Avrasiya": (
        "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQfury12IJi1Je1XBtGEuY7Ek0vmzPSXTrpyCGwWEhMhw&s=10"
    ),
    "Bank of Baku": (
        "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTaaMVQm3ycgwrVN2YzgWoliL7kYIybSOqsH7IsJJBnzRUJ-QbEGSXb_U4&s=10"
    ),
    "Expressbank": (
        "https://media.licdn.com/dms/image/v2/D4E0BAQGzlDI_jqJrLA/"
        "company-logo_200_200/company-logo_200_200/0/1735368988500/expressbankasc_logo"
        "?e=2147483647&v=beta&t=E0GYTlJwpe1pnW0n_LjCPtBul1ba6WTjM--7kfNmguw"
    ),
    "VTB Bank": (
        "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRySgUJlX5_MNfgnyEKkTU40p0cYZP9PqF-Ru3QjikUcaGgLTjHIUTwCbE&s=10"
    ),
    "e-Vergi": "https://www.google.com/s2/favicons?domain=www.taxes.gov.az&sz=64",
}


def _safe(sql: str) -> None:
    op.execute(sql)


def upgrade() -> None:
    op.execute("DELETE FROM portals WHERE name = 'e-Vergi (köhnə)'")
    op.execute("UPDATE portals SET name = 'e-Vergi' WHERE name = 'e-Vergi (yeni)'")

    for name, logo_url in _LOGO_UPDATES.items():
        _safe(f"UPDATE portals SET logo_url = '{logo_url}' WHERE name = '{name}'")


def downgrade() -> None:
    op.execute("UPDATE portals SET name = 'e-Vergi (yeni)' WHERE name = 'e-Vergi'")

    old_id = str(uuid.uuid5(uuid.NAMESPACE_URL, "https://taxpro.portal/e-Vergi (köhnə)"))
    mapping_json = json.dumps(_STANDARD_MAPPING, ensure_ascii=False).replace("'", "''")
    _safe(
        "INSERT INTO portals (id, name, url, category, logo_url, selector_mapping, auth_type, sort_order) "
        f"VALUES ('{old_id}', 'e-Vergi (köhnə)', 'https://www.e-taxes.gov.az', 'dövlət', NULL, "
        f"'{mapping_json}'::jsonb, 'unknown', 6) "
        "ON CONFLICT (name) DO NOTHING"
    )

    for name in _LOGO_UPDATES:
        _safe(f"UPDATE portals SET logo_url = NULL WHERE name = '{name}'")
