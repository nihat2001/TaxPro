"""e-Sosial ve Asan Imza login url + MES-in silinmesi

Revision ID: 0006
Revises: 0005
Create Date: 2026-08-11
"""

import json
import uuid

from alembic import op

revision = "0006"
down_revision = "0005"
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

_MES_NAME = "MƏS"
_MES_URL = "https://portal.asxm.gov.az/login"


def upgrade() -> None:
    op.execute("DELETE FROM portals WHERE name = 'MƏS'")
    op.execute(
        "UPDATE portals SET url = 'https://mygovid.gov.az/auth' WHERE name = 'e-Sosial'"
    )
    op.execute(
        "UPDATE portals SET url = 'https://portal.asxm.gov.az/login' "
        "WHERE name = 'Asan İmza Portalı'"
    )


def downgrade() -> None:
    op.execute(
        "UPDATE portals SET url = 'https://e-sosial.az/login' WHERE name = 'e-Sosial'"
    )
    op.execute(
        "UPDATE portals SET url = 'https://asan-imza.az' WHERE name = 'Asan İmza Portalı'"
    )

    mes_id = str(uuid.uuid5(uuid.NAMESPACE_URL, f"https://taxpro.portal/{_MES_NAME}"))
    mapping_json = json.dumps(_STANDARD_MAPPING, ensure_ascii=False).replace("'", "''")
    op.execute(
        "INSERT INTO portals (id, name, url, category, logo_url, selector_mapping, sort_order) "
        f"VALUES ('{mes_id}', '{_MES_NAME}', '{_MES_URL}', 'dövlət', NULL, "
        f"'{mapping_json}'::jsonb, 11) "
        "ON CONFLICT (name) DO NOTHING"
    )
