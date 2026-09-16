"""Bütün portallardan toxum (seed) selector-larının təmizlənməsi

Revision ID: 0019
Revises: 0018
Create Date: 2026-09-16

Sxem dəyişmir — yalnız data.

0001 miqrasiyası bütün portallara eyni generic `selector_mapping`-i yazmışdı
(`input[name='phone']` və `input[name='asanid']`). Bu, real qeyd deyildi və
müasir saytların DOM-una uyğun gəlmir — nəticədə autofill "uğursuz" olur.
Bu miqrasiya yalnız həmin toxum dəyərlərini daşıyan `nomre`/`asan_id`
açarlarını silir (real Record Mode qeydlərinə toxunmur).
"""

import sqlalchemy as sa

from alembic import op

revision = "0019"
down_revision = "0018"
branch_labels = None
depends_on = None

_SEED_PHONE = "input[name='phone']"
_SEED_ASAN = "input[name='asanid']"


def upgrade() -> None:
    op.execute(
        sa.text(
            "UPDATE portals SET selector_mapping = selector_mapping - 'nomre' "
            "WHERE selector_mapping->>'nomre' = :seed"
        ).bindparams(seed=_SEED_PHONE)
    )
    op.execute(
        sa.text(
            "UPDATE portals SET selector_mapping = selector_mapping - 'asan_id' "
            "WHERE selector_mapping->>'asan_id' = :seed"
        ).bindparams(seed=_SEED_ASAN)
    )


def downgrade() -> None:
    # Toxum selectorlar süni idi — geri qaytarılmır.
    pass
