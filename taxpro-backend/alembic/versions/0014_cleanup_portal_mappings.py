"""portal selector_mapping data təmizliyi + yanlış auth_type='none' düzəlişi

Revision ID: 0014
Revises: 0013
Create Date: 2026-09-13

Sxem dəyişmir — yalnız data düzəlişi:

1) 0013-də profil modeli 3 sahəyə (ad_soyad, nomre, asan_id) endirildi, amma
   portal `selector_mapping`-lərində köhnə `voen/pin1/pin2/puk/ad_soyad`
   açarları qalmışdı. Autofill yalnız `nomre` + `asan_id` yazdığı üçün bu
   açarlar ölü datadır və qarışıqlıq yaradır — silinir.

2) `e-Vergi` sətri səhv `auth_type='none'` daşıyırdı (real login formu var).
   'none' dəyəri autofill-i UI-də bloklayırdı; 'unknown'-a çevrilir.
   Real selector-lar yalnız "Qeydiyyat rejimi" ilə tutula bilər — bu miqrasiya
   onları uydurmur.
"""

import sqlalchemy as sa

from alembic import op

revision = "0014"
down_revision = "0013"
branch_labels = None
depends_on = None

# 0013-dən sonra istifadəsiz qalan legacy mapping açarları.
_DEAD_KEYS = ("ad_soyad", "voen", "pin1", "pin2", "puk")


def upgrade() -> None:
    expr = "selector_mapping"
    for key in _DEAD_KEYS:
        expr += f" - '{key}'"
    op.execute(f"UPDATE portals SET selector_mapping = {expr}")

    # 'none' = "login formu yoxdur" mənasını daşıyırdı, lakin səhv data idi.
    op.execute("UPDATE portals SET auth_type = 'unknown' WHERE auth_type = 'none'")


def downgrade() -> None:
    # Data təmizliyi geri qaytarıla bilməz (hansı portalda hansı legacy açarın
    # olduğu bilinmir). Auth_type-un 'none'-a qaytarılması da yanlış olardı.
    pass
