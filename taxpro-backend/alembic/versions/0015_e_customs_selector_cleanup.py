"""e-customs: səhv (qeyri-sabit) asan_id selector-unun təmizlənməsi

Revision ID: 0015
Revises: 0014
Create Date: 2026-09-13

Sxem dəyişmir — yalnız data düzəlişi.

`e-customs` portalı üçün Qeydiyyat rejimi zamanı Asan ID kimi Mantine-in
generasiya etdiyi qeyri-sabit ID (`#mantine-...`) tutulmuşdu. Həmin ID növbəti
yükləmələrdə prefiks/operator sahəsinə düşür və autofill Asan ID dəyərini
prefiks sahəsinə yazırdı (6 rəqəmli dəyər 5 rəqəmə kəsilirdi). Yanlış seçici
silinir; düzgün Asan ID + prefiks seçiciləri Record Mode ilə yenidən tutulur.
"""

import sqlalchemy as sa

from alembic import op

revision = "0015"
down_revision = "0014"
branch_labels = None
depends_on = None


def upgrade() -> None:
    # Yalnız QEYRİ-SABİT (framework generasiyalı) asan_id selectorunu sil.
    # İstifadəçi sonradan düzgün stabil selector (məs. input[name="userId"])
    # qeyd edibsə, ona toxunma.
    op.execute(
        sa.text(
            "UPDATE portals SET selector_mapping = selector_mapping - 'asan_id' "
            "WHERE name = 'e-customs' AND ("
            "selector_mapping->>'asan_id' ILIKE :mantine "
            "OR strpos(selector_mapping->>'asan_id', :react) > 0)"
        ).bindparams(mantine="%mantine%", react=":r")
    )


def downgrade() -> None:
    # Səhv (qeyri-sabit) selector geri qaytarılmır — mənasızdır.
    pass
