import logging
import uuid

from sqlalchemy.ext.asyncio import AsyncSession

from app.core.security import decrypt_value
from app.models.portal import Portal
from app.schemas.portal import AutoFillPayload
from app.schemas.profile import ProfileAutoFill
from app.services.deterministic.phone import normalize_phone
from app.services.deterministic.profile_service import get_profile

logger = logging.getLogger(__name__)


def _decrypted_values(profile) -> dict[str, str | None]:
    return {
        "ad_soyad": profile.ad_soyad,
        "voen": profile.voen,
        "nomre": normalize_phone(profile.nomre),
        "asan_id": profile.asan_id,
        "bitme_tarixi": (
            profile.bitme_tarixi.isoformat() if profile.bitme_tarixi else None
        ),
        "qeyd": profile.qeyd,
        "pin1": decrypt_value(profile.pin1),
        "pin2": decrypt_value(profile.pin2),
        "puk": decrypt_value(profile.puk),
    }


async def get_autofill_profile(session: AsyncSession, profile_id: uuid.UUID) -> ProfileAutoFill | None:
    profile = await get_profile(session, profile_id)
    if profile is None:
        return None
    return ProfileAutoFill(
        id=profile.id,
        ad_soyad=profile.ad_soyad,
        voen=profile.voen,
        nomre=normalize_phone(profile.nomre),
        asan_id=profile.asan_id,
        bitme_tarixi=profile.bitme_tarixi,
        qeyd=profile.qeyd,
        pin1=decrypt_value(profile.pin1),
        pin2=decrypt_value(profile.pin2),
        puk=decrypt_value(profile.puk),
    )


async def build_autofill_payload(
    session: AsyncSession, profile_id: uuid.UUID, portal_id: uuid.UUID
) -> AutoFillPayload:
    profile = await get_profile(session, profile_id)
    if profile is None:
        raise ValueError("Profil tapılmadı.")

    portal = await session.get(Portal, portal_id)
    if portal is None:
        raise ValueError("Portal tapılmadı.")

    values = _decrypted_values(profile)
    selector_mapping = portal.selector_mapping or {}

    fields: dict[str, str] = {}
    field_map: dict[str, str] = {}
    for field, selector in selector_mapping.items():
        value = values.get(field)
        if value is None:
            continue
        if not isinstance(selector, str):
            selector = str(selector)
        fields[selector] = value
        field_map[selector] = field

    logger.info(
        "Auto-fill payload hazırlandı",
        extra={"profile_id": profile_id, "portal_id": portal_id},
    )
    return AutoFillPayload(
        profile_id=profile_id,
        portal_id=portal_id,
        portal_name=portal.name,
        fields=fields,
        field_map=field_map,
    )
