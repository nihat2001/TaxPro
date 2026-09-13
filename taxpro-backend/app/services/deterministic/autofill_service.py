import logging
import uuid

from sqlalchemy.ext.asyncio import AsyncSession

from app.models.portal import Portal
from app.schemas.portal import AutoFillPayload
from app.schemas.profile import ProfileAutoFill
from app.services.deterministic.phone import normalize_phone
from app.services.deterministic.profile_service import get_profile

logger = logging.getLogger(__name__)

# Yalnız bu profil sahələri saytlara doldurulur.
# Ad Soyad yalnız profil identifikasiyası üçündür — saytda doldurulmur.
_AUTOFILL_FIELDS = ("nomre", "asan_id")

# Portal konfiqində (selector_mapping) prefiks select üçün ehtiyat açar.
_PHONE_PREFIX_KEY = "_phone_prefix_selector"


def _autofill_values(profile) -> dict[str, str | None]:
    return {
        "nomre": normalize_phone(profile.nomre),
        "asan_id": profile.asan_id,
    }


async def get_autofill_profile(session: AsyncSession, profile_id: uuid.UUID) -> ProfileAutoFill | None:
    profile = await get_profile(session, profile_id)
    if profile is None:
        return None
    return ProfileAutoFill(
        id=profile.id,
        ad_soyad=profile.ad_soyad,
        nomre=normalize_phone(profile.nomre),
        asan_id=profile.asan_id,
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

    values = _autofill_values(profile)
    selector_mapping = portal.selector_mapping or {}

    fields: dict[str, str] = {}
    field_map: dict[str, str] = {}
    # Selector-u qeyd olunmayan sahələr yenə də payload-a düşür ki, content
    # script-in heuristik fallback zənciri (attr → label → placeholder) işə
    # düşə bilsin. Belə sahələr ayrıca siyahıda qaytarılır ki, UI istifadəçini
    # Qeydiyyat rejimi ilə selector əlavə etməyə yönləndirsin.
    fallback_fields: dict[str, str] = {}
    missing_selectors: list[str] = []

    for field in _AUTOFILL_FIELDS:
        value = values.get(field)
        if value is None:
            continue
        selector = selector_mapping.get(field)
        if selector is None:
            fallback_fields[field] = value
            missing_selectors.append(field)
            continue
        if not isinstance(selector, str):
            selector = str(selector)
        fields[selector] = value
        field_map[selector] = field

    phone_prefix_selector = selector_mapping.get(_PHONE_PREFIX_KEY)
    if phone_prefix_selector is not None and not isinstance(phone_prefix_selector, str):
        phone_prefix_selector = str(phone_prefix_selector)

    logger.info(
        "Auto-fill payload hazırlandı",
        extra={
            "profile_id": profile_id,
            "portal_id": portal_id,
            "missing_selectors": missing_selectors,
        },
    )
    return AutoFillPayload(
        profile_id=profile_id,
        portal_id=portal_id,
        portal_name=portal.name,
        fields=fields,
        field_map=field_map,
        fallback_fields=fallback_fields,
        missing_selectors=missing_selectors,
        phone_prefix_selector=phone_prefix_selector,
    )
