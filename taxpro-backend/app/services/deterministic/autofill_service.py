import logging
import uuid

from sqlalchemy.ext.asyncio import AsyncSession

from app.models.portal import Portal
from app.schemas.portal import AutoFillPayload
from app.services.deterministic.profile_service import (
    decrypted_values,
    get_profile,
)

logger = logging.getLogger(__name__)


async def build_autofill_payload(
    session: AsyncSession, profile_id: uuid.UUID, portal_id: uuid.UUID
) -> AutoFillPayload:
    profile = await get_profile(session, profile_id)
    if profile is None:
        raise ValueError("Profil tapılmadı.")

    portal = await session.get(Portal, portal_id)
    if portal is None:
        raise ValueError("Portal tapılmadı.")

    values = decrypted_values(profile)
    selector_mapping = portal.selector_mapping or {}

    fields: dict[str, str] = {}
    for field, selector in selector_mapping.items():
        value = values.get(field)
        if value is None:
            continue
        if not isinstance(selector, str):
            selector = str(selector)
        fields[selector] = value

    logger.info(
        "Auto-fill payload hazırlandı",
        extra={"profile_id": profile_id, "portal_id": portal_id},
    )
    return AutoFillPayload(
        profile_id=profile_id,
        portal_id=portal_id,
        portal_name=portal.name,
        fields=fields,
    )
