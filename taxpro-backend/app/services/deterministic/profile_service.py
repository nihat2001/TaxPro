import logging
import uuid

from sqlalchemy import delete, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.security import decrypt_value, encrypt_value
from app.models.profile import Profile
from app.schemas.profile import ProfileCreate, ProfileUpdate

logger = logging.getLogger(__name__)


async def create_profile(session: AsyncSession, data: ProfileCreate) -> Profile:
    profile = Profile(
        ad_soyad=data.ad_soyad,
        voen=data.voen,
        nomre=data.nomre,
        asan_id=data.asan_id,
        bitme_tarixi=data.bitme_tarixi,
        qeyd=data.qeyd,
        pin1=encrypt_value(data.pin1),
        pin2=encrypt_value(data.pin2),
        puk=encrypt_value(data.puk),
        reng=data.reng,
        ikon=data.ikon,
    )
    session.add(profile)
    await session.commit()
    await session.refresh(profile)
    logger.info("Yeni profil yaradıldı", extra={"profile_id": profile.id})
    return profile


async def list_profiles(session: AsyncSession) -> list[Profile]:
    result = await session.execute(select(Profile).order_by(Profile.created_at.desc()))
    return list(result.scalars().all())


async def get_profile(session: AsyncSession, profile_id: uuid.UUID) -> Profile | None:
    return await session.get(Profile, profile_id)


async def update_profile(
    session: AsyncSession, profile_id: uuid.UUID, data: ProfileUpdate
) -> Profile | None:
    profile = await session.get(Profile, profile_id)
    if profile is None:
        return None

    fields = data.model_dump(exclude_unset=True)
    if "pin1" in fields:
        profile.pin1 = encrypt_value(fields.pop("pin1"))
    if "pin2" in fields:
        profile.pin2 = encrypt_value(fields.pop("pin2"))
    if "puk" in fields:
        profile.puk = encrypt_value(fields.pop("puk"))
    for field, value in fields.items():
        setattr(profile, field, value)

    await session.commit()
    await session.refresh(profile)
    logger.info("Profil yeniləndi", extra={"profile_id": profile_id})
    return profile


async def delete_profile(session: AsyncSession, profile_id: uuid.UUID) -> bool:
    result = await session.execute(
        delete(Profile).where(Profile.id == profile_id)
    )
    await session.commit()
    return bool(result.rowcount)


def decrypted_values(profile: Profile) -> dict[str, str | None]:
    return {
        "ad_soyad": profile.ad_soyad,
        "voen": profile.voen,
        "nomre": profile.nomre,
        "asan_id": profile.asan_id,
        "bitme_tarixi": (
            profile.bitme_tarixi.isoformat() if profile.bitme_tarixi else None
        ),
        "qeyd": profile.qeyd,
        "pin1": decrypt_value(profile.pin1),
        "pin2": decrypt_value(profile.pin2),
        "puk": decrypt_value(profile.puk),
    }
