from uuid import UUID

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.schemas.profile import (
    ProfileAutoFill,
    ProfileCreate,
    ProfileRead,
    ProfileUpdate,
)
from app.services.deterministic.autofill_service import get_autofill_profile
from app.services.deterministic.profile_service import (
    create_profile,
    delete_profile,
    get_profile,
    list_profiles,
    update_profile,
)

router = APIRouter(prefix="/profiles", tags=["profiles"])


@router.get("", response_model=list[ProfileRead])
async def get_profiles(db: AsyncSession = Depends(get_db)):
    return await list_profiles(db)


@router.post("", response_model=ProfileRead, status_code=status.HTTP_201_CREATED)
async def post_profile(data: ProfileCreate, db: AsyncSession = Depends(get_db)):
    return await create_profile(db, data)


@router.get("/{profile_id}", response_model=ProfileRead)
async def get_one_profile(profile_id: UUID, db: AsyncSession = Depends(get_db)):
    profile = await get_profile(db, profile_id)
    if profile is None:
        raise HTTPException(status_code=404, detail="Profil tapılmadı.")
    return profile


@router.get("/{profile_id}/autofill", response_model=ProfileAutoFill)
async def get_profile_autofill(
    profile_id: UUID, db: AsyncSession = Depends(get_db)
):
    payload = await get_autofill_profile(db, profile_id)
    if payload is None:
        raise HTTPException(status_code=404, detail="Profil tapılmadı.")
    return payload


@router.patch("/{profile_id}", response_model=ProfileRead)
async def patch_profile(
    profile_id: UUID, data: ProfileUpdate, db: AsyncSession = Depends(get_db)
):
    profile = await update_profile(db, profile_id, data)
    if profile is None:
        raise HTTPException(status_code=404, detail="Profil tapılmadı.")
    return profile


@router.delete("/{profile_id}", status_code=status.HTTP_204_NO_CONTENT)
async def remove_profile(profile_id: UUID, db: AsyncSession = Depends(get_db)):
    deleted = await delete_profile(db, profile_id)
    if not deleted:
        raise HTTPException(status_code=404, detail="Profil tapılmadı.")
