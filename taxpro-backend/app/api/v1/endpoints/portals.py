from uuid import UUID

from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.schemas.portal import AutoFillPayload, PortalRead, SelectorMappingUpdate
from app.services.deterministic.autofill_service import build_autofill_payload
from app.services.deterministic.portal_service import get_portal, list_portals

router = APIRouter(prefix="/portals", tags=["portals"])


@router.get("", response_model=list[PortalRead])
async def get_portals(db: AsyncSession = Depends(get_db)):
    return await list_portals(db)


@router.get("/{portal_id}", response_model=PortalRead)
async def get_one_portal(portal_id: UUID, db: AsyncSession = Depends(get_db)):
    portal = await get_portal(db, portal_id)
    if portal is None:
        raise HTTPException(status_code=404, detail="Portal tapılmadı.")
    return portal


@router.get("/{portal_id}/autofill/{profile_id}", response_model=AutoFillPayload)
async def get_autofill(
    portal_id: UUID, profile_id: UUID, db: AsyncSession = Depends(get_db)
):
    try:
        return await build_autofill_payload(db, profile_id, portal_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc))


@router.patch("/{portal_id}/selector-mapping", response_model=PortalRead)
async def update_selector_mapping(
    portal_id: UUID,
    body: SelectorMappingUpdate,
    db: AsyncSession = Depends(get_db),
):
    portal = await get_portal(db, portal_id)
    if portal is None:
        raise HTTPException(status_code=404, detail="Portal tapılmadı.")
    portal.selector_mapping = body.selector_mapping
    portal.auth_type = body.auth_type
    await db.commit()
    await db.refresh(portal)
    return portal
