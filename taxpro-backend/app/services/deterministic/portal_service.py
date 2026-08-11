from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models.portal import Portal


async def list_portals(session: AsyncSession) -> list[Portal]:
    result = await session.execute(
        select(Portal).order_by(Portal.sort_order, Portal.name)
    )
    return list(result.scalars().all())


async def get_portal(session: AsyncSession, portal_id) -> Portal | None:
    return await session.get(Portal, portal_id)
