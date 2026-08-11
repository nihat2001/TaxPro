from sqlalchemy import or_, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models.profile import Profile
from app.models.vector_store import VectorDocument


async def search_profiles(
    session: AsyncSession, query: str, profile_id=None, limit: int = 5
) -> list[Profile]:
    tokens = [t for t in query.lower().split() if t]
    stmt = select(Profile)
    if profile_id is not None:
        stmt = stmt.where(Profile.id == profile_id)
    if tokens:
        conditions = [Profile.ad_soyad.ilike(f"%{tok}%") for tok in tokens]
        stmt = stmt.where(or_(*conditions))
    result = await session.execute(stmt.limit(limit))
    return list(result.scalars().all())


async def search_documents(
    session: AsyncSession, query: str, profile_id=None, limit: int = 5
) -> list[VectorDocument]:
    stmt = select(VectorDocument)
    if profile_id is not None:
        stmt = stmt.where(VectorDocument.profile_id == profile_id)
    stmt = stmt.where(VectorDocument.content.ilike(f"%{query}%"))
    result = await session.execute(stmt.limit(limit))
    return list(result.scalars().all())
