import pytest

from app.schemas.profile import ProfileCreate, ProfileUpdate
from app.services.deterministic.profile_service import (
    create_profile,
    delete_profile,
    get_profile,
    list_profiles,
    update_profile,
)


@pytest.fixture
async def session():
    from app.core.database import async_session_maker

    async with async_session_maker() as s:
        yield s


def _data(**overrides) -> ProfileCreate:
    base = dict(
        ad_soyad="Test Profil",
        nomre="+994501234567",
        asan_id="TT001",
    )
    base.update(overrides)
    return ProfileCreate(**base)


async def test_create_profile(session):
    profile = await create_profile(session, _data())
    assert profile.id is not None
    assert profile.ad_soyad == "Test Profil"
    assert profile.nomre == "+994501234567"
    assert profile.asan_id == "TT001"


async def test_list_and_get_profile(session):
    profile = await create_profile(session, _data())
    items = await list_profiles(session)
    assert any(p.id == profile.id for p in items)
    found = await get_profile(session, profile.id)
    assert found is not None
    assert found.ad_soyad == "Test Profil"


async def test_update_profile(session):
    profile = await create_profile(session, _data())
    updated = await update_profile(
        session, profile.id, ProfileUpdate(asan_id="TT999", nomre="0551234567")
    )
    assert updated is not None
    assert updated.asan_id == "TT999"
    assert updated.nomre == "0551234567"
    assert updated.ad_soyad == "Test Profil"


async def test_update_profile_missing_returns_none(session):
    import uuid

    assert await update_profile(session, uuid.uuid4(), ProfileUpdate()) is None


async def test_delete_profile(session):
    profile = await create_profile(session, _data())
    assert await delete_profile(session, profile.id) is True
    assert await get_profile(session, profile.id) is None