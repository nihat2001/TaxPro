import pytest

from app.core.security import decrypt_value, encrypt_value
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
        voen="1234567890",
        nomre="+994501234567",
        asan_id="TT001",
        pin1="1111",
        pin2="2222",
        puk="3333",
    )
    base.update(overrides)
    return ProfileCreate(**base)


async def test_create_profile_encrypts_secrets(session):
    profile = await create_profile(session, _data())
    assert profile.id is not None
    assert profile.pin1.startswith("enc:")
    assert decrypt_value(profile.pin1) == "1111"
    assert decrypt_value(profile.puk) == "3333"


async def test_list_and_get_profile(session):
    profile = await create_profile(session, _data())
    items = await list_profiles(session)
    assert any(p.id == profile.id for p in items)
    found = await get_profile(session, profile.id)
    assert found is not None
    assert found.ad_soyad == "Test Profil"


async def test_update_profile_reencrypts_pin(session):
    profile = await create_profile(session, _data())
    updated = await update_profile(
        session, profile.id, ProfileUpdate(pin1="9999", asan_id="TT999")
    )
    assert updated is not None
    assert updated.asan_id == "TT999"
    assert decrypt_value(updated.pin1) == "9999"


async def test_update_profile_missing_returns_none(session):
    import uuid

    assert await update_profile(session, uuid.uuid4(), ProfileUpdate()) is None


async def test_delete_profile(session):
    profile = await create_profile(session, _data())
    assert await delete_profile(session, profile.id) is True
    assert await get_profile(session, profile.id) is None


def test_encrypt_roundtrip():
    token = encrypt_value("secret-pin")
    assert token.startswith("enc:")
    assert decrypt_value(token) == "secret-pin"


def test_decrypt_plain_value_passthrough():
    assert decrypt_value("not-encrypted") == "not-encrypted"
