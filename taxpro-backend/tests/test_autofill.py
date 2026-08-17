import pytest

from app.models.portal import Portal
from app.schemas.profile import ProfileCreate
from app.services.deterministic.autofill_service import (
    build_autofill_payload,
    get_autofill_profile,
)
from app.services.deterministic.profile_service import create_profile


@pytest.fixture
async def session():
    from app.core.database import async_session_maker

    async with async_session_maker() as s:
        yield s


async def _profile(session):
    return await create_profile(
        session,
        ProfileCreate(
            ad_soyad="Ehmed Eliyev",
            voen="1234567890",
            nomre="+994501234567",
            asan_id="AA123",
            pin1="1111",
            pin2="2222",
            puk="3333",
        ),
    )


async def _portal(session) -> Portal:
    portal = Portal(
        name="Test Bank",
        url="https://bank.test",
        category="bank",
        selector_mapping={
            "ad_soyad": "input[name=fullname]",
            "voen": "input[name=voen]",
            "asan_id": "input[name=asan]",
            "pin1": "input[name=pin1]",
        },
    )
    session.add(portal)
    await session.commit()
    await session.refresh(portal)
    return portal


async def test_get_autofill_profile_decrypts(session):
    profile = await _profile(session)
    payload = await get_autofill_profile(session, profile.id)
    assert payload is not None
    assert payload.pin1 == "1111"
    assert payload.pin2 == "2222"
    assert payload.puk == "3333"


async def test_get_autofill_profile_missing(session):
    import uuid

    assert await get_autofill_profile(session, uuid.uuid4()) is None


async def test_build_autofill_payload_maps_selectors(session):
    profile = await _profile(session)
    portal = await _portal(session)

    payload = await build_autofill_payload(session, profile.id, portal.id)
    assert payload.portal_name == "Test Bank"
    assert payload.fields["input[name=fullname]"] == "Ehmed Eliyev"
    assert payload.fields["input[name=voen]"] == "1234567890"
    assert payload.fields["input[name=asan]"] == "AA123"
    assert payload.fields["input[name=pin1]"] == "1111"


async def test_build_autofill_payload_missing_profile(session):
    portal = await _portal(session)
    import uuid

    with pytest.raises(ValueError):
        await build_autofill_payload(session, uuid.uuid4(), portal.id)


async def test_build_autofill_payload_missing_portal(session):
    profile = await _profile(session)
    import uuid

    with pytest.raises(ValueError):
        await build_autofill_payload(session, profile.id, uuid.uuid4())


async def test_build_autofill_payload_normalizes_phone(session):
    profile = await _profile(session)
    portal = Portal(
        name="Phone Bank",
        url="https://phone.test",
        category="bank",
        selector_mapping={"nomre": "input[name=phone]"},
    )
    session.add(portal)
    await session.commit()
    await session.refresh(portal)

    payload = await build_autofill_payload(session, profile.id, portal.id)
    assert payload.fields["input[name=phone]"] == "501234567"
