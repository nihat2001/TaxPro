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
            nomre="+994501234567",
            asan_id="AA123",
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
            "nomre": "input[name=phone]",
            "pin1": "input[name=pin1]",
        },
    )
    session.add(portal)
    await session.commit()
    await session.refresh(portal)
    return portal


async def test_get_autofill_profile(session):
    profile = await _profile(session)
    payload = await get_autofill_profile(session, profile.id)
    assert payload is not None
    assert payload.ad_soyad == "Ehmed Eliyev"
    assert payload.nomre == "501234567"
    assert payload.asan_id == "AA123"


async def test_get_autofill_profile_missing(session):
    import uuid

    assert await get_autofill_profile(session, uuid.uuid4()) is None


async def test_build_autofill_payload_only_nomre_asan_id(session):
    profile = await _profile(session)
    portal = await _portal(session)

    payload = await build_autofill_payload(session, profile.id, portal.id)
    assert payload.portal_name == "Test Bank"
    # Yalnız nomre + asan_id doldurulur; ad_soyad/voen/pin1 saytda yazılmır.
    assert payload.fields["input[name=phone]"] == "501234567"
    assert payload.fields["input[name=asan]"] == "AA123"
    assert "input[name=fullname]" not in payload.fields
    assert "input[name=voen]" not in payload.fields
    assert "input[name=pin1]" not in payload.fields


async def test_build_autofill_payload_phone_prefix_selector(session):
    profile = await _profile(session)
    portal = await _portal(session)
    portal.selector_mapping = {
        **portal.selector_mapping,
        "_phone_prefix_selector": "select[name=prefix]",
    }
    await session.commit()
    await session.refresh(portal)

    payload = await build_autofill_payload(session, profile.id, portal.id)
    assert payload.phone_prefix_selector == "select[name=prefix]"


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