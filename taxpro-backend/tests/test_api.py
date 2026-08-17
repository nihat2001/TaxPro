import io

import pandas as pd

from app.core.database import async_session_maker
from app.models.portal import Portal


async def _seed_profile(client, **overrides) -> dict:
    data = {
        "ad_soyad": "API Test Profili",
        "voen": "1234567890",
        "nomre": "+994501234567",
        "asan_id": "AP01",
        "pin1": "1111",
        "pin2": "2222",
        "puk": "3333",
    }
    data.update(overrides)
    resp = await client.post("/api/v1/profiles", json=data)
    assert resp.status_code == 201
    return resp.json()


async def _seed_portal(client) -> dict:
    portal = Portal(
        name="API Test Bank",
        url="https://bank.api.test",
        category="bank",
        selector_mapping={
            "ad_soyad": "input#fullname",
            "voen": "input#voen",
            "asan_id": "input#asan",
            "pin1": "input#pin1",
        },
    )
    async with async_session_maker() as session:
        session.add(portal)
        await session.commit()
        await session.refresh(portal)
        portal_id = portal.id
    resp = await client.get(f"/api/v1/portals/{portal_id}")
    assert resp.status_code == 200
    return resp.json()


def _excel_bytes(headers: list[str], rows: list[list]) -> bytes:
    df = pd.DataFrame(rows, columns=headers)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf.read()


def _profile_excel() -> bytes:
    return _excel_bytes(
        ["Ad Soyad", "VÖEN", "Nömrə", "Asan ID", "PIN1", "PIN2", "PUK"],
        [
            ["Ehmed Eliyev", 1234567890, "0501234567", "AA123", "1234", "5678", "9999"],
            ["Aysel Quliyeva", 9876543210, "0551234567", "BB456", "4321", "8765", "1111"],
        ],
    )


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


async def test_health(client):
    resp = await client.get("/health")
    assert resp.status_code == 200
    assert resp.json() == {"status": "ok"}


async def test_profiles_crud(client):
    created = await _seed_profile(client)
    pid = created["id"]
    assert created["pin1"] == "****"
    assert created["pin2"] == "****"

    listed = await client.get("/api/v1/profiles")
    assert listed.status_code == 200
    assert any(p["id"] == pid for p in listed.json())

    one = await client.get(f"/api/v1/profiles/{pid}")
    assert one.status_code == 200
    assert one.json()["ad_soyad"] == "API Test Profili"

    updated = await client.patch(f"/api/v1/profiles/{pid}", json={"asan_id": "AP99"})
    assert updated.status_code == 200
    assert updated.json()["asan_id"] == "AP99"

    deleted = await client.delete(f"/api/v1/profiles/{pid}")
    assert deleted.status_code == 204

    missing = await client.get(f"/api/v1/profiles/{pid}")
    assert missing.status_code == 404


async def test_profile_autofill_decrypts(client):
    created = await _seed_profile(client)
    resp = await client.get(f"/api/v1/profiles/{created['id']}/autofill")
    assert resp.status_code == 200
    body = resp.json()
    assert body["pin1"] == "1111"
    assert body["puk"] == "3333"


async def test_portals_autofill(client):
    profile = await _seed_profile(client)
    portal = await _seed_portal(client)

    resp = await client.get(
        f"/api/v1/portals/{portal['id']}/autofill/{profile['id']}"
    )
    assert resp.status_code == 200
    body = resp.json()
    assert body["portal_name"] == "API Test Bank"
    assert body["fields"]["input#fullname"] == "API Test Profili"
    assert body["fields"]["input#pin1"] == "1111"


async def test_portal_selector_mapping_update(client):
    portal = await _seed_portal(client)
    assert portal["auth_type"] == "unknown"

    resp = await client.patch(
        f"/api/v1/portals/{portal['id']}/selector-mapping",
        json={
            "selector_mapping": {"asan_id": "input#asanid", "nomre": "input#phone"},
            "auth_type": "asan_id_form",
        },
    )
    assert resp.status_code == 200
    body = resp.json()
    assert body["auth_type"] == "asan_id_form"
    assert body["selector_mapping"]["asan_id"] == "input#asanid"

    resp = await client.patch(
        f"/api/v1/portals/{portal['id']}/selector-mapping",
        json={"selector_mapping": {}, "auth_type": "bogus"},
    )
    assert resp.status_code == 422


async def test_import_excel_standard(client):
    files = {"file": ("profiles.xlsx", _profile_excel(), XLSX_MIME)}
    resp = await client.post("/api/v1/import/excel", files=files)
    assert resp.status_code == 200
    body = resp.json()
    assert body["status"] == "done"
    assert body["inserted"] == 2

    listed = await client.get("/api/v1/profiles")
    assert len(listed.json()) == 2


async def test_import_excel_invalid_headers(client):
    files = {"file": ("bad.xlsx", _excel_bytes(["foo", "bar"], [[1, 2]]), XLSX_MIME)}
    resp = await client.post("/api/v1/import/excel", files=files)
    assert resp.status_code == 422
