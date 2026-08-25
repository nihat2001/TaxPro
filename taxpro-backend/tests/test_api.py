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


def _profile_csv() -> bytes:
    df = pd.DataFrame(
        [
            ["Ehmed Eliyev", 1234567890, "0501234567", "AA123", "1234", "5678", "9999"],
            ["Aysel Quliyeva", 9876543210, "0551234567", "BB456", "4321", "8765", "1111"],
        ],
        columns=["Ad Soyad", "VÖEN", "Nömrə", "Asan ID", "PIN1", "PIN2", "PUK"],
    )
    buf = io.BytesIO()
    df.to_csv(buf, index=False)
    buf.seek(0)
    return buf.read()


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CSV_MIME = "text/csv"


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
    assert "Əskik sütunlar" in resp.json()["detail"]


async def test_import_excel_csv(client):
    files = {"file": ("profiles.csv", _profile_csv(), CSV_MIME)}
    resp = await client.post("/api/v1/import/excel", files=files)
    assert resp.status_code == 200
    body = resp.json()
    assert body["status"] == "done"
    assert body["inserted"] == 2

    listed = await client.get("/api/v1/profiles")
    assert len(listed.json()) == 2


async def test_import_excel_template(client):
    resp = await client.get("/api/v1/import/excel-template")
    assert resp.status_code == 200
    assert "spreadsheetml" in resp.headers["content-type"]
    assert "attachment" in resp.headers.get("content-disposition", "")


# ===== Bank çıxarışı kateqoriyalaşdırma (HTTP) =====


def _bank_statement_bytes() -> bytes:
    import io

    from openpyxl import Workbook

    headers = [
        "Əməliyyat Tarixi",
        "Həyata Keçmə Tarixi",
        "Göndərən/Alıcı Məlumatları",
        "Əməliyyat Detalları",
        "Əməliyyat No",
        "Kredit",
        "Debit",
        "Balans",
        "Kateqoriya",
    ]
    rows = [
        ["01.01.2026", "01.01.2026", "", "ATM XIRDALAN BIZIM M", "1", "", "-1000.00", "9000.00", ""],
        ["02.01.2026", "02.01.2026", "", "PIN YENI UCRET", "2", "", "-2.00", "8998.00", ""],
        ["03.01.2026", "03.01.2026", "", "Com. bank komissiyası", "3", "", "-5.50", "8992.50", ""],
    ]
    wb = Workbook()
    ws = wb.active
    ws.append(["Qurum Adı", "X Bank"])
    ws.append(headers)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


XLSX_MIME_BS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


async def test_categorize_bank_statement_json(client):
    resp = await client.post(
        "/api/v1/import/bank-statement/categorize",
        files={"file": ("bs.xlsx", _bank_statement_bytes(), XLSX_MIME_BS)},
    )
    assert resp.status_code == 200
    body = resp.json()
    assert "KOMİSİYA" in body["summary"]
    assert body["summary"]["KOMİSİYA"]["cem"] == 5.5
    assert len(body["unclassified"]) == 2
    assert body["unclassified"][0]["row_index"] == 3
    assert body["unclassified"][0]["detail"] == "ATM XIRDALAN BIZIM M"
    # file_base64 → valid xlsx
    import base64

    raw = base64.b64decode(body["file_base64"])
    assert raw[:2] == b"PK"


async def test_finalize_bank_statement(client):
    import base64
    import json

    raw = _bank_statement_bytes()
    r1 = await client.post(
        "/api/v1/import/bank-statement/categorize",
        files={"file": ("bs.xlsx", raw, XLSX_MIME_BS)},
    )
    unclassified = r1.json()["unclassified"]
    atm = next(u for u in unclassified if "ATM" in u["detail"])["row_index"]
    pin = next(u for u in unclassified if "PIN" in u["detail"])["row_index"]

    selections = json.dumps({atm: "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ", pin: ""})
    resp = await client.post(
        "/api/v1/import/bank-statement/finalize",
        files={"file": ("bs.xlsx", raw, XLSX_MIME_BS)},
        data={"selections": selections},
    )
    assert resp.status_code == 200
    assert "spreadsheetml" in resp.headers["content-type"]
    summary = json.loads(base64.b64decode(resp.headers["x-summary"]).decode("utf-8"))
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"]["cem"] == 1000.0
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"]["say"] == 1
    assert summary["KOMİSİYA"]["say"] == 1
    # override sütunlara tətbiq olunub
    from openpyxl import load_workbook

    wb = load_workbook(__import__("io").BytesIO(resp.content))
    main = wb[wb.sheetnames[0]]
    assert main.cell(row=atm, column=9).value == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    assert main.cell(row=pin, column=9).value is None


async def test_finalize_bank_statement_invalid(client):
    import json

    resp = await client.post(
        "/api/v1/import/bank-statement/finalize",
        files={"file": ("bs.xlsx", _bank_statement_bytes(), XLSX_MIME_BS)},
        data={"selections": json.dumps({999: "KOMİSİYA"})},
    )
    assert resp.status_code == 422

    resp = await client.post(
        "/api/v1/import/bank-statement/finalize",
        files={"file": ("bs.xlsx", _bank_statement_bytes(), XLSX_MIME_BS)},
        data={"selections": "not-json"},
    )
    assert resp.status_code == 422
