import io

import pandas as pd
import pytest

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


def _bank_excel() -> bytes:
    return _excel_bytes(
        ["Məbləğ", "Təyinat", "ƏDV"],
        [
            [120.50, "ƏDV ödənişi", 20.50],
            [300.00, "Ofis icarəsi", 0.0],
            [500.00, "Satışdan medaxil", 0.0],
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


async def test_bank_statement_analyze(client):
    files = {"file": ("bank.xlsx", _bank_excel(), XLSX_MIME)}
    resp = await client.post("/api/v1/bank-statements/analyze", files=files)
    assert resp.status_code == 200
    body = resp.json()
    assert body["status"] == "done"
    assert body["analysis"]["total_amount"] == 920.5
    assert body["analysis"]["vat_amount"] == 20.5


async def test_import_smart_interrupt_then_resume(client):
    files = {"file": ("smart.xlsx", _profile_excel(), XLSX_MIME)}
    resp = await client.post("/api/v1/import/smart", files=files)
    assert resp.status_code == 200
    body = resp.json()
    assert body["status"] == "interrupt"
    assert body["interrupt"]["action"] == "create_profiles"
    thread_id = body["thread_id"]

    cancelled = await client.post(
        "/api/v1/resume",
        json={"thread_id": thread_id, "decision": "cancel"},
    )
    assert cancelled.status_code == 200
    assert cancelled.json()["status"] == "done"
    assert cancelled.json()["inserted"] is None


async def test_import_smart_resume_confirm_inserts(client):
    files = {"file": ("smart.xlsx", _profile_excel(), XLSX_MIME)}
    resp = await client.post("/api/v1/import/smart", files=files)
    thread_id = resp.json()["thread_id"]

    confirmed = await client.post(
        "/api/v1/resume",
        json={"thread_id": thread_id, "decision": "confirm"},
    )
    assert confirmed.status_code == 200
    assert confirmed.json()["inserted"] == 2

    listed = await client.get("/api/v1/profiles")
    assert len(listed.json()) == 2


async def test_chat_retriever(client):
    resp = await client.post("/api/v1/chat", json={"message": "Salam"})
    assert resp.status_code == 200
    body = resp.json()
    assert body["status"] == "done"
    assert body["thread_id"]


async def test_chat_thread_continuity(client):
    from langchain_core.messages import AIMessage
    from langgraph.checkpoint.memory import MemorySaver

    import app.services.ai.graph as graph_mod
    from app.services.ai.graph import build_graph
    from tests.fakes import ScriptedModel

    class CountingModel(ScriptedModel):
        def _decide(self, messages):
            agent = self._identify(messages)
            called = [m.name for m in messages if getattr(m, "type", "") == "tool"]
            if agent == "supervisor":
                if any(n.startswith("transfer_to_") for n in called):
                    human_count = sum(
                        1 for m in messages if getattr(m, "type", "") in ("user", "human")
                    )
                    return AIMessage(content=f"Mesaj sayi: {human_count}")
                user_text = " ".join(
                    (m.content or "")
                    for m in messages
                    if getattr(m, "type", "") in ("user", "human")
                    and isinstance(m.content, str)
                ).lower()
                if any(k in user_text for k in ("bank", "cixaris", "çıxarış")):
                    name = "transfer_to_bank_vat_agent"
                elif any(k in user_text for k in ("excel", "fayl", "idxal")):
                    name = "transfer_to_unstructured_parser_agent"
                else:
                    name = "transfer_to_profile_data_retriever_agent"
                return AIMessage(
                    content="",
                    tool_calls=[{"name": name, "args": {}, "id": "call_sup"}],
                )
            return super()._decide(messages)

    graph_mod._graph = build_graph(
        checkpointer=MemorySaver(), model=CountingModel()
    )

    try:
        first = await client.post("/api/v1/chat", json={"message": "Salam"})
        assert first.status_code == 200
        body1 = first.json()
        thread_id = body1["thread_id"]
        assert body1["status"] == "done"
        assert "1" in body1["message"]

        second = await client.post(
            "/api/v1/chat",
            json={"message": "Necəsiniz?", "thread_id": thread_id},
        )
        assert second.status_code == 200
        body2 = second.json()
        assert body2["thread_id"] == thread_id
        assert "2" in body2["message"]

        fresh = await client.post("/api/v1/chat", json={"message": "Salam"})
        body3 = fresh.json()
        assert body3["thread_id"] != thread_id
        assert "1" in body3["message"]
    finally:
        graph_mod._graph = None


async def test_chat_multiturn_clarification(client):
    first = await client.post("/api/v1/chat", json={"message": "salam nömrəni yaz"})
    assert first.status_code == 200
    body1 = first.json()
    assert body1["status"] == "done"
    thread_id = body1["thread_id"]
    assert thread_id

    second = await client.post(
        "/api/v1/chat",
        json={"message": "Orxan Oğuzlu", "thread_id": thread_id},
    )
    assert second.status_code == 200
    body2 = second.json()
    assert body2["thread_id"] == thread_id
    assert body2["status"] == "done"
    assert body2["message"]

    third = await client.post(
        "/api/v1/chat",
        json={"message": "Orxan Oğuzlunun nömrəsini yaz", "thread_id": thread_id},
    )
    assert third.status_code == 200
    body3 = third.json()
    assert body3["thread_id"] == thread_id
    assert body3["status"] == "done"
    assert body3["message"]


async def test_chat_multiturn_retriever_re_searches_db(client, monkeypatch):
    import app.services.ai.graph as graph_mod
    import app.services.ai.tools.retriever_tools as retriever_tools
    from langchain_core.messages import AIMessage
    from langgraph.checkpoint.memory import MemorySaver
    from app.services.ai.graph import build_graph
    from tests.fakes import ScriptedModel

    class ReDelegateModel(ScriptedModel):
        def _decide(self, messages):
            agent = self._identify(messages)
            if agent != "supervisor":
                return super()._decide(messages)
            last_transfer_idx = max(
                (
                    i
                    for i, m in enumerate(messages)
                    if getattr(m, "type", "") == "tool"
                    and str(getattr(m, "name", "")).startswith("transfer_")
                ),
                default=-1,
            )
            last_human_idx = max(
                (
                    i
                    for i, m in enumerate(messages)
                    if getattr(m, "type", "") in ("human", "user")
                ),
                default=-1,
            )
            if last_human_idx > last_transfer_idx:
                return AIMessage(
                    content="",
                    tool_calls=[
                        {
                            "name": "transfer_to_profile_data_retriever_agent",
                            "args": {},
                            "id": "call_sup",
                        }
                    ],
                )
            return AIMessage(content="Yekun cavab")

    calls = {"n": 0}
    original = retriever_tools.search_profiles

    async def counting_search(session, query, profile_id=None, limit=5):
        calls["n"] += 1
        return await original(session, query, profile_id, limit)

    monkeypatch.setattr(retriever_tools, "search_profiles", counting_search)

    graph_mod._graph = build_graph(
        checkpointer=MemorySaver(), model=ReDelegateModel()
    )
    try:
        thread_id = None
        for message in ("salam nömrəni yaz", "Orxan Oğuzlu", "Orxan Oğuzlunun nömrəsini yaz"):
            payload = {"message": message}
            if thread_id:
                payload["thread_id"] = thread_id
            resp = await client.post("/api/v1/chat", json=payload)
            assert resp.status_code == 200
            body = resp.json()
            assert body["status"] == "done"
            thread_id = body["thread_id"]
        assert calls["n"] >= 3, (
            "Hər turda axtarış yenidən icra olunmalıdır; try_count limiti "
            "keçmiş turn-dan daşınmamalıdır"
        )
    finally:
        graph_mod._graph = None
