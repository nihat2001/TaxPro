import base64
import io

import pandas as pd
from langchain_core.messages import HumanMessage
from langgraph.checkpoint.memory import MemorySaver
from langgraph.types import Command

from app.services.ai.graph import build_graph


def _bank_excel_b64() -> str:
    df = pd.DataFrame(
        [
            {"Məbləğ": 120.50, "Təyinat": "ƏDV ödənişi", "ƏDV": 20.50},
            {"Məbləğ": 300.00, "Təyinat": "Ofis icarəsi", "ƏDV": 0.0},
            {"Məbləğ": 500.00, "Təyinat": "Satışdan medaxil", "ƏDV": 0.0},
        ]
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()


def _profile_excel_b64() -> str:
    df = pd.DataFrame(
        [
            {"Ad Soyad": "Ehmed Eliyev", "VÖEN": 1234567890, "Nömrə": "0501234567",
             "Asan ID": "AA123", "PIN1": "1234", "PIN2": "5678", "PUK": "9999"},
            {"Ad Soyad": "Aysel Quliyeva", "VÖEN": 9876543210, "Nömrə": "0551234567",
             "Asan ID": "BB456", "PIN1": "4321", "PIN2": "8765", "PUK": "1111"},
        ]
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()


def _config(thread_id: str) -> dict:
    return {"configurable": {"thread_id": thread_id}}


async def test_bank_analysis_flow(fake_model):
    graph = build_graph(checkpointer=MemorySaver(), model=fake_model)
    state = {
        "messages": [HumanMessage(content="Bank çıxarışını təhlil et")],
        "bank": {"file_bytes": _bank_excel_b64(), "source": "excel"},
        "parser": {},
        "retriever": {},
        "try_count": 0,
        "max_try": 3,
    }
    result = await graph.ainvoke(state, config=_config("t_bank"))
    last = result["messages"][-1]
    assert "tamamlandı" in last.content
    analysis = result["bank"]["analysis"]
    assert analysis["total_amount"] == 920.5
    assert analysis["vat_amount"] == 20.5
    assert analysis["income_amount"] == 500.0


async def test_parser_interrupt_before_write(fake_model):
    graph = build_graph(checkpointer=MemorySaver(), model=fake_model)
    state = {
        "messages": [HumanMessage(content="Excel faylını idxal et")],
        "parser": {"file_bytes": _profile_excel_b64(), "source_name": "test.xlsx"},
        "bank": {},
        "retriever": {},
        "try_count": 0,
        "max_try": 3,
    }
    result = await graph.ainvoke(state, config=_config("t_import"))
    assert result.get("__interrupt__")
    payload = result["__interrupt__"][0].value
    assert payload["action"] == "create_profiles"
    assert payload["profile_count"] == 2
    assert not (result.get("parser") or {}).get("inserted")


async def test_parser_resume_cancel_writes_nothing(fake_model):
    graph = build_graph(checkpointer=MemorySaver(), model=fake_model)
    state = {
        "messages": [HumanMessage(content="Excel faylını idxal et")],
        "parser": {"file_bytes": _profile_excel_b64(), "source_name": "test.xlsx"},
        "bank": {},
        "retriever": {},
        "try_count": 0,
        "max_try": 3,
    }
    await graph.ainvoke(state, config=_config("t_cancel"))
    result2 = await graph.ainvoke(
        Command(resume="cancel"), config=_config("t_cancel")
    )
    last = result2["messages"][-1]
    assert "tamamlandı" in last.content
    assert not (result2.get("parser") or {}).get("inserted")


async def test_parser_resume_confirm_inserts(fake_model):
    graph = build_graph(checkpointer=MemorySaver(), model=fake_model)
    state = {
        "messages": [HumanMessage(content="Excel faylını idxal et")],
        "parser": {"file_bytes": _profile_excel_b64(), "source_name": "test.xlsx"},
        "bank": {},
        "retriever": {},
        "try_count": 0,
        "max_try": 3,
    }
    await graph.ainvoke(state, config=_config("t_confirm"))
    result2 = await graph.ainvoke(
        Command(resume="confirm"), config=_config("t_confirm")
    )
    assert (result2.get("parser") or {}).get("inserted") == 2


async def test_retriever_returns_db_answer(fake_model):
    from app.core.database import async_session_maker
    from app.schemas.profile import ProfileCreate
    from app.services.deterministic.profile_service import create_profile

    async with async_session_maker() as session:
        await create_profile(
            session,
            ProfileCreate(
                ad_soyad="Ehmed Eliyev",
                voen="1234567890",
                nomre="+994501234567",
                asan_id="AA123",
                pin1="1234",
                pin2="5678",
                puk="9999",
            ),
        )

    graph = build_graph(checkpointer=MemorySaver(), model=fake_model)
    state = {
        "messages": [HumanMessage(content="Ehmed Asan ID-si nədir?")],
        "retriever": {"query": "Ehmed Asan ID-si nədir?", "profile_id": None},
        "parser": {},
        "bank": {},
        "try_count": 0,
        "max_try": 3,
    }
    result = await graph.ainvoke(state, config=_config("t_retriever"))
    retriever = result.get("retriever") or {}
    assert retriever.get("answer")
    assert "AA123" in retriever["answer"]


async def test_no_synthetic_handoff_back_in_multiturn_state(fake_model):
    graph = build_graph(checkpointer=MemorySaver(), model=fake_model)
    thread = "t_multiturn_clean"
    for text in (
        "salam nömrəni yaz",
        "Orxan Oğuzlu",
        "Orxan Oğuzlunun nömrəsini yaz",
    ):
        state = {
            "messages": [HumanMessage(content=text)],
            "parser": {},
            "bank": {},
            "retriever": {},
            "try_count": 0,
            "max_try": 3,
        }
        result = await graph.ainvoke(state, config=_config(thread))

    for message in result["messages"]:
        tool_calls = getattr(message, "tool_calls", None) or []
        for call in tool_calls:
            assert call["name"] != "transfer_back_to_supervisor"
