import logging
from typing import Annotated

from langchain_core.tools import InjectedToolCallId, tool
from langgraph.prebuilt import InjectedState
from langgraph.types import Command

from app.core.database import async_session_maker
from app.services.ai.state import AgentState
from app.services.ai.tools._common import check_try, tool_message
from app.services.ai.tools.db_retriever_tool import search_documents, search_profiles

logger = logging.getLogger(__name__)

_FIELD_KEYWORDS = {
    "asan_id": ["asan id", "asan", "asandid"],
    "voen": ["voen", "vergi nomresi"],
    "pin1": ["pin1", "pin 1", "pin"],
    "puk": ["puk"],
    "nomre": ["nomre", "nomer", "nomresi", "telefon", "tel", "mobil"],
    "bitme_tarixi": ["bitme tarixi", "bitirilme", "etibarliliq"],
    "qeyd": ["qeyd"],
}

_SECRET_FIELDS = {"pin1", "pin2", "puk"}


def _normalize(text: str) -> str:
    result = text.lower()
    for char, repl in {"ə": "e", "ö": "o", "ü": "u", "ı": "i", "ş": "s", "ç": "c", "ğ": "g"}.items():
        result = result.replace(char, repl)
    return result


def _extract_field(query: str) -> str | None:
    normalized = _normalize(query)
    for field, keywords in _FIELD_KEYWORDS.items():
        if any(keyword in normalized for keyword in keywords):
            return field
    return None


def _profile_answer(profile, field: str | None) -> str:
    label = profile.ad_soyad
    if field == "asan_id":
        return f"{label} — Asan ID: {profile.asan_id}"
    if field == "voen":
        return f"{label} — VÖEN: {profile.voen}"
    if field == "nomre":
        return f"{label} — Telefon: {profile.nomre}"
    if field == "bitme_tarixi":
        date_text = profile.bitme_tarixi.isoformat() if profile.bitme_tarixi else "məlumat yoxdur"
        return f"{label} — Bitmə tarixi: {date_text}"
    if field == "qeyd":
        return f"{label} — Qeyd: {profile.qeyd or 'məlumat yoxdur'}"
    if field in _SECRET_FIELDS:
        return f"{label} — bu məlumat çat vasitəsilə açıqlanmır."
    return f"{label} — VÖEN: {profile.voen}, Asan ID: {profile.asan_id}, Telefon: {profile.nomre}"


@tool
async def search_profile_data(
    state: Annotated[AgentState, InjectedState],
    tool_call_id: Annotated[str, InjectedToolCallId],
) -> Command:
    """Profil məlumatlarını bazada axtarır və yalnız verilənlər bazasına əsaslanan dəqiq cavab hazırlayır. Gizli sahələr (PIN/PUK) açıqlanmır."""
    error, new_try = check_try(state, tool_call_id)
    if error is not None:
        return error

    retriever = state.get("retriever") or {}
    query = str(retriever.get("query", ""))
    profile_id = retriever.get("profile_id")

    async with async_session_maker() as session:
        profiles = await search_profiles(session, query, profile_id)
        field = _extract_field(query)
        facts = [_profile_answer(p, field) for p in profiles]

        if not facts:
            documents = await search_documents(session, query, profile_id)
            facts = [f"Qeyd: {d.content[:300]}" for d in documents]

        sources = [str(p.id) for p in profiles]
    logger.info(
        "search_profile_data query=%r profile_id=%r tapildi=%d",
        query,
        profile_id,
        len(profiles),
    )

    if not facts:
        answer = "Bazada uyğun məlumat tapılmadı."
    else:
        answer = facts[0]

    message = tool_message(answer, tool_call_id)
    return Command(
        update={
            "retriever": {"answer": answer, "sources": sources},
            "try_count": new_try,
            "messages": [message],
        }
    )
