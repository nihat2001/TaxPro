import uuid

from langchain_core.messages import HumanMessage

from app.schemas.ai import AgentRunResponse
from app.services.ai.graph import invoke_graph, resume_graph


def new_thread_id() -> str:
    return uuid.uuid4().hex


def _base_state(message: str) -> dict:
    return {
        "messages": [HumanMessage(content=message)],
        "parser": {},
        "bank": {},
        "retriever": {},
        "try_count": 0,
        "max_try": 3,
    }


async def run_import(
    thread_id: str, file_bytes_b64: str, source_name: str
) -> dict:
    state = _base_state("Excel faylını idxal et.")
    state["parser"] = {"file_bytes": file_bytes_b64, "source_name": source_name}
    return await invoke_graph(state, thread_id)


async def run_bank_analysis(thread_id: str, file_bytes_b64: str) -> dict:
    state = _base_state("Bank çıxarışını təhlil et.")
    state["bank"] = {"file_bytes": file_bytes_b64, "source": "excel"}
    return await invoke_graph(state, thread_id)


async def run_chat(thread_id: str, message: str, profile_id: str | None) -> dict:
    state = _base_state(message)
    state["retriever"] = {"query": message, "profile_id": profile_id}
    return await invoke_graph(state, thread_id)


async def resume(thread_id: str, decision: str) -> dict:
    return await resume_graph(thread_id, decision)


def _interrupt_payload(state: dict) -> dict | None:
    interrupts = state.get("__interrupt__")
    if not interrupts:
        return None
    return interrupts[0].value


def _final_message(state: dict) -> str:
    for message in reversed(state.get("messages", [])):
        content = getattr(message, "content", None)
        if content:
            return str(content)
    return ""


def to_run_response(thread_id: str, state: dict) -> AgentRunResponse:
    interrupt = _interrupt_payload(state)
    return AgentRunResponse(
        thread_id=thread_id,
        status="interrupt" if interrupt else "done",
        message=_final_message(state),
        interrupt=interrupt,
        analysis=(state.get("bank") or {}).get("analysis"),
        inserted=(state.get("parser") or {}).get("inserted"),
        sources=(state.get("retriever") or {}).get("sources", []),
    )
