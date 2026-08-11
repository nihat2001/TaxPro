import logging

from langgraph.checkpoint.memory import MemorySaver
from langgraph.types import Command

from app.core.config import get_settings
from app.services.ai.agents.bank_vat import build_bank_vat_agent
from app.services.ai.agents.parser import build_parser_agent
from app.services.ai.agents.retriever import build_retriever_agent
from app.services.ai.agents.supervisor import build_supervisor
from app.services.ai.checkpointer import get_checkpointer

logger = logging.getLogger(__name__)

_graph = None


def build_graph(checkpointer=None, model=None):
    parser_agent = build_parser_agent(model=model)
    bank_agent = build_bank_vat_agent(model=model)
    retriever_agent = build_retriever_agent(model=model)
    builder = build_supervisor([parser_agent, bank_agent, retriever_agent], model=model)

    if checkpointer is None:
        checkpointer = get_checkpointer()
    if checkpointer is None:
        settings = get_settings()
        if settings.app_env == "production":
            raise RuntimeError(
                "Production mühitində checkpointer yoxdur — MemorySaver fallback-i qadağandır. "
                "App start oluna bilməz."
            )
        logger.warning(
            "Checkpointer yoxdur, MemorySaver istifadə olunur (yalnız development). "
            "HITL/resume state-i restart-da itir."
        )
        checkpointer = MemorySaver()
    return builder.compile(checkpointer=checkpointer)


def get_graph(checkpointer=None):
    global _graph
    if checkpointer is not None:
        return build_graph(checkpointer=checkpointer)
    if _graph is None:
        _graph = build_graph()
    return _graph


def _config(thread_id: str) -> dict:
    return {
        "recursion_limit": get_settings().recursion_limit,
        "configurable": {"thread_id": thread_id},
    }


def _log_history(state: dict, action: str) -> None:
    messages = state.get("messages") or []
    for message in messages:
        msg_type = type(message).__name__
        tool_calls = getattr(message, "tool_calls", None)
        if tool_calls:
            for call in tool_calls:
                logger.info(
                    "%s | AI tool_call -> %s(%s)",
                    action,
                    call.get("name"),
                    str(call.get("args"))[:200],
                )
        content = str(getattr(message, "content", "") or "").strip()
        if content and msg_type in ("AIMessage", "HumanMessage", "ToolMessage"):
            logger.info(
                "%s | %s -> %s", action, msg_type, content[:200].replace("\n", " ")
            )


async def invoke_graph(initial_state: dict, thread_id: str, checkpointer=None) -> dict:
    graph = get_graph(checkpointer=checkpointer)
    state = {**initial_state}
    state.setdefault("remaining_steps", 25)
    result = await graph.ainvoke(state, config=_config(thread_id))
    _log_history(result, f"graph:{thread_id}")
    return result


async def resume_graph(thread_id: str, value, checkpointer=None) -> dict:
    graph = get_graph(checkpointer=checkpointer)
    result = await graph.ainvoke(Command(resume=value), config=_config(thread_id))
    _log_history(result, f"resume:{thread_id}")
    return result
