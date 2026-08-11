from langchain_core.messages import ToolMessage
from langgraph.types import Command

from app.services.ai.state import AgentState

SENSITIVE_METADATA_KEY = "sensitive"
DEFAULT_MAX_TRIES = 3


def mark_sensitive(tool) -> None:
    tool.metadata = {
        **(tool.metadata or {}),
        SENSITIVE_METADATA_KEY: True,
    }


def bump_try(state: AgentState) -> dict:
    return {"try_count": state.get("try_count", 0) + 1}


def try_limit_reached(state: AgentState) -> bool:
    max_try = state.get("max_try", DEFAULT_MAX_TRIES)
    return state.get("try_count", 0) >= max_try


def check_try(state: AgentState, tool_call_id: str) -> tuple[Command | None, int]:
    new_try = state.get("try_count", 0) + 1
    if new_try > state.get("max_try", DEFAULT_MAX_TRIES):
        return (
            Command(
                update={
                    "messages": [
                        tool_message(
                            "Cəhd limiti aşıldı, əməliyyat dayandırıldı.",
                            tool_call_id,
                        )
                    ]
                }
            ),
            new_try,
        )
    return (None, new_try)


def tool_message(content: str, tool_call_id: str) -> ToolMessage:
    return ToolMessage(content=content, tool_call_id=tool_call_id)


def section_command(section: str, update: dict, message: ToolMessage) -> Command:
    return Command(update={section: update, "messages": [message]})
