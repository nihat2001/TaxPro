from langchain.agents.middleware import AgentMiddleware
from langchain_core.messages import ToolMessage

_HANDOFF_PREFIXES = ("transfer_to_", "transfer_back_to_")


def _is_handoff_noise(message) -> bool:
    if isinstance(message, ToolMessage) and str(
        getattr(message, "name", "")
    ).startswith(_HANDOFF_PREFIXES):
        return True
    tool_calls = getattr(message, "tool_calls", None) or []
    if tool_calls and all(
        str(call.get("name", "")).startswith(_HANDOFF_PREFIXES) for call in tool_calls
    ):
        return True
    return False


class FreshContextMiddleware(AgentMiddleware):
    """Sub-aqentin model görüntüsünü yalnız cari tura endirir.

    Supervizorun handoff (transfer_*) mesajları və keçmiş turların konteksti
    sub-aqentin model çağırışına daxil edilmir. Beləcə hər handoff-da sub-aqent
    tool-larını yenidən çağırır və əvvəlki turun nəticəsini (məs. "tapılmadı")
    təkrarlayaraq uydurma cavab vermir.
    """

    def before_model(self, state, runtime):
        messages = list(state.get("messages") or [])
        last_human = -1
        for index, message in enumerate(messages):
            if getattr(message, "type", "") in ("human", "user"):
                last_human = index
        if last_human < 0:
            return None
        kept = [
            message
            for message in messages[last_human:]
            if not _is_handoff_noise(message)
        ]
        return {"messages": kept}
