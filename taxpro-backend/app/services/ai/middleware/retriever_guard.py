import logging
import re

from langchain.agents.middleware import AgentMiddleware
from langchain_core.messages import AIMessage, ToolMessage

from app.core.config import get_settings

logger = logging.getLogger(__name__)

SEARCH_TOOL_NAME = "search_profile_data"
MIN_DIGIT_LEN = 5
SAFE_MESSAGE = (
    "Təhlükəsizlik xəbərdarlığı: Axtarış nəticəsində təsdiqlənməyən məlumat bloklandı. "
    "Zəhmət olmasa sorğunuzu dəqiqləşdirin."
)


def _last_human_index(messages) -> int:
    last_human = -1
    for index, message in enumerate(messages):
        if getattr(message, "type", "") in ("human", "user"):
            last_human = index
    return last_human


def _fresh_search_facts(messages) -> str | None:
    """Cari turda çağırılmış ən son search_profile_data nəticəsini qaytarır.

    Keçmiş turlardan qalma nəticələr təsdiq sayılmır: yalnız son HumanMessage-dən
    sonrakı ToolMessage hesab olunur. ToolMessage.name boş ola bilər, ona görə
    tool_call_id vasitəsilə əvvəlki AI tool-call adı ilə əlaqələndirilir.
    Tool yoxdursa və ya səhv qaytarıbsa, None.
    """
    last_human = _last_human_index(messages)
    if last_human < 0:
        return None
    call_names: dict[str, str] = {}
    for message in messages[last_human + 1 :]:
        for call in getattr(message, "tool_calls", None) or []:
            call_names[call.get("id")] = call.get("name")
        if isinstance(message, ToolMessage):
            if call_names.get(message.tool_call_id) != SEARCH_TOOL_NAME:
                continue
            if getattr(message, "status", "") == "error":
                return None
            return str(getattr(message, "content", "") or "")
    return None


def _last_text_answer(messages) -> tuple[int, AIMessage] | None:
    """Cari model çıxışını (ən son mesaj) mətn cavabı kimi qaytarır.

    Keçmiş turlardan qalma AI mesajları nəzərə alınmır — yalnız ən son mesaj
    doğrulana bilər, çünki after_model model çağırışından dərhal sonra işləyir.
    """
    if not messages:
        return None
    index = len(messages) - 1
    message = messages[index]
    if not isinstance(message, AIMessage):
        return None
    if getattr(message, "tool_calls", None):
        return None
    if not str(getattr(message, "content", "") or "").strip():
        return None
    return index, message


def _block_answer(messages, index: int, last_ai: AIMessage) -> dict:
    updated = AIMessage(
        content=SAFE_MESSAGE,
        id=last_ai.id,
        name=last_ai.name,
        tool_calls=last_ai.tool_calls,
    )
    new_messages = list(messages)
    new_messages[index] = updated
    return {"messages": new_messages}


def _provider_tool_choice() -> str:
    settings = get_settings()
    if settings.llm_provider == "anthropic":
        return "any"
    return "required"


class RetrieverGuard(AgentMiddleware):
    """Retriever agent üçün qoruyucu.

    - tool_choice: LLM provider-ə görə dinamik seçilir (groq → "required",
      anthropic → "any"). Yalnız cari turda hələ search_profile_data nəticəsi
      yoxkən məcbur edilir — nəticə gəldikdən sonra model sərbəst cavab verə bilir.
    - Post-hoc doğrulama: cavabdan ≥5 rəqəmli ardıcıllıqlar cari turdakı tool
      nəticəsi ilə müqayisə olunur. Uyğun gəlməyən (uydurma) və ya cari turda
      heç bir tool çağırışı olmayan cavab bloklanır və loglanır.
    """

    def wrap_model_call(self, request, handler):
        state = request.state or {}
        messages = list(state.get("messages") or [])
        if _fresh_search_facts(messages) is None:
            request = request.override(tool_choice=_provider_tool_choice())
        return handler(request)

    async def awrap_model_call(self, request, handler):
        state = request.state or {}
        messages = list(state.get("messages") or [])
        if _fresh_search_facts(messages) is None:
            request = request.override(tool_choice=_provider_tool_choice())
        return await handler(request)

    def after_model(self, state, runtime):
        messages = list(state.get("messages") or [])
        found = _last_text_answer(messages)
        if found is None:
            return None
        index, last_ai = found
        content = str(last_ai.content)

        facts = _fresh_search_facts(messages)
        if facts is None:
            logger.warning(
                "RetrieverGuard: cari turda %s çağırışı tapılmadı; cavab bloklandı. "
                "AI cavab: %s",
                SEARCH_TOOL_NAME,
                content[:300],
            )
            return _block_answer(messages, index, last_ai)

        answer_numbers = re.findall(r"\d{%d,}" % MIN_DIGIT_LEN, content)
        if not answer_numbers:
            return None

        fact_numbers = set(re.findall(r"\d+", facts))
        hallucinated = [number for number in answer_numbers if number not in fact_numbers]
        if hallucinated:
            logger.warning(
                "RetrieverGuard: uydurma rəqəmlər bloklandı: %s. AI cavab: %s",
                hallucinated,
                content[:300],
            )
            return _block_answer(messages, index, last_ai)

        return None

    async def aafter_model(self, state, runtime):
        return self.after_model(state, runtime)
