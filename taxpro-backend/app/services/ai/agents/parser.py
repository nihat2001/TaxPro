from langchain.agents import create_agent

from app.services.ai.agents.context_middleware import FreshContextMiddleware
from app.services.ai.llm import get_chat_model
from app.services.ai.state import AgentState
from app.services.ai.tools.parser_tools import create_profiles_from_rows, parse_excel_file

_SYSTEM_PROMPT = (
    "Sən qarışıq və ya sütunları pozulmuş Excel faylını emal edən mütəxəssissən.\n"
    "İş ardıcıllığı:\n"
    "1. parse_excel_file tool-u ilə faylı oxu və sütunları standart sahələrə uyğunlaşdır.\n"
    "2. Profillər yaradılmalıdırsa, create_profiles_from_rows tool-unu çağır. Bu tool həssasdır — "
    "istifadəçi təsdiqi olmadan yazma baş verməz.\n"
    "İşini bitirəndə cavabı birbaşa mətn kimi yaz. Yadda saxla: transfer_to_* və ya "
    "transfer_back_to_* kimi ötürmə tool-ları sənin üçün mövcud deyil — onları heç vaxt çağırma."
)


def build_parser_agent(model=None):
    return create_agent(
        model=model or get_chat_model(),
        tools=[parse_excel_file, create_profiles_from_rows],
        system_prompt=_SYSTEM_PROMPT,
        state_schema=AgentState,
        middleware=[FreshContextMiddleware()],
        name="unstructured_parser_agent",
    )
