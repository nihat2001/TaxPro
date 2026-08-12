from langchain.agents import create_agent

from app.services.ai.agents.context_middleware import FreshContextMiddleware
from app.services.ai.llm import get_chat_model
from app.services.ai.middleware import RetrieverGuard
from app.services.ai.state import AgentState
from app.services.ai.tools.retriever_tools import search_profile_data

_SYSTEM_PROMPT = (
    "Sən profil məlumatları haqqında suallara cavab verən mütəxəssissən.\n"
    "İş ardıcıllığı:\n"
    "1. Hər yeni istifadəçi mesajı üçün search_profile_data tool-u ilə bazada MÜTLƏQ yenidən axtarış et — "
    "əvvəlki axtarışın nəticəsini əsas götürüb cavab yazma.\n"
    "2. YALNIZ search_profile_data tool-unun qaytardığı dəqiq məlumatı istifadəçiyə təqdim et. "
    "Bazada olmayan heç bir rəqəmi, nömrəni, ID-ni və ya məlumatı HEÇ VAXT uydurma və ya ehtimal etmə.\n"
    "3. PIN/PUK kimi gizli sahələr çat vasitəsilə açıqlanmır.\n"
    "İşini bitirəndə cavabı birbaşa mətn kimi yaz. Yadda saxla: transfer_to_* və ya "
    "transfer_back_to_* kimi ötürmə tool-ları sənin üçün mövcud deyil — onları heç vaxt çağırma, "
    "yalnız search_profile_data istifadə et."
)


def build_retriever_agent(model=None):
    return create_agent(
        model=model or get_chat_model(),
        tools=[search_profile_data],
        system_prompt=_SYSTEM_PROMPT,
        state_schema=AgentState,
        middleware=[FreshContextMiddleware(), RetrieverGuard()],
        name="profile_data_retriever_agent",
    )
