from langchain.agents import create_agent

from app.services.ai.llm import get_chat_model
from app.services.ai.state import AgentState
from app.services.ai.tools.retriever_tools import search_profile_data

_SYSTEM_PROMPT = (
    "Sən profil məlumatları haqqında suallara cavab verən mütəxəssissən.\n"
    "İş ardıcıllığı:\n"
    "1. search_profile_data tool-u ilə bazada axtarış et.\n"
    "2. Tool-un qaytardığı dəqiq cavabı istifadəçiyə təqdim et; heç vaxt uydurma məlumat vermə.\n"
    "PIN/PUK kimi gizli sahələr çat vasitəsilə açıqlanmır.\n"
    "İşini bitirəndə supervisor-a qayıt."
)


def build_retriever_agent(model=None):
    return create_agent(
        model=model or get_chat_model(),
        tools=[search_profile_data],
        system_prompt=_SYSTEM_PROMPT,
        state_schema=AgentState,
        name="profile_data_retriever_agent",
    )
