from langchain.agents import create_agent

from app.services.ai.llm import get_chat_model
from app.services.ai.state import AgentState
from app.services.ai.tools.bank_tools import analyze_bank_statement, update_profile_from_analysis

_SYSTEM_PROMPT = (
    "Sən bank çıxarışlarını analiz edən mütəxəssissən.\n"
    "İş ardıcıllığı:\n"
    "1. analyze_bank_statement tool-u ilə çıxarışı analiz et: ümumi məbləğ, ƏDV, xərclər, gəlirlər.\n"
    "2. Yalnız istifadəçi açıq şəkildə xahiş edərsə, update_profile_from_analysis tool-u ilə nəticəni "
    "yaddaşa yaz (həssasdır — təsdiq gözlənilir).\n"
    "İşini bitirəndə supervisor-a qayıt."
)


def build_bank_vat_agent(model=None):
    return create_agent(
        model=model or get_chat_model(),
        tools=[analyze_bank_statement, update_profile_from_analysis],
        system_prompt=_SYSTEM_PROMPT,
        state_schema=AgentState,
        name="bank_vat_agent",
    )
