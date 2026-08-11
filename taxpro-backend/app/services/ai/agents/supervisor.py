from langgraph_supervisor import create_supervisor

from app.services.ai.llm import get_chat_model
from app.services.ai.state import AgentState

_SUPERVISOR_PROMPT = (
    "Sən TaxPro çox-aqentli sisteminin supervizorsan. Komandadakı aqentlər:\n"
    "- unstructured_parser_agent: qarışıq və ya sütunları pozulmuş Excel faylının idxalı üçün.\n"
    "- bank_vat_agent: bank çıxarışının (Excel/PDF) analizi üçün.\n"
    "- profile_data_retriever_agent: profillərin məlumatları haqqında suallara cavab üçün.\n"
    "Gələn sorğunu məzmununa görə düzgün aqentə yönləndir. Aqent işini bitirdikdən sonra "
    "onun nəticəsini istifadəçiyə yekun cavab kimi təqdim et."
)


def build_supervisor(agents, model=None):
    return create_supervisor(
        agents,
        model=model or get_chat_model(),
        prompt=_SUPERVISOR_PROMPT,
        state_schema=AgentState,
    )
