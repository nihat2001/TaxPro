from langgraph_supervisor import create_supervisor

from app.services.ai.llm import get_chat_model
from app.services.ai.state import AgentState

_SUPERVISOR_PROMPT = (
    "Sən TaxPro çox-aqentli sisteminin supervizorsan. Komandadakı aqentlər:\n"
    "- unstructured_parser_agent: qarışıq və ya sütunları pozulmuş Excel faylının idxalı üçün.\n"
    "- bank_vat_agent: bank çıxarışının (Excel/PDF) analizi üçün.\n"
    "- profile_data_retriever_agent: bazadakı profillərin konkret məlumatları (ad/soyad, telefon/nömrə, VÖEN, Asan ID, bitmə tarixi, qeyd) haqqında suallar üçün.\n"
    "\n"
    "Yönləndirmə qaydaları:\n"
    "1. İstifadəçi xüsusi bir profilin konkret məlumatını soruşduqda (məsələn: 'Orxanın nömrəsi nədir?', 'VÖEN-i yaz', 'Asan ID nədir?') "
    "mütləq transfer_to_profile_data_retriever_agent aqentinə yönləndir və öz biliyinə əsaslanıb cavab yazma.\n"
    "2. Excel idxalı sorğusunda unstructured_parser_agent-ə, bank çıxarışı analizində isə bank_vat_agent-ə yönləndir.\n"
    "3. Aqent nəticəsini gözlə və onu yekun cavab kimi təqdim et. Aqent çağırılmadan final cavab yazma.\n"
    "4. Ümumi salamlaşma və söhbət mesajları (məsələn: 'Salam', 'Necəsiz?', 'Təşəkkürlər', 'Kimsiniz?') üçün heç bir aqenti çağırma — "
    "özbaşına qısa və mehriban cavab ver, bazaya sorğu göndərmə."
)


def build_supervisor(agents, model=None):
    return create_supervisor(
        agents,
        model=model or get_chat_model(),
        prompt=_SUPERVISOR_PROMPT,
        state_schema=AgentState,
        add_handoff_back_messages=False,
    )
