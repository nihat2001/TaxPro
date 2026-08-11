from langchain_core.language_models.chat_models import BaseChatModel
from langchain_core.messages import AIMessage
from langchain_core.outputs import ChatGeneration, ChatResult


class ScriptedModel(BaseChatModel):
    mode: str = "bank"

    @property
    def _llm_type(self) -> str:
        return "scripted"

    def bind_tools(self, tools, **kwargs):
        return self

    def _identify(self, messages):
        system = " ".join(
            (m.content or "")
            for m in messages
            if getattr(m, "type", "") == "system"
        )
        if "supervizorsan" in system:
            return "supervisor"
        if "Excel faylını emal edən" in system:
            return "parser"
        if "bank çıxarışlarını analiz" in system:
            return "bank"
        if "profil məlumatları haqqında suallara cavab" in system:
            return "retriever"
        return "unknown"

    def _decide(self, messages):
        agent = self._identify(messages)
        called = [m.name for m in messages if getattr(m, "type", "") == "tool"]

        if agent == "supervisor":
            if any(n.startswith("transfer_to_") for n in called):
                return AIMessage(content="Yekun cavab: əməliyyat tamamlandı.")
            user_text = " ".join(
                (m.content or "")
                for m in messages
                if getattr(m, "type", "") in ("user", "human") and isinstance(m.content, str)
            ).lower()
            if "bank" in user_text or "cixaris" in user_text or "çıxarış" in user_text:
                name = "transfer_to_bank_vat_agent"
            elif "excel" in user_text or "fayl" in user_text or "idxal" in user_text:
                name = "transfer_to_unstructured_parser_agent"
            else:
                name = "transfer_to_profile_data_retriever_agent"
            return AIMessage(content="", tool_calls=[{"name": name, "args": {}, "id": "call_sup"}])

        if agent == "parser":
            if "parse_excel_file" not in called:
                name = "parse_excel_file"
            elif "create_profiles_from_rows" not in called:
                name = "create_profiles_from_rows"
            else:
                return AIMessage(content="Excel faylı emal edildi.")
            return AIMessage(content="", tool_calls=[{"name": name, "args": {}, "id": f"call_{name}"}])

        if agent == "bank":
            if "analyze_bank_statement" not in called:
                name = "analyze_bank_statement"
            elif self.mode == "hitl" and "update_profile_from_analysis" not in called:
                name = "update_profile_from_analysis"
            else:
                return AIMessage(content="Bank çıxarışı analiz edildi.")
            return AIMessage(content="", tool_calls=[{"name": name, "args": {}, "id": f"call_{name}"}])

        if agent == "retriever":
            if "search_profile_data" not in called:
                return AIMessage(content="", tool_calls=[{"name": "search_profile_data", "args": {}, "id": "call_ret"}])
            return AIMessage(content="Sual cavablandırıldı.")

        return AIMessage(content="Anlaşılmadı.")

    def _generate(self, messages, stop=None, run_manager=None, **kwargs):
        return ChatResult(generations=[ChatGeneration(message=self._decide(messages))])

    async def _agenerate(self, messages, stop=None, run_manager=None, **kwargs):
        return self._generate(messages, stop=stop, run_manager=run_manager, **kwargs)
