from langchain_anthropic import ChatAnthropic
from langchain_groq import ChatGroq

from app.core.config import get_settings


def get_chat_model():
    settings = get_settings()
    if settings.llm_provider == "anthropic":
        return ChatAnthropic(
            model=settings.anthropic_model,
            api_key=settings.anthropic_api_key or None,
            temperature=0,
        )
    return ChatGroq(
        model=settings.groq_model,
        api_key=settings.groq_api_key or None,
        temperature=0,
    )
