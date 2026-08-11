from langchain_anthropic import ChatAnthropic
from langchain_openai import ChatOpenAI

from app.core.config import get_settings


def get_chat_model():
    settings = get_settings()
    if settings.llm_provider == "anthropic":
        return ChatAnthropic(
            model=settings.anthropic_model,
            api_key=settings.anthropic_api_key or None,
            temperature=0,
        )
    return ChatOpenAI(
        model=settings.llm_model,
        api_key=settings.openai_api_key or None,
        temperature=0,
    )
