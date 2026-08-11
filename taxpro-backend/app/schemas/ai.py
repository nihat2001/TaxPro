from typing import Any, Literal

from pydantic import BaseModel, Field


class AgentRunResponse(BaseModel):
    thread_id: str
    status: Literal["interrupt", "done"]
    message: str
    interrupt: dict[str, Any] | None = None
    analysis: dict[str, Any] | None = None
    inserted: int | None = None
    sources: list[str] = []


class ResumeRequest(BaseModel):
    thread_id: str = Field(min_length=1)
    decision: Literal["confirm", "cancel"]
