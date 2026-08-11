from uuid import UUID

from pydantic import BaseModel, Field


class ChatRequest(BaseModel):
    message: str = Field(min_length=1, max_length=4000)
    profile_id: UUID | None = None


class ChatResponse(BaseModel):
    answer: str
    sources: list[str] = []


class ErrorResponse(BaseModel):
    detail: str
    error_code: str | None = None
