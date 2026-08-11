from fastapi import APIRouter

from app.schemas.ai import AgentRunResponse
from app.schemas.chat import ChatRequest
from app.services.ai.runner import new_thread_id, run_chat, to_run_response

router = APIRouter(prefix="/chat", tags=["chat"])


@router.post("", response_model=AgentRunResponse)
async def chat(payload: ChatRequest):
    thread_id = new_thread_id()
    profile_id = str(payload.profile_id) if payload.profile_id else None
    state = await run_chat(thread_id, payload.message, profile_id)
    return to_run_response(thread_id, state)
