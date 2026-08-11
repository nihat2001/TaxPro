from fastapi import APIRouter

from app.schemas.ai import AgentRunResponse, ResumeRequest
from app.services.ai.runner import resume, to_run_response

router = APIRouter(prefix="/resume", tags=["resume"])


@router.post("", response_model=AgentRunResponse)
async def resume_run(payload: ResumeRequest):
    state = await resume(payload.thread_id, payload.decision)
    return to_run_response(payload.thread_id, state)
