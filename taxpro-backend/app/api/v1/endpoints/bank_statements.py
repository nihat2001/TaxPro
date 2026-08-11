import base64

from fastapi import APIRouter, File, UploadFile

from app.schemas.ai import AgentRunResponse
from app.services.ai.runner import new_thread_id, run_bank_analysis, to_run_response

router = APIRouter(prefix="/bank-statements", tags=["bank-statements"])


@router.post("/analyze", response_model=AgentRunResponse)
async def analyze_bank_statement(file: UploadFile = File(...)):
    raw = await file.read()
    thread_id = new_thread_id()
    file_b64 = base64.b64encode(raw).decode()
    state = await run_bank_analysis(thread_id, file_b64)
    return to_run_response(thread_id, state)
