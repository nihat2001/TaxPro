import base64

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.schemas.ai import AgentRunResponse
from app.schemas.profile import ProfileCreate
from app.services.ai.runner import new_thread_id, run_import, to_run_response
from app.services.deterministic.excel_parser import parse_standard
from app.services.deterministic.profile_service import create_profile

router = APIRouter(prefix="/import", tags=["import"])


@router.post("/excel", response_model=AgentRunResponse)
async def import_standard_excel(
    file: UploadFile = File(...), db: AsyncSession = Depends(get_db)
):
    raw = await file.read()
    try:
        rows = parse_standard(raw)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    inserted = 0
    errors: list[str] = []
    for idx, row in enumerate(rows, start=2):
        try:
            data = ProfileCreate(
                ad_soyad=row["ad_soyad"],
                voen=row["voen"],
                nomre=row.get("nomre", ""),
                asan_id=row.get("asan_id", ""),
                bitme_tarixi=row.get("bitme_tarixi"),
                qeyd=row.get("qeyd"),
                pin1=row.get("pin1", ""),
                pin2=row.get("pin2", ""),
                puk=row.get("puk", ""),
                reng=row.get("reng", "#3b82f6"),
                ikon=row.get("ikon", "user"),
            )
        except Exception as exc:
            errors.append(f"Sətir {idx}: {exc}")
            continue
        await create_profile(db, data)
        inserted += 1

    thread_id = new_thread_id()
    return AgentRunResponse(
        thread_id=thread_id,
        status="done",
        message=f"{inserted} profil idxal edildi; {len(errors)} sətir atlandı.",
        inserted=inserted,
        sources=errors,
    )


@router.post("/smart", response_model=AgentRunResponse)
async def import_smart_excel(
    file: UploadFile = File(...),
):
    raw = await file.read()
    thread_id = new_thread_id()
    file_b64 = base64.b64encode(raw).decode()
    state = await run_import(
        thread_id,
        file_b64,
        source_name=file.filename or "upload.xlsx",
    )
    return to_run_response(thread_id, state)
