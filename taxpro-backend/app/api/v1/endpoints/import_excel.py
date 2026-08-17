import io

import pandas as pd
from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.schemas.imports import ImportExcelResponse
from app.schemas.profile import ProfileCreate
from app.services.deterministic.excel_parser import parse_standard
from app.services.deterministic.profile_service import create_profile

router = APIRouter(prefix="/import", tags=["import"])

_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _row_error(idx: int, exc: Exception) -> str:
    errors = getattr(exc, "errors", None)
    if callable(errors):
        try:
            parts = []
            for e in errors():
                loc = ".".join(str(x) for x in e.get("loc", ()))
                msg = e.get("msg", "")
                parts.append(f"{loc}: {msg}" if loc else msg)
            return f"Sətir {idx}: {'; '.join(parts)}"
        except Exception:
            pass
    return f"Sətir {idx}: {exc}"


@router.post("/excel", response_model=ImportExcelResponse)
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
            errors.append(_row_error(idx, exc))
            continue
        await create_profile(db, data)
        inserted += 1

    return ImportExcelResponse(
        message=f"{inserted} profil idxal edildi; {len(errors)} sətir atlandı.",
        inserted=inserted,
        sources=errors,
    )


@router.get("/excel-template")
async def excel_template():
    df = pd.DataFrame(
        [
            {
                "Ad Soyad": "Ehmed Eliyev",
                "VÖEN": 1234567890,
                "Nömrə": "501234567",
                "Asan ID": "AA123",
                "PIN1": "1234",
                "PIN2": "5678",
                "PUK": "9999",
                "Bitmə tarixi": "01.01.2026",
                "Qeyd": "Nümunə sətir",
            }
        ]
    )
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type=_XLSX_MIME,
        headers={
            "Content-Disposition": 'attachment; filename="taxpro_profil_sablonu.xlsx"'
        },
    )
