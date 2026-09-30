import base64
import io
import json

import pandas as pd
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.schemas.imports import (
    BankStatementCategorizeResult,
    BankStatementSummary,
    BankStatementUnclassifiedRow,
    ImportExcelResponse,
)
from app.schemas.profile import ProfileCreate
from app.services.deterministic.bank_statement_categorizer import (
    categorize_statement_verbose,
    finalize_statement,
)
from app.services.deterministic.creditors import (
    MAX_BANK_FILES,
    generate_creditors_workbook,
    generate_debitors_workbook,
)
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
                nomre=row["nomre"],
                asan_id=row["asan_id"],
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


@router.post("/bank-statement/categorize", response_model=BankStatementCategorizeResult)
async def categorize_bank_statement(file: UploadFile = File(...)):
    raw = await file.read()
    try:
        out_bytes, summary, unclassified = categorize_statement_verbose(raw)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    validated = BankStatementSummary.model_validate(summary)
    return BankStatementCategorizeResult(
        summary=validated,
        unclassified=[BankStatementUnclassifiedRow(**row) for row in unclassified],
        file_base64=base64.b64encode(out_bytes).decode("ascii"),
    )


@router.post("/bank-statement/finalize")
async def finalize_bank_statement(
    file: UploadFile = File(...),
    selections: str = Form(...),
):
    raw = await file.read()
    try:
        parsed = json.loads(selections or "{}")
        if not isinstance(parsed, dict):
            raise ValueError("Selections JSON obyekti olmalıdır.")
        selections_map = {int(k): str(v) for k, v in parsed.items()}
    except (ValueError, TypeError, json.JSONDecodeError) as exc:
        raise HTTPException(status_code=422, detail=f"Selections keçərsizdir: {exc}")

    try:
        out_bytes, new_summary = finalize_statement(raw, selections_map)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    validated = BankStatementSummary.model_validate(new_summary)
    encoded = base64.b64encode(
        json.dumps(validated.model_dump(by_alias=True), ensure_ascii=False).encode(
            "utf-8"
        )
    ).decode("ascii")

    return StreamingResponse(
        io.BytesIO(out_bytes),
        media_type=_XLSX_MIME,
        headers={
            "Content-Disposition": (
                'attachment; filename="taxpro_bank_cizarsisi_yekun.xlsx"'
            ),
            "X-Summary": encoded,
        },
    )


async def _read_banks(bank: list[UploadFile]) -> list[tuple[str, bytes]]:
    """Çoxlu bank faylını oxuyur; say limitini yoxlayır."""
    if not bank:
        raise HTTPException(
            status_code=422, detail="Bank çıxarışı faylı tələb olunur."
        )
    if len(bank) > MAX_BANK_FILES:
        raise HTTPException(
            status_code=422,
            detail=(
                f"Maksimum {MAX_BANK_FILES} bank çıxarışı faylı göndərilə bilər "
                f"(göndərilən: {len(bank)})."
            ),
        )
    result: list[tuple[str, bytes]] = []
    for f in bank:
        result.append((f.filename or "bank", await f.read()))
    return result


@router.post("/creditors/generate")
async def generate_creditors(
    gelen: UploadFile = File(...),
    depozit: UploadFile = File(...),
    bank: list[UploadFile] = File(...),
):
    gelen_raw = await gelen.read()
    depozit_raw = await depozit.read()
    banks = await _read_banks(bank)
    try:
        out = generate_creditors_workbook(gelen_raw, depozit_raw, banks)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    return StreamingResponse(
        io.BytesIO(out),
        media_type=_XLSX_MIME,
        headers={
            "Content-Disposition": 'attachment; filename="taxpro_kreditorlar.xlsx"'
        },
    )


@router.post("/debitors/generate")
async def generate_debitors(
    gedeni: UploadFile = File(...),
    depozit: UploadFile = File(...),
    bank: list[UploadFile] = File(...),
):
    geden_raw = await gedeni.read()
    depozit_raw = await depozit.read()
    banks = await _read_banks(bank)
    try:
        out = generate_debitors_workbook(geden_raw, depozit_raw, banks)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    return StreamingResponse(
        io.BytesIO(out),
        media_type=_XLSX_MIME,
        headers={
            "Content-Disposition": 'attachment; filename="taxpro_debitorlar.xlsx"'
        },
    )


@router.get("/excel-template")
async def excel_template():
    df = pd.DataFrame(
        [
            {
                "Ad Soyad": "Ehmed Eliyev",
                "Nömrə": "501234567",
                "Asan ID": "AA123",
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
