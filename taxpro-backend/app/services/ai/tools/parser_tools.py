import base64
import io
from typing import Annotated

import pandas as pd
from langchain_core.tools import InjectedToolCallId, tool
from langgraph.prebuilt import InjectedState
from langgraph.types import Command, interrupt
from pydantic import ValidationError

from app.core.database import async_session_maker
from app.schemas.profile import ProfileCreate
from app.services.ai.state import AgentState
from app.services.ai.tools._common import (
    check_try,
    mark_sensitive,
    tool_message,
)
from app.services.deterministic.excel_parser import _clean_row, _map_headers
from app.services.deterministic.profile_service import create_profile

SOURCE_NAME = "source_name"
FILE_BYTES = "file_bytes"


def _decode_file(state: AgentState) -> bytes:
    return base64.b64decode((state.get("parser") or {}).get(FILE_BYTES, ""))


@tool
async def parse_excel_file(
    state: Annotated[AgentState, InjectedState],
    tool_call_id: Annotated[str, InjectedToolCallId],
) -> Command:
    """Excel faylını oxuyur, sütun adlarını standart profil sahələrinə uyğunlaşdırır və işlənmiş sətirləri parser state-ə yazır."""
    error, new_try = check_try(state, tool_call_id)
    if error is not None:
        return error

    file_bytes = _decode_file(state)
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as exc:
        return Command(
            update={
                "messages": [
                    tool_message(f"Fayl oxunmadı: {exc}", tool_call_id)
                ]
            }
        )

    headers = [str(col) for col in df.columns]
    field_map = _map_headers(headers)

    rows: list[dict] = []
    skipped: list[str] = []
    for idx, raw in df.iterrows():
        row = _clean_row(dict(raw), field_map)
        if not row.get("ad_soyad") or not row.get("voen"):
            skipped.append(f"Sətir {idx + 2}: ad_soyad/voen əskikdir — atlandı.")
            continue
        rows.append(row)

    message = tool_message(
        f"{len(rows)} sətir uğurla oxundu, {len(skipped)} sətir atlandı.",
        tool_call_id,
    )
    return Command(
        update={
            "parser": {"rows": rows, "skipped": skipped, "source_name": (state.get("parser") or {}).get(SOURCE_NAME, "")},
            "try_count": new_try,
            "messages": [message],
        }
    )


@tool
async def create_profiles_from_rows(
    state: Annotated[AgentState, InjectedState],
    tool_call_id: Annotated[str, InjectedToolCallId],
) -> Command:
    """Yeni profilləri verilənlər bazasına yazır. Həssas əməliyyatdır — istifadəçi təsdiqi olmadan yazma baş vermir."""
    error, new_try = check_try(state, tool_call_id)
    if error is not None:
        return error

    rows = (state.get("parser") or {}).get("rows", [])
    if not rows:
        return Command(
            update={
                "messages": [
                    tool_message("Yazılacaq profil sətiri tapılmadı.", tool_call_id)
                ]
            }
        )

    valid: list[ProfileCreate] = []
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
        except ValidationError as exc:
            errors.append(f"Sətir {idx}: {exc}")
            continue
        valid.append(data)

    if not valid:
        return Command(
            update={
                "messages": [
                    tool_message(
                        "Heç bir sətir validasiyadan keçmədi, heç nə yazılmadı.",
                        tool_call_id,
                    )
                ]
            }
        )

    decision = interrupt(
        {
            "action": "create_profiles",
            "message": f"{len(valid)} yeni profil yaradılacaq. Təsdiqləyirsiniz?",
            "profile_count": len(valid),
        }
    )

    if decision != "confirm":
        return Command(
            update={
                "messages": [
                    tool_message(
                        "Profil yaradılması istifadəçi tərəfindən rədd edildi, heç nə yazılmadı.",
                        tool_call_id,
                    )
                ]
            }
        )

    inserted = 0
    async with async_session_maker() as session:
        for data in valid:
            await create_profile(session, data)
            inserted += 1

    message = tool_message(
        f"{inserted} profil yaradıldı; {len(errors)} sətir atlandı.",
        tool_call_id,
    )
    return Command(
        update={
            "parser": {"inserted": inserted, "skipped": errors},
            "try_count": new_try,
            "messages": [message],
        }
    )


mark_sensitive(create_profiles_from_rows)
