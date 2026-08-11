import base64
import io
from typing import Annotated

import pandas as pd
from langchain_core.tools import InjectedToolCallId, tool
from langgraph.prebuilt import InjectedState
from langgraph.types import Command, interrupt

from app.core.database import async_session_maker
from app.models.vector_store import VectorDocument
from app.services.ai.state import AgentState
from app.services.ai.tools._common import check_try, mark_sensitive, tool_message

_VAT_KEYWORDS = ["edv", "vat", "vergi", "18%"]
_EXPENSE_KEYWORDS = [
    "xerc",
    "ohtiyac",
    "odemis",
    "komissiya",
    "kredit",
    "sigorta",
    "telefon",
    "icare",
    "elektrik",
    "su",
    "internet",
]
_INCOME_KEYWORDS = ["medaxil", "gelir", "daxilolma", "transfer", "satis", "sifaris"]

_AMOUNT_ALIASES = ["mebl", "tutar", "amount", "cemi", "sum"]
_DESC_ALIASES = [
    "teyinat",
    "izahat",
    "aciqlama",
    "description",
    "info",
    "qeyd",
    "kontragent",
]
_VAT_ALIASES = ["edv", "vat", "vergi"]


def _normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).lower()
    for char, repl in {"ə": "e", "ö": "o", "ü": "u", "ı": "i", "ş": "s", "ç": "c", "ğ": "g"}.items():
        text = text.replace(char, repl)
    return text


def _first_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    for col in df.columns:
        normalized = _normalize_text(str(col))
        for alias in aliases:
            if alias in normalized:
                return col
    return None


def _to_amount(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def _categorize(description: str) -> str:
    normalized = _normalize_text(description)
    if any(k in normalized for k in _VAT_KEYWORDS):
        return "edv"
    if any(k in normalized for k in _EXPENSE_KEYWORDS):
        return "xerc"
    if any(k in normalized for k in _INCOME_KEYWORDS):
        return "medaxil"
    return "diger"


def analyze_excel(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    if df.empty:
        return {
            "total_amount": 0.0,
            "vat_amount": 0.0,
            "expense_amount": 0.0,
            "income_amount": 0.0,
            "categories": {},
        }

    amount_col = _first_column(df, _AMOUNT_ALIASES)
    desc_col = _first_column(df, _DESC_ALIASES)
    vat_col = _first_column(df, _VAT_ALIASES)

    categories: dict[str, float] = {}
    total = 0.0
    vat_total = 0.0
    expense_total = 0.0
    income_total = 0.0

    for _, row in df.iterrows():
        amount = _to_amount(row.get(amount_col)) if amount_col else 0.0
        total += amount

        if vat_col is not None:
            vat_total += _to_amount(row.get(vat_col))

        if desc_col is not None:
            category = _categorize(str(row.get(desc_col, "")))
        else:
            category = "medaxil" if amount > 0 else "diger"

        categories[category] = categories.get(category, 0.0) + amount
        if category == "xerc":
            expense_total += amount
        if category == "medaxil":
            income_total += amount

    return {
        "total_amount": round(total, 2),
        "vat_amount": round(vat_total, 2),
        "expense_amount": round(expense_total, 2),
        "income_amount": round(income_total, 2),
        "categories": {k: round(v, 2) for k, v in categories.items()},
    }


@tool
async def analyze_bank_statement(
    state: Annotated[AgentState, InjectedState],
    tool_call_id: Annotated[str, InjectedToolCallId],
) -> Command:
    """Bank çıxarışını analiz edir: ümumi məbləğ, ƏDV, xərclər və gəlirlər kateqoriyalaşdırılır."""
    error, new_try = check_try(state, tool_call_id)
    if error is not None:
        return error

    bank = state.get("bank") or {}
    file_bytes = base64.b64decode(bank.get("file_bytes", ""))
    source = str(bank.get("source", "excel")).lower()

    try:
        if source == "pdf":
            analysis = {
                "message": "PDF mətn çıxarışı üçün xüsusi emal tələb olunur; Excel/CSV formatı dəstəklənir.",
                "total_amount": 0.0,
                "vat_amount": 0.0,
                "expense_amount": 0.0,
                "income_amount": 0.0,
                "categories": {},
            }
        else:
            analysis = analyze_excel(file_bytes)
    except Exception as exc:
        return Command(
            update={
                "messages": [
                    tool_message(f"Çıxarış analiz edilə bilmədi: {exc}", tool_call_id)
                ]
            }
        )

    summary = (
        f"Ümumi: {analysis['total_amount']} AZN, "
        f"ƏDV: {analysis['vat_amount']} AZN, "
        f"Xərc: {analysis['expense_amount']} AZN, "
        f"Gəlir: {analysis['income_amount']} AZN"
    )
    message = tool_message(summary, tool_call_id)
    return Command(
        update={
            "bank": {"analysis": analysis},
            "try_count": new_try,
            "messages": [message],
        }
    )


@tool
async def update_profile_from_analysis(
    state: Annotated[AgentState, InjectedState],
    tool_call_id: Annotated[str, InjectedToolCallId],
) -> Command:
    """Bank analizinin nəticəsini profilə/yaddaşa yazır. Həssas əməliyyatdır — istifadəçi təsdiqi olmadan yazma baş vermir."""
    error, new_try = check_try(state, tool_call_id)
    if error is not None:
        return error

    bank = state.get("bank") or {}
    analysis = bank.get("analysis") or {}

    decision = interrupt(
        {
            "action": "update_profile_from_analysis",
            "message": "Bank analizinin nəticəsi profillə əlaqələndirilib yaddaşa yazılacaq. Təsdiqləyirsiniz?",
            "analysis_preview": {k: analysis.get(k) for k in ("total_amount", "vat_amount", "expense_amount", "income_amount")},
        }
    )

    if decision != "confirm":
        return Command(
            update={
                "messages": [
                    tool_message(
                        "Analiz nəticəsinin yazılması rədd edildi, heç nə yazılmadı.",
                        tool_call_id,
                    )
                ]
            }
        )

    content = (
        "Bank çıxarışı analizi: "
        f"ümumi {analysis.get('total_amount', 0)} AZN, "
        f"ƏDV {analysis.get('vat_amount', 0)} AZN, "
        f"xərc {analysis.get('expense_amount', 0)} AZN, "
        f"gəlir {analysis.get('income_amount', 0)} AZN."
    )

    async with async_session_maker() as session:
        document = VectorDocument(
            profile_id=None,
            content=content,
            embedding=None,
            source_type="bank_statement",
        )
        session.add(document)
        await session.commit()

    message = tool_message("Bank analizi yaddaşa yazıldı.", tool_call_id)
    return Command(
        update={
            "bank": {"analysis": analysis},
            "try_count": new_try,
            "messages": [message],
        }
    )


mark_sensitive(update_profile_from_analysis)
