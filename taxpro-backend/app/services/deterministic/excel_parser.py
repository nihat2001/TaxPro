import io
import re
import unicodedata
from datetime import date
from typing import Any

import pandas as pd

from app.services.deterministic.phone import normalize_phone

REQUIRED_FIELDS = {
    "ad_soyad",
    "nomre",
    "asan_id",
}

_REQUIRED_ORDER = ["ad_soyad", "nomre", "asan_id"]

_XLSX_MAGIC = b"\x50\x4b"  # PK (zip/xlsx)
_XLS_MAGIC = b"\xd0\xcf\x11\xe0"  # OLE2 (xls)

_HEADER_ALIASES = {
    "adsoyad": "ad_soyad",
    "adsoyadsoyad": "ad_soyad",
    "soyadad": "ad_soyad",
    "tam ad": "ad_soyad",
    "tamadi": "ad_soyad",
    "istifadeci": "ad_soyad",
    "istifadeciadsoyad": "ad_soyad",
    "nomre": "nomre",
    "telefon": "nomre",
    "telefonnomresi": "nomre",
    "tel": "nomre",
    "mobil": "nomre",
    "asanid": "asan_id",
    "asandid": "asan_id",
    "asan id": "asan_id",
    "asan": "asan_id",
}

_DIACRITIC_MAP = {
    "ə": "e",
    "ö": "o",
    "ü": "u",
    "ı": "i",
    "ş": "s",
    "ç": "c",
    "ğ": "g",
}


def normalize_text(value: str) -> str:
    """Aşağı registr + diaqritiksiz + yalnız alfasayısal formata çevirir.

    Həm başlıq uyğunlaşdırmasında, həm də deterministik pattern-matching-də
    (bank çıxarışı kateqoriyalaşdırması) istifadə olunur.
    """
    lowered = value.strip().lower()
    for char, repl in _DIACRITIC_MAP.items():
        lowered = lowered.replace(char, repl)
    return re.sub(r"[^a-z0-9]", "", lowered)


def _normalize_header(value: str) -> str:
    return normalize_text(value)


def _map_header(value: str) -> str | None:
    normalized = _normalize_header(value)
    return _HEADER_ALIASES.get(normalized)


def _map_headers(headers: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for header in headers:
        field = _map_header(header)
        if field and field not in mapping:
            mapping[header] = field
    return mapping


def missing_required_fields(headers: list[str]) -> list[str]:
    mapped = set(_map_headers(headers).values())
    return [field for field in _REQUIRED_ORDER if field not in mapped]


def detect_standard(headers: list[str]) -> bool:
    return not missing_required_fields(headers)


def _normalize_phone(value: str) -> str:
    return normalize_phone(value)


def _to_scalar(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    return value


def _normalize_date(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, date):
        return value
    if isinstance(value, pd.Timestamp):
        return value.date()
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return pd.to_datetime(text, format=fmt).date()
        except (ValueError, TypeError):
            continue
    try:
        return pd.to_datetime(text, dayfirst=True).date()
    except (ValueError, TypeError):
        return None


def _clean_row(row: dict[str, Any], field_map: dict[str, str]) -> dict[str, Any]:
    cleaned: dict[str, Any] = {}
    for header, field in field_map.items():
        value = row.get(header)
        if value is None:
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        if field in {"asan_id"}:
            value = _to_scalar(value)
            value = str(value).strip()
        if field == "nomre":
            value = _normalize_phone(str(_to_scalar(value)).strip())
        cleaned[field] = value
    return cleaned


def detect_format(file_bytes: bytes) -> str:
    """Bytes-imzası ilə fayl formatını təyin edir: "xlsx" | "xls" | "csv"."""
    if file_bytes.startswith(_XLSX_MAGIC):
        return "xlsx"
    if file_bytes.startswith(_XLS_MAGIC):
        return "xls"
    return "csv"


def _read_dataframe(file_bytes: bytes) -> pd.DataFrame:
    if detect_format(file_bytes) in {"xlsx", "xls"}:
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    return pd.read_csv(io.BytesIO(file_bytes))


def parse_standard(file_bytes: bytes) -> list[dict[str, Any]]:
    try:
        data = _read_dataframe(file_bytes)
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f"Fayl oxunmadı (Excel/CSV gözlənilir): {exc}")

    if data.empty:
        return []

    headers = [str(col) for col in data.columns]
    missing = missing_required_fields(headers)
    if missing:
        raise ValueError(
            "Fayl standart şablona uyğun deyil. "
            f"Əskik sütunlar: {', '.join(missing)}"
        )

    field_map = _map_headers(headers)
    rows: list[dict[str, Any]] = []
    for _, raw in data.iterrows():
        cleaned = _clean_row(dict(raw), field_map)
        if cleaned:
            rows.append(cleaned)
    return rows
