import io
import re
import unicodedata
from datetime import date
from typing import Any

import pandas as pd

REQUIRED_FIELDS = {
    "ad_soyad",
    "voen",
    "nomre",
    "asan_id",
    "pin1",
    "pin2",
    "puk",
}

_HEADER_ALIASES = {
    "adsoyad": "ad_soyad",
    "adsoyadsoyad": "ad_soyad",
    "soyadad": "ad_soyad",
    "tam ad": "ad_soyad",
    "tamadi": "ad_soyad",
    "istifadeci": "ad_soyad",
    "istifadeciadsoyad": "ad_soyad",
    "voen": "voen",
    "voenkodu": "voen",
    "vergi nomresi": "voen",
    "verginomresi": "voen",
    "vergino": "voen",
    "nomre": "nomre",
    "telefon": "nomre",
    "telefonnomresi": "nomre",
    "tel": "nomre",
    "mobil": "nomre",
    "asanid": "asan_id",
    "asandid": "asan_id",
    "asan id": "asan_id",
    "asan": "asan_id",
    "pin1": "pin1",
    "pin": "pin1",
    "pin2": "pin2",
    "puk": "puk",
    "bitme tarixi": "bitme_tarixi",
    "bitmetarixi": "bitme_tarixi",
    "bitirilme tarixi": "bitme_tarixi",
    "etibarliliq tarixi": "bitme_tarixi",
    "qeyd": "qeyd",
    "qeydler": "qeyd",
    "qeydiyyat": "qeyd",
    "reng": "reng",
    "ikon": "ikon",
    "icon": "ikon",
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


def _normalize_header(value: str) -> str:
    lowered = value.strip().lower()
    for char, repl in _DIACRITIC_MAP.items():
        lowered = lowered.replace(char, repl)
    return re.sub(r"[^a-z0-9]", "", lowered)


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


def detect_standard(headers: list[str]) -> bool:
    mapped = set(_map_headers(headers).values())
    return REQUIRED_FIELDS.issubset(mapped)


def _normalize_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value)
    if digits.startswith("994") and len(digits) == 12:
        return "+" + digits
    if len(digits) == 10 and digits.startswith("0"):
        return "+994" + digits[1:]
    if len(digits) == 9:
        return "+994" + digits
    return value


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
        if field in {"pin1", "pin2", "puk", "voen", "asan_id"}:
            value = _to_scalar(value)
            value = str(value).strip()
        if field == "nomre":
            value = _normalize_phone(str(_to_scalar(value)).strip())
        if field == "bitme_tarixi":
            value = _normalize_date(value)
        cleaned[field] = value
    return cleaned


def parse_standard(file_bytes: bytes) -> list[dict[str, Any]]:
    data = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    if data.empty:
        return []

    headers = [str(col) for col in data.columns]
    field_map = _map_headers(headers)
    if not detect_standard(headers):
        raise ValueError(
            "Fayl standart şablona uyğun deyil: tələb olunan sütunlar tam deyil."
        )

    rows: list[dict[str, Any]] = []
    for _, raw in data.iterrows():
        cleaned = _clean_row(dict(raw), field_map)
        if cleaned:
            rows.append(cleaned)
    return rows
