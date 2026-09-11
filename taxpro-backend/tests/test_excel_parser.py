import io

import pandas as pd
import pytest

from app.services.deterministic.excel_parser import (
    detect_standard,
    missing_required_fields,
    parse_standard,
)
from app.services.deterministic.phone import normalize_phone


def _bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf.read()


def _csv_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_csv(buf, index=False)
    buf.seek(0)
    return buf.read()


def test_detect_standard_full_headers():
    headers = ["Ad Soyad", "Nömrə", "Asan ID"]
    assert detect_standard(headers) is True


def test_detect_standard_missing_headers():
    assert detect_standard(["Ad Soyad", "Nömrə"]) is False


def test_missing_required_fields_lists_columns():
    assert missing_required_fields(["Ad Soyad"]) == ["nomre", "asan_id"]


def test_normalize_phone_variants():
    assert normalize_phone("0501234567") == "501234567"
    assert normalize_phone("+994501234567") == "501234567"
    assert normalize_phone("994701234567") == "701234567"
    assert normalize_phone("703889246") == "703889246"
    assert normalize_phone("050-123-45-67") == "501234567"
    assert normalize_phone(None) == ""
    assert normalize_phone("012345678") == "012345678"


def test_detect_standard_aliases():
    headers = ["Ad Soyad", "Telefon", "AsanID"]
    assert detect_standard(headers) is True


def test_parse_standard_normalizes_values():
    df = pd.DataFrame(
        [
            {
                "Ad Soyad": "Ehmed Eliyev",
                "Nömrə": "0501234567",
                "Asan ID": "AA123",
            }
        ]
    )
    rows = parse_standard(_bytes(df))
    assert len(rows) == 1
    row = rows[0]
    assert row["ad_soyad"] == "Ehmed Eliyev"
    assert row["nomre"] == "501234567"
    assert row["asan_id"] == "AA123"


def test_parse_standard_skips_empty_rows():
    df = pd.DataFrame(
        [
            {"Ad Soyad": "A", "Nömrə": "0501111111", "Asan ID": "X"},
            {"Ad Soyad": None, "Nömrə": None, "Asan ID": None},
        ]
    )
    rows = parse_standard(_bytes(df))
    assert len(rows) == 1


def test_parse_standard_invalid_raises():
    df = pd.DataFrame([{"foo": 1, "bar": 2}])
    with pytest.raises(ValueError):
        parse_standard(_bytes(df))


def test_parse_standard_missing_columns_message():
    df = pd.DataFrame([{"Ad Soyad": "A"}])
    with pytest.raises(ValueError, match="Əskik sütunlar: nomre, asan_id"):
        parse_standard(_bytes(df))


def test_parse_standard_csv():
    df = pd.DataFrame(
        [
            {
                "Ad Soyad": "Ehmed Eliyev",
                "Nömrə": "0501234567",
                "Asan ID": "AA123",
            }
        ]
    )
    rows = parse_standard(_csv_bytes(df))
    assert len(rows) == 1
    assert rows[0]["ad_soyad"] == "Ehmed Eliyev"
    assert rows[0]["nomre"] == "501234567"
    assert rows[0]["asan_id"] == "AA123"