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
    headers = ["Ad Soyad", "VÖEN", "Nömrə", "Asan ID", "PIN1", "PIN2", "PUK"]
    assert detect_standard(headers) is True


def test_detect_standard_missing_headers():
    assert detect_standard(["Ad Soyad", "VÖEN"]) is False


def test_missing_required_fields_lists_columns():
    assert missing_required_fields(["Ad Soyad", "VÖEN"]) == [
        "nomre",
        "asan_id",
        "pin1",
        "pin2",
        "puk",
    ]


def test_normalize_phone_variants():
    assert normalize_phone("0501234567") == "501234567"
    assert normalize_phone("+994501234567") == "501234567"
    assert normalize_phone("994701234567") == "701234567"
    assert normalize_phone("703889246") == "703889246"
    assert normalize_phone("050-123-45-67") == "501234567"
    assert normalize_phone(None) == ""
    assert normalize_phone("012345678") == "012345678"


def test_detect_standard_aliases():
    headers = ["Ad Soyad", "VERGİ NÖMRƏSİ", "Telefon", "AsanID", "PIN", "PIN2", "PUK"]
    assert detect_standard(headers) is True


def test_parse_standard_normalizes_values():
    df = pd.DataFrame(
        [
            {
                "Ad Soyad": "Ehmed Eliyev",
                "VÖEN": 1234567890,
                "Nömrə": "0501234567",
                "Asan ID": "AA123",
                "PIN1": 1234,
                "PIN2": "5678",
                "PUK": "9999",
            }
        ]
    )
    rows = parse_standard(_bytes(df))
    assert len(rows) == 1
    row = rows[0]
    assert row["voen"] == "1234567890"
    assert row["nomre"] == "501234567"
    assert row["asan_id"] == "AA123"
    assert row["pin1"] == "1234"
    assert row["pin2"] == "5678"


def test_parse_standard_skips_empty_rows():
    df = pd.DataFrame(
        [
            {"Ad Soyad": "A", "VÖEN": 1, "Nömrə": "0501111111", "Asan ID": "X", "PIN1": "1", "PIN2": "2", "PUK": "3"},
            {"Ad Soyad": None, "VÖEN": None, "Nömrə": None, "Asan ID": None, "PIN1": None, "PIN2": None, "PUK": None},
        ]
    )
    rows = parse_standard(_bytes(df))
    assert len(rows) == 1


def test_parse_standard_invalid_raises():
    df = pd.DataFrame([{"foo": 1, "bar": 2}])
    with pytest.raises(ValueError):
        parse_standard(_bytes(df))


def test_parse_standard_missing_columns_message():
    df = pd.DataFrame([{"Ad Soyad": "A", "VÖEN": 1}])
    with pytest.raises(ValueError, match="Əskik sütunlar: nomre, asan_id, pin1, pin2, puk"):
        parse_standard(_bytes(df))


def test_parse_standard_csv():
    df = pd.DataFrame(
        [
            {
                "Ad Soyad": "Ehmed Eliyev",
                "VÖEN": 1234567890,
                "Nömrə": "0501234567",
                "Asan ID": "AA123",
                "PIN1": 1234,
                "PIN2": "5678",
                "PUK": "9999",
            }
        ]
    )
    rows = parse_standard(_csv_bytes(df))
    assert len(rows) == 1
    assert rows[0]["voen"] == "1234567890"
    assert rows[0]["nomre"] == "501234567"
    assert rows[0]["pin1"] == "1234"


def test_parse_standard_date_normalization():
    df = pd.DataFrame(
        [
            {
                "Ad Soyad": "A",
                "VÖEN": 1,
                "Nömrə": "0501111111",
                "Asan ID": "X",
                "PIN1": "1",
                "PIN2": "2",
                "PUK": "3",
                "Bitme Tarixi": "12.08.2026",
            }
        ]
    )
    rows = parse_standard(_bytes(df))
    assert rows[0]["bitme_tarixi"].isoformat() == "2026-08-12"
