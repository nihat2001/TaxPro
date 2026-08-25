import io

import pytest
from openpyxl import Workbook, load_workbook

from app.services.deterministic.bank_statement_categorizer import (
    _match_category,
    categorize_statement,
)

STD_HEADER = [
    "Əməliyyat Tarixi",
    "Həyata Keçmə Tarixi",
    "Göndərən/Alıcı Məlumatları",
    "Əməliyyat Detalları",
    "Əməliyyat No",
    "Kredit",
    "Debit",
    "Balans",
    "Kateqoriya",
]

META = [
    ["Qurum Adı", "X Bank ASC"],
    ["Hesab No", "AZ1234567"],
    ["IBAN", "AZ00BA1234567890"],
    ["Dövr", "01.01.2026 - 31.01.2026"],
]


def _xlsx_bytes(meta, header, rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    for m in meta:
        ws.append(m)
    ws.append(header)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _row(detail, debit, no="1"):
    return [
        "01.01.2026",
        "01.01.2026",
        "Alıcı/Satıcı",
        detail,
        no,
        "",
        debit,
        "1000.00",
        "",
    ]


def _reload(out: bytes) -> Workbook:
    return load_workbook(io.BytesIO(out))


# ===== Pattern qaydaları =====


@pytest.mark.parametrize(
    ("detail", "expected"),
    [
        ("Com. bank xidmət haqqı", "KOMİSİYA"),
        ("COMM- 012345", "KOMİSİYA"),
        ("DK KOM", "KOMİSİYA"),
        ("HESAB ACILIS KOM", "KOMİSİYA"),
        ("KREDIT TAHSIS KOM", "KOMİSİYA"),
        ("KOMISSIYA 0.50 AZN", "KOMİSİYA"),
        ("NAKIT CEKME KOMISYONU", "Nağd çəkim komissiyası"),
        ("NAĞD ÇƏKİM KOMİSYONU", "Nağd çəkim komissiyası"),
        ("1 FAIZ NAGD CEKIM", "Nağd çəkim komissiyası"),
        ("1 FAİZ NAĞD ÇƏKİM", "Nağd çəkim komissiyası"),
        ("Taksit Faiz 5", "BANK FAİZ XƏRCİ"),
        ("Faiz+Cərimə", "BANK FAİZ XƏRCİ"),
        ("faiz + cərimə ödənişi", "BANK FAİZ XƏRCİ"),
        ("ALIS VERIS", None),
        ("ADIDI TRANSFER", None),
        ("", None),
        (None, None),
    ],
)
def test_match_category_patterns(detail, expected):
    assert _match_category(detail) == expected


def test_match_category_case_insensitive():
    assert _match_category("nakıt çekme komisyonu") == "Nağd çəkim komissiyası"
    assert _match_category("Kom. 123") == "KOMİSİYA"
    assert _match_category("FAİZ+CƏRİMƏ") == "BANK FAİZ XƏRCİ"


def test_priority_nakit_cekme_not_komisiya():
    """Öncəlik: NAKIT CEKME KOMISYONU KOMİSİYA-ya düşməməli."""
    assert _match_category("NAKIT CEKME KOMISYONU") == "Nağd çəkim komissiyası"
    assert _match_category("NAKIT CEKME KOMISYONU") != "KOMİSİYA"


# ===== Məbləğ parslama + abs cəm =====


def test_categorize_xlsx_sums_absolute_debits():
    rows = [
        _row("Com. bank komissiyası", "-1234.56", "1"),
        _row("Com. bank komissiyası", "1 234,56", "2"),
        _row("Com. bank komissiyası", "", "3"),  # debit yoxdur — say sayılır
        _row("ALIS VERIS", "-500.00", "4"),
    ]
    out, summary = categorize_statement(_xlsx_bytes(META, STD_HEADER, rows))
    assert summary["KOMİSİYA"] == {"cem": 2469.12, "say": 3}
    assert summary["Nağd çəkim komissiyası"] == {"cem": 0.0, "say": 0}
    assert summary["BANK FAİZ XƏRCİ"] == {"cem": 0.0, "say": 0}


def test_categorize_xlsx_summary_and_meta_preserved():
    rows = [
        _row("Com. bank komissiyası", "-5.50", "1"),
        _row("DK KOM", "-10.00", "2"),
        _row("NAKIT CEKME KOMISYONU", "-2.00", "3"),
        _row("1 FAIZ NAGD CEKIM", "-3.00", "4"),
        _row("Taksit Faiz", "-7.25", "5"),
        _row("Faiz+Cərimə", "-8.75", "6"),
        _row("ALIS VERIS", "-100.00", "7"),
    ]
    out, summary = categorize_statement(_xlsx_bytes(META, STD_HEADER, rows))

    assert summary["KOMİSİYA"] == {"cem": 15.5, "say": 2}
    assert summary["Nağd çəkim komissiyası"] == {"cem": 5.0, "say": 2}
    assert summary["BANK FAİZ XƏRCİ"] == {"cem": 16.0, "say": 2}

    wb = _reload(out)
    ws = wb.active
    # Meta sətirlər dəyişməz saxlanılıb
    assert ws.cell(row=1, column=1).value == "Qurum Adı"
    assert ws.cell(row=1, column=2).value == "X Bank ASC"
    assert ws.cell(row=3, column=2).value == "AZ00BA1234567890"

    # Dinamik başlıq sətri (5-ci sətir) + Kateqoriya sütunu (9-cu)
    assert ws.cell(row=5, column=4).value == "Əməliyyat Detalları"
    assert ws.cell(row=5, column=9).value == "Kateqoriya"
    assert ws.cell(row=6, column=9).value == "KOMİSİYA"
    assert ws.cell(row=7, column=9).value == "KOMİSİYA"
    assert ws.cell(row=8, column=9).value == "Nağd çəkim komissiyası"
    assert ws.cell(row=9, column=9).value == "Nağd çəkim komissiyası"
    assert ws.cell(row=10, column=9).value == "BANK FAİZ XƏRCİ"
    assert ws.cell(row=11, column=9).value == "BANK FAİZ XƏRCİ"
    # Kateqoriya sütununda yalnız sinif adı var; uyğun gəlməyən boşdur
    assert ws.cell(row=12, column=9).value is None
    assert ws.cell(row=8, column=9).value == "Nağd çəkim komissiyası"

    # Xülasə sheet-i
    assert "Xülasə" in wb.sheetnames
    xs = wb["Xülasə"]
    rows_in_sheet = list(xs.iter_rows(values_only=True))
    assert rows_in_sheet[0] == ("Kateqoriya", "Əməliyyat sayı", "Cəm (Debit, abs)")
    assert ("KOMİSİYA", 2, 15.5) in rows_in_sheet
    assert ("Nağd çəkim komissiyası", 2, 5.0) in rows_in_sheet
    assert ("BANK FAİZ XƏRCİ", 2, 16.0) in rows_in_sheet
    assert rows_in_sheet[-1][0] == "CƏMİ"
    assert rows_in_sheet[-1][1] == 6
    assert rows_in_sheet[-1][2] == 36.5


def test_categorize_adds_kateqoriya_column_if_missing():
    header_without_kat = STD_HEADER[:-1]
    rows = [
        ["01.01.2026", "01.01.2026", "X", "Com. komissiya", "1", "", "-5.50", "994.50"],
        ["02.01.2026", "02.01.2026", "X", "NAKIT CEKME KOMISYONU", "2", "", "-2.00", "992.50"],
        ["03.01.2026", "03.01.2026", "X", "ALIS VERIS", "3", "", "-50.00", "942.50"],
    ]
    out, summary = categorize_statement(_xlsx_bytes(META, header_without_kat, rows))
    wb = _reload(out)
    ws = wb.active
    assert ws.cell(row=5, column=9).value == "Kateqoriya"
    assert ws.cell(row=6, column=9).value == "KOMİSİYA"
    assert ws.cell(row=7, column=9).value == "Nağd çəkim komissiyası"
    assert ws.cell(row=8, column=9).value is None
    assert summary["KOMİSİYA"] == {"cem": 5.5, "say": 1}
    assert summary["Nağd çəkim komissiyası"] == {"cem": 2.0, "say": 1}


def test_categorize_csv_to_xlsx():
    rows_csv = "\n".join(
        [
            "Qurum Adi,X Bank",
            "IBAN,AZ00BA123",
            "Əməliyyat Tarixi,Həyata Keçmə Tarixi,Göndərən/Alıcı Məlumatları,Əməliyyat Detalları,Əməliyyat No,Kredit,Debit,Balans,Kateqoriya",
            *[
                ",".join(
                    [
                        "01.01.2026",
                        "01.01.2026",
                        "X",
                        d,
                        no,
                        "",
                        debit,
                        "1000.00",
                        "",
                    ]
                )
                for (d, no, debit) in [
                    ("Taksit Faiz", "1", "-7.25"),
                    ("Faiz+Cərimə", "2", "-8.75"),
                    ("ADIDI TRANSFER", "3", "-50.00"),
                ]
            ],
        ]
    )
    out, summary = categorize_statement(rows_csv.encode("utf-8"))
    assert summary["BANK FAİZ XƏRCİ"] == {"cem": 16.0, "say": 2}
    assert summary["KOMİSİYA"] == {"cem": 0.0, "say": 0}

    wb = _reload(out)
    assert wb.sheetnames == ["Çıxarış", "Xülasə"]
    ws = wb.active
    assert ws.cell(row=1, column=1).value == "Qurum Adi"
    assert ws.cell(row=3, column=4).value == "Əməliyyat Detalları"
    assert ws.cell(row=4, column=9).value == "BANK FAİZ XƏRCİ"
    assert ws.cell(row=5, column=9).value == "BANK FAİZ XƏRCİ"
    assert ws.cell(row=6, column=9).value is None


def test_categorize_header_not_found_raises():
    bad = _xlsx_bytes(META, ["a", "b"], [[1, 2]])
    with pytest.raises(ValueError, match="Başlıq sətri tapılmadı"):
        categorize_statement(bad)


def test_categorize_xls_not_supported():
    with pytest.raises(ValueError, match="dəstəklənmir"):
        categorize_statement(b"\xd0\xcf\x11\xe0" + b"ole2fake")