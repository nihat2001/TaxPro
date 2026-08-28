import io

import pytest
from openpyxl import Workbook, load_workbook

from app.services.deterministic.bank_statement_categorizer import (
    _extract_sender_name,
    _match_category,
    _split_signed_amount,
    categorize_statement,
    categorize_statement_verbose,
    finalize_statement,
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


def _row(detail, debit, no="1", sender="", kredit=""):
    return [
        "01.01.2026",
        "01.01.2026",
        sender,
        detail,
        no,
        kredit,
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
        ("PAYMENT ORDER COMMISSION", "KOMİSİYA"),
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
    assert _match_category("", detail, None, None) == expected


def test_match_category_case_insensitive():
    assert _match_category("", "nakıt çekme komisyonu", None, None) == "Nağd çəkim komissiyası"
    assert _match_category("", "Kom. 123", None, None) == "KOMİSİYA"
    assert _match_category("", "FAİZ+CƏRİMƏ", None, None) == "BANK FAİZ XƏRCİ"


def test_priority_nakit_cekme_not_komisiya():
    """Öncəlik: NAKIT CEKME KOMISYONU KOMİSİYA-ya düşməməli."""
    assert _match_category("", "NAKIT CEKME KOMISYONU", None, None) == "Nağd çəkim komissiyası"
    assert _match_category("", "NAKIT CEKME KOMISYONU", None, None) != "KOMİSİYA"


# ===== Yeni kateqoriyalar (prioritet 4-10) =====


def test_priority_maas_odemesi_komissiya():
    """MAAS ODEMESI KOMISSIYA KOMİSİYA-ya düşür (MAAŞ-a yox)."""
    assert (
        _match_category("", "MAAS ODEMESI KOMISSIYA", None, "-50.00")
        == "KOMİSİYA"
    )
    assert (
        _match_category("", "MAAS ODEMESI KOMISSIYA", None, "-50.00")
        != "MAAŞ MƏXARİC"
    )


def test_edv_odenisi():
    assert _match_category("", "EDV ODENILIR", None, "-75.60") == "ƏDV ÖDƏNİŞİ"
    assert _match_category("", "ƏDV ödənilir", None, "-2.88") == "ƏDV ÖDƏNİŞİ"
    assert _match_category("", "EDV", None, "-10.00") == "ƏDV ÖDƏNİŞİ"


def test_edv_odenisi_requires_debit():
    # Kredit sətirində EDV olsa da (pul daxil) ƏDV ÖDƏNİŞİ deyil.
    assert _match_category("", "EDV", "100.00", None) is None


def test_rusum():
    assert _match_category("", "DÖVLƏT RÜSUMU", None, "-10.00") == "RÜSUM"
    assert _match_category("", "DOVLET RUSUMU", None, "-15.00") == "RÜSUM"
    assert _match_category("", "RÜSUM", None, "-5.00") == "RÜSUM"


def test_priority_rusum_over_vergi():
    """Eyni sətirdə həm RUSUM, həm ümumi vergi açar sözü → RÜSUM qalib."""
    assert (
        _match_category("VN Milli Gəlir.Departamenti", "DOVLET RUSUMU", None, "-10.00")
        == "RÜSUM"
    )
    assert (
        _match_category("VN Milli Gəlir.Departamenti", "DOVLET RUSUMU", None, "-10.00")
        != "VERGİ MƏXARİC"
    )


def test_vergi_meharic_from_sender_keyword():
    assert (
        _match_category("Azərbaycan Respublikası DSMF", "TAM ODEME", None, "-14.71")
        == "VERGİ MƏXARİC"
    )
    assert (
        _match_category("Mərk.Sos.Sığ.FUİ (işsiz.yığım)", "126300 TAM ODEME", None, "-5.00")
        == "VERGİ MƏXARİC"
    )
    assert (
        _match_category("VN Milli Gəlir.Departamenti", "TAM ODEME", None, "-10.00")
        == "VERGİ MƏXARİC"
    )


def test_vergi_meharic_from_detail_keyword():
    assert (
        _match_category("", "MİLLİ GƏLİR DEPARTAMENTİ", None, "-15.00")
        == "VERGİ MƏXARİC"
    )


def test_vergi_meharic_requires_debit():
    # Kredit sətiri (pul daxil), dövlət açar sözü olsa da VERGİ MƏXARİC deyil.
    assert _match_category("", "TAM ODEME", "100.00", None) is None


def test_priority_vergi_vs_komisiya_same_sender():
    """Eyni dövlət qurumundan komissiya sətri → KOMİSİYA (prioritet 2 > 4)."""
    assert _match_category("DSMF", "DK KOM", None, "-2.00") == "KOMİSİYA"


def test_maas_meharic():
    assert _match_category("", "MAAS PAKET", None, "-100.00") == "MAAŞ MƏXARİC"
    assert _match_category("", "ƏMƏKHAQQ hesablanması", None, "-250.00") == "MAAŞ MƏXARİC"
    assert _match_category("", "MAAŞ ÖDƏMƏSİ", None, "-120.00") == "MAAŞ MƏXARİC"


def test_tesisciden_borc_three_conditions():
    # 3 şərt: kredit dolu + NAGD YATIRMA/CASH DEPOSIT + tesisciden token.
    assert (
        _match_category(
            "NAGD YATIRMA/CASH DEPOSIT",
            "CSS001/23 TESİSCİDEN BORC\\DEMİROVA",
            "60000",
            None,
        )
        == "TƏSİSÇİDƏN BORC"
    )
    # BORC sözü ayrı dayana bilər — tək tesisciden token kifayətdir.
    assert (
        _match_category(
            "NAGD YATIRMA/CASH DEPOSIT",
            "CSS001/23SAYLI BORC MÜQ TƏSİSÇİDƏN",
            "4200",
            None,
        )
        == "TƏSİSÇİDƏN BORC"
    )


def test_tesisciden_borc_requires_sender_token():
    # Sender-də NAGD YATIRMA/CASH DEPOSIT yoxdursa, kredit + non-empty sender → DAXİL.
    assert (
        _match_category("X Investment MMC", "TESİSCİDEN BORC", "1000", None)
        == "DİGƏR ŞİRKƏT DAXİL OLMALAR"
    )


def test_tesisciden_borc_debit_direction():
    # Yeni qayda: NAGD YATIRMA + Debit (borc qaytarma) → TƏSİSÇİDƏN BORC.
    assert (
        _match_category("NAGD YATIRMA/CASH DEPOSIT", "TESİSCİDEN BORC", None, "-60000")
        == "TƏSİSÇİDƏN BORC"
    )


def test_tesisciden_borc_without_token_word():
    # tesisciden sözü olmadan da NAGD YATIRMA sətri TƏSİSÇİDƏN BORC-dur.
    assert (
        _match_category("NAGD YATIRMA/CASH DEPOSIT", "DƏMİR ZÖHRAB ƏKBƏR", None, "-1000")
        == "TƏSİSÇİDƏN BORC"
    )


def test_diger_sirket_daxil_olmalar():
    assert (
        _match_category("TRANSTONE MMC", "Muqavile esas en", "5132", None)
        == "DİGƏR ŞİRKƏT DAXİL OLMALAR"
    )
    assert (
        _match_category("Atesgah sigorta sirketi asc", "SIGORTA HAQQI", "50", None)
        == "DİGƏR ŞİRKƏT DAXİL OLMALAR"
    )


def test_diger_sirket_odenisleri():
    assert (
        _match_category("Avtonax MR MMC", "EBDG 503767 esas en", None, "-4216.59")
        == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    )


def test_catchall_borusan_no_suffix():
    """Suffiks yoxdur → yenə də catch-all: sender boş deyil + debit → ÖDƏNİŞLƏRİ."""
    assert (
        _match_category("BORUSAN MAKINA VE GUC", "EBDO sayli EQF esasen odenilir", None, "-450.86")
        == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    )


def test_catchall_universal_spare_parts():
    assert (
        _match_category("UNIVERSAL SPARE PARTS", "26 1023UC MUQAVILE", None, "-335.00")
        == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    )


def test_catchall_empty_sender_stays_uncategorized():
    """Sender boşdursa catch-all işləməməlidir (ATM/PIN regression)."""
    assert _match_category("", "ATM XIRDALAN BIZIM M", None, "-1000.00") is None
    assert _match_category("", "PIN YENI UCRET", None, "-2.00") is None
    assert _match_category("", "ATM 28 MAY", None, "-300.00") is None


def test_priority_vergi_over_company_payment():
    """Dövlət qurumu debit + tax açar sözü → VERGİ (şirkət formatı olsa da)."""
    assert (
        _match_category("VN Milli Gəlir.Departamenti", "TAM ODEME", None, "-15.00")
        == "VERGİ MƏXARİC"
    )


def test_summary_has_ten_categories():
    out, summary = categorize_statement(
        _xlsx_bytes(META, STD_HEADER, [_row("Com. bank komissiyası", "-5.50", "1")])
    )
    assert set(summary.keys()) == {
        "Nağd çəkim komissiyası",
        "KOMİSİYA",
        "BANK FAİZ XƏRCİ",
        "ƏDV ÖDƏNİŞİ",
        "RÜSUM",
        "VERGİ MƏXARİC",
        "MAAŞ MƏXARİC",
        "TƏSİSÇİDƏN BORC",
        "DİGƏR ŞİRKƏT DAXİL OLMALAR",
        "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ",
    }


def test_recategorize_own_output():
    """Öz çıxışımızı (Xülasə aktiv sheet ilə) yenidən kateqoriyalaşdırmaq."""
    rows = [
        _row("Com. bank komissiyası", "-5.50", "1"),
        _row("MAAŞ PAKET", "-100.00", "2"),
        _row("EDV ODENILIR", "-75.60", "3"),
    ]
    out1, summary1 = categorize_statement(_xlsx_bytes(META, STD_HEADER, rows))
    out2, summary2 = categorize_statement(out1)  # aktiv sheet Xülasə-dir
    assert summary2 == summary1
    assert summary2["KOMİSİYA"] == {"cem": 5.5, "say": 1}
    assert summary2["MAAŞ MƏXARİC"] == {"cem": 100.0, "say": 1}
    assert summary2["ƏDV ÖDƏNİŞİ"] == {"cem": 75.6, "say": 1}


def test_categorize_integration_sender_sutunu_oxunur():
    """End-to-end: sender sütunu (Göndərən/Alıcı) düzgün oxunmalıdır."""
    rows = [
        ["01.01.2026", "01.01.2026", "NAGD YATIRMA/CASH DEPOSIT", "TESİSCİDEN BORC", "1", "60000", "", "60000", ""],
        ["02.01.2026", "02.01.2026", "TRANSTONE MMC", "Muqavile esas", "2", "5132", "", "5132", ""],
        ["03.01.2026", "03.01.2026", "Avtonax MR MMC", "EBDG esas", "3", "", "-4216.59", "4216.59", ""],
        ["04.01.2026", "04.01.2026", "", "ALIS VERIS", "4", "", "-50.00", "50.00", ""],
    ]
    out, summary = categorize_statement(_xlsx_bytes(META, STD_HEADER, rows))
    assert summary["TƏSİSÇİDƏN BORC"] == {"cem": 60000.0, "say": 1}
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 5132.0, "say": 1}
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 4216.59, "say": 1}


def test_xulase_sheet_shows_only_present_categories():
    rows = [
        _row("Com. bank komissiyası", "-5.50", "1"),
        _row("MAAŞ PAKET", "-100.00", "2"),
        _row("ALIS VERIS", "-50.00", "3"),
    ]
    out, _ = categorize_statement(_xlsx_bytes(META, STD_HEADER, rows))
    wb = _reload(out)
    xs = wb["Xülasə"]
    data = list(xs.iter_rows(values_only=True))
    cats = [row[0] for row in data]
    assert "KOMİSİYA" in cats
    assert "MAAŞ MƏXARİC" in cats
    assert "Nağd çəkim komissiyası" not in cats  # sıfır → göstərilmir
    assert "VERGİ MƏXARİC" not in cats
    assert cats[-1] == "CƏMİ"


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

    # Sətir-üzrə İcmal sütunu (Kateqoriya-nın dərhal sağı, col10)
    wb = _reload(out)
    ws = wb["Sheet"]
    assert ws.cell(row=5, column=10).value == "İcmal"
    assert ws.cell(row=6, column=10).value == 1234.56   # -1234.56 → abs
    assert ws.cell(row=7, column=10).value == 1234.56   # "1 234,56" → abs
    assert ws.cell(row=8, column=10).value is None      # debit boş → boş
    assert ws.cell(row=9, column=10).value is None      # kateqoriyasız → boş


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
    ws = wb["Sheet"]
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

    # İcmal sütunu: Kateqoriya-nın dərhal sağı (col10), abs(debit)
    assert ws.cell(row=5, column=10).value == "İcmal"
    assert ws.cell(row=6, column=10).value == 5.5
    assert ws.cell(row=7, column=10).value == 10.0
    assert ws.cell(row=8, column=10).value == 2.0
    assert ws.cell(row=9, column=10).value == 3.0
    assert ws.cell(row=10, column=10).value == 7.25
    assert ws.cell(row=11, column=10).value == 8.75
    assert ws.cell(row=12, column=10).value is None

    # Xülasə sheet-i
    assert "Xülasə" in wb.sheetnames
    xs = wb["Xülasə"]
    rows_in_sheet = list(xs.iter_rows(values_only=True))
    assert rows_in_sheet[0] == ("Kateqoriya", "Əməliyyat sayı", "Cəm (abs)")
    assert ("KOMİSİYA", 2, 15.5) in rows_in_sheet
    assert ("Nağd çəkim komissiyası", 2, 5.0) in rows_in_sheet
    assert ("BANK FAİZ XƏRCİ", 2, 16.0) in rows_in_sheet
    assert rows_in_sheet[-1][0] == "CƏMİ"
    assert rows_in_sheet[-1][1] == 6
    assert rows_in_sheet[-1][2] == 36.5


def test_categorize_adds_kateqoriya_column_if_missing():
    header_without_kat = STD_HEADER[:-1]
    rows = [
        ["01.01.2026", "01.01.2026", "", "Com. komissiya", "1", "", "-5.50", "994.50"],
        ["02.01.2026", "02.01.2026", "", "NAKIT CEKME KOMISYONU", "2", "", "-2.00", "992.50"],
        ["03.01.2026", "03.01.2026", "", "ALIS VERIS", "3", "", "-50.00", "942.50"],
    ]
    out, summary = categorize_statement(_xlsx_bytes(META, header_without_kat, rows))
    wb = _reload(out)
    ws = wb["Sheet"]
    assert ws.cell(row=5, column=9).value == "Kateqoriya"
    assert ws.cell(row=5, column=10).value == "İcmal"
    assert ws.cell(row=6, column=9).value == "KOMİSİYA"
    assert ws.cell(row=6, column=10).value == 5.5
    assert ws.cell(row=7, column=9).value == "Nağd çəkim komissiyası"
    assert ws.cell(row=7, column=10).value == 2.0
    assert ws.cell(row=8, column=9).value is None
    assert ws.cell(row=8, column=10).value is None
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
                        "",
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
    assert wb.active.title == "Xülasə"
    ws = wb.active
    assert ws.title == "Xülasə"
    main = wb["Çıxarış"]
    assert main.cell(row=1, column=1).value == "Qurum Adi"
    assert main.cell(row=3, column=4).value == "Əməliyyat Detalları"
    assert main.cell(row=3, column=10).value == "İcmal"
    assert main.cell(row=4, column=9).value == "BANK FAİZ XƏRCİ"
    assert main.cell(row=4, column=10).value == 7.25
    assert main.cell(row=5, column=9).value == "BANK FAİZ XƏRCİ"
    assert main.cell(row=5, column=10).value == 8.75
    assert main.cell(row=6, column=9).value is None
    assert main.cell(row=6, column=10).value is None


def test_categorize_header_not_found_raises():
    bad = _xlsx_bytes(META, ["a", "b"], [[1, 2]])
    with pytest.raises(ValueError, match="Başlıq sətri tapılmadı"):
        categorize_statement(bad)


def test_categorize_xls_not_supported():
    with pytest.raises(ValueError, match="dəstəklənmir"):
        categorize_statement(b"\xd0\xcf\x11\xe0" + b"ole2fake")


def test_output_formatting():
    rows = [
        _row("Com. bank komissiyası", "-5.50", "1"),
        _row("ALIS VERIS", "-50.00", "2"),
    ]
    out, _ = categorize_statement(_xlsx_bytes(META, STD_HEADER, rows))
    wb = _reload(out)

    # Default aktiv sheet Xülasə-dir
    assert wb.active.title == "Xülasə"

    main = wb["Sheet"]
    # Header sətri bold + açıq boz fon
    hdr = main.cell(row=5, column=1)
    assert hdr.font.bold is True
    assert str(hdr.fill.start_color.rgb).endswith("E8E8E8")
    # İcmal başlıq + number format + sağa düzləndirmə
    assert main.cell(row=5, column=10).value == "İcmal"
    amount_cell = main.cell(row=6, column=10)
    assert amount_cell.number_format == "#,##0.00"
    assert amount_cell.alignment.horizontal == "right"

    xs = wb["Xülasə"]
    # Xülasə başlıq bold + fon
    assert xs.cell(row=1, column=1).font.bold is True
    assert str(xs.cell(row=1, column=1).fill.start_color.rgb).endswith("E8E8E8")
    # Cəm sütunu format
    assert xs.cell(row=2, column=3).number_format == "#,##0.00"
    # CƏMİ sətri bold + yuxarı border
    total_row = xs.max_row
    assert xs.cell(row=total_row, column=1).font.bold is True
    assert xs.cell(row=total_row, column=1).border.top.style is not None
    assert xs.cell(row=total_row, column=1).value == "CƏMİ"


# ===== Human-in-the-loop: verbose + finalize =====


def test_categorize_verbose_unclassified():
    rows = [
        _row("Com. bank komissiyası", "-5.50", "1"),
        _row("ATM XIRDALAN BIZIM M", "-1000.00", "2"),
        _row("PIN YENI UCRET", "-2.00", "3"),
        _row("MAAŞ PAKET", "-100.00", "4"),
    ]
    out, summary, unclassified = categorize_statement_verbose(
        _xlsx_bytes(META, STD_HEADER, rows)
    )
    assert summary["KOMİSİYA"]["say"] == 1
    assert summary["MAAŞ MƏXARİC"]["say"] == 1
    idxs = {r["row_index"] for r in unclassified}
    assert idxs == {7, 8}  # meta 1-4, header 5, data 6-9
    atm = next(r for r in unclassified if r["row_index"] == 7)
    assert atm["detail"] == "ATM XIRDALAN BIZIM M"
    assert atm["debit"] == "-1000.00"
    assert atm["kredit"] is None
    assert atm["sender"] == ""


def test_all_recognized_unclassified_empty():
    rows = [
        _row("Com. bank komissiyası", "-5.50", "1"),
        _row("MAAŞ PAKET", "-100.00", "2"),
    ]
    _, _, unclassified = categorize_statement_verbose(_xlsx_bytes(META, STD_HEADER, rows))
    assert unclassified == []


def test_finalize_applies_overrides():
    rows = [
        _row("ATM XIRDALAN BIZIM M", "-1000.00", "1"),
        _row("PIN YENI UCRET", "-2.00", "2"),
        _row("Com. bank komissiyası", "-5.50", "3"),
    ]
    raw = _xlsx_bytes(META, STD_HEADER, rows)
    _, _, unclassified = categorize_statement_verbose(raw)
    row_atm = next(r for r in unclassified if "ATM" in r["detail"])["row_index"]
    row_pin = next(r for r in unclassified if "PIN" in r["detail"])["row_index"]

    out, summary = finalize_statement(
        raw,
        {
            row_atm: "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ",
            row_pin: "",  # "kateqoriyasız qalsın"
        },
    )
    wb = _reload(out)
    main = wb["Sheet"]
    assert main.cell(row=row_atm, column=9).value == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    assert main.cell(row=row_atm, column=10).value == 1000.0
    assert main.cell(row=row_pin, column=9).value is None
    assert main.cell(row=row_pin, column=10).value is None
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 1000.0, "say": 1}
    assert summary["KOMİSİYA"] == {"cem": 5.5, "say": 1}
    # Xülasə sheet-i yenidən hesablanmış
    xs = wb["Xülasə"]
    cats = [r[0] for r in xs.iter_rows(values_only=True)]
    assert "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ" in cats
    assert cats[-1] == "CƏMİ"


def test_finalize_invalid_row_raises():
    raw = _xlsx_bytes(META, STD_HEADER, [_row("ATM XIRDALAN", "-100.00", "1")])
    with pytest.raises(ValueError, match="tanınmamış sətir"):
        finalize_statement(raw, {999: "KOMİSİYA"})


def test_finalize_invalid_category_raises():
    raw = _xlsx_bytes(META, STD_HEADER, [_row("ATM XIRDALAN", "-100.00", "1")])
    _, _, unclassified = categorize_statement_verbose(raw)
    idx = unclassified[0]["row_index"]
    with pytest.raises(ValueError, match="Keçərsiz kateqoriya"):
        finalize_statement(raw, {idx: "BOGUS"})


def test_finalize_empty_selections():
    raw = _xlsx_bytes(META, STD_HEADER, [_row("Com. bank komissiyası", "-5.50", "1")])
    out, summary = finalize_statement(raw, {})
    assert summary["KOMİSİYA"] == {"cem": 5.5, "say": 1}
    wb = _reload(out)
    assert wb.active.title == "Xülasə"


# ===== Kapital Bank format dəstəyi =====

KAPITAL_HEADER = [
    "№",
    "Tarix",
    "Sənədin nömrəsi",
    "Əks hesabın nömrəsi, adı və vöen-i",
    "Məbləğ",
    "Valyuta",
    "AZN dəyəri",
    "Balans",
    "Təyinat",
]


def _kapital_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(KAPITAL_HEADER)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def test_kapital_extract_sender_name():
    assert (
        _extract_sender_name("AZ10AIIB00000100100001944221 / KAPITAL / 9900003611")
        == "KAPITAL"
    )
    assert (
        _extract_sender_name("AZ95AIIB410450D9441426 / FAHRADOV TALA T")
        == "FAHRADOV TALA T"
    )
    assert _extract_sender_name("AZ95AIIB410450D9441426") == "AZ95AIIB410450D9441426"
    assert _extract_sender_name("") == ""
    assert _extract_sender_name(None) == ""


def test_kapital_split_signed_amount():
    assert _split_signed_amount("-1750") == ("1750", None)
    assert _split_signed_amount("2080") == (None, "2080")
    assert _split_signed_amount("-17.5") == ("17.5", None)
    assert _split_signed_amount("0") == (None, None)
    assert _split_signed_amount("") == (None, None)
    assert _split_signed_amount(None) == (None, None)


def test_match_nagd_cekim_flexible():
    assert (
        _match_category(
            "", "hesabindan naglashdirma ucun 1 faiz vergi meblegi", None, "-17.50"
        )
        == "Nağd çəkim komissiyası"
    )


def test_match_dovlat_vergi_xidmati():
    assert (
        _match_category("", "Dovlat Vergi Xidmati", None, "-90.40")
        == "VERGİ MƏXARİC"
    )


def test_kapital_end_to_end():
    rows = [
        [1, "2022-12-15", "4351627793", "AZ10AIIB00000100100001944221 / KAPITAL / 9900003611", 2080, "AZN", 2080, 2080, "mal uchun muqavile"],
        [2, "2022-12-16", "4359350403", "AZ95AIIB410450D9441426 / FAHRADOV TALA T", -1750, "AZN", -1750, 330, "MAL VƏ XİDMƏT ÜZRƏ ÖDƏ"],
        [3, "2022-12-16", "4359350411", "AZ64AIIB00000451090004 / Dovlat Vergi Xidmati / 9900001111", -17.5, "AZN", -17.5, 312.5, "hesabindan naglashdirma ucun 1 faiz vergi meblegi"],
        [4, "2022-12-16", "4359350433", "AZ11AIIB00000670100001 / Dovlat Vergi Xidmati / 9900001111", -8.75, "AZN", -8.75, 303.75, "Dovlat Vergi Xidmati"],
    ]
    out, summary, _ = categorize_statement_verbose(_kapital_xlsx(rows))
    assert summary["Nağd çəkim komissiyası"] == {"cem": 17.5, "say": 1}
    assert summary["VERGİ MƏXARİC"] == {"cem": 8.75, "say": 1}
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 2080.0, "say": 1}
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 1750.0, "say": 1}

    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    assert main.cell(row=2, column=10).value == "DİGƏR ŞİRKƏT DAXİL OLMALAR"
    assert main.cell(row=3, column=10).value == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    assert main.cell(row=4, column=10).value == "Nağd çəkim komissiyası"
    assert main.cell(row=5, column=10).value == "VERGİ MƏXARİC"
    assert main.cell(row=2, column=11).value == 2080.0  # İcmal (kredit)
    assert main.cell(row=3, column=11).value == 1750.0  # İcmal (debit)


def test_kapital_csv():
    import csv as csv_mod

    buf = io.StringIO()
    writer = csv_mod.writer(buf)
    writer.writerow(KAPITAL_HEADER)
    writer.writerow([1, "2022-12-15", "x", "AZ / KAPITAL / 1", 100, "AZN", 100, 100, "mal uchun"])
    writer.writerow([2, "2022-12-16", "y", "AZ / FAHRADOV / 2", -50, "AZN", -50, 50, "odenis"])
    raw = buf.getvalue().encode("utf-8")
    _, summary, _ = categorize_statement_verbose(raw)
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 100.0, "say": 1}
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 50.0, "say": 1}


def test_formats_independent_sequential():
    """YapıKredi (standard) + Kapital eyni sessiyada ardıcıl — təsirsiz."""
    std = _xlsx_bytes(
        META,
        STD_HEADER,
        [_row("Com. bank komissiyası", "-5.50", "1"), _row("ALIS VERIS", "-100.00", "2")],
    )
    _, s1, _ = categorize_statement_verbose(std)
    assert s1["KOMİSİYA"] == {"cem": 5.5, "say": 1}

    kap = _kapital_xlsx([[1, "2022-12-15", "x", "AZ / KAPITAL / 1", 100, "AZN", 100, 100, "mal"]])
    _, s2, _ = categorize_statement_verbose(kap)
    assert s2["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 100.0, "say": 1}

    _, s1b, _ = categorize_statement_verbose(std)
    assert s1b["KOMİSİYA"] == {"cem": 5.5, "say": 1}
    assert s1b["DİGƏR ŞİRKƏT DAXİL OLMALAR"]["say"] == 0