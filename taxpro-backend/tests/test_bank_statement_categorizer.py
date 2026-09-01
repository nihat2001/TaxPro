import io

import pytest
from openpyxl import Workbook, load_workbook

from app.services.deterministic.bank_statement_categorizer import (
    _extract_sender_name,
    _extract_terminal_fee,
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


def test_summary_has_eleven_categories():
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
        "Terminal-Komissiya",
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


def test_kapital_xidmet_haqqi_komisiya():
    assert (
        _match_category(
            "", "530392404 saylı əməliyyat üzrə xidmət haqqı", None, "-2.00",
            profile="kapital",
        )
        == "KOMİSİYA"
    )
    assert (
        _match_category(
            "", "Nağdsız köçürmə üzrə xidmət haqqı 4.00(AZN)", None, "-4.00",
            profile="kapital",
        )
        == "KOMİSİYA"
    )


def test_kapital_emek_haqqi_xidmet_haqqi_stays_maas():
    """'Emek haqqi kartlarinin xidmet haqqi' MAAŞ MƏXARİC-də qalmalıdır."""
    assert (
        _match_category(
            "", "Emek haqqi kartlarinin xidmet haqqi. BatchId = 123", None, "-5.00",
            profile="kapital",
        )
        == "MAAŞ MƏXARİC"
    )


def test_standard_profile_ignores_xidmet_haqqi():
    """Standard profilində 'xidmət haqqı' KOMİSİYA-ya düşmür (izolyasiya)."""
    result = _match_category(
        "Azpetrol MMC", "Nağdsız köçürmə üzrə xidmət haqqı", None, "-4.00"
    )
    assert result != "KOMİSİYA"
    assert result == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"


def test_kapital_xidmet_haqqi_end_to_end():
    rows = [
        [1, "2022-12-15", "x", "AZ / KAPITAL / 1", -2.0, "AZN", -2.0, 100, "530392404 saylı əməliyyat üzrə xidmət haqqı"],
        [2, "2022-12-15", "y", "AZ / KAPITAL / 1", -4.0, "AZN", -4.0, 96, "Nağdsız köçürmə üzrə xidmət haqqı 4.00(AZN)"],
        [3, "2022-12-15", "z", "AZ / KAPITAL / 1", -5.0, "AZN", -5.0, 91, "Emek haqqi kartlarinin xidmet haqqi. BatchId = 1"],
    ]
    _, summary, _ = categorize_statement_verbose(_kapital_xlsx(rows))
    assert summary["KOMİSİYA"] == {"cem": 6.0, "say": 2}
    assert summary["MAAŞ MƏXARİC"] == {"cem": 5.0, "say": 1}


# ===== Terminal-Komissiya (kapital) =====

POS_EXAMPLE = (
    "260915324944221 Pos-Terminal üzrə əməliyyatlar. "
    "Ödəniş məbləği: 1.02 AZN, Xidmət haqqı: -.04 AZN"
)


def test_extract_terminal_fee():
    assert _extract_terminal_fee(POS_EXAMPLE) == 0.04
    assert _extract_terminal_fee("Xidmət haqqı: -0.30 AZN PGDR...") == 0.3
    assert _extract_terminal_fee("Xidmət haqqı: -1.28 AZN") == 1.28
    assert _extract_terminal_fee("Xidmət haqqı: 1,28 AZN") == 1.28
    assert _extract_terminal_fee("Xidmət haqqı: 1.1 AZN") == 1.1
    assert _extract_terminal_fee("Xidmət haqqı: -0.30 AZN") == 0.3
    assert _extract_terminal_fee("adi əməliyyat") is None
    assert _extract_terminal_fee("") is None
    assert _extract_terminal_fee(None) is None


def test_terminal_komissiya_kapital():
    assert (
        _match_category("", POS_EXAMPLE, None, "-0.98", profile="kapital")
        == "Terminal-Komissiya"
    )


def test_terminal_komissiya_standard_inactive():
    """Standard profildə Pos-Terminal ∧ Xidmət haqqı → Terminal-Komissiya deyil."""
    assert (
        _match_category("", POS_EXAMPLE, None, "-0.98") != "Terminal-Komissiya"
    )


def test_terminal_komissiya_no_pos_terminal_stays_komisiya():
    """Pos-Terminal yoxdursa adi 'xidmət haqqı' KOMİSİYA-da qalır."""
    assert (
        _match_category(
            "", "Nağdsız köçürmə üzrə xidmət haqqı 4.00(AZN)", None, "-4.00",
            profile="kapital",
        )
        == "KOMİSİYA"
    )


def test_terminal_komissiya_end_to_end_icmal_fee():
    rows = [
        [1, "2022-12-15", "x", "AZ / KAPITAL / 1", 0.98, "AZN", 0.98, 100, POS_EXAMPLE],
        [2, "2022-12-15", "y", "AZ / KAPITAL / 1", 3.94, "AZN", 3.94, 96, "260915324944221 Pos-Terminal üzrə əməliyyatlar. Ödəniş məbləği: 4 AZN, Xidmət haqqı: -0.06 AZN"],
    ]
    out, summary, _ = categorize_statement_verbose(_kapital_xlsx(rows))
    assert summary["Terminal-Komissiya"] == {"cem": 0.1, "say": 2}
    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    assert main.cell(row=2, column=10).value == "Terminal-Komissiya"
    assert main.cell(row=2, column=11).value == 0.04  # İcmal = fee, Məbləğ 0.98 deyil
    assert main.cell(row=3, column=10).value == "Terminal-Komissiya"
    assert main.cell(row=3, column=11).value == 0.06
    xs = wb["Xülasə"]
    cats = [r[0] for r in xs.iter_rows(values_only=True)]
    assert "Terminal-Komissiya" in cats


def test_terminal_komissiya_fallback_abs_mebleg():
    """Regex uyğun gəlməzsə abs(Məbləğ) istifadə olunur."""
    row = [1, "2022-12-15", "x", "AZ / KAPITAL / 1", -7.5, "AZN", -7.5, 92.5, "Pos-Terminal üzrə əməliyyat. Xidmət haqqı yoxdur"]
    _, summary, _ = categorize_statement_verbose(_kapital_xlsx([row]))
    assert summary["Terminal-Komissiya"] == {"cem": 7.5, "say": 1}


# ===== Yelo Bank format dəstəyi (çoxsətirli header, Mədaxil/Məxaric) =====

YELO_HEADER = [
    "Əməl.tarixi/\ntr.date",
    "İcra tarixi/\nreal date",
    "Təyinat/\nDescription",
    "Ödəyən/alan\npayer/reciever",
    "Qarşı hesab/\ncounter account",
    "Mədaxil/\nincome",
    "Məxaric/\noutcome",
    "Balans/\nbalance",
    "Kurs/\nrate",
    "AZN/\nekv",
]


def _yelo_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["Hesab açılma tarixi", "26.01"])
    ws.append(["IBAN", "AZ00YELO000000"])
    ws.append(YELO_HEADER)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _yelo_row(sender, detail, medaxil, mexaric):
    return ["08.01.26", "07.01.26", detail, sender, "qarsi", medaxil, mexaric, "100", "1", "0"]


def test_yelo_header_substring_detected():
    raw = _yelo_xlsx([_yelo_row("YELO BANK ASC", "FAIZ ÖDƏMƏ", 0, 5)])
    _, summary, _ = categorize_statement_verbose(raw)
    assert summary["BANK FAİZ XƏRCİ"] == {"cem": 5.0, "say": 1}


def test_yelo_patterns():
    # BANK FAİZ: faiz ∧ odeme
    assert _match_category("YELO BANK ASC", "FAIZ ÖDƏMƏ", "0", "5.00", profile="yelo") == "BANK FAİZ XƏRCİ"
    assert _match_category("YELO BANK ASC", "GECİKMİŞ FAIZ ÖDƏMƏ", "0", "4.23", profile="yelo") == "BANK FAİZ XƏRCİ"
    # KOMİSİYA: DƏBBƏ ÖDƏMƏ / ƏMƏLİYYAT MƏSRƏFİ / xidmət haqqı
    assert _match_category("YELO BANK ASC", "DƏBBƏ ÖDƏMƏ", "0", "1.68", profile="yelo") == "KOMİSİYA"
    assert _match_category("YELO BANK ASC", "ƏMƏLİYYAT MƏSRƏFİ", "0", "1.98", profile="yelo") == "KOMİSİYA"
    assert _match_category("YELO BANK ASC", "xidmət haqqı", "0", "2.00", profile="yelo") == "KOMİSİYA"
    # Nağd çəkim: Nağdlaşdırma vergisi
    assert _match_category("YELO BANK ASC", "Nağdlaşdırma vergisi", "0", "118", profile="yelo") == "Nağd çəkim komissiyası"
    # VERGİ: gəlir vergisi
    assert _match_category("", "Fiziki şəxslərin gəlir vergisi", "0", "50.00", profile="yelo") == "VERGİ MƏXARİC"


def test_yelo_asas_odeme_uncategorized():
    """Əsas ödəmə / GECİKMİŞ ƏSAS ÖDƏMƏ → kateqoriyasız (izolyasiya)."""
    assert _match_category("YELO BANK ASC", "ƏSAS ÖDƏMƏ", "0", "100.00", profile="yelo") is None
    assert _match_category("YELO BANK ASC", "GECİKMİŞ ƏSAS ÖDƏMƏ", "0", "1281.22", profile="yelo") is None
    # "ESASEN" (müqaviləyə əsasən) düşməməlidir
    assert _match_category("YELO BANK ASC", "MUQAVILEYE ESASEN ODENIS", "0", "50.00", profile="yelo") is not None


def test_yelo_patterns_standard_inactive():
    """Yelo pattern-ləri standard profildə aktiv deyil (izolyasiya)."""
    assert _match_category("", "DƏBBƏ ÖDƏMƏ", None, "-1.68") != "KOMİSİYA"
    assert _match_category("", "FAIZ ÖDƏMƏ", None, "-5.00") != "BANK FAİZ XƏRCİ"
    assert _match_category("", "Nağdlaşdırma vergisi", None, "-10.00") != "Nağd çəkim komissiyası"


def test_yelo_end_to_end_no_k_column():
    """K sütunu yoxdur — kateqoriyalaşdırma Təyinat/Ödəyən/Mədaxil/Məxaric-dən asılı."""
    rows = [
        _yelo_row("YELO BANK ASC", "FAIZ ÖDƏMƏ", 0, 5),
        _yelo_row("YELO BANK ASC", "GECİKMİŞ ƏSAS ÖDƏMƏ", 0, 100),
        _yelo_row("YELO BANK ASC", "DƏBBƏ ÖDƏMƏ", 0, 2),
        _yelo_row("YELO BANK ASC", "Əməliyyat Məsrəfi", 0, 3),
        _yelo_row("YELO BANK ASC", "Nağdlaşdırma vergisi", 0, 10),
        _yelo_row("", "Fiziki şəxslərin gəlir vergisi", 0, 50),
        _yelo_row("", "B44:Verilmiş kreditler", 5000, 0),
        _yelo_row("SOME MMC", "MAL", 100, 0),
    ]
    out, summary, unclassified = categorize_statement_verbose(_yelo_xlsx(rows))
    assert summary["BANK FAİZ XƏRCİ"] == {"cem": 5.0, "say": 1}
    assert summary["KOMİSİYA"] == {"cem": 5.0, "say": 2}
    assert summary["Nağd çəkim komissiyası"] == {"cem": 10.0, "say": 1}
    assert summary["VERGİ MƏXARİC"] == {"cem": 50.0, "say": 1}
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 100.0, "say": 1}
    # əsas ödəmə + kredit daxilolması kateqoriyasız
    idx = {u["row_index"] for u in unclassified}
    assert idx == {5, 10}  # meta 1-2, header 3, data 4..; əsas ödəmə row5, B44 row10

    wb = _reload(out)
    assert wb.sheetnames[0] != "Xülasə"
    main = wb[wb.sheetnames[0]]
    # Kateqoriya/İcmal əlavə olunub (11/12) — K sütunu yoxdur
    assert main.max_column == 12


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


# ===== Access Bank format dəstəyi (Header 2-ci sətirdə; Mədaxil/Məxaric ayrı) =====

ACCESS_HEADER = [
    "",
    "Hesablaşma tarixi\nValue Date",
    "Əməliyyat tarixi\nPosting Date",
    "Əməliyyatın №-si\nTransaction ID",
    "Mədaxil\nIn",
    "Məxaric\nOut",
    "Balans\nBalance",
    "Göndərənin/Alanın Adı\nSender / Beneficiary Name",
    "Göndərən/Alanin Tax ID\nSender/Beneficiary TAX ID",
    "Təyinat\nDescription",
]


def _access_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append([""])  # 1-ci sətir boş
    ws.append(ACCESS_HEADER)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _access_row(detail, medaxil, mexaric, sender="", tax="", no="1"):
    return [
        "",
        "01.04.2025",
        "01.04.2025",
        no,
        medaxil,
        mexaric,
        "1000",
        sender,
        tax,
        detail,
    ]


def test_access_profile_detected_header_on_second_row():
    """Access başlığı 2-ci sətirdədir; boş 1-ci sətir meta kimi qəbul edilməlidir."""
    raw = _access_xlsx(
        [_access_row("Muqavileye esasen Mart ayi xidmet haqqi odenilir", 700, "", "HASANLI-H MMC")]
    )
    _, summary, _ = categorize_statement_verbose(raw)
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 700.0, "say": 1}


def test_access_yelo_do_not_mix():
    """Access başlığı yelo kimi tanınmamalıdır (gondereninalanin ≠ odeyenalan)."""
    access_raw = _access_xlsx(
        [_access_row("FAIZ ÖDƏMƏ", 0, 5, "X MMC")]
    )
    _, s_access, _ = categorize_statement_verbose(access_raw)
    # yelo-ya xas 'FAIZ ÖDƏMƏ' → BANK FAİZ access-də aktiv olmamalıdır
    assert s_access["BANK FAİZ XƏRCİ"]["say"] == 0
    assert s_access["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 5.0, "say": 1}

    # Yelo faylı access kimi tanınmamalıdır
    yelo_raw = _yelo_xlsx([_yelo_row("YELO BANK ASC", "FAIZ ÖDƏMƏ", 0, 5)])
    _, s_yelo, _ = categorize_statement_verbose(yelo_raw)
    assert s_yelo["BANK FAİZ XƏRCİ"] == {"cem": 5.0, "say": 1}


def test_access_yelo_access_sequential_no_mix():
    """Eyni sessiyada access + yelo ardıcıl — bir-birinə qarışmamalıdır."""
    a = _access_xlsx(
        [_access_row("Muqavileye esasen Mart ayi xidmet haqqi odenilir", 700, "", "HASANLI-H MMC")]
    )
    _, s1, _ = categorize_statement_verbose(a)
    assert s1["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 700.0, "say": 1}

    y = _yelo_xlsx([_yelo_row("YELO BANK ASC", "GECİKMİŞ ƏSAS ÖDƏMƏ", 0, 100)])
    _, s2, _ = categorize_statement_verbose(y)
    # əsas ödəmə kateqoriyasız (yelo-xas); access token-u bura təsir etməməli
    assert s2["DİGƏR ŞİRKƏT DAXİL OLMALAR"]["say"] == 0
    assert s2["KOMİSİYA"]["say"] == 0

    a2 = _access_xlsx(
        [_access_row("Məxaric üçün komm.haqqı", "", 14.7)]
    )
    _, s3, _ = categorize_statement_verbose(a2)
    assert s3["KOMİSİYA"] == {"cem": 14.7, "say": 1}


def test_access_xidmet_haqqi_not_komisiya():
    """KRİTİK: 'xidmət haqqı' access-də KOMİSİYA YARATMAMALIDIR (gəlir sətri)."""
    result = _match_category(
        "HASANLI-H MMC",
        "Muqavileye esasen Mart ayi xidmet haqqi odenilir",
        "700",
        "",
        profile="access",
    )
    assert result != "KOMİSİYA"
    assert result == "DİGƏR ŞİRKƏT DAXİL OLMALAR"


def test_access_xidmet_haqqi_debit_empty_sender_stays_unclassified():
    """Boş sender + 'xidmət haqqı' debit — KOMİSİYA-ya düşməməli (catch-all üçün sender lazım)."""
    assert _match_category("", "Nağdsız köçürmə üzrə xidmət haqqı", "", "4.00", profile="access") is None


def test_access_menfeet_vergisi():
    """MƏNFƏƏT VERGİSİ/MENFEET VERGISI → VERGİ MƏXARİC (qlo bal)."""
    assert _match_category("", "111213 Menfeet vergisi", "", "300.00", profile="access") == "VERGİ MƏXARİC"
    assert _match_category("", "MƏNFƏƏT VERGİSİ", "", "50.00", profile="kapital") == "VERGİ MƏXARİC"


def test_access_nagdlasdirma():
    """Access-xas 'NAGDLASDIRMA' → Nağd çəkim komissiyası."""
    assert _match_category("", "Nağdlaşdırma üçün S/V", "", "49.00", profile="access") == "Nağd çəkim komissiyası"


def test_access_nagdlasdirma_standard_inactive():
    """Access Nağd çəkim token-i standard profildə aktiv deyil (izolyasiya)."""
    assert _match_category("", "Nağdlaşdırma üçün S/V", "", "49.00") != "Nağd çəkim komissiyası"


def test_access_end_to_end():
    """Real Access strukturuna uyğun end-to-end — sender ad+IBAN birlikdə ötürülür."""
    rows = [
        _access_row("Muqavileye esasen Mart ayi xidmet haqqi odenilir", 700, "", "HASANLI-H MMC\n1605270611/"),
        _access_row("Nağdlaşdırma üçün S/V", "", 49),
        _access_row("Maaş haqqqı kommissiyası AZN3.82", "", 3.82),
        _access_row("1\n111213\nMenfeet vergisi", "", 300, "VN Milli Gəlir Departamenti"),
        _access_row("Əmək haqqı\nəmək haqqı - salary", "", 100, "X MMC"),
        _access_row("Xüsusi əməliyyat - Məxaric (SOP)\nOPEN.SALARY.CARD.FEE", "", 10),
        _access_row("İcarə haqqı", "", 1000, "Huseynova Ramide Yusif qizi"),
    ]
    out, summary, unclassified = categorize_statement_verbose(_access_xlsx(rows))
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 700.0, "say": 1}
    assert summary["Nağd çəkim komissiyası"] == {"cem": 49.0, "say": 1}
    assert summary["KOMİSİYA"] == {"cem": 3.82, "say": 1}
    assert summary["VERGİ MƏXARİC"] == {"cem": 300.0, "say": 1}
    assert summary["MAAŞ MƏXARİC"] == {"cem": 100.0, "say": 1}
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 1000.0, "say": 1}
    idx = {u["row_index"] for u in unclassified}
    assert idx == {8}  # meta 1, header 2, data 3..9 → OPEN.SALARY.CARD.FEE row8

    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    # Kateqoriya/İcmal 11/12 (access sütunları 1-10)
    assert main.cell(row=2, column=11).value == "Kateqoriya"
    assert main.cell(row=2, column=12).value == "İcmal"
    assert main.cell(row=4, column=11).value == "Nağd çəkim komissiyası"
    assert main.cell(row=4, column=12).value == 49.0


def test_access_finalize_unclassified():
    """Access unclassified sətirinə manual təsnifat + finalize."""
    raw = _access_xlsx(
        [_access_row("Xüsusi əməliyyat - Məxaric (SOP)\nCHEQUE.BOOK", "", 10)]
    )
    _, _, unclassified = categorize_statement_verbose(raw)
    assert len(unclassified) == 1
    idx = unclassified[0]["row_index"]
    out, summary = finalize_statement(raw, {idx: "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"})
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 10.0, "say": 1}


def test_standard_and_kapital_regression_with_access():
    """YapıKredi (standard) + Kapital access əlavə olunduqdan sonra dəyişməz."""
    std = _xlsx_bytes(
        META,
        STD_HEADER,
        [_row("Com. bank komissiyası", "-5.50", "1"), _row("NAKIT CEKME KOMISYONU", "-2.00", "2")],
    )
    _, s1, _ = categorize_statement_verbose(std)
    assert s1["KOMİSİYA"] == {"cem": 5.5, "say": 1}
    assert s1["Nağd çəkim komissiyası"] == {"cem": 2.0, "say": 1}

    kap = _kapital_xlsx([[1, "2022-12-15", "x", "AZ / KAPITAL / 1", -17.5, "AZN", -17.5, 100, "hesabindan naglashdirma ucun 1 faiz vergi meblegi"]])
    _, s2, _ = categorize_statement_verbose(kap)
    assert s2["Nağd çəkim komissiyası"] == {"cem": 17.5, "say": 1}


# ===== VTB Bank format dəstəyi (Header 11-ci sətirdə; Debet/Kredit ayrı) =====

VTB_HEADER = [
    "Tarix",
    "Qarşı hesab",
    "Adı",
    "VÖEN",
    "Debet",
    "Kredit",
    "Qalıq",
    "Təyinat",
]


def _vtb_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["BANK VTB (AZƏRBAYCAN) ASC"])
    ws.append(["VÖEN: 1400117231"])
    ws.append(["Müştərinin adı", '"GLOBAL EXPORTS" MMC'])
    ws.append(["Hesab", "AZ19VTBA04006004375789AZN001"])
    ws.append(["VÖEN", "1302070301"])
    ws.append(["Valyuta", "AZN"])
    ws.append(["Çıxarışın müddəti", "2025-01-01 - 2025-12-31"])
    ws.append(["Dövrün əvvəlinə qalıq", "171.99"])
    ws.append(["Dövrün sonuna qalıq", "1058.46"])
    ws.append(["Dövr", "2025"])
    ws.append(VTB_HEADER)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _vtb_row(detail, debet=None, kredit=None, adi="FKT MMC", voen="2905379861"):
    return [
        "24.09.2025 17:50",
        "AZ29VTBA04016004375789RUB001",
        adi,
        voen,
        debet,
        kredit,
        "1000",
        detail,
    ]


def test_vtb_profile_detected_header_on_row_11():
    """VTB başlığı 11-ci sətirdədir; 10 meta sətir skan edilməlidir."""
    raw = _vtb_xlsx([_vtb_row("Outward remittance commission", debet=16.56, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231")])
    _, summary, _ = categorize_statement_verbose(raw)
    assert summary["KOMİSİYA"] == {"cem": 16.56, "say": 1}


def test_vtb_outward_remittance_commission():
    """'Outward remittance commission' ümumi 'commission' qaydası ilə → KOMİSİYA."""
    assert _match_category(
        "BANK VTB (AZƏRBAYCAN) ASC",
        "Outward remittance commission",
        None,
        "16.56",
        profile="vtb",
    ) == "KOMİSİYA"


def test_vtb_cyrillic_komissiya():
    """KRİTİK: kiril 'Комиссия за срочность' normalize ilə boşalır; KOMİSİYA olmalı."""
    assert _match_category(
        "BANK VTB (AZƏRBAYCAN) ASC",
        "Комиссия за срочность",
        None,
        "5",
        profile="vtb",
    ) == "KOMİSİYA"


def test_vtb_cyrillic_komissiya_other_profiles_inactive():
    """Kiril KOMİSİYA token-i yalnız vtb profildə aktivdir (izolyasiya)."""
    assert _match_category(
        "BANK VTB (AZƏRBAYCAN) ASC", "Комиссия за срочность", None, "5"
    ) != "KOMİSİYA"


def test_vtb_vergi_meharic():
    """DSMF / İcbari Tibbi Sığorta → VERGİ MƏXARİC (ümumi TAX keyword)."""
    assert _match_category(
        "DSMF Iri Sigortaedanlarla Is uzra Idara",
        "27.10.2025-ci il tarixli 25959703534AS saylı DSMF üzrə sərəncamın tam ödənişi",
        None,
        "403.39",
        profile="vtb",
    ) == "VERGİ MƏXARİC"
    assert _match_category(
        "Icbari Tibbi Sigorta uzra DA",
        "27.10.2025-ci il tarixli 25959703553AT saylı Icbari tibbi sığorta üzrə sərəncamın tam ödənişi",
        None,
        "0.21",
        profile="vtb",
    ) == "VERGİ MƏXARİC"


def test_vtb_edv_odenisi():
    """'esasen EDV odenilir' → ƏDV ÖDƏNİŞİ (ümumi VAT pattern)."""
    assert _match_category(
        "FKT MMC", "e-qaime MT2510 11711243 11711206 11711252 esasen EDV odenilir",
        None, "9600.52", profile="vtb",
    ) == "ƏDV ÖDƏNİŞİ"


def test_vtb_daxil_and_odenis_catchall():
    """Kredit → DAXİL OLMALAR; debit + sender → ÖDƏNİŞLƏRİ."""
    assert _match_category(
        '"GLOBAL EXPORTS" MMC',
        "Customer's transfer FCY INPUT_OUTPUT Currency Conversion N 189041",
        "87409", None, profile="vtb",
    ) == "DİGƏR ŞİRKƏT DAXİL OLMALAR"
    assert _match_category(
        "FKT MMC", "e-qaime MT2509 10887993 10887995 10887991 10887994 10887987",
        None, "16558.19", profile="vtb",
    ) == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"


def test_vtb_not_mistaken_for_yelo_access():
    """VTB başlığı yelo/access kimi tanınmamalıdır (token izolyasiyası)."""
    raw = _vtb_xlsx([_vtb_row("Outward remittance commission", debet=16.56, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231")])
    _, summary, _ = categorize_statement_verbose(raw)
    # yelo/access-ə xas 'xidmət haqqı'/'DƏBBƏ ÖDƏMƏ' yoxdur; profil düzgün seçilməlidir
    assert summary["KOMİSİYA"] == {"cem": 16.56, "say": 1}


def test_vtb_end_to_end_counts():
    """Real VTB strukturuna uyğun end-to-end — proqnoz 52/2/4/19/36, unclassified 0."""
    rows = []
    for _ in range(42):
        rows.append(_vtb_row("Outward remittance commission", debet=1.0, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231"))
    for _ in range(10):
        rows.append(_vtb_row("Комиссия за срочность", debet=5.0, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231"))
    rows.append(_vtb_row("e-qaime MT2510 11711243 esasen EDV odenilir", debet=9600.52))
    rows.append(_vtb_row("e-qaime MT2512 10065609 esasen EDV odenilir", debet=6549.68))
    rows.append(_vtb_row("27.10.2025 DSMF üzrə sərəncamın tam ödənişi", debet=403.39, adi="DSMF Iri Sigortaedanlarla Is uzra Idara", voen="1300115511"))
    rows.append(_vtb_row("27.10.2025 DSMF üzrə sərəncamın tam ödənişi", debet=41.08, adi="DSMF Iri Sigortaedanlarla Is uzra Idara", voen="1300115511"))
    rows.append(_vtb_row("27.10.2025 Icbari tibbi sığorta üzrə sərəncam", debet=0.21, adi="Icbari Tibbi Sigorta uzra DA", voen="1701911851"))
    rows.append(_vtb_row("27.10.2025 Icbari tibbi sığorta üzrə sərəncam", debet=53.14, adi="Icbari Tibbi Sigorta uzra DA", voen="1701911851"))
    for _ in range(19):
        rows.append(_vtb_row("Customer's transfer FCY INPUT_OUTPUT Currency Conversion", kredit=1000, adi='"GLOBAL EXPORTS" MMC', voen=""))
    for _ in range(36):
        rows.append(_vtb_row("e-qaime MT2509 10887993 esasen odenilir", debet=100))
    raw = _vtb_xlsx(rows)
    out, summary, unclassified = categorize_statement_verbose(raw)
    assert summary["KOMİSİYA"] == {"cem": 92.0, "say": 52}
    assert summary["ƏDV ÖDƏNİŞİ"] == {"cem": 16150.2, "say": 2}
    assert summary["VERGİ MƏXARİC"] == {"cem": pytest.approx(497.82), "say": 4}
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 19000.0, "say": 19}
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 3600.0, "say": 36}
    assert unclassified == []

    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    # header 11-ci sətir (1-əsaslı); Kateqoriya/İcmal 9/10
    assert main.cell(row=11, column=1).value == "Tarix"
    assert main.cell(row=11, column=9).value == "Kateqoriya"
    assert main.cell(row=11, column=10).value == "İcmal"
    # 42 outward → row 12 KOMİSİYA, İcmal 1.0
    assert main.cell(row=12, column=9).value == "KOMİSİYA"
    assert main.cell(row=12, column=10).value == 1.0
    # ilk kiril sətir → row 12+42+1 = 55 KOMİSİYA
    assert main.cell(row=55, column=9).value == "KOMİSİYA"
    # Xülasə sheet aktiv
    assert wb.active.title == "Xülasə"


def test_vtb_finalize():
    """VTB profili finalize (manual override) ilə uyğun işləyir."""
    raw = _vtb_xlsx([_vtb_row("hesablar arasi kocurme edilir", debet=5, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231")])
    _, _, unclassified = categorize_statement_verbose(raw)
    assert len(unclassified) == 0  # catch-all → ÖDƏNİŞLƏRİ, unclassified yoxdur
    out, summary = finalize_statement(raw, {})
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 5.0, "say": 1}


def test_all_profiles_sequential_no_mix():
    """standard + kapital + yelo + access + vtb ardıcıl — bir-birinə qarışmır."""
    std = _xlsx_bytes(META, STD_HEADER, [_row("Com. bank komissiyası", "-5.50", "1")])
    _, s1, _ = categorize_statement_verbose(std)
    assert s1["KOMİSİYA"] == {"cem": 5.5, "say": 1}

    kap = _kapital_xlsx([[1, "2022-12-15", "x", "AZ / KAPITAL / 1", 100, "AZN", 100, 100, "mal"]])
    _, s2, _ = categorize_statement_verbose(kap)
    assert s2["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 100.0, "say": 1}

    yelo = _yelo_xlsx([_yelo_row("YELO BANK ASC", "FAIZ ÖDƏMƏ", 0, 5)])
    _, s3, _ = categorize_statement_verbose(yelo)
    assert s3["BANK FAİZ XƏRCİ"] == {"cem": 5.0, "say": 1}

    access = _access_xlsx([_access_row("Muqavileye esasen Mart ayi xidmet haqqi odenilir", 700, "", "HASANLI-H MMC")])
    _, s4, _ = categorize_statement_verbose(access)
    assert s4["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 700.0, "say": 1}

    vtb = _vtb_xlsx([_vtb_row("Комиссия за срочность", debet=5.0, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231")])
    _, s5, _ = categorize_statement_verbose(vtb)
    assert s5["KOMİSİYA"] == {"cem": 5.0, "say": 1}


# ===== Turan Bank format dəstəyi (Header 13-cü sətirdə; tək "Məbləğ", sender yoxdur) =====

TURAN_HEADER = [
    "№",
    "Tarix",
    "Hesab nömrəsi",
    "Məbləğ",
    "Məbləğin AZN dəyəri",
    "Qalıq",
    "Qalığın AZN dəyəri",
    "Valyuta",
    "Təyinat",
]


def _turan_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append([None])  # 1-ci sətir boş (real fayla uyğun)
    ws.append(["Hesabdan çıxarış / Account statement"])
    ws.append(["Çıxarışın tarixi / Statement date:", None, None, None, "27-08-2026"])
    ws.append([""])
    ws.append(["Müştərinin Adı / Client's Name", None, None, None, "CR CONSTRUCTION MMC"])
    ws.append([""])
    ws.append(["Hesab adı", None, None, "CR CONSTRUCTION MMC", "Hesab nömrəsi / IBAN", None, "AZ09TURA40050096495200107944"])
    ws.append(["Hesabın valyutası / Account Currency", None, None, "AZN", "VÖEN", None, "2008334861"])
    ws.append(["İlk əməliyyat tarixi", None, None, "2026-01-01", "Son əməliyyat tarixi", None, "2026-08-27"])
    ws.append(["Günün əvvəlinə qalıq", None, None, 0, "Günün sonuna qalıq", None, 17.63])
    ws.append([""])
    ws.append([""])
    ws.append(TURAN_HEADER)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _turan_row(no, tarix, mebleg, teyinat):
    return [
        no,
        tarix,
        "AZ09TURA40050096495200107944",
        mebleg,
        mebleg,
        17.63,
        17.63,
        "AZN",
        teyinat,
    ]


def test_turan_profile_detected_header_on_row_13():
    """Turan başlığı 13-cü sətirdədir; 12 meta sətir skan edilməlidir."""
    raw = _turan_xlsx([_turan_row(0, "24-08-2026", -2162.25, "əsas borc üzrə ödəniş")])
    _, summary, _ = categorize_statement_verbose(raw)
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 2162.25, "say": 1}


def test_turan_not_stolen_by_kapital():
    """Konflikt reqressiyası: kapital (tarix+mebleg+teyinat) Turan başlığını oğurlamamalı."""
    raw = _turan_xlsx(
        [_turan_row(0, "24-08-2026", -2162.25, "əsas borc üzrə ödəniş")]
    )
    _, summary, _ = categorize_statement_verbose(raw)
    # kapital ssenarisi olsaydı sender _extract_sender_name eks_hesab-dan → boş → unclassified qalardı
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 2162.25, "say": 1}


def test_kapital_not_stolen_by_turan():
    """Kapital başlığı turan required token-lərinə (hesabnomresi/qaliq/mebleginazndeyeri) uyğun gəlmir."""
    kap = _kapital_xlsx([[1, "2022-12-15", "x", "AZ / KAPITAL / 1", 100, "AZN", 100, 100, "mal"]])
    _, s1, _ = categorize_statement_verbose(kap)
    assert s1["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 100.0, "say": 1}


def test_turan_skip_empty_tarix_teyinat():
    """Yalnız qalıq göstərən yekun sətri (Tarix/Təyinat boş) atlanmalıdır."""
    raw = _turan_xlsx(
        [
            _turan_row(0, "24-08-2026", -2162.25, "əsas borc üzrə ödəniş"),
            [None, None, None, 17.63, None, None, None, None, None],
        ]
    )
    out, summary, unclassified = categorize_statement_verbose(raw)
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 2162.25, "say": 1}
    assert len(unclassified) == 0
    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    assert main.cell(row=15, column=10).value is None  # yekun sətir kateqoriyasız
    assert main.cell(row=15, column=11).value is None


def test_turan_nagd_cekim_komissiyasi():
    """'Nəğd vəsaitə görə 1 % vergi' → Nağd çəkim komissiyası."""
    assert _match_category(
        "", "Nəğd vəsaitə görə 1 % vergi - Məbləğ 30000.00 - Vergi məbləği 300.00",
        None, "-300.00", profile="turan",
    ) == "Nağd çəkim komissiyası"


def test_turan_komisiya_patterns():
    """AZIPS xidmət haqqı / çek kitabçası rüsumu / hesab açılması / kassa komisyonu → KOMİSİYA."""
    assert _match_category(
        "", "AZIPS xidmət haqqı, VESAIT HESABIN MOHKEMLENDIRILMESI UCUN KOCURULUR",
        None, "-68.61", profile="turan",
    ) == "KOMİSİYA"
    assert _match_category(
        "", "çek kitabçası rüsumu (CR CONSTRUCTION MMC №__283976 dən__284000 dək)",
        None, "-10", profile="turan",
    ) == "KOMİSİYA"
    assert _match_category("", "hesab açılması", None, "-10", profile="turan") == "KOMİSİYA"
    assert _match_category(
        "", "Kreditlər üzrə kassa komisyonu - 1%", None, "-1000", profile="turan"
    ) == "KOMİSİYA"


def test_turan_cek_rusumu_not_rusum():
    """'çek kitabçası rüsumu' KOMİSİYA-ya düşməli, RÜSUM-a deyil (prioritet)."""
    assert _match_category("", "çek kitabçası rüsumu", None, "-10", profile="turan") == "KOMİSİYA"


def test_turan_catchall_without_sender():
    """Sender sütunu yoxdur — debit → ÖDƏNİŞLƏRİ, kredit → DAXİL OLMALAR."""
    assert _match_category("", "əsas borc üzrə ödəniş", None, "-2162.25", profile="turan") == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    assert _match_category("", "kredit tranşı", "100000", None, profile="turan") == "DİGƏR ŞİRKƏT DAXİL OLMALAR"


def test_turan_catchall_without_sender_not_in_standard():
    """Sender-siz catch-all yalnız turan üçündür — standard profildə boş sender unclassified qalır."""
    assert _match_category("", "ATM XIRDALAN", None, "-1000") is None


def test_turan_end_to_end_counts():
    """Real Turan strukturuna uyğun end-to-end — proqnoz 1/4/3/6, unclassified 0, skip 1."""
    rows = [
        _turan_row(0, "24-08-2026", -2162.25, "L032404/0001 - nömrəli müqavilə üçün əsas borc üzrə ödəniş"),
        _turan_row(1, "24-08-2026", -1479.13, "L032404/0001 - nömrəli müqavilə üçün faiz borc üzrə ödəniş"),
        _turan_row(2, "24-08-2026", 3600, "03 saylı 24.08.2026 taixli Borc müq əsasən təsisçidən - Cabbarov Əfqan Rəşid oğlu"),
        _turan_row(3, "24-07-2026", -1391.38, "L032404/0001 - nömrəli müqavilə üçün əsas borc üzrə ödəniş"),
        _turan_row(4, "24-07-2026", -2250, "L032404/0001 - nömrəli müqavilə üçün faiz borc üzrə ödəniş"),
        _turan_row(5, "23-07-2026", 3700, "ƏDV ödəyicisi AZIPS, HESABIN MOHKEMLENMESI KREDIT ODENISI UCUN"),
        _turan_row(6, "11-06-2026", -68611, "AZIPS, VESAIT HESABIN MOHKEMLENDIRILMESI UCUN KOCURULUR"),
        _turan_row(7, "11-06-2026", -68.61, "AZIPS xidmət haqqı, VESAIT HESABIN MOHKEMLENDIRILMESI UCUN KOCURULUR"),
        _turan_row(8, "09-06-2026", -30000, "09/06/2026 tarixli L032404/0001 nömrəli müqavilə üzrə kredit məxaric"),
        _turan_row(9, "09-06-2026", -300, "Nəğd vəsaitə görə 1 % vergi - Məbləğ 30000.00 - Vergi məbləği 300.00"),
        _turan_row(10, "09-06-2026", -10, "çek kitabçası rüsumu (CR CONSTRUCTION MMC №__283976 dən__284000 dək)"),
        _turan_row(11, "09-06-2026", 100000, "L032404/0001- müqavilə üzrə kredit tranşı"),
        _turan_row(12, "09-06-2026", -10, "hesab açılması"),
        _turan_row(13, "09-06-2026", -1000, "Kreditlər üzrə kassa komisyonu - 1%"),
        [None, None, None, 17.63, None, None, None, None, None],  # yekun sətir — skip
    ]
    raw = _turan_xlsx(rows)
    out, summary, unclassified = categorize_statement_verbose(raw)
    assert summary["Nağd çəkim komissiyası"] == {"cem": 300.0, "say": 1}
    assert summary["KOMİSİYA"] == {"cem": 1088.61, "say": 4}
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 107300.0, "say": 3}
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 105893.76, "say": 6}
    assert unclassified == []

    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    # header 13-cü sətir (1-əsaslı); Kateqoriya/İcmal 10/11
    assert main.cell(row=13, column=1).value == "№"
    assert main.cell(row=13, column=10).value == "Kateqoriya"
    assert main.cell(row=13, column=11).value == "İcmal"
    assert main.cell(row=14, column=10).value == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    assert main.cell(row=14, column=11).value == 2162.25
    assert main.cell(row=16, column=10).value == "DİGƏR ŞİRKƏT DAXİL OLMALAR"
    assert main.cell(row=23, column=10).value == "Nağd çəkim komissiyası"
    assert main.cell(row=28, column=10).value is None  # skip sətir
    # Xülasə sheet aktiv
    assert wb.active.title == "Xülasə"


def test_turan_finalize():
    """Turan profili finalize (işarəli Məbləğ) ilə uyğun işləyir."""
    raw = _turan_xlsx([_turan_row(0, "09-06-2026", -1000, "Kreditlər üzrə kassa komisyonu - 1%")])
    _, _, unclassified = categorize_statement_verbose(raw)
    assert len(unclassified) == 0
    out, summary = finalize_statement(raw, {})
    assert summary["KOMİSİYA"] == {"cem": 1000.0, "say": 1}


def test_all_six_profiles_sequential_no_mix():
    """standard + turan + kapital + yelo + access + vtb ardıcıl — bir-birinə qarışmır."""
    std = _xlsx_bytes(META, STD_HEADER, [_row("Com. bank komissiyası", "-5.50", "1")])
    _, s1, _ = categorize_statement_verbose(std)
    assert s1["KOMİSİYA"] == {"cem": 5.5, "say": 1}

    tur = _turan_xlsx([_turan_row(0, "09-06-2026", -1000, "Kreditlər üzrə kassa komisyonu - 1%")])
    _, s2, _ = categorize_statement_verbose(tur)
    assert s2["KOMİSİYA"] == {"cem": 1000.0, "say": 1}

    kap = _kapital_xlsx([[1, "2022-12-15", "x", "AZ / KAPITAL / 1", 100, "AZN", 100, 100, "mal"]])
    _, s3, _ = categorize_statement_verbose(kap)
    assert s3["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 100.0, "say": 1}

    yelo = _yelo_xlsx([_yelo_row("YELO BANK ASC", "FAIZ ÖDƏMƏ", 0, 5)])
    _, s4, _ = categorize_statement_verbose(yelo)
    assert s4["BANK FAİZ XƏRCİ"] == {"cem": 5.0, "say": 1}

    access = _access_xlsx([_access_row("Muqavileye esasen Mart ayi xidmet haqqi odenilir", 700, "", "HASANLI-H MMC")])
    _, s5, _ = categorize_statement_verbose(access)
    assert s5["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 700.0, "say": 1}

    vtb = _vtb_xlsx([_vtb_row("Комиссия за срочность", debet=5.0, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231")])
    _, s6, _ = categorize_statement_verbose(vtb)
    assert s6["KOMİSİYA"] == {"cem": 5.0, "say": 1}


# ===== Ziraat Bank format dəstəyi (Header 13-cü sətirdə; Debet/Kredit ayrı, sender = Müxbirin adı) =====

ZIRAAT_HEADER = [
    "Tarix",
    "Sənədin N-si",
    "Müxbirin hesabı",
    "Müxbirin adı",
    "Debet",
    "Kredit",
    "Ara qaliq",
    "Təyinat",
    "Əməliyat saatı",
]


def _ziraat_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["27.05.25 16:55"])
    ws.append(["Ziraat Bank Azərbaycan ASC"])
    ws.append(["Kod  - 513218,M/H - AZ42NABZ01350100000000095944"])
    ws.append([""])
    ws.append([None, None, None, "HESABDAN ÇIXARIŞ"])
    ws.append([None, None, None, "01.09.2021 - 31.12.2022 tarixləri arası"])
    ws.append([""])
    ws.append(["Hesablaşma hesabı: ", "AZ53TCZB40030944006598700102"])
    ws.append(["Hesabın valyutası: ", "AZN"])
    ws.append(["Hesabın Adı: ", "İNTEKS MMC"])
    ws.append(["Giriş qalığı: ", 0])
    ws.append([""])
    ws.append(ZIRAAT_HEADER)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _ziraat_row(tarix, sened, mux_hesab, mux_adi, debet, kredit, teyinat):
    return [
        tarix,
        sened,
        mux_hesab,
        mux_adi,
        debet,
        kredit,
        "1 000.00",
        teyinat,
        "15:36:18",
    ]


def test_ziraat_profile_detected_header_on_row_13():
    """Ziraat başlığı 13-cü sətirdədir; 12 meta sətir skan edilməlidir."""
    raw = _ziraat_xlsx(
        [_ziraat_row("30.09.2021", "CODE", "AZ75TCZB40130978006598700102", "İNTEKS MMC", 0, 1198804.1, "Valyuta Satışı 1.963")]
    )
    _, summary, _ = categorize_statement_verbose(raw)
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 1198804.1, "say": 1}


def test_ziraat_not_stolen_by_other_profiles():
    """Mövcud profillər Ziraat başlığını tanımır — catch-all davranışı düzgündür."""
    raw = _ziraat_xlsx(
        [_ziraat_row("30.09.2021", "CODE", "AZ75TCZB40130978006598700102", "İNTEKS MMC", 0, 1000, "Valyuta Satışı")]
    )
    _, summary, _ = categorize_statement_verbose(raw)
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 1000.0, "say": 1}


def test_ziraat_skip_summary_rows():
    """CƏMİ / Çıxış qalıq sətirləri (Təyinat + sender boş) atlanmalıdır."""
    raw = _ziraat_xlsx(
        [
            _ziraat_row("30.09.2021", "CODE", "AZ75TCZB40130978006598700102", "İNTEKS MMC", 0, 1000, "Valyuta Satışı"),
            ["CƏMİ:", None, None, None, "1000", "1000", None, None, None],
            ["Çıxış qalıq:", 0, None, None, None, None, None, None, None],
        ]
    )
    out, summary, unclassified = categorize_statement_verbose(raw)
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 1000.0, "say": 1}
    assert len(unclassified) == 0
    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    assert main.cell(row=15, column=10).value is None  # CƏMİ sətri kateqoriyasız
    assert main.cell(row=16, column=10).value is None  # Çıxış qalıq sətri kateqoriyasız


def test_ziraat_keeps_sms_fee_rows():
    """Boş Təyinat AMA dolu sender (SMS xidmət haqqı) — skip edilməməlidir (catch-all)."""
    rows = [
        _ziraat_row("21.04.2022", "6795", "AZ37TCZB67017944000000001002", "CARDS:SMS NOTİFİCATİONI ÜZRƏ XİDMƏT HAQQI", 1.18, 0, None)
    ]
    out, summary, unclassified = categorize_statement_verbose(_ziraat_xlsx(rows))
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": 1.18, "say": 1}
    assert len(unclassified) == 0


def test_ziraat_komisiya():
    """'Əməliyyat üzrə komissiya' → KOMİSİYA (ümumi 'kom')."""
    assert _match_category(
        "İçərişəhər filialı",
        "Bank tariflərinə uyğun olaraq Əməliyyat üzrə komissiya \"Sadə köçürmə\" məbləğ 120.00 (AZN). ƏDV daxil deyil",
        None, "120.00", profile="ziraat",
    ) == "KOMİSİYA"


def test_ziraat_vergi():
    """VN Milli Gəlir Departamenti sender → VERGİ MƏXARİC."""
    assert _match_category(
        "VN Milli Gəlir.Departamenti - AZ17CTRE00000000000002117131 - 1602473621",
        "3 rub 2021 il uzre sadelew vergi",
        None, "1000000", profile="ziraat",
    ) == "VERGİ MƏXARİC"


def test_ziraat_edv():
    """'AVANS ODENIWI UZRE EDV' → ƏDV ÖDƏNİŞİ."""
    assert _match_category(
        "BAKU STEEL COMPANY MMC", "AVANS ODENIWI UZRE EDV", None, "18000", profile="ziraat"
    ) == "ƏDV ÖDƏNİŞİ"


def test_ziraat_catchall_directions():
    """Kredit → DAXİL OLMALAR; debit → ÖDƏNİŞLƏRİ."""
    assert _match_category(
        "İNTEKS MMC", "Valyuta Satışı 1.963", "1198804.1", None, profile="ziraat"
    ) == "DİGƏR ŞİRKƏT DAXİL OLMALAR"
    assert _match_category(
        "BAKU STEEL COMPANY MMC", "AVANS ODENIWI", None, "18000", profile="ziraat"
    ) == "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"


def test_ziraat_end_to_end_counts():
    """Real Ziraat strukturuna uyğun end-to-end — proqnoz 342/42/42/46/331 = 803, skip 2."""
    rows = []
    for _ in range(342):
        rows.append(_ziraat_row("15.10.2021", "x", "AZ97TCZB67013944000000000102", "İçərişəhər filialı", 1.0, 0, "Bank tariflərinə uyğun olaraq Əməliyyat üzrə komissiya"))
    for _ in range(42):
        rows.append(_ziraat_row("06.10.2021", "02", "AZ17CTRE00000000000002117131", "VN Milli Gəlir.Departamenti", 1000, 0, "sadelew vergi"))
    for _ in range(42):
        rows.append(_ziraat_row("11.10.2021", "04", "AZ48CTRE00000000000000014001", "BAKU STEEL COMPANY MMC", 100, 0, "AVANS ODENIWI UZRE EDV"))
    for _ in range(46):
        rows.append(_ziraat_row("30.09.2021", "CODE", "AZ75TCZB40130978006598700102", "İNTEKS MMC", 0, 1000, "Valyuta Satışı"))
    for _ in range(323):
        rows.append(_ziraat_row("13.10.2021", "04", "AZ62AIIB40060019440419948105", "BAKI MEXANIKA", 100, 0, "EBBF qaliz mebleq"))
    for _ in range(8):
        rows.append(_ziraat_row("21.04.2022", "6795", "AZ37TCZB67017944000000001002", "CARDS:SMS NOTİFİCATİONI ÜZRƏ XİDMƏT HAQQI", 1.18, 0, None))
    rows.append(["CƏMİ:", None, None, None, "1000", "1000", None, None, None])
    rows.append(["Çıxış qalıq:", 0, None, None, None, None, None, None, None])
    raw = _ziraat_xlsx(rows)
    out, summary, unclassified = categorize_statement_verbose(raw)
    assert summary["KOMİSİYA"] == {"cem": 342.0, "say": 342}
    assert summary["VERGİ MƏXARİC"] == {"cem": 42000.0, "say": 42}
    assert summary["ƏDV ÖDƏNİŞİ"] == {"cem": 4200.0, "say": 42}
    assert summary["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 46000.0, "say": 46}
    assert summary["DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"] == {"cem": pytest.approx(32309.44), "say": 331}
    assert unclassified == []

    wb = _reload(out)
    main = wb[wb.sheetnames[0]]
    assert main.cell(row=13, column=1).value == "Tarix"
    assert main.cell(row=13, column=10).value == "Kateqoriya"
    assert main.cell(row=13, column=11).value == "İcmal"
    assert main.cell(row=14, column=10).value == "KOMİSİYA"
    assert main.cell(row=14, column=11).value == 1.0
    # yekun sətirlər (CƏMİ 817, Çıxış qalıq 818) kateqoriyasız
    assert main.cell(row=817, column=10).value is None
    assert main.cell(row=818, column=10).value is None
    assert wb.active.title == "Xülasə"


def test_ziraat_finalize():
    """Ziraat profili finalize ilə uyğun işləyir."""
    raw = _ziraat_xlsx(
        [_ziraat_row("11.10.2021", "04", "AZ48CTRE00000000000000014001", "BAKU STEEL COMPANY MMC", 18000, 0, "AVANS ODENIWI UZRE EDV")]
    )
    _, _, unclassified = categorize_statement_verbose(raw)
    assert len(unclassified) == 0
    out, summary = finalize_statement(raw, {})
    assert summary["ƏDV ÖDƏNİŞİ"] == {"cem": 18000.0, "say": 1}


def test_all_seven_profiles_sequential_no_mix():
    """standard + turan + kapital + yelo + access + vtb + ziraat ardıcıl — qarışmır."""
    std = _xlsx_bytes(META, STD_HEADER, [_row("Com. bank komissiyası", "-5.50", "1")])
    _, s1, _ = categorize_statement_verbose(std)
    assert s1["KOMİSİYA"] == {"cem": 5.5, "say": 1}

    tur = _turan_xlsx([_turan_row(0, "09-06-2026", -1000, "Kreditlər üzrə kassa komisyonu - 1%")])
    _, s2, _ = categorize_statement_verbose(tur)
    assert s2["KOMİSİYA"] == {"cem": 1000.0, "say": 1}

    kap = _kapital_xlsx([[1, "2022-12-15", "x", "AZ / KAPITAL / 1", 100, "AZN", 100, 100, "mal"]])
    _, s3, _ = categorize_statement_verbose(kap)
    assert s3["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 100.0, "say": 1}

    yelo = _yelo_xlsx([_yelo_row("YELO BANK ASC", "FAIZ ÖDƏMƏ", 0, 5)])
    _, s4, _ = categorize_statement_verbose(yelo)
    assert s4["BANK FAİZ XƏRCİ"] == {"cem": 5.0, "say": 1}

    access = _access_xlsx([_access_row("Muqavileye esasen Mart ayi xidmet haqqi odenilir", 700, "", "HASANLI-H MMC")])
    _, s5, _ = categorize_statement_verbose(access)
    assert s5["DİGƏR ŞİRKƏT DAXİL OLMALAR"] == {"cem": 700.0, "say": 1}

    vtb = _vtb_xlsx([_vtb_row("Комиссия за срочность", debet=5.0, adi="BANK VTB (AZƏRBAYCAN) ASC", voen="1400117231")])
    _, s6, _ = categorize_statement_verbose(vtb)
    assert s6["KOMİSİYA"] == {"cem": 5.0, "say": 1}

    zir = _ziraat_xlsx([_ziraat_row("11.10.2021", "04", "AZ48CTRE00000000000000014001", "BAKU STEEL COMPANY MMC", 100, 0, "AVANS ODENIWI UZRE EDV")])
    _, s7, _ = categorize_statement_verbose(zir)
    assert s7["ƏDV ÖDƏNİŞİ"] == {"cem": 100.0, "say": 1}