import io
import re
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from app.services.deterministic.creditors import (
    MAIN_SHEET_NAME,
    UNMATCHED_SHEET_NAME,
    _clean_counterparty_name,
    _detect_bank_header,
    _extract_voen_and_name,
    _extract_voen_from_text,
    _load_rows,
    _split_counter_account,
    build_workbook,
    generate_creditors_workbook,
    match_by_voen,
    parse_bank_statement,
    parse_depozit,
    parse_gelen,
    sanitize_sheet_name,
    unique_sheet_name,
)

FIXTURES = Path(__file__).parent / "fixtures"


def _save(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ===== Sintetik mənbə faylları =====

_GELEN_HEADER = [
    "№", "VÖEN", "Adı", "Tipi", "Vəziyyəti", "Qaimə tarixi",
    "Qaimə seriyası", "Qaimə nömrəsi", "Qaimə/Akt növləri",
    "Əsas qeyd", "Əlavə qeyd", "Aksiz məbləği",
    "Malın ƏDV-siz ümumi dəyəri", "Malın ƏDV məbləği",
]

_DEPOZIT_HEADER = [
    "№", "VÖEN", "Vergi ödəyicisinin adı", "Ödəniş edən vətəndaşın FİN-i",
    "Əməliyyat tarixi", "Mədaxil", "Məxaric", "Hesabın növü",
    "Əməliyyatın növü", "Sənədin nömrəsi", "Əməliyyatın təyinatı",
]

_BANK_HEADER = [
    "№", "Tarix", "Sənədin nömrəsi", "Əks hesabın nömrəsi, adı və vöen-i",
    "Məbləğ", "Valyuta", "AZN dəyəri", "Balans", "Təyinat",
]


def _gelen_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_GELEN_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _depozit_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_DEPOZIT_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _bank_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_BANK_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _gelen_row(voen, ad, tarix, seriya, nomre, esas, edv):
    return [
        "", voen, ad, "Cari", "Təsdiqləndi", tarix, seriya, nomre,
        "", "", "", 0, esas, edv,
    ]


def _depozit_row(voen, ad, tarix, medaxil, mexaric):
    return ["", voen, ad, "", tarix, medaxil, mexaric, "", "", "", ""]


def _bank_row(tarix, eks_hesab, mebleg, azn):
    return ["", tarix, "", eks_hesab, mebleg, "AZN", azn, "", ""]


_PASHA_HEADER = [
    "Əməliyyat tarixi", "İcra tarixi", "Ödəyən/Benefisiar", "Təyinat",
    "İstinad No", "VÖEN", "Mədaxil", "Məxaric", "Balans",
]


def _pasha_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_PASHA_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _pasha_row(tarix, odeyen, voen, medaxil, mexaric):
    return [tarix, tarix, odeyen, "Urgent payment", "", voen, medaxil, mexaric, ""]


_ABB_HEADER = [
    "Əməliyyatın tarixi / Operation date",
    "Əməliyyatın referensi / Reference İD",
    "Əks hesab üzrə məlumatlar / Counterparty account details",
    "Debit məbləğ / Debit amount",
    "Kredit məbləğ / Credit amount",
    "Ödənişin təyinatı / Description",
    "Qarşı tərəfin VÖEN-i / Counterparty TAX ID",
]


def _abb_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_ABB_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _abb_row(tarix, debit, kredit, voen):
    return [tarix, "ref", "eks", debit, kredit, "teyinat", voen]


_UNIBANK_HEADER = [
    "VÖEN\n(TIN)",
    "Əməliyyat tarixi\n(Transaction date)",
    "Mədaxil\n(Paid In)",
    "Məxaric\n(Paid Out)",
    "Balans\n(Balance)",
    "Əməliyyat Tərəfdaşı\n(Counterparty Name)",
    "Təyinat\n(Description)",
    "POS Transaction ID",
    "Transaction ID",
]


def _unibank_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_UNIBANK_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _unibank_row(voen, tarix, medaxil, mexaric, ad, teyinat):
    return [voen, tarix, medaxil, mexaric, "", ad, teyinat, "", ""]


_TURAN_HEADER = [
    "№", "Tarix", "Hesab nömrəsi", "Məbləğ", "Məbləğin AZN dəyəri",
    "Qalıq", "Qalığın AZN dəyəri", "Valyuta", "Təyinat",
]


def _turan_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_TURAN_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _turan_row(hesab, tarix, mebleg, azn, teyinat):
    return ["", tarix, hesab, mebleg, azn, "", "", "AZN", teyinat]


_RABITE_HEADER = [
    "Əməliyyat tarixi", "Silinmə tarixi", "Benefisiarın bankı",
    "Qarşı tərəfin hesab nömrəsi", "Qarşı tərəfin adı",
    "Debet", "Kredit", "Balans", "Təsvir",
]


def _rabitabank_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_RABITE_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _rabitabank_row(tarix, hesab, ad, debet, kredit, tesvir=""):
    return [tarix, tarix, "RABİTƏBANK ASC", hesab, ad, debet, kredit, "", tesvir]


_ACCESS_HEADER = [
    "", "Hesablaşma tarixi\nValue Date", "Əməliyyat tarixi\nPosting Date",
    "Əməliyyatın №-si\nTransaction ID", "Mədaxil\nIn", "Məxaric\nOut",
    "Balans\nBalance", "Göndərənin/Alanın Adı\nSender / Beneficiary Name",
    "Göndərən/Alanin Tax ID\nSender/Beneficiary TAX ID", "Təyinat\nDescription",
]


def _accessbank_xlsx(rows, meta_rows=1) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append([""])
    ws.append(_ACCESS_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _accessbank_row(tarix, tid, mexaric, ad, taxid, teyinat, medaxil=""):
    return ["", tarix, tarix, tid, medaxil, mexaric, "", ad, taxid, teyinat]


_BANKRES_HEADER = [
    "", "Tarix", "Sənədin N", "Göndərənin adı/Alanın adı", "VÖEN",
    "Hesab", "Ödənişin təyinatı", "Kredit", "Debet", "Qalıq",
]


def _bankrespublika_xlsx(rows, meta_rows=2) -> bytes:
    wb = Workbook()
    ws = wb.active
    for _ in range(meta_rows):
        ws.append(["meta"])
    ws.append(_BANKRES_HEADER)
    for r in rows:
        ws.append(r)
    return _save(wb)


def _bankrespublika_row(tarix, ad, voen, kredit, debet, teyinat=""):
    return ["", tarix, "DOC1", ad, voen, "AZ00BRES...", teyinat, kredit, debet, ""]


# ===== Parser testləri =====


def test_parse_gelen_fields():
    raw = _gelen_xlsx(
        [
            _gelen_row("2902062861", "EQUIPMENT MMC", datetime(2026, 1, 6), "MT2601", "10149359", 6730.77, 1211.5386),
            _gelen_row("1701560031", "TOPCONTROL MMC", datetime(2026, 1, 7), "MT2601", "10279782", 60, 0),
        ]
    )
    rows = parse_gelen(raw)
    assert len(rows) == 2
    assert rows[0]["voen"] == "2902062861"
    assert rows[0]["ad"] == "EQUIPMENT MMC"
    assert rows[0]["tarix"] == "06.01.2026"
    assert rows[0]["seriya"] == "MT2601"
    assert rows[0]["nomre"] == "10149359"
    assert rows[0]["esas"] == 6730.77
    assert rows[0]["edv"] == 1211.5386
    assert rows[1]["edv"] == 0.0


def test_parse_gelen_skips_empty_voen():
    raw = _gelen_xlsx(
        [
            ["", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            _gelen_row("1701560031", "TOPCONTROL MMC", datetime(2026, 1, 7), "MT2601", "10279782", 60, 0),
        ]
    )
    rows = parse_gelen(raw)
    assert len(rows) == 1


def test_parse_gelen_missing_header_raises():
    wb = Workbook()
    ws = wb.active
    ws.append(["a", "b"])
    with pytest.raises(ValueError, match="GƏLƏN"):
        parse_gelen(_save(wb))


def test_parse_depozit_only_mexaric():
    raw = _depozit_xlsx(
        [
            _depozit_row("1601005511", "IMA ENERGY MMC", datetime(2026, 8, 31), 9000, 0),
            _depozit_row("1403427061", "SAR-2016 MMC", datetime(2026, 8, 19), 0, 1296),
            _depozit_row("2004739712", "ABDULLAYEVA R", datetime(2026, 8, 19), 0, 208.23),
        ]
    )
    rows = parse_depozit(raw)
    assert len(rows) == 2  # medaxil sətri (mexaric=0) atlanır
    assert {r["voen"] for r in rows} == {"1403427061", "2004739712"}
    assert {r["mexaric"] for r in rows} == {1296.0, 208.23}


def test_parse_bank_statement_only_debit():
    raw = _bank_xlsx(
        [
            _bank_row(datetime(2026, 3, 13), "AZ37AIIB / IMA ENERGY MMC / 1601005511", 100000, 100000),
            _bank_row(datetime(2026, 3, 16), "AZ43AIIB / VƏLİYEV MOBİL / 3201279492", -11400, -11400),
            _bank_row(datetime(2026, 3, 16), "AZ09AIIB / KAPITAL / 9900003611", -9.12, -9.12),
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 2  # kredit (müsbət) atlanır
    assert rows[0]["voen"] == "3201279492"
    assert rows[0]["mebleg"] == 11400.0
    assert rows[1]["voen"] == "9900003611"
    assert rows[1]["mebleg"] == 9.12


def test_parse_bank_statement_missing_header_raises():
    wb = Workbook()
    ws = wb.active
    ws.append(["a", "b"])
    with pytest.raises(ValueError, match="başlıq sətri tapılmadı"):
        parse_bank_statement(_save(wb))


def test_clean_counterparty_name():
    assert (
        _clean_counterparty_name(
            "/AZ55KABA00000000000006526675_x000D_CAUCASIAN SUPPLY AND SERVICES MMC_x000D_/"
        )
        == "CAUCASIAN SUPPLY AND SERVICES MMC"
    )
    assert _clean_counterparty_name("Pasha Bank") == "Pasha Bank"
    assert _clean_counterparty_name("") == ""


def test_extract_voen_and_name():
    # xalis rəqəm (ayrıca VÖEN sütunu) + ad ayrıca
    assert _extract_voen_and_name("1601005511", "IMA ENERGY MMC") == (
        "1601005511",
        "IMA ENERGY MMC",
    )
    # qarışıq mətn (Kapital) — VÖEN + ad eyni hüceyrədə
    assert _extract_voen_and_name(
        "AZ37AIIB / IMA ENERGY MMC / 1601005511", ""
    ) == ("1601005511", "IMA ENERGY MMC")
    # boş VÖEN + PAŞA `_x000D_` adı
    assert _extract_voen_and_name(
        "", "/AZ03KABA_x000D_CENTRO MMC_x000D_/"
    ) == (None, "CENTRO MMC")


def test_parse_bank_statement_pasha():
    raw = _pasha_xlsx(
        [
            _pasha_row("09.06.2026", "CAUCASIAN MMC", "2006687041", "", 5500),
            _pasha_row("09.06.2026", "Pasha Bank", "", "", 1000),  # bank haqqı (boş VÖEN → skip)
            _pasha_row("10.06.2026", "IMA ENERGY MMC", "1601005511", 100000, ""),  # mədaxil
            _pasha_row("11.06.2026", "CENTRO MMC", "1202255201", "", 2600),
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 2  # mədaxil + boş-VÖEN sətirləri atlanır
    assert {r["voen"] for r in rows} == {"2006687041", "1202255201"}
    by_voen = {r["voen"]: r for r in rows}
    assert by_voen["2006687041"]["mebleg"] == 5500.0
    assert by_voen["2006687041"]["ad"] == "CAUCASIAN MMC"
    assert by_voen["1202255201"]["mebleg"] == 2600.0


def test_parse_bank_statement_pasha_real():
    raw = (FIXTURES / "PAŞA BANK ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    assert len(rows) == 43  # 67 debit - 24 boş-VÖEN bank haqqı sətri (skip)

    totals: dict[str, float] = {}
    empty_voen = 0
    for r in rows:
        if r["voen"]:
            totals[r["voen"]] = totals.get(r["voen"], 0.0) + r["mebleg"]
        else:
            empty_voen += 1

    assert round(totals["2006687041"], 2) == 23371.72
    assert round(totals["1202255201"], 2) == 14643.0
    assert empty_voen == 0  # boş-VÖEN sətirləri skip edildi
    # DÖVRÜN SONUNA BALANS sətri (349999.79) heç bir yerdə debit kimi düşməməlidir
    assert all(r["mebleg"] != 349999.79 for r in rows)


def test_parse_bank_statement_abb():
    raw = (FIXTURES / "ABB ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    # 9 debit sətri daxili əməliyyatdır (kart sifarişi/əmək haqqı/bank haqqı),
    # hamısının VÖEN-i boş → skip edilir.
    assert rows == []


def test_parse_bank_statement_abb_synthetic():
    raw = _abb_xlsx(
        [
            _abb_row("13.07.2026", 100, 0, " 1506275021"),  # dolu VÖEN (trim) → debit
            _abb_row("14.07.2026", 50, 0, ""),  # boş VÖEN → skip
            _abb_row("15.07.2026", 0, 200, "1506275021"),  # kredit (debit=0) → skip
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 1
    assert rows[0]["voen"] == "1506275021"
    assert rows[0]["mebleg"] == 100.0


def test_parse_bank_statement_unibank():
    raw = (FIXTURES / "UNİBANK ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    assert len(rows) == 74  # 95 debit - 20 boş-VÖEN - 1 Yekun(summary)
    assert all(r["voen"] for r in rows)  # boş-VÖEN sətirləri skip edildi
    assert all(r["mebleg"] != 965154.48 for r in rows)  # Yekun (Summary) skip

    voens = {r["voen"] for r in rows}
    assert "1102475261" in voens  # şirkətin öz VÖEN-i (Uyğunsuza düşəcək, skip deyil)
    assert "1300017201" in voens  # filial kodu

    dsmf = next(r for r in rows if r["voen"] == "1300115511")
    assert dsmf["ad"] == "Hüquqi şəxslər üzrə Adi Merkezi Sosial sig.FUI"
    assert dsmf["mebleg"] == 4500.0


def test_parse_bank_statement_unibank_synthetic():
    raw = _unibank_xlsx(
        [
            _unibank_row("1102475261", "05.01.2026", "32700", "-", "ÖZ ŞİRKƏT", "kassa"),  # Məxaric "-" → skip
            _unibank_row("1102475261", "06.01.2026", "-", "5000", "ÖZ ŞİRKƏT", "hesablararasi"),  # öz VÖEN debit → dolu
            _unibank_row("3103655452", "07.01.2026", "-", "4500", "SUPPLIER X", "ödəniş"),  # string debit
            _unibank_row("", "08.01.2026", "-", "100", "", "daxili"),  # boş VÖEN → skip
            _unibank_row("Yekun (Summary)", "", "-", "965154.48", "", ""),  # summary → skip
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 2
    by_voen = {r["voen"]: r for r in rows}
    assert by_voen["3103655452"]["mebleg"] == 4500.0
    assert by_voen["1102475261"]["mebleg"] == 5000.0


def test_parse_bank_statement_turanbank():
    raw = (FIXTURES / "TURAN BANK ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    assert len(rows) == 11  # mənfi Məbləğ/Məbləğin AZN dəyəri → debit
    assert all(r["voen"] is None for r in rows)  # VÖEN sütunu yoxdur → Uyğunsuz
    amounts = sorted(r["mebleg"] for r in rows)
    assert amounts == sorted(
        [2162.25, 1479.13, 1391.38, 2250, 68611, 68.61, 30000, 300, 10, 10, 1000]
    )
    # "Cəmi" (Hesab nömrəsi sütununda, E=17.63) sətri daxil deyil
    assert all(r["mebleg"] != 17.63 for r in rows)
    # ad = Təyinat
    first = next(r for r in rows if r["mebleg"] == 2162.25)
    assert first["ad"] == "L032404/0001 - nömrəli müqavilə üçün əsas borc üzrə ödəniş"


def test_parse_bank_statement_turanbank_synthetic():
    raw = _turan_xlsx(
        [
            _turan_row("AZ09...", "24-08-2026", -2162.25, -2162.25, "əsas borc ödənişi"),
            _turan_row("AZ09...", "25-08-2026", -100, -85, "valyuta fərqli"),  # AZN dəyəri (E) üstün
            _turan_row("Cəmi", "", -9999.00, -9999.00, ""),  # MƏNFİ cəm sətri → ref sütunundan skip
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 2
    by_date = {r["tarix"]: r for r in rows}
    assert by_date["24-08-2026"]["mebleg"] == 2162.25
    assert by_date["24-08-2026"]["ad"] == "əsas borc ödənişi"
    # "azndeyeri" (E) üstünlüyü: D=-100, E=-85 → mebleg 85 (mebleg 100 deyil)
    assert by_date["25-08-2026"]["mebleg"] == 85.0


def test_extract_voen_from_text():
    assert _extract_voen_from_text("Ödəniş VOEN: 1403610931 bitdi") == "1403610931"
    assert _extract_voen_from_text("VÖEN 9900001231") == "9900001231"
    assert _extract_voen_from_text("adi mətn 1403610931 yox") == "1403610931"
    # Ayırıcı variasiyaları: boşluq/kolon/hash/alt-xətt (Rabitə `VOEN___`)
    assert _extract_voen_from_text("AZ82UBAZ001 VOEN___1405631661") == "1405631661"
    assert _extract_voen_from_text("IBAN VOEN_1503971131") == "1503971131"
    assert _extract_voen_from_text("VOEN#1600855111") == "1600855111"
    # IBAN içindəki rəqəm ardıcıllığı VÖEN deyil
    assert _extract_voen_from_text("IBAN AZ55KABA00000000000006526675") is None
    assert _extract_voen_from_text("") is None


def test_parse_bank_statement_yapikredi():
    raw = (FIXTURES / "YAPIKREDİ BANK ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    assert len(rows) == 3  # yalnız Debit (mənfi) sətirləri; Kredit/mədaxil atlanır
    amounts = sorted(r["mebleg"] for r in rows)
    assert amounts == [4000.0, 8532.0, 31000.0]
    assert all(r["voen"] is None for r in rows)  # VÖEN yoxdur → uyğunsuz
    ads = {r["ad"] for r in rows}
    assert ads == {"HAFİZ ASLANLI", "BUBBLE GUM BAKU MMC"}


def test_parse_bank_statement_rabitabank():
    raw = (FIXTURES / "RABİTƏ BANK ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    # 279 sətirdən 129 debit; 76 VÖEN-siz (56 komissiya + 17 HÖP + 3 şəxs)
    # skip, 149 kredit daxil edilmir → 53 debit sətri.
    assert len(rows) == 53
    assert all(r["voen"] and re.fullmatch(r"\d{7,12}", r["voen"]) for r in rows)
    assert round(sum(r["mebleg"] for r in rows), 2) == 197296.38
    assert sum(1 for r in rows if r["voen"] == "1306715831") == 10  # öz VÖEN

    first = next(r for r in rows if r["voen"] == "1405631661")
    assert first["ad"] == "Digital Classifieds MMC"
    assert first["mebleg"] == 127.12
    assert first["tarix"] == "01.07.2025"

    # ad "Qarşı tərəfin adı" (E) sütunundan oxunur
    assert any(r["ad"] == "PARROTIAS MMC" for r in rows)

    # VÖEN-siz daxili əməliyyatlar (komissiya/HÖP) skip olunur
    assert all("komissiya" not in r["ad"].lower() for r in rows)
    # Kredit sətirləri (pul girişi) daxil edilmir — 1304456331 yalnız kreditdə
    assert all(r["voen"] != "1304456331" for r in rows)
    # Son cəm sətri (289) heç bir qeyd yaratmır
    assert all(r["mebleg"] != 235233.29 for r in rows)


def test_parse_bank_statement_rabitabank_synthetic():
    raw = _rabitabank_xlsx(
        [
            # VÖEN___ ayırıcı (real format) → daxil
            _rabitabank_row(
                "01.07.2025",
                "AZ82UBAZ04564842440060AZN001 VOEN___1405631661",
                "Digital Classifieds MMC",
                127.12, 0,
            ),
            # VOEN_ və VOEN# / VOEN: ayırıcı variasiyaları → daxil
            _rabitabank_row(
                "02.07.2025", "AZ09RBTA0001 VOEN_1503971131", "GLOLYNX MMC", 100, 0,
            ),
            _rabitabank_row(
                "03.07.2025", "AZ09RBTA0002 VOEN: 1700232971", "AR SHURA", 50, 0,
            ),
            _rabitabank_row(
                "04.07.2025", "AZ09RBTA0003 VOEN#1600855111", "ATS FEED MMC", 20, 0,
            ),
            # VÖEN-siz daxili əməliyyat (komissiya) → skip
            _rabitabank_row(
                "05.07.2025", "AZ68RBTA02670130100944414555 ",
                "Haqq və komissiya gəlirləri", 0.5, 0,
            ),
            # Kredit sətri (pul girişi) → skip
            _rabitabank_row(
                "06.07.2025", "AZ09RBTA0004 VOEN___1304456331",
                "MÜŞTƏRİ MMC", 0, 499,
            ),
            # Tarixi boş footer/cəm sətri (mətn yoxdur) → skip
            _rabitabank_row("", "", "", 235233.29, 240603.66),
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 4
    assert {r["voen"] for r in rows} == {
        "1405631661", "1503971131", "1700232971", "1600855111",
    }
    assert all(r["mebleg"] != 235233.29 for r in rows)  # footer skip
    assert all(r["mebleg"] != 499 for r in rows)  # kredit skip
    assert all("komissiya" not in r["ad"].lower() for r in rows)
    by_voen = {r["voen"]: r for r in rows}
    assert by_voen["1405631661"]["ad"] == "Digital Classifieds MMC"


def test_parse_bank_statement_accessbank():
    raw = (FIXTURES / "ACCES BANK ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    # 706 sətirdən 429 debit; 275 boş-Tax ID skip, 277 kredit daxil edilmir
    # → 154 debit sətri.
    assert len(rows) == 154
    assert all(r["voen"] and re.fullmatch(r"\d{7,12}", r["voen"]) for r in rows)
    assert round(sum(r["mebleg"] for r in rows), 2) == 31939.75
    assert all(r["mebleg"] > 0 for r in rows)
    assert rows[0]["tarix"] == "07.04.2025"

    # Tarix: C ("Əməliyyat tarixi / Posting Date"), B ("Hesablaşma") deyil
    found = None
    for sheet_rows in _load_rows(raw):
        found = _detect_bank_header(sheet_rows)
        if found is not None:
            break
    assert found is not None
    _, col_map = found
    assert col_map["date"] == 2  # C sütunu
    assert col_map["voen"] == 8  # I (Tax ID)
    assert col_map["name"] == 7  # H

    # Ad "Göndərənin/Alanın Adı" (H) sütunundan, çoxsətirdən ilk sətir
    assert any(r["ad"] == "Mərkəzi Sos.Sığ.FUİ (yığım)" for r in rows)
    # Kredit sətirləri (Mədaxil) daxil edilmir — 1605270611 yalnız kreditdə
    assert all(r["voen"] != "1605270611" for r in rows)


def test_parse_bank_statement_accessbank_synthetic():
    raw = _accessbank_xlsx(
        [
            # Tax ID + çoxsətirli ad → daxil; ad ilk sətirdir
            _accessbank_row(
                "01.07.2025", "FT1", 127.12,
                "HASANLI-H MMC\n1605270611/\n200004/9900003611",
                "1605270611", "odəniş",
            ),
            # boş Tax ID (komissiya) → skip
            _accessbank_row(
                "02.07.2025", "FT2", 0.9,
                "Ölkədaxili köç komm.haqqı", "", "komm.haqqı",
            ),
            # kredit (Mədaxil>0, Məxaric boş) → skip
            _accessbank_row(
                "03.07.2025", "FT3", "", "MÜŞTƏRİ MMC", "1503971131",
                "daxilolma", medaxil=700,
            ),
            # Tax ID var, ad boş → daxil; ad Təyinatdan
            _accessbank_row(
                "04.07.2025", "FT4", 38, "", "1701911851", "İssizlik odenisi",
            ),
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 2
    by_voen = {r["voen"]: r for r in rows}
    assert set(by_voen) == {"1605270611", "1701911851"}
    assert by_voen["1605270611"]["ad"] == "HASANLI-H MMC"
    assert by_voen["1605270611"]["mebleg"] == 127.12
    assert by_voen["1701911851"]["ad"] == "İssizlik odenisi"
    assert all(r["mebleg"] != 700 for r in rows)  # kredit skip
    assert all(r["voen"] for r in rows)  # boş Tax ID skip

    # Alias əlavələri başqa formatların mappingini pozmur (PAŞA nümunəsi)
    found = _detect_bank_header(
        [
            _PASHA_HEADER,
            _pasha_row("09.06.2026", "CAUCASIAN MMC", "2006687041", "", 5500),
        ]
    )
    assert found is not None
    _, cm = found
    assert cm.get("name") == 2  # Ödəyən/Benefisiar
    assert cm.get("voen") == 5  # VÖEN


def test_parse_bank_statement_bankrespublika():
    raw = (FIXTURES / "BANK RESPUBLİKA ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    # 53 sətirdən 43 debet (I>0, "Debet dövriyyəsi"); 10 kredit atlanır.
    assert len(rows) == 43
    assert round(sum(r["mebleg"] for r in rows), 2) == 96238.80
    assert all(r["voen"] and re.fullmatch(r"\d{7,12}", r["voen"]) for r in rows)
    assert all(r["mebleg"] > 0 for r in rows)

    first = rows[0]
    assert first["voen"] == "1301650741"
    assert first["ad"] == "LUBRISTAR LLC"
    assert first["mebleg"] == 2102.0
    assert first["tarix"] == "06.01.2025"

    # öz VÖEN (Bank Respublika) debit sətirləri saxlanılır (Unibank məntiqi)
    assert sum(1 for r in rows if r["voen"] == "9900001901") == 5

    # Sütun xəritəsi: date=B(1), voen=E(4), debit=I(8), name=D(3)
    found = None
    for sheet_rows in _load_rows(raw):
        found = _detect_bank_header(sheet_rows)
        if found is not None:
            break
    assert found is not None
    _, col_map = found
    assert col_map["date"] == 1
    assert col_map["voen"] == 4
    assert col_map["debit"] == 8
    assert col_map["name"] == 3


def test_parse_bank_statement_bankrespublika_synthetic():
    raw = _bankrespublika_xlsx(
        [
            # debet → daxil
            _bankrespublika_row(
                "06.01.2025", "LUBRISTAR LLC", "1301650741", 0, 2102, "odəniş"
            ),
            # kredit (Debet=0) → skip
            _bankrespublika_row(
                "07.01.2025", "AGM MMC", "7201226221", 2000, 0, "mədaxil"
            ),
            # sıfır/boş → skip
            _bankrespublika_row(
                "08.01.2025", "BOŞ MMC", "1234567890", 0, 0, ""
            ),
            # boş VÖEN debet → skip
            _bankrespublika_row(
                "09.01.2025", "ADSIZ", "", 0, 550, "odəniş"
            ),
            # debet (string məbləğ) → daxil
            _bankrespublika_row(
                "10.01.2025", "F AYDINOGLU MMC", "2005568521", 0, "1726", "odəniş"
            ),
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 2
    by_voen = {r["voen"]: r for r in rows}
    assert set(by_voen) == {"1301650741", "2005568521"}
    assert by_voen["1301650741"]["ad"] == "LUBRISTAR LLC"
    assert by_voen["1301650741"]["mebleg"] == 2102.0
    assert by_voen["2005568521"]["mebleg"] == 1726.0
    assert all(r["mebleg"] > 0 for r in rows)

    # Alias əlavələri Access Bank başlığının mappingini pozmur
    found = _detect_bank_header(
        [
            _ACCESS_HEADER,
            _accessbank_row("01.07.2025", "FT1", 1.0, "X MMC", "1605270611", "t"),
        ]
    )
    assert found is not None
    _, cm = found
    assert cm.get("name") == 7  # H
    assert cm.get("voen") == 8  # I (Tax ID)


def test_parse_bank_statement_described_format():
    # Təsvir olunan format: Tarix / Əməliyyatın adı / Təyinat / Məbləğ (AZN)
    header = ["Tarix", "Əməliyyatın adı", "", "Təyinat", "", "", "", "Məbləğ (AZN)"]
    wb = Workbook()
    ws = wb.active
    ws.append(["meta"])
    ws.append(header)
    ws.append(["01.01.2026", "Ödəniş", "", "Muqavile VOEN: 1403610931", "", "", "", -1000.50])
    ws.append(["02.01.2026", "Daxilolma", "", "Muqavile VÖEN 9900001231", "", "", "", 500.00])
    ws.append(["03.01.2026", "Yekun", "", "CƏM", "", "", "", -9999.00])
    rows = parse_bank_statement(_save(wb))
    assert len(rows) == 1  # mədaxil (müsbət) və Yekun/CƏM sətri atlanır
    assert rows[0]["voen"] == "1403610931"
    assert rows[0]["mebleg"] == 1000.5
    assert rows[0]["tarix"] == "01.01.2026"


def test_split_counter_account():
    assert _split_counter_account("AZ76AIIB00000450190087944110 / KAPITAL / 9900003611") == (
        "9900003611",
        "KAPITAL",
    )
    assert _split_counter_account("AZ43AIIB410400G9444652368186 / VƏLİYEV MOBİL MƏHƏMMƏD OĞLU / 3201279492") == (
        "3201279492",
        "VƏLİYEV MOBİL MƏHƏMMƏD OĞLU",
    )
    assert _split_counter_account("AZ95AIIB410450D9441426 / FAHRADOV TALA T") == (
        None,
        "FAHRADOV TALA T",
    )
    assert _split_counter_account("") == (None, "")


# ===== VÖEN uzlaşması =====


def test_match_by_voen():
    gelen = [
        {"voen": "1111111111", "ad": "SUPPLIER A", "tarix": "01.01.2026", "seriya": "MT", "nomre": "1", "esas": 1000.0, "edv": 180.0},
        {"voen": "1111111111", "ad": "SUPPLIER A", "tarix": "02.01.2026", "seriya": "MT", "nomre": "2", "esas": 500.0, "edv": 90.0},
        {"voen": "2222222222", "ad": "SUPPLIER B", "tarix": "03.01.2026", "seriya": "MT", "nomre": "3", "esas": 200.0, "edv": 36.0},
    ]
    bank = [
        {"voen": "1111111111", "ad": "SUPPLIER A", "tarix": "10.01.2026", "mebleg": 1000.0},
        {"voen": "9999999999", "ad": "UNKNOWN", "tarix": "11.01.2026", "mebleg": 55.0},
    ]
    depozit = [
        {"voen": "1111111111", "ad": "SUPPLIER A", "tarix": "12.01.2026", "mexaric": 270.0},
        {"voen": "8888888888", "ad": "OTHER", "tarix": "13.01.2026", "mexaric": 40.0},
    ]
    result = match_by_voen(gelen, depozit, bank)
    assert len(result["suppliers"]) == 2
    a = result["suppliers"][0]  # SUPPLIER A
    assert a["voen"] == "1111111111"
    assert a["hesablanib_esas"] == 1500.0
    assert a["hesablanib_edv"] == 270.0
    assert a["odenilib_esas"] == 1000.0
    assert a["odenilib_edv"] == 270.0
    assert a["qaliq_esas"] == 500.0
    assert a["qaliq_edv"] == 0.0
    assert len(a["qaimeler"]) == 2

    unmatched = {(u["menbe"], u["voen"]) for u in result["unmatched"]}
    assert unmatched == {("Bank", "9999999999"), ("Depozit", "8888888888")}


# ===== Sheet adı utils =====


def test_sanitize_sheet_name():
    assert sanitize_sheet_name('"ADIS" MMC') == "ADIS MMC"
    assert sanitize_sheet_name("A/B:C?D*E[F]G") == "ABCDEFG"
    assert sanitize_sheet_name("  çox   boşluq  ") == "çox boşluq"
    assert sanitize_sheet_name("///") == "Tedarukcu"


def test_sanitize_sheet_name_truncates_31():
    name = "A" * 50
    assert len(sanitize_sheet_name(name)) == 50  # sanitize özü kəsmir


def test_unique_sheet_name_collision():
    used = set()
    n1 = unique_sheet_name("Eyni Ad", "1234567890", used)
    n2 = unique_sheet_name("Eyni Ad", "1234567890", used)
    assert n1 != n2
    assert n2.endswith("7890")  # VÖEN son 4 rəqəmi


# ===== Workbook struktur testləri =====


def _sample_match():
    return {
        "suppliers": [
            {
                "voen": "1111111111",
                "ad": "SUPPLIER A",
                "qaimeler": [
                    {"tarix": "02.01.2026", "seriya": "MT2601", "nomre": "2", "esas": 500.0, "edv": 90.0},
                    {"tarix": "01.01.2026", "seriya": "MT2601", "nomre": "1", "esas": 1000.0, "edv": 180.0},
                ],
                "odenis_esas": [
                    {"tarix": "10.01.2026", "mebleg": 1000.0},
                    {"tarix": "20.01.2026", "mebleg": 800.0},
                ],
                "odenis_edv": [{"tarix": "12.01.2026", "mebleg": 270.0}],
                "hesablanib_esas": 1500.0,
                "hesablanib_edv": 270.0,
                "odenilib_esas": 1800.0,
                "odenilib_edv": 270.0,
                "qaliq_esas": -300.0,
                "qaliq_edv": 0.0,
            },
            {
                "voen": "2222222222",
                "ad": "SUPPLIER B",
                "qaimeler": [
                    {"tarix": "03.01.2026", "seriya": "MT2601", "nomre": "3", "esas": 200.0, "edv": 36.0},
                ],
                "odenis_esas": [],
                "odenis_edv": [],
                "hesablanib_esas": 200.0,
                "hesablanib_edv": 36.0,
                "odenilib_esas": 0.0,
                "odenilib_edv": 0.0,
                "qaliq_esas": 200.0,
                "qaliq_edv": 36.0,
            },
        ],
        "unmatched": [
            {"voen": "9999999999", "ad": "UNKNOWN", "tarix": "11.01.2026", "menbe": "Bank", "mebleg": 55.0},
        ],
    }


def test_build_workbook_main_sheet():
    out = build_workbook(_sample_match())
    wb = load_workbook(io.BytesIO(out))
    assert MAIN_SHEET_NAME in wb.sheetnames
    assert UNMATCHED_SHEET_NAME in wb.sheetnames
    assert len(wb.sheetnames) == 4  # main + 2 tədarükçü + uyğunsuz

    main = wb[MAIN_SHEET_NAME]
    assert main["B1"].value == "KREDİTORLAR"
    assert main["D1"].value == "Hesablanıb"
    assert main["F1"].value == "Ödənilib"
    assert main["H1"].value == "Qalıq"
    assert main["J1"].value == "Qeyd"
    assert main["D2"].value == "Əsas"
    assert main["E2"].value == "ƏDV"

    # A sütunu: 1, 2 (sətir 3,4)
    assert main["A3"].value == 1
    assert main["A4"].value == 2

    # CƏMİ sətri: sətir 5
    assert main["B5"].value == "CƏMİ"
    assert main["D5"].value == "=SUM(D3:D4)"


def test_build_workbook_hyperlink_and_formula():
    out = build_workbook(_sample_match())
    wb = load_workbook(io.BytesIO(out))
    main = wb[MAIN_SHEET_NAME]

    b3 = main["B3"]
    assert b3.value == "SUPPLIER A"
    assert b3.hyperlink is not None
    assert b3.hyperlink.location == "'SUPPLIER A'!A1"

    # D3 = tədarükçü sheet-in SUM sətrinə istinad (E sütunu; 4 data sətiri → SUM=10)
    assert main["D3"].value == "='SUPPLIER A'!E10"
    assert main["I3"].value == "='SUPPLIER A'!J10"

    # Qeyd flag: qalıq əsas mənfi → Artıq ödəmə
    assert main["J3"].value == "Artıq ödəmə: Əsas"
    assert main["J4"].value in (None, "")


def test_build_workbook_supplier_sheet_layout():
    out = build_workbook(_sample_match())
    wb = load_workbook(io.BytesIO(out))
    ws = wb["SUPPLIER A"]

    assert ws["A1"].value == "SUPPLIER A"
    assert ws["A2"].value == 1111111111
    assert ws["K2"].hyperlink.location == MAIN_SHEET_NAME + "!A1"

    assert ws["A4"].value == "Date"
    assert ws["B4"].value == "Elektron Qaimə"
    assert ws["D4"].value == "Ödəniş tarixi"
    assert ws["E5"].value == "Əsas"
    assert ws["F5"].value == "ƏDV"

    # 2 qaimə + (2 ödəniş - 1 birləşmiş) = 3 data sətiri, sonra SUM
    # qaimələr tarixə görə sıralanır: 01.01.2026, 02.01.2026
    assert ws["A6"].value == "01.01.2026"  # ilk qaimə
    assert ws["B6"].value == "MT2601"
    assert ws["C6"].value == "1"
    assert ws["E6"].value == 1000.0
    # ilk ödəniş (10.01.2026, 1000) birləşir
    assert ws["D6"].value == "10.01.2026"
    assert ws["G6"].value == 1000.0
    assert ws["I6"].value == "=E6-G6"
    assert ws["J6"].value == "=F6-H6"

    # ikinci qaimə
    assert ws["A7"].value == "02.01.2026"
    assert ws["E7"].value == 500.0
    # qalan ödənişlər ayrı sətir: 20.01.2026 (esas 800), 12.01.2026 (edv 270)
    assert ws["D8"].value == "12.01.2026"
    assert ws["H8"].value == 270.0
    assert ws["D9"].value == "20.01.2026"
    assert ws["G9"].value == 800.0

    # SUM sətri (10)
    assert ws["E10"].value == "=SUM(E6:E9)"
    assert ws["F10"].value == "=SUM(F6:F9)"


def test_build_workbook_unmatched_sheet():
    out = build_workbook(_sample_match())
    wb = load_workbook(io.BytesIO(out))
    ws = wb[UNMATCHED_SHEET_NAME]
    assert ws["A1"].value == "VÖEN"
    assert ws["D1"].value == "Mənbə"
    assert ws["A2"].value == "9999999999"
    assert ws["D2"].value == "Bank"
    assert ws["E2"].value == 55.0


def test_build_workbook_main_row_hyperlinks():
    out = build_workbook(_sample_match())
    wb = load_workbook(io.BytesIO(out))
    main = wb[MAIN_SHEET_NAME]
    for col in ("B", "D", "E", "F", "G", "H", "I"):
        cell = main[f"{col}3"]
        assert cell.hyperlink is not None, col
        assert cell.hyperlink.location == "'SUPPLIER A'!A1", col


def test_build_workbook_supplier_back_hyperlinks():
    out = build_workbook(_sample_match())
    wb = load_workbook(io.BytesIO(out))
    ws = wb["SUPPLIER A"]
    assert ws["A1"].hyperlink is not None
    assert ws["A1"].hyperlink.location == MAIN_SHEET_NAME + "!A1"
    assert ws["K2"].value == "← KREDİTORLAR"
    assert ws["K2"].hyperlink.location == MAIN_SHEET_NAME + "!A1"


def test_build_workbook_styling():
    out = build_workbook(_sample_match())
    wb = load_workbook(io.BytesIO(out))
    main = wb[MAIN_SHEET_NAME]

    h = main["D1"]
    assert h.font.bold is True
    assert h.alignment.horizontal == "center"
    assert h.fill.fill_type == "solid"

    m = main["D3"]
    assert m.number_format == "#,##0.00"
    assert m.alignment.horizontal == "right"

    assert main["B3"].alignment.horizontal == "left"
    assert main["B3"].border.left.style == "thin"

    assert main.column_dimensions["B"].width is not None
    assert main.column_dimensions["B"].width >= 8

    ws = wb["SUPPLIER A"]
    assert ws["A1"].hyperlink is not None
    assert ws["I6"].alignment.horizontal == "right"
    assert ws["A6"].alignment.horizontal == "left"
    assert ws["A6"].border.bottom.style == "thin"


# ===== Real nümunə fayllarla inteqrasiya =====


def test_integration_real_files():
    gelen_raw = (FIXTURES / "GƏLƏN.xlsx").read_bytes()
    depozit_raw = (FIXTURES / "DEPOZİT.xlsx").read_bytes()
    bank_raw = (FIXTURES / "KAPİTAL BANK ÇIXARIŞI.xlsx").read_bytes()

    gelen = parse_gelen(gelen_raw)
    depozit = parse_depozit(depozit_raw)
    bank = parse_bank_statement(bank_raw)
    assert len(gelen) == 228
    assert len(depozit) == 122
    assert len(bank) == 50

    matched = match_by_voen(gelen, depozit, bank)
    assert len(matched["suppliers"]) == 57

    out = generate_creditors_workbook(gelen_raw, depozit_raw, bank_raw)
    wb = load_workbook(io.BytesIO(out))
    assert len(wb.sheetnames) == 59  # main + 57 tədarükçü + uyğunsuz
    assert wb.sheetnames[0] == MAIN_SHEET_NAME
    assert wb.sheetnames[-1] == UNMATCHED_SHEET_NAME

    main = wb[MAIN_SHEET_NAME]
    assert main.max_row == 60  # başlıq 2 + 57 tədarükçü + CƏMİ
    assert main["B3"].hyperlink is not None
    # CƏMİ sətri bütün tədarükçüləri əhatə edir
    assert main[f"D{main.max_row}"].value == "=SUM(D3:D59)"
    # hər tədarükçü sheet-də geri hiperlink var
    first_supplier_sheet = wb[wb.sheetnames[1]]
    assert first_supplier_sheet["K2"].hyperlink.location == MAIN_SHEET_NAME + "!A1"


# ===== Endpoint (HTTP) =====

XLSX_MIME_CR = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


async def test_creditors_generate_endpoint(client):
    gelen_raw = (FIXTURES / "GƏLƏN.xlsx").read_bytes()
    depozit_raw = (FIXTURES / "DEPOZİT.xlsx").read_bytes()
    bank_raw = (FIXTURES / "KAPİTAL BANK ÇIXARIŞI.xlsx").read_bytes()

    resp = await client.post(
        "/api/v1/import/creditors/generate",
        files={
            "gelen": ("GƏLƏN.xlsx", gelen_raw, XLSX_MIME_CR),
            "depozit": ("DEPOZİT.xlsx", depozit_raw, XLSX_MIME_CR),
            "bank": ("BANK.xlsx", bank_raw, XLSX_MIME_CR),
        },
    )
    assert resp.status_code == 200
    assert "spreadsheetml" in resp.headers["content-type"]
    assert "attachment" in resp.headers["content-disposition"]

    wb = load_workbook(io.BytesIO(resp.content))
    assert wb.sheetnames[0] == MAIN_SHEET_NAME
    assert wb.sheetnames[-1] == UNMATCHED_SHEET_NAME


async def test_creditors_generate_invalid_file(client):
    resp = await client.post(
        "/api/v1/import/creditors/generate",
        files={
            "gelen": ("g.xlsx", b"notxlsx", XLSX_MIME_CR),
            "depozit": ("d.xlsx", b"notxlsx", XLSX_MIME_CR),
            "bank": ("b.xlsx", b"notxlsx", XLSX_MIME_CR),
        },
    )
    assert resp.status_code == 422
