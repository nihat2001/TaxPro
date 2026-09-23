import io
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from app.services.deterministic.creditors import (
    MAIN_SHEET_NAME,
    UNMATCHED_SHEET_NAME,
    _clean_counterparty_name,
    _extract_voen_and_name,
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
            _pasha_row("09.06.2026", "Pasha Bank", "", "", 1000),  # bank haqqı (boş VÖEN)
            _pasha_row("10.06.2026", "IMA ENERGY MMC", "1601005511", 100000, ""),  # mədaxil
            _pasha_row("11.06.2026", "CENTRO MMC", "1202255201", "", 2600),
        ]
    )
    rows = parse_bank_statement(raw)
    assert len(rows) == 3  # mədaxil sətri atlanır
    assert {r["voen"] for r in rows} == {"2006687041", None, "1202255201"}
    by_voen = {r["voen"]: r for r in rows}
    assert by_voen["2006687041"]["mebleg"] == 5500.0
    assert by_voen["2006687041"]["ad"] == "CAUCASIAN MMC"
    assert by_voen["1202255201"]["mebleg"] == 2600.0


def test_parse_bank_statement_pasha_real():
    raw = (FIXTURES / "PAŞA BANK ÇIXARIŞI.xlsx").read_bytes()
    rows = parse_bank_statement(raw)
    assert len(rows) == 67  # 67 debit sətri (balans/summary atlanır)

    totals: dict[str, float] = {}
    empty_voen = 0
    for r in rows:
        if r["voen"]:
            totals[r["voen"]] = totals.get(r["voen"], 0.0) + r["mebleg"]
        else:
            empty_voen += 1

    assert round(totals["2006687041"], 2) == 23371.72
    assert round(totals["1202255201"], 2) == 14643.0
    assert empty_voen == 24  # Pasha Bank daxili haqq sətirləri
    # DÖVRÜN SONUNA BALANS sətri (349999.79) heç bir yerdə debit kimi düşməməlidir
    assert all(r["mebleg"] != 349999.79 for r in rows)


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
