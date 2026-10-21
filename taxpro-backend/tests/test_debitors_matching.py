# -*- coding: utf-8 -*-
"""DEBİTORLAR (GEDƏN / mədaxil / KREDİT) — paritet və xüsusi qaydalar testləri.

Real fixture-lar: GEDƏN.xlsx (göndərilən qaimələr), DEPOZİT.xlsx (mədaxil),
KAPİTAL/PAŞA/ABB/Yapı kredi bank çıxarışları.
"""

import io
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app.services.deterministic import creditors as C
from app.services.deterministic.creditors import (
    BANK_SUMMARY_SHEET_NAME,
    DEBITOR_MAIN_SHEET_NAME,
    REVERSALS_SHEET_NAME,
    REMOVED_INVOICES_SHEET_NAME,
    UNMATCHED_SHEET_NAME,
    _allocate_supplier,
    _extract_invoice_refs,
    _filter_invoices,
    generate_creditors_workbook,
    generate_debitors_workbook,
    parse_bank_statement,
    parse_depozit,
    parse_gelen,
)
from tests.test_creditors_matching import (
    _formula_eval,
    _pasha_row,
    _pasha_xlsx,
)

FIXTURES = Path(__file__).parent / "fixtures"
GEDEN = (FIXTURES / "GEDƏN.xlsx").read_bytes()
DEPOZIT = (FIXTURES / "DEPOZİT.xlsx").read_bytes()

# Task-ın "KAPİTAL_2026 + PASHA_2026 + YAPİKREDİ_2026 + ABB_2026" dəstinə
# uyğun mövcud fixture-lar.
BANKS4 = [
    ("KAPİTAL BANK ÇIXARIŞI.xlsx", (FIXTURES / "KAPİTAL BANK ÇIXARIŞI.xlsx").read_bytes()),
    ("PAŞA BANK ÇIXARIŞI.xlsx", (FIXTURES / "PAŞA BANK ÇIXARIŞI.xlsx").read_bytes()),
    ("ABB ÇIXARIŞI.xlsx", (FIXTURES / "ABB ÇIXARIŞI.xlsx").read_bytes()),
    ("Yapı kredi.xlsx", (FIXTURES / "Yapı kredi.xlsx").read_bytes()),
]

BANK_FIXTURES = [
    "ABB ÇIXARIŞI.xlsx", "ACCES BANK ÇIXARIŞI.xlsx", "ATB BANK ÇIXARIŞI.xlsx",
    "BANK OF BAKU ÇIXARIŞI.xlsx", "BANK RESPUBLİKA ÇIXARIŞI.xlsx",
    "KAPİTAL BANK ÇIXARIŞI.xlsx", "PAŞA BANK ÇIXARIŞI.xlsx",
    "RABİTƏ BANK ÇIXARIŞI.xlsx", "TURAN BANK ÇIXARIŞI.xlsx",
    "UNİBANK ÇIXARIŞI.xlsx", "VTB BANK ÇIXARIŞI.xlsx",
    "YAPIKREDİ BANK ÇIXARIŞI.xlsx", "YAPI KREDİ (YENİ).xlsx",
    "YELO BANK ÇIXARIŞI.xlsx", "ZİRAAT BANK ÇIXARIŞI.xlsx",
]

_SPECIAL = {
    DEBITOR_MAIN_SHEET_NAME, UNMATCHED_SHEET_NAME, REVERSALS_SHEET_NAME,
    REMOVED_INVOICES_SHEET_NAME, BANK_SUMMARY_SHEET_NAME,
}


def _wb(out):
    return load_workbook(io.BytesIO(out))


def _sheet_by_prefix(wb, prefix):
    return wb[next(n for n in wb.sheetnames if prefix.upper() in n.upper())]


def _cell_for_ref(ws, seriya, nomre):
    """Qaimə sətrinin nömrəsini qaytarır (A/B/C = tarix/seriya/nömre)."""
    for r in range(6, ws.max_row + 1):
        if ws.cell(row=r, column=2).value == seriya and str(
            ws.cell(row=r, column=3).value
        ) == nomre:
            return r
    return None


# ===== A/B: qaimə süzgəci =====


def test_geden_qaime_sayi_ve_cem():
    kept, removed, warnings = _filter_invoices(parse_gelen(GEDEN))
    assert len(kept) == 42
    assert round(sum(g["esas"] for g in kept), 2) == 1473790.37
    assert round(sum(g["edv"] for g in kept), 2) == 265282.27
    # 3 AA (hamısı Passiv) çıxarılır
    assert len(removed) == 3
    assert all(r["sebeb"] == "Passiv qaimə" for r in removed)
    assert {r["seriya"] for r in removed} == {"AA2602", "AA2603", "AA2604"}
    assert {r["voen"] for r in removed} == {"1605283051"}
    # MQ məlum seriyadır → "Naməlum seriya" xəbərdarlığı YOXDUR
    assert not any("Naməlum qaimə seriyası" in w for w in warnings)


def test_mq_qaytarma_qaimesi_saxlanilir():
    out = generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4)
    wb = _wb(out)
    # İRON AVTO: MQ2607/12709447 — tam ödənilib
    iron = _sheet_by_prefix(wb, "İRON AVTO")
    r = _cell_for_ref(iron, "MQ2607", "12709447")
    assert r is not None
    assert iron.cell(row=r, column=5).value == 560
    assert iron.cell(row=r, column=6).value == 100.8
    assert "Qaytarma qaiməsi" in str(iron.cell(row=r, column=12).value)
    # bank 31.07/560 (PAŞA) + depozit 30.07/100.80 → Qalıq 0
    sr = iron.max_row
    assert abs(_formula_eval(iron, f"I{sr}") - 0) < 0.01
    assert abs(_formula_eval(iron, f"J{sr}") - 0) < 0.01
    # ARAZ MAKİNA: MQ2605/11502913 ödənişsiz
    araz = _sheet_by_prefix(wb, "ARAZ MAKİNA")
    r = _cell_for_ref(araz, "MQ2605", "11502913")
    assert r is not None
    assert araz.cell(row=r, column=5).value == 13000
    assert araz.cell(row=r, column=6).value == 2340
    assert "Qaytarma qaiməsi" in str(araz.cell(row=r, column=12).value)
    # MQ üçün xəbərdarlıq yoxdur
    ws = wb[BANK_SUMMARY_SHEET_NAME]
    texts = [str(ws.cell(row=rr, column=1).value or "") for rr in range(1, ws.max_row + 1)]
    assert not any("Naməlum qaimə seriyası" in t for t in texts)


# ===== E: DEBİTORLAR bölüşdürmə / depozit mədaxil =====


def test_depozit_medaxil_cemi():
    out = generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4)
    wb = _wb(out)
    hsum = 0.0
    for nm in wb.sheetnames:
        if nm in _SPECIAL:
            continue
        ws = wb[nm]
        for r in range(6, ws.max_row):
            v = ws.cell(row=r, column=8).value
            if isinstance(v, (int, float)):
                hsum += v
    assert round(hsum, 2) == 150104.00

    # CAUCASIAN 2089 və ƏLİNCƏ QALA 550 cütləri "Nəzərə alınmayan"-da
    rev = wb[REVERSALS_SHEET_NAME]
    pairs = sorted(
        (rev.cell(row=r, column=1).value, rev.cell(row=r, column=6).value)
        for r in range(2, rev.max_row + 1)
        if rev.cell(row=r, column=8).value == "Səhv köçürmə — qarşılıqlı ləğv olundu"
    )
    assert ("2006687041", 2089) in pairs
    assert ("6001731321", 550) in pairs
    # FASTTECH (müştəri deyil) Uyğunsuz-da
    unm = wb[UNMATCHED_SHEET_NAME]
    voens = {str(unm.cell(row=r, column=1).value) for r in range(2, unm.max_row + 1)}
    assert "1006665241" in voens


def test_famn_avans_zenciri():
    out = generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4)
    wb = _wb(out)
    famn = _sheet_by_prefix(wb, "FAMN")
    expectations = {
        ("MT2604", "12637200"): 1800.0,
        ("MT2604", "12638199"): 360.0,
        ("MT2604", "12638742"): 2700.0,
    }
    for (ser, nom), amount in expectations.items():
        r = _cell_for_ref(famn, ser, nom)
        assert r is not None
        # alt ödəniş sətrində H (ƏDV) = amount, qeyd "Avans zənciri"
        found = False
        rr = r + 1
        while rr < famn.max_row and not (
            isinstance(famn.cell(row=rr, column=2).value, str)
            and famn.cell(row=rr, column=2).value.strip()
        ):
            h = famn.cell(row=rr, column=8).value
            q = famn.cell(row=rr, column=12).value or ""
            if isinstance(h, (int, float)) and abs(h - amount) < 0.01 and "Avans zənciri" in q:
                found = True
            rr += 1
        assert found, (ser, nom, amount)


# ===== İMA-ENERGY: bank kredit + çox istinad bölünməsi =====


def test_ima_energy_bank_kredit():
    out = generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4)
    wb = _wb(out)
    ima = _sheet_by_prefix(wb, "İMA-ENERGY")
    kapital = pasa = 0.0
    for r in range(6, ima.max_row):
        g = ima.cell(row=r, column=7).value
        if not isinstance(g, (int, float)):
            continue
        bank = str(ima.cell(row=r, column=11).value or "")
        if "KAPİTAL" in bank:
            kapital += g
        if "PAŞA" in bank:
            pasa += g
    # KAPİTAL: 100000 (13.03) + 70000 (13.05); PAŞA: 150000 (10.07)
    assert round(kapital, 2) == 170000.00
    assert round(pasa, 2) == 150000.00


def test_ima_energy_uc_istinad_split():
    """10.07/27000 depozit ödənişi 3 MT2601 qaiməyə bölünür.

    QEYD: canlı fixture-da istinadlar TAM 12 rəqəmdir
    (MT260110651608 / MT260110651087 / MT260110649695) — task təsvirindəki
    "sonuncu 11 rəqəm" truncation real faylda yoxdur.
    """
    out = generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4)
    wb = _wb(out)
    ima = _sheet_by_prefix(wb, "İMA-ENERGY")
    for ser, nom in (
        ("MT2601", "10651608"),
        ("MT2601", "10651087"),
        ("MT2601", "10649695"),
    ):
        r = _cell_for_ref(ima, ser, nom)
        assert r is not None, (ser, nom)
        has_split = False
        rr = r + 1
        while rr < ima.max_row and not (
            isinstance(ima.cell(row=rr, column=2).value, str)
            and ima.cell(row=rr, column=2).value.strip()
        ):
            h = ima.cell(row=rr, column=8).value
            q = str(ima.cell(row=rr, column=12).value or "")
            if isinstance(h, (int, float)) and "bölündü" in q:
                has_split = True
            rr += 1
        assert has_split, (ser, nom)


def test_amazone_intec_rapid_tek_qaime_istinadi():
    out = generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4)
    wb = _wb(out)
    checks = {
        "AMAZONE": [("MT2607", "10331130"), ("MT2606", "10603608"), ("MT2605", "10774557")],
        "İNTEC": [("MT2605", "10693177"), ("MT2607", "11633826")],
        "RAPID": [("MT2601", "10609924")],
    }
    for prefix, refs in checks.items():
        ws = _sheet_by_prefix(wb, prefix)
        for ser, nom in refs:
            assert _cell_for_ref(ws, ser, nom) is not None, (prefix, ser, nom)


# ===== E: tie-break (istiqamətdən asılı) =====


def test_tie_break_debitor_en_kohne_kreditor_en_yaxin():
    def _make():
        return {
            "voen": "1", "ad": "X",
            "qaimeler": [
                {"tarix": "01.01.2026", "tarix_dt": date(2026, 1, 1),
                 "seriya": "MT2601", "nomre": "11111111", "esas": 5000.0, "edv": 0.0},
                {"tarix": "01.06.2026", "tarix_dt": date(2026, 6, 1),
                 "seriya": "MT2601", "nomre": "22222222", "esas": 5000.0, "edv": 0.0},
            ],
            "odenis_esas": [
                {"tarix": "01.07.2026", "tarix_dt": date(2026, 7, 1),
                 "mebleg": 5000.0, "text": ""},
            ],
            "odenis_edv": [],
            "hesablanib_esas": 10000.0, "hesablanib_edv": 0.0,
            "odenilib_esas": 5000.0, "odenilib_edv": 0.0,
            "qaliq_esas": 5000.0, "qaliq_edv": 0.0,
        }

    deb = _allocate_supplier(_make(), direction="debitor")
    # Ən köhnə (tarix ≤ ödəniş): 01.01.2026
    assert any(g["q"]["nomre"] == "11111111" and g["payments"] for g in deb["groups"])
    assert all(not g["payments"] for g in deb["groups"] if g["q"]["nomre"] == "22222222")

    kre = _allocate_supplier(_make(), direction="kreditor")
    # Ödəniş tarixinə ən yaxın: 01.06.2026
    assert any(g["q"]["nomre"] == "22222222" and g["payments"] for g in kre["groups"])
    assert all(not g["payments"] for g in kre["groups"] if g["q"]["nomre"] == "11111111")


# ===== H: sintetik (prefiks, status, avans zənciri bağlı) =====


def test_prefiks_tamamlama_sintetik():
    known = {("MT2601", "10649695")}
    assert _extract_invoice_refs("MT26011064969", {}, known) == [("MT2601", "10649695")]
    # Çox namizəd → həll olunmur
    assert _extract_invoice_refs(
        "MT26011064969", {}, {("MT2601", "10649695"), ("MT2601", "10649699")}
    ) == []
    # 7 rəqəmdən az kəsilmiş → həll olunmur
    assert _extract_invoice_refs("MT2602114", {}, {("MT2602", "11462198")}) == []


def test_status_tesdiq_gozleyir_ve_legve_xeberdarligi():
    rows = [
        {"voen": "1", "ad": "X", "tarix": "01.01.2026", "tarix_dt": date(2026, 1, 1),
         "seriya": "MT2601", "nomre": "1", "veziyyet": "Təsdiq gözləyir",
         "nov": "", "esas": 100.0, "edv": 18.0},
        {"voen": "2", "ad": "Y", "tarix": "01.01.2026", "tarix_dt": date(2026, 1, 1),
         "seriya": "MT2601", "nomre": "2", "veziyyet": "Ləğvə göndərildi",
         "nov": "", "esas": 50.0, "edv": 9.0},
    ]
    kept, removed, warnings = _filter_invoices(rows)
    assert len(kept) == 2 and removed == []
    assert any("Təsdiq gözləyən" in w for w in warnings)
    assert any("Ləğvə göndərilmiş" in w for w in warnings)


def test_avans_zenciri_aciqdir_kreditorda_baglidir():
    """KREDİTORLAR default BAĞLI; açar açıq olsa AA ödənişləri MT-yə keçir."""
    gelen = (FIXTURES / "GƏLƏN.xlsx").read_bytes()
    depozit = DEPOZIT
    banks = [
        ("Kapital.xlsx", (FIXTURES / "KAPİTAL BANK ÇIXARIŞI.xlsx").read_bytes()),
        ("Pasha.xlsx", (FIXTURES / "PAŞA BANK ÇIXARIŞI.xlsx").read_bytes()),
        ("Yapi.xlsx", (FIXTURES / "Yapı kredi.xlsx").read_bytes()),
        ("ABB.xlsx", (FIXTURES / "ABB ÇIXARIŞI.xlsx").read_bytes()),
    ]

    # 1) Default: açar bağlı → AZ BUILD 9000 "Nəzərə alınmayan"-da qalır
    assert C._AVANS_ZENCIRI_AKTIV == {"debitor": True, "kreditor": False}
    wb = _wb(generate_creditors_workbook(gelen, depozit, banks))
    rev = wb[REVERSALS_SHEET_NAME]
    assert any(
        rev.cell(row=r, column=6).value == 9000
        and "Passiv/avans" in str(rev.cell(row=r, column=8).value)
        for r in range(2, rev.max_row + 1)
    )
    az = _sheet_by_prefix(wb, "AZ BUILD")
    r = _cell_for_ref(az, "MT2604", "11338143")
    assert r is not None
    # alt sətirlərdə 9000 bank ödənişi YOXDUR (açar bağlı)
    has_9000 = any(
        az.cell(row=rr, column=7).value == 9000 for rr in range(r + 1, az.max_row)
    )
    assert not has_9000

    # 2) Açar KREDİTORLAR üçün də açılır → 9000 MT2604/11338143-ə keçir
    old = C._AVANS_ZENCIRI_AKTIV
    try:
        C._AVANS_ZENCIRI_AKTIV = {"debitor": True, "kreditor": True}
        wb2 = _wb(generate_creditors_workbook(gelen, depozit, banks))
        az2 = _sheet_by_prefix(wb2, "AZ BUILD")
        r2 = _cell_for_ref(az2, "MT2604", "11338143")
        assert r2 is not None
        found = False
        rr = r2 + 1
        while rr < az2.max_row and not (
            isinstance(az2.cell(row=rr, column=2).value, str)
            and az2.cell(row=rr, column=2).value.strip()
        ):
            g = az2.cell(row=rr, column=7).value
            q = str(az2.cell(row=rr, column=12).value or "")
            if isinstance(g, (int, float)) and g == 9000 and "Avans zənciri" in q:
                found = True
            rr += 1
        assert found
    finally:
        C._AVANS_ZENCIRI_AKTIV = old


# ===== G: invariantlar (DEBİTORLAR) =====


def _debit_wb():
    return _wb(generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4))


def test_debitor_qaliq_invariant_her_varaq():
    wb = _debit_wb()
    for nm in wb.sheetnames:
        if nm in _SPECIAL:
            continue
        ws = wb[nm]
        sr = ws.max_row
        ev = lambda c: _formula_eval(ws, f"{c}{sr}")
        assert abs(ev("I") - (ev("E") - ev("G"))) < 0.01, (nm, "I")
        assert abs(ev("J") - (ev("F") - ev("H"))) < 0.01, (nm, "J")


def test_debitor_qaima_odenis_sutun_qaydasi():
    """Qaimə sətirində G/H boş; ödəniş sətirində E/F boş."""
    wb = _debit_wb()
    for nm in wb.sheetnames:
        if nm in _SPECIAL:
            continue
        ws = wb[nm]
        for r in range(6, ws.max_row):
            ser = ws.cell(row=r, column=2).value
            is_invoice = isinstance(ser, str) and ser.strip()
            if is_invoice:
                assert ws.cell(row=r, column=7).value in (None, "")
                assert ws.cell(row=r, column=8).value in (None, "")
            else:
                if ws.cell(row=r, column=1).value == "Qaimə ilə əlaqələndirilməyən ödənişlər":
                    continue
                assert ws.cell(row=r, column=5).value in (None, "")
                assert ws.cell(row=r, column=6).value in (None, "")


@pytest.mark.parametrize("bank_name", BANK_FIXTURES)
def test_butun_banklar_debitor_invariant(bank_name):
    raw = (FIXTURES / bank_name).read_bytes()
    out = generate_debitors_workbook(GEDEN, DEPOZIT, [(bank_name, raw)])
    wb = _wb(out)
    assert wb.sheetnames[0] == DEBITOR_MAIN_SHEET_NAME
    # Bank xülasə invariantı: Uyğunlaşan+Uyğunsuz+Adsız+Nəzərə alınmayan == Ümumi
    ws = wb[BANK_SUMMARY_SHEET_NAME]
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value in (None, "ÜMUMİ"):
            continue
        cem = ws.cell(row=r, column=4).value or 0
        total = (
            (ws.cell(row=r, column=5).value or 0)
            + (ws.cell(row=r, column=6).value or 0)
            + (ws.cell(row=r, column=7).value or 0)
            + (ws.cell(row=r, column=10).value or 0)
        )
        assert abs(total - cem) < 0.01, (bank_name, r, total, cem)


def test_debitor_basliq_ve_back_link_pariteti():
    wb = _debit_wb()
    main = wb[DEBITOR_MAIN_SHEET_NAME]
    assert main["B1"].value == "DEBİTORLAR"
    sup_name = next(
        n for n in wb.sheetnames[1:] if n not in _SPECIAL
    )
    sup = wb[sup_name]
    assert sup["K2"].value == "← DEBİTORLAR"
    assert sup["K2"].hyperlink.location == DEBITOR_MAIN_SHEET_NAME + "!A1"
    # Vərəqdə "KREDİTORLAR" mətninə bağlı qalıq yoxdur
    texts = {main["C1"].value, main["B1"].value}
    assert not any("KREDİTOR" in str(t or "") for t in texts)


def test_bank_debet_debitorlarda_sayilmir():
    """Bank DEBET (müştəriyə geri ödəmə) DEBİTORLAR-da ödəniş sayılmır."""
    # PAŞA 03.07/1350 İRON AVTO Məxaric (debet) — İRON vərəqində G kimi görünməməli
    out = generate_debitors_workbook(GEDEN, DEPOZIT, BANKS4)
    wb = _wb(out)
    iron = _sheet_by_prefix(wb, "İRON AVTO")
    # MQ2607 560-dan başqa 1350 heç bir yerdə yoxdur
    for r in range(6, iron.max_row):
        for c in (7, 8):
            v = iron.cell(row=r, column=c).value
            assert v != 1350, (r, c)
