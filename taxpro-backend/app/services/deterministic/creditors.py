"""Kreditorlar/Debitorlar — "Gələn" (Kreditorlar) modulu.

3 mənbə faylı (GƏLƏN.xlsx, DEPOZİT.xlsx, BANK ÇIXARIŞI.xlsx) deterministik
alqoritmlə VÖEN üzrə uzlaşdırılır və TƏK nəticə workbook-u yaranır:

  - "KREDİTORLAR" əsas sheet (hər tədarükçü bir sətir, hiperlink + formula)
  - hər tədarükçü üçün ayrı sheet (qaimə/ödəniş sətirləri + SUM formulları)
  - "Uyğunsuz ödənişlər" sheet (GƏLƏN-də olmayan VÖEN-ə ödənişlər)

V1: stateless — DB-yə yazılmır; yalnız .xlsx generasiya olunur.
BANK ÇIXARIŞI parseri dinamik auto-detection ilə çoxbanklıdır (Kapital,
PAŞA, Yapı Kredi, ABB, Unibank, TuranBank, Rabitəbank, Access Bank,
Bank Respublika, Yelo Bank, Bank of Baku, VTB Bank, ATB Bank, Ziraat Bank;
digər banklar alias siyahıları ilə genişləndirilə bilər).
AI/ML yoxdur.
"""

import difflib
import hashlib
import io
import os
import re
import unicodedata
from datetime import date, datetime
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from app.services.deterministic.excel_parser import normalize_text

# ===== Sheet / sütun sabitləri =====
MAIN_SHEET_NAME = "KREDİTORLAR"
REVERSALS_SHEET_NAME = "Nəzərə alınmayan əməliyyatlar"
REMOVED_INVOICES_SHEET_NAME = "Nəzərə alınmayan qaimələr"
BANK_SUMMARY_SHEET_NAME = "Banklar üzrə xülasə"
# Çoxlu bank çıxarışı: maksimum fayl sayı (endpoint səviyyəsində yoxlanılır)
MAX_BANK_FILES = 20
# Depozit ödənişlərinin "Bank" sütununda göstərilən mənbə etiketi
DEPOZIT_BANK_LABEL = "Depozit (ƏDV)"
# Bank preamble ailələri üçün oxunaqlı etiketlər
_FAMILY_LABELS = {"vtb": "VTB", "atb": "ATB", "ziraat": "Ziraat"}
UNMATCHED_SHEET_NAME = "Uyğunsuz ödənişlər"
_ACIQLAMA = "Mal material satışı"

# ===== Sənəd tipləri (Hesablanıb / Ödənilib sərt ayrılığı) =====
# INVOICE — yalnız GƏLƏN (elektron qaimə): "Hesablanıb" (E/F) formalaşdırır.
# PAYMENT — yalnız Bank çıxarışı + DEPOZİT: "Ödənilib" (G/H) formalaşdırır.
# GƏLƏN heç bir halda birbaşa ödəniş sətri yaratmır; bank/depozit heç bir halda
# Hesablanıb sütununa yazılmır.
DOC_INVOICE = "INVOICE"
DOC_PAYMENT = "PAYMENT"

# Çıxarılmış (passiv/AA) qaiməyə istinad edən ödənişin rejimi:
#   "cixar" (default) — hesabdan çıxarılır, "Nəzərə alınmayan əməliyyatlar"-da.
#   "avans" — istinadsız avans kimi işlənir (tam məbləğ → FIFO → blok), Qeyd-də
#             "Avans (AA qaiməyə istinad)" yazılır.
_CIXARILMIS_QAIME_ODENIS_REJIMI = "cixar"

# Avans zənciri (Bölmə C): çıxarılmış AA avansına istinad edən ödənişin, həmin
# AA-nı "Avans Seriası/Nömrəsi" sütunlarında göstərən MT qaiməyə yazılması.
# İstiqamət üzrə açar: DEBİTORLAR aktivdir, KREDİTORLAR default BAĞLIDIR
# (KREDİTORLAR nəticələri dəyişməsin).
_AVANS_ZENCIRI_AKTIV = {"debitor": True, "kreditor": False}
# (2) rejimi: istinadsız ödənişin məbləği AA-nın komponentinə bərabər və tarixi
# AA tarixinin ±N günündədirsə istinad kimi qəbul olunur.
_AVANS_TARIX_PENCERESI_GUN = 7

# Transliterasiya ad-həlli (Bölmə 2): Yapı Kredi ad-only sətirlərində ad ASCII
# transliterasiya ilə gəlir ("FAMN MAHDUD MASULIYYATLI CAMIYYATI"); AZ hərfli
# formalarla uyğunlaşdırmaq üçün ayrıca qatlama. İstiqamət üzrə açar:
# DEBİTORLAR aktiv, KREDİTORLAR default BAĞLI (reqressiya qorunur).
_TRANSLIT_AD_HELLI_AKTIV = {"debitor": True, "kreditor": False}

# Excel sheet adı qadağanları (spec: : \ / ? * [ ]) + formula/hiperlink
# təhlükəsizliyi üçün əlavə " ' simvolları da atılır.
_INVALID_SHEET_CHARS = set(':\\/?*[]\'"')
_SHEET_NAME_MAX = 31

# ===== Formatlaşdırma =====
_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
_HEADER_FONT = Font(bold=True)
_MONEY_FORMAT = "#,##0.00"
_DATE_NUMBER_FORMAT = "DD.MM.YYYY"
_MONEY_ALIGNMENT = Alignment(horizontal="right", vertical="center")
_CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
_TEXT_ALIGNMENT = Alignment(horizontal="left", vertical="center")
_THIN_BORDER = Border(
    left=Side(style="thin", color="B0B7C3"),
    right=Side(style="thin", color="B0B7C3"),
    top=Side(style="thin", color="B0B7C3"),
    bottom=Side(style="thin", color="B0B7C3"),
)
_AUTO_FIT_MIN = 8
_AUTO_FIT_MAX = 50
_AUTO_FIT_PADDING = 2
_MONEY_FIT_WIDTH = 13

# ===== Başlıq aşkarlanması (normalize edilmiş sütun adları) =====
# GƏLƏN.xlsx: başlıq 11-ci sətirdə (VÖEN, Adı, Qaimə tarixi/seriyası/nömrəsi,
# Malın ƏDV-siz ümumi dəyəri = Əsas, Malın ƏDV məbləği = ƏDV).
_GELEN_FIELDS = {
    "voen": "voen",
    "adi": "ad",
    "qaimetarixi": "tarix",
    "qaimeseriyasi": "seriya",
    "qaimenomresi": "nomre",
    "veziyyeti": "veziyyet",
    "qaimeaktnovleri": "nov",
    "malinedvsizumumideyeri": "esas",
    "malinedvmeblegi": "edv",
    # Avans sütunları (GƏLƏN/GEDƏN): MT qaimə hansı AA avansına bağlıdır.
    "avansserias": "avans_seriya",
    "avansnomresi": "avans_nomre",
    "avansmeblegi": "avans_mebleg",
}
_GELEN_REQUIRED = ("voen", "qaimetarixi")

# ===== Qaimə süzgəci (A) =====
# Vəziyyət normallaşdırılmış halda bu açar ilə başlayırsa → çıxarılır.
_QAIME_PASSIV_ACAR = "passiv"
# Növ mətnində bu açar varsa → avans qaiməsi (çıxarılır).
_QAIME_AVANS_ACAR = "avans"
# Seriya bu prefikslərdən biri ilə başlamalıdır; "AA" həmişə çıxarılır.
# "MQ" (Malların qaytarılması) SAXLANILIR və "Qaytarma qaiməsi" kimi işarələnir.
_QAIME_AA_PREFIX = "AA"
_QAIME_MT_PREFIX = "MT"
_QAIME_MQ_PREFIX = "MQ"
_QAIME_MELUM_PREFIKSLER = (_QAIME_MT_PREFIX, _QAIME_AA_PREFIX, _QAIME_MQ_PREFIX)
_QAIME_QAYTARMA_QEYD = "Qaytarma qaiməsi"
# Vəziyyət bu açar ilə UYĞUN GƏLƏRSƏ saxlanılır, lakin xabərdarlıq yazılır.
# "Təsdiq gözləyən" və "Təsdiq gözləyir" hər ikisi normallaşdırılmış
# "tesdiqgozle" prefiksi ilə tutulur.
_QAIME_VEZIYYET_GOZLEYEN = "tesdiqgozle"
# "Ləğvə göndərildi" → saxlanılır + xəbərdarlıq (məruzədə say).
_QAIME_VEZIYYET_LEGVE = "legvegonderildi"

# İstinad nişanəsi: MT/AA + rəqəm (natamam/dövr xarici aşkarlaması üçün).
_INVOICE_HINT_RE = re.compile(r"(?<![A-Za-z])(MT|AA)\s?\d")
# Əsas istinad regex-i: MT/AA + 4 rəqəm (seriya) + 8 rəqəm (nömrə).
_INVOICE_REF_RE = re.compile(
    r"(?<![A-Za-z0-9])([A-Z]{2})\s?(\d{4})\s*(\d{8})(?!\d)"
)
# Fallback: 8 rəqəmli nömrə (yalnız həmin VÖEN-in qaimə nömrələri arasında).
_BARE_INVOICE_NO_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")
# Kəsilmiş istinad: prefiks (MT/AA/MQ) + 4 rəqəm (seriya) + 7 rəqəm (nömrənin
# ilk hissəsi). Cəm 11 rəqəm — tam 12-yə çatmır; yalnız həmin VÖEN-in həmin
# seriyalı TƏK qaiməsi bu prefikslə başlayırsa həll olunur.
_PARTIAL_REF_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z]{2})\s?(\d{4})\s?(\d{7})(?!\d)")

# DEPOZİT.xlsx: başlıq 10-11-ci sətirdə (VÖEN, Vergi ödəyicisinin adı,
# Əməliyyat tarixi, Mədaxil, Məxaric). "Məxaric" = VÖEN-ə ödənilən ƏDV.
_DEPOZIT_FIELDS = {
    "voen": "voen",
    "vergioddeyicisininadi": "ad",
    "emeliyyattarixi": "tarix",
    "medaxil": "medaxil",
    "mexaric": "mexaric",
    "emeliyyatinteyinati": "teyinat",
}
_DEPOZIT_REQUIRED = ("voen", "mexaric")

# Səhv köçürmə / geri qaytarma açar sözləri (compact normallaşdırma ilə).
# GÜCLÜ: cüt tapılmasa da hədəf sətir ödənişdən çıxarılır.
# ZƏİF: yalnız cüt tapılanda ləğv olunur (təkbaşına "qaytar" YOX).
_SEHV_ACARLARI_GUCLU = ("sehv", "yanlis")
_SEHV_ACARLARI_ZEIF = ("geriqaytar", "iade")

# BANK ÇIXARIŞI — çoxbanklı dinamik aşkarlanma (alias əsaslı).
# Kapital: başlıq 13-cü sətirdə (Tarix, Əks hesabın nömrəsi/adı/vöen-i,
#   Məbləğ, AZN dəyəri, Təyinat) — VÖEN qarışıq mətndə, debit = mənfi Məbləğ/AZN.
# PAŞA: başlıq 13-cü sətirdə (Əməliyyat tarixi, İcra tarixi, Ödəyən/Benefisiar,
#   Təyinat, İstinad No, VÖEN, Mədaxil, Məxaric, Balans) — VÖEN ayrıca sütunda,
#   debit = müsbət Məxaric.
# Yapı Kredi ("standard"): başlıq 15-ci sətirdə (Əməliyyat Tarixi,
#   Həyata Keçmə Tarixi, Göndərən/Alıcı adı, Göndərən/Alıcı Hesabı,
#   Əməliyyat Detalları, Kredit, Debit, Balans) — VÖEN yoxdur, debit = mənfi
#   "Debit". VÖEN təyinat mətnindən regex-lə çıxarıla bilər (label forması).
# ABB (Kapital Bank Business, AZ/EN bilinqval): başlıq 18-ci sətirdə
#   (Əməliyyatın tarixi / Operation date, Debit məbləğ / Debit amount,
#   Kredit məbləğ / Credit amount, Ödənişin təyinatı / Description,
#   Qarşı tərəfin VÖEN-i / Counterparty TAX ID) — VÖEN ayrıca sütunda,
#   debit = işarəsiz müsbət "Debit məbləğ".
# Unibank: başlıq 13-cü sətirdə (VÖEN\n(TIN), Əməliyyat tarixi\n(Transaction
#   date), Mədaxil\n(Paid In), Məxaric\n(Paid Out), Balans\n(Balance),
#   Əməliyyat Tərəfdaşı\n(Counterparty Name), Təyinat\n(Description), ...) —
#   VÖEN ayrıca sütunda, debit = işarəsiz müsbət "Məxaric"; məbləğlər string,
#   boş xana "-" işarəsidir; "Yekun (Summary)" sətri VÖEN sütunundadır.
# TuranBank: başlıq 13-cü sətirdə (№, Tarix, Hesab nömrəsi, Məbləğ,
#   Məbləğin AZN dəyəri, Qalıq, Qalığın AZN dəyəri, Valyuta, Təyinat) —
#   VÖEN sütunu YOXDUR, debit = mənfi Məbləğ/Məbləğin AZN dəyəri (azndeyeri
#   üstünlüklü); "Cəmi" sətri "Hesab nömrəsi" sütunundadır.
_BANK_VOEN_ALIASES = (
    "muxbirinadi",  # Ziraat: "Müxbirin adı" (D) — "muxbir"-dən əvvəl
    "muxbir",  # ATB: "Müxbir" sütununda AD + VÖEN (voen-dən əvvəl)
    "voen",
    "terefinvoeni",
    "eksterefvoen",
    "ekshesab",
    "hesablasmahesabivoeen",
    "qarsiterefinhesabnomresi",
    "taxid",
    "odeyenalan",  # Yelo: "Ödəyən/alan" kompozitində VÖEN (AD<boşluq>VÖEN)
    "creditaccount",  # Bank of Baku (debet sətri qarşı tərəfi)
    "debetaccount",   # Bank of Baku (kredit sətri qarşı tərəfi)
)
_BANK_DEBIT_ALIASES = ("mexaric", "debitmebleg", "expenditure", "debet")
_BANK_SIGNED_ALIASES = ("debit", "azndeyeri", "mebleg")
# Debitorlar (GEDƏN): əks istiqamət — kredit/mədaxil sütunu
_BANK_CREDIT_ALIASES = ("kreditmebleg", "kredit", "medaxil", "profit", "income")
_BANK_DATE_ALIASES = ("emeliyyatintarixi", "emeliyyattarixi", "tarix", "date")
_BANK_NAME_ALIASES = (
    "muxbirinadi",     # Ziraat: "Müxbirin adı" (D)
    "muxbir",          # ATB: "Müxbir" (qarşı tərəf + VÖEN)
    "creditaccount",   # Bank of Baku ("AD / VÖEN / IBAN")
    "debetaccount",
    "gondereninadialaninadi",
    "gondereninadi",
    "senderbeneficiaryname",
    "terefinadi",
    "odeyenbenefisiar",
    "odeyen",
    "benefisiar",
    "qarsiterefinadi",
    "tarafdasi",
    "counterpartyname",
    "adi",
)
_BANK_DESC_ALIASES = (
    "teyinat",
    "emeliyyatdetallari",
    "emeliyyatinadi",
    "emeliyyatinteyinati",
    "description",  # Bank of Baku (Təyinat)
    "tesvir",       # ATB / Rabitə (Təsvir)
)
_BANK_REF_ALIASES = (
    "hesabnomresi",
    "senedinnomresi",
    "istinadno",
    "referens",
    "emeliyyatinreferensi",
    "gonderenalicihesabi",
)
_BANK_BALANCE_TOKENS = ("balans", "dovrun", "movcud", "qaliq", "yekun", "cem")

_VOEN_RE = re.compile(r"^\d{7,12}$")
_IBAN_RE = re.compile(r"^[A-Za-z]{2}\d")
_VOEN_LABEL_RE = re.compile(r"(?:voen|vöen)\s*[:#_\s]*\s*(\d{7,12})", re.IGNORECASE)
_BARE_VOEN_RE = re.compile(r"(?<!\w)(\d{10})(?!\w)")


def _is_summary_row(combined: str) -> bool:
    """Balans/yekun/cəm sətirini tanıyır (yalan-pozitivdən qorunaraq).

    `cem` token-i "cəmiyyət" (şirkət) sözünün daxilindədirsə summary sayılmır.
    """
    for tok in _BANK_BALANCE_TOKENS:
        if tok not in combined:
            continue
        if tok == "cem" and "cemiyyet" in combined:
            continue
        return True
    return False


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if value != value:  # NaN
            return ""
    except Exception:
        pass
    return str(value).strip()


def _voen_str(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, float):
        if value != value:
            return None
        if value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    return text or None


def _to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        result = float(value)
        return result if result == result else None
    text = (
        str(value)
        .strip()
        .replace(" ", "")
        .replace("\u00a0", "")
        .replace("AZN", "")
        .replace("₼", "")
    )
    if not text:
        return None
    if text == "-":
        return 0.0  # boş xana markeri (Unibank) → 0
    try:
        if "," in text:
            text = text.replace(".", "").replace(",", ".")
        return float(text)
    except ValueError:
        return None


def _format_date(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    try:
        if isinstance(value, float) and value != value:
            return ""
    except Exception:
        pass
    return _cell_text(value)


_DATE_INPUT_FORMATS = ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")


def _parse_date(value: Any) -> date | None:
    """Tarixi `datetime.date`-a çevirir (xronoloji sıralama/yazma üçün).

    `datetime`/`date` birbaşa; mətn üçün bir neçə format sınanır. Oxunmayan
    və ya boş dəyər → None (xəta atmır; belə sətirlər sıralamada sona düşür).
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        if isinstance(value, float) and value != value:
            return None
    except Exception:
        pass
    text = _cell_text(value)
    if not text:
        return None
    for fmt in _DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


# Tarixsiz sətirləri sıralamada SONA atmaq üçün "maksimum" tarix açarı.
_DATE_MAX = date.max


def _date_sort_key(value_dt: date | None, fallback: Any, fb_key: Any):
    """(tarix_dt, ...) açarı.

    `tarix_dt` varsa onunla; yoxdursa `fallback` mətn tarixi parse olunur
    (köhnə `_sample_match` kimi testlər yalnız mətn verir). Yenə yoxdursa
    sətir blokun sonuna düşür.
    """
    if value_dt is not None:
        return (value_dt, fb_key)
    parsed = _parse_date(fallback)
    if parsed is not None:
        return (parsed, fb_key)
    return (_DATE_MAX, fb_key)


def _date_cell_value(entry: dict[str, Any]) -> Any:
    """Workbook xanasına yazılan dəyər: real tarix varsa `date`, yoxsa mətn."""
    dt = entry.get("tarix_dt")
    return dt if dt is not None else entry.get("tarix", "")


def _get(row: list[Any], col_map: dict[str, int], field: str) -> Any:
    idx = col_map.get(field)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _find_header(
    rows: list[list[Any]], required: tuple[str, ...], fields: dict[str, str]
) -> tuple[int, dict[str, int]] | None:
    for idx, row in enumerate(rows):
        raw_map: dict[str, int] = {}
        for col_idx, value in enumerate(row):
            key = normalize_text(_cell_text(value))
            if key:
                raw_map.setdefault(key, col_idx)
        if not raw_map:
            continue
        raw_keys = list(raw_map)
        if not all(any(req in k for k in raw_keys) for req in required):
            continue
        col_map: dict[str, int] = {}
        for norm, field in fields.items():
            for key in raw_keys:
                if norm in key:
                    col_map[field] = raw_map[key]
                    break
        return idx, col_map
    return None


# Eyni fayl (bayt məzmunu) bir generasiya ərzində bir neçə funksiya tərəfindən
# oxunur (`_bank_source_label`, `parse_bank_statement`, `count_adonly_self_skips`,
# `_apply_reversals`). `openpyxl` böyük fayllarda çox baha olduğu üçün nəticə
# keşlənir. Keş açarı → SHA-256; qaytarılan struktur YALNIZ oxunur (mutasiya yox).
_LOAD_ROWS_CACHE: dict[str, list[list[list[Any]]]] = {}


def _clear_load_rows_cache() -> None:
    """Testlər/generasiya arası keşi təmizləyir (yaddaş saxlamamaq üçün)."""
    _LOAD_ROWS_CACHE.clear()


def _load_rows(raw: bytes) -> list[list[list[Any]]]:
    key = _sha256(raw)
    cached = _LOAD_ROWS_CACHE.get(key)
    if cached is not None:
        return cached
    try:
        wb = load_workbook(io.BytesIO(raw))
    except Exception as exc:
        raise ValueError(f"Fayl oxunmadı (Excel gözlənilir): {exc}")
    result = [
        [[cell.value for cell in row] for row in ws.iter_rows()]
        for ws in wb.worksheets
    ]
    _LOAD_ROWS_CACHE[key] = result
    return result


def _sha256(raw: bytes) -> str:
    """Fayl məzmununun SHA-256 heşi (dublikat aşkarlaması üçün)."""
    return hashlib.sha256(raw).hexdigest()


def _bank_source_label(filename: str, raw: bytes, used: set[str]) -> str:
    """Bank faylı üçün mənbə etiketi qaytarır.

    Preamble-dən VTB/ATB/Ziraat ailəsi aşkarlanarsa onun adı, əks halda
    fayl adı (uzantısız, os.path.basename). Ad boş olarsa "bank_N" ehtiyatı.
    Eyni etiket təkrarlanarsa sonuna " (2)", " (3)" ... əlavə olunur.
    """
    label = ""
    try:
        rows_list = _load_rows(raw)
    except ValueError:
        rows_list = []  # korlanmış fayl → etiket fayl adından götürülür
    for rows in rows_list:
        header = _detect_bank_header(rows)
        if header is None:
            continue
        family = _detect_preamble_family(rows[: header[0]])
        if family in _FAMILY_LABELS:
            label = _FAMILY_LABELS[family]
        break
    if not label:
        base = os.path.basename(filename or "").strip()
        if base:
            label = os.path.splitext(base)[0].strip()
    if not label:
        label = f"bank_{len(used) + 1}"
    # Təkrar etiketlərə " (2)", " (3)" əlavə et
    candidate = label
    idx = 2
    while candidate in used:
        candidate = f"{label} ({idx})"
        idx += 1
    used.add(candidate)
    return candidate


def _split_counter_account(text: str) -> tuple[str | None, str]:
    """'IBAN / AD / VÖEN' formatından (voen, ad) ayırır.

    IBAN iki hərf + rəqəmlə başlayır, VÖEN yalnız rəqəmdir; qalan ilk
    seqment ad sayılır. VÖEN tapılmasa None qaytarır.
    """
    parts = [p.strip() for p in text.split("/") if p.strip()]
    voen: str | None = None
    ad = ""
    for part in parts:
        if _VOEN_RE.match(part) and voen is None:
            voen = part
        elif not _IBAN_RE.match(part) and not ad:
            ad = part
    return voen, ad


def _clean_counterparty_name(text: str) -> str:
    """Ödəyən/Benefisiar mətnindən adı təmizləyir (PAŞA `_x000D_`, Access `\\n`).

    `/AZ03KABA..._x000D_CAUCASIAN SUPPLY..._x000D_/` → `CAUCASIAN SUPPLY...`;
    Access `İFARƏ...\\nVÖEN\\nHesab` → `İFARƏ...`. IBAN-a bənzər və boş
    seqmentlər atlanır; ilk mənalı seqment ad sayılır.
    """
    parts = [p.strip() for p in re.split(r"_x000D_|\r\n|\r|\n|/", text) if p.strip()]
    for part in parts:
        if not _IBAN_RE.match(part):
            return part
    return ""


def _strip_trailing_voen(text: str, voen: str) -> str:
    """Kompozit `AD<boşluq(lar)>VÖEN[/]` mətnindən son VÖEN-i silir (Yelo/ATB)."""
    if not text or not voen:
        return text
    return re.sub(r"[\s()]*" + re.escape(voen) + r"\s*/?\s*$", "", text).strip()


# Ad müqayisəsi üçün diakritik/replasman cədvəli (AZ hərfləri → latın)
_AD_CHAR_MAP = {
    "ı": "i", "İ": "i", "I": "i",
    "ə": "e", "Ə": "e",
    "ö": "o", "Ö": "o",
    "ü": "u", "Ü": "u",
    "ş": "s", "Ş": "s",
    "ç": "c", "Ç": "c",
    "ğ": "g", "Ğ": "g",
}
# Hüquqi şəkilçilər (normallaşdırmada atılır)
_LEGAL_STOPWORDS = {
    "mmc", "llc", "asc", "qsc", "ojsc", "cjsc", "ltd", "jsc",
    "mehdud", "mesuliyyetli", "cemiyyeti", "sirketi",
}


def _norm_ad(text: str) -> str:
    """Adı müqayisə üçün normallaşdırır.

    İ/ı/Ə/ə/Ö/Ü/Ş/Ç/Ğ → latın, NFKD ilə diakritiklər silinir, kiçik hərf,
    durğu/dırnaq silinir, hüquqi şəkilçilər (MMC, LLC, ASC, QSC, "Məhdud
    Məsuliyyətli Cəmiyyəti" ...) atılır, boşluqlar sıxılır.
    """
    if not text:
        return ""
    replaced = "".join(_AD_CHAR_MAP.get(ch, ch) for ch in text)
    decomposed = unicodedata.normalize("NFKD", replaced)
    ascii_like = "".join(
        ch for ch in decomposed if not unicodedata.combining(ch)
    ).lower()
    cleaned = re.sub(r"[^a-z0-9 ]", " ", ascii_like)
    tokens = [t for t in cleaned.split() if t and t not in _LEGAL_STOPWORDS]
    return " ".join(tokens)


# Transliterasiya qatlaması: ə həm "a" həm "e" kimi gələ bilər (MAHDUD vs
# MEHDUD); hər iki variant ayrı-ayrı hesablanır və ad-həllində hər ikisi sınanır.
_TRANSLIT_CHAR_MAP = {
    "ı": "i", "İ": "i", "I": "i",
    "ö": "o", "Ö": "o",
    "ü": "u", "Ü": "u",
    "ş": "s", "Ş": "s",
    "ç": "c", "Ç": "c",
    "ğ": "g", "Ğ": "g",
}
# Hüquqi stopword-ların transliterasiya variantları.
_LEGAL_STOPWORDS_TRANSLIT = {
    "mmc", "llc", "asc", "qsc", "ojsc", "cjsc", "ltd", "jsc",
    "mehdud", "mahdud",
    "mesuliyyetli", "masuliyyatli", "masuliyyetli",
    "cemiyyeti", "camiyyati", "camiyyeti",
    "sirketi", "sirket", "association",
}
# Kəsilmiş son token: hüquqi stopword-ün ≥3 simvollu prefiksidirsə və ad
# kifayət qədər uzundursa atılır (məs. "CAMIY", "MASUL").
_TRANSLIT_TRUNC_MIN_TOKEN = 3
_TRANSLIT_TRUNC_MIN_AD = 30


def _norm_ad_translit_variants(text: str) -> list[str]:
    """Ad üçün transliterasiya-uyğun normallaşdırılmış variant(lar) qaytarır.

    AZ hərfləri latına çevrilir (ə → "a" və "e" variantları), hüquqi
    stopword-lar və onların ≥3 simvollu kəsilmiş prefiksləri (ad uzundursa)
    atılır. Müqayisə üçün TƏK namizəd tələb olunur; çağıran tərəf bütün
    variantları sınayıb unikal VÖEN seçir.
    """
    if not text:
        return []
    variants: list[str] = []
    for e_repl in ("a", "e"):
        cmap = dict(_TRANSLIT_CHAR_MAP)
        cmap["ə"] = e_repl
        cmap["Ə"] = e_repl
        replaced = "".join(cmap.get(ch, ch) for ch in text)
        decomposed = unicodedata.normalize("NFKD", replaced)
        ascii_like = "".join(
            ch for ch in decomposed if not unicodedata.combining(ch)
        ).lower()
        cleaned = re.sub(r"[^a-z0-9 ]", " ", ascii_like)
        tokens = [
            t for t in cleaned.split() if t and t not in _LEGAL_STOPWORDS_TRANSLIT
        ]
        if len(cleaned) >= _TRANSLIT_TRUNC_MIN_AD and tokens:
            last = tokens[-1]
            if (
                len(last) >= _TRANSLIT_TRUNC_MIN_TOKEN
                and last not in _LEGAL_STOPWORDS_TRANSLIT
                and any(
                    s.startswith(last) and len(s) > len(last)
                    for s in _LEGAL_STOPWORDS_TRANSLIT
                )
            ):
                tokens = tokens[:-1]
        norm = " ".join(tokens)
        if norm and norm not in variants:
            variants.append(norm)
    return variants


def _compact(s: str) -> str:
    """Boşluqları tamamilə silir (compact müqayisə üçün)."""
    return re.sub(r"\s+", "", s or "")


def _token_sorted(s: str) -> str:
    """Sözləri əlifba sırası ilə düzür (söz sırasından asılı olmayan müqayisə)."""
    return " ".join(sorted((s or "").split()))


def _extract_voen_and_name(voen_cell: Any, name_cell: Any) -> tuple[str | None, str]:
    """Ağıllı VÖEN + ad çıxarışı: xalis rəqəm → birbaşa; qarışıq mətn → split.

    Split uğursuz olarsa (məs. Rabitə `IBAN VOEN___<rəqəm>`) regex fallback
    işə düşür. VÖEN və ad eyni kompozit xanadadırsa (məs. Yelo
    `AD<boşluq>VÖEN`), addan son VÖEN silinir.
    """
    voen_text = _cell_text(voen_cell)
    if _VOEN_RE.match(voen_text):
        return voen_text, _clean_counterparty_name(_cell_text(name_cell))
    voen, ad = _split_counter_account(voen_text)
    if voen is None:
        voen = _extract_voen_from_text(voen_text)
    if voen is not None and _cell_text(name_cell) == voen_text:
        # VÖEN və ad eyni xanada (kompozit) → addan VÖEN-i təmizlə
        ad = _clean_counterparty_name(_strip_trailing_voen(voen_text, voen)) or ad
    return voen, ad or _clean_counterparty_name(_cell_text(name_cell))


def _extract_voen_from_text(text: str) -> str | None:
    """Təyinat/əməliyyat mətnindən VÖEN-i regex-lə çıxarır.

    Əvvəlcə label forması (`VOEN: 1403610931`, `VÖEN 9900001231`,
    `VOEN___1405631661`), sonra boş dayanan 10 rəqəmli pattern. IBAN
    substring-i (bitişik rəqəmlər) tutulmur.
    """
    if not text:
        return None
    match = _VOEN_LABEL_RE.search(text)
    if match:
        return match.group(1)
    match = _BARE_VOEN_RE.search(text)
    if match:
        return match.group(1)
    return None


_SELF_VOEN_RE = re.compile(r"(?<!\d)(\d{7,12})(?!\d)")


def _extract_self_identity(raw: bytes) -> tuple[set[str], set[str]]:
    """Qaimə faylının preamble-indən öz şirkətin VÖEN-i və adını çıxarır.

    GƏLƏN/GEDƏN preamble sətrində `1506275021/"ASLANLI-H GROUP" ...` formasında
    öz VÖEN və adı olur. Qaytarır: (voen set, normallaşdırılmış ad set).
    """
    voens: set[str] = set()
    names: set[str] = set()
    try:
        rows_list = _load_rows(raw)
    except ValueError:
        return voens, names
    for rows in rows_list:
        found = _find_header(rows, _GELEN_REQUIRED, _GELEN_FIELDS)
        if found is None:
            continue
        header_idx, _ = found
        for row in rows[:header_idx]:
            for cell in row:
                text = _cell_text(cell)
                m = re.match(r"^\s*(\d{7,12})\s*/\s*(.+)$", text)
                if m:
                    voens.add(m.group(1))
                    nm = _norm_ad(m.group(2))
                    if nm:
                        names.add(nm)
    return voens, names


def _detect_preamble_family(preamble_rows: list[list[Any]]) -> str | None:
    """Preamble-dən bank ailəsini təyin edir: `"vtb"` | `"atb"` | `"ziraat"` | `None`."""
    normalized = " ".join(
        normalize_text(_cell_text(cell))
        for row in preamble_rows
        for cell in row
    )
    if "bankvtb" in normalized:
        return "vtb"
    if "azerturkbank" in normalized or "bankatb" in normalized:
        return "atb"
    if "ziraatbank" in normalized or "tczb" in normalized:
        return "ziraat"
    return None


def _is_adonly_format(
    preamble_rows: list[list[Any]], col_map: dict[str, int]
) -> bool:
    """Yapı Kredi "VÖEN-siz ad-only" rejimini aşkarlayır.

    Şərtlər: VÖEN sütunu YOX + ad sütunu VAR + IBAN/hesab sütunu (ref) VAR +
    preamble-də "Qurum Adı". Yalnız Yapı Kredi fayllarını tutur; digər
    banklarda (xüsusən VÖEN sütunlu) heç vaxt aktivləşmir.
    """
    if "voen" in col_map:
        return False
    if "name" not in col_map or "ref" not in col_map:
        return False
    normalized = " ".join(
        normalize_text(_cell_text(cell))
        for row in preamble_rows
        for cell in row
    )
    return "qurumadi" in normalized


def _extract_qurum_ad(preamble_rows: list[list[Any]]) -> str:
    """Preamble-dən "Qurum Adı" (öz şirkət adı) dəyərini çıxarır."""
    for row in preamble_rows:
        cells = [_cell_text(c) for c in row]
        for i, cell in enumerate(cells):
            if "qurumadi" in normalize_text(cell):
                for nxt in cells[i + 1:]:
                    if nxt:
                        return nxt
    return ""


def count_adonly_self_skips(raw: bytes, direction: str) -> tuple[int, float]:
    """Adonly faylda öz şirkət köçürmələrinin (self-skip) sayı və cəmi.

    Parse zamanı `continue` edilən self-skip sətirləri xaricə çıxmadığı üçün
    eyni şərtlərlə ayrıca hesablanır (xülasə xəbərdarlığı üçün).
    """
    count = 0
    total = 0.0
    for rows in _load_rows(raw):
        found = _detect_bank_header(rows)
        if found is None:
            continue
        header_idx, col_map = found
        preamble = rows[:header_idx]
        if not _is_adonly_format(preamble, col_map):
            return 0, 0.0
        qurum_norm = _norm_ad(_extract_qurum_ad(preamble))
        qurum_compact = _compact(qurum_norm)
        if not qurum_norm:
            return 0, 0.0
        for row in rows[header_idx + 1 :]:
            if direction == "credit":
                if "credit" in col_map:
                    amount = _to_float(_get(row, col_map, "credit"))
                    if amount is None or amount <= 0:
                        continue
                elif "signed" in col_map:
                    amount = _to_float(_get(row, col_map, "signed"))
                    if amount is None or amount <= 0:
                        continue
                else:
                    continue
            elif "debit" in col_map:
                amount = _to_float(_get(row, col_map, "debit"))
                if amount is None or amount <= 0:
                    continue
            else:
                amount = _to_float(_get(row, col_map, "signed"))
                if amount is None or amount >= 0:
                    continue
                amount = abs(amount)
            if not _cell_text(_get(row, col_map, "date")):
                continue
            name = _cell_text(_get(row, col_map, "name"))
            if not name:
                continue
            name_norm = _norm_ad(name)
            if name_norm == qurum_norm or _compact(name_norm) == qurum_compact:
                count += 1
                total += amount
        return count, round(total, 4)
    return count, round(total, 4)


def _detect_self_voens(preamble_rows: list[list[Any]]) -> set[str]:
    """VTB/ATB/Ziraat kimi fayllarda öz (kreditor olmayan) VÖEN-ləri çıxarır.

    Preamble VTB (`bankvtb`), ATB (`azerturkbank`/`bankatb`) və ya Ziraat
    (`ziraatbank`/`tczb`) kimi aşkarlanarsa, `voen` sözlü preamble
    sətirlərindəki 7-12 rəqəmli tokenlər toplanır (bankın öz VÖEN-i + hesab
    sahibinin öz VÖEN-i). Digər banklarda boş set qaytarılır — təsir yoxdur.
    """
    family = _detect_preamble_family(preamble_rows)
    if family is None:
        return set()
    self_voens: set[str] = set()
    for row in preamble_rows:
        row_norm = " ".join(normalize_text(_cell_text(cell)) for cell in row)
        if "voen" not in row_norm:
            continue
        for cell in row:
            self_voens.update(_SELF_VOEN_RE.findall(_cell_text(cell)))
    return self_voens


def _first_col(
    keys: list[str], key_to_col: dict[str, int], aliases: tuple[str, ...]
) -> int | None:
    """Alias siyahısındakı ilk uyğun (dəqiq və ya substring) sütun indeksini qaytarır."""
    for alias in aliases:
        exact = key_to_col.get(alias)
        if exact is not None:
            return exact
        for key in keys:
            if alias in key:
                return key_to_col[key]
    return None


def _detect_bank_header(
    rows: list[list[Any]],
) -> tuple[int, dict[str, int]] | None:
    """Bank çıxarışında başlıq sətrini dinamik tapır və sütunları xəritələyir.

    Qaytarır: (header_idx, col_map) — col_map açarları: date, voen (məcburi
    deyil), debit, signed, name, desc. Debit rejimi üçün `debit` (müsbət),
    işarəli rejim üçün `signed` (mənfi).
    """
    for idx, row in enumerate(rows):
        key_to_col: dict[str, int] = {}
        for col_idx, value in enumerate(row):
            key = normalize_text(_cell_text(value))
            if key:
                key_to_col.setdefault(key, col_idx)
        if not key_to_col:
            continue
        keys = list(key_to_col)

        voen_col = _first_col(keys, key_to_col, _BANK_VOEN_ALIASES)
        date_col = _first_col(keys, key_to_col, _BANK_DATE_ALIASES)
        debit_col = _first_col(keys, key_to_col, _BANK_DEBIT_ALIASES)
        signed_col = _first_col(keys, key_to_col, _BANK_SIGNED_ALIASES)
        credit_col = _first_col(keys, key_to_col, _BANK_CREDIT_ALIASES)
        name_col = _first_col(keys, key_to_col, _BANK_NAME_ALIASES)
        desc_col = _first_col(keys, key_to_col, _BANK_DESC_ALIASES)
        ref_col = _first_col(keys, key_to_col, _BANK_REF_ALIASES)

        if date_col is None:
            continue
        if debit_col is None and signed_col is None:
            continue
        if (
            debit_col is None
            and signed_col is not None
            and _detect_preamble_family(rows[:idx]) == "atb"
        ):
            # ATB "Debit" sütunu işarəsiz müsbətdir (Yapı Kredi signed-dən fərqli)
            debit_col = signed_col
            signed_col = None

        col_map: dict[str, int] = {"date": date_col}
        if voen_col is not None:
            col_map["voen"] = voen_col
        if debit_col is not None:
            col_map["debit"] = debit_col
        if signed_col is not None:
            col_map["signed"] = signed_col
        if credit_col is not None:
            col_map["credit"] = credit_col
        if name_col is not None:
            col_map["name"] = name_col
        if desc_col is not None:
            col_map["desc"] = desc_col
        if ref_col is not None:
            col_map["ref"] = ref_col
        return idx, col_map
    return None


def parse_gelen(raw: bytes) -> list[dict[str, Any]]:
    rows_list = _load_rows(raw)
    for rows in rows_list:
        found = _find_header(rows, _GELEN_REQUIRED, _GELEN_FIELDS)
        if found is None:
            continue
        header_idx, col_map = found
        result: list[dict[str, Any]] = []
        for row in rows[header_idx + 1 :]:
            voen = _voen_str(_get(row, col_map, "voen"))
            if not voen:
                continue
            ad = _cell_text(_get(row, col_map, "ad"))
            if not ad:
                continue
            tarix_raw = _get(row, col_map, "tarix")
            result.append(
                {
                    "doc_type": DOC_INVOICE,
                    "voen": voen,
                    "ad": ad,
                    "tarix": _format_date(tarix_raw),
                    "tarix_dt": _parse_date(tarix_raw),
                    "seriya": _cell_text(_get(row, col_map, "seriya")),
                    "nomre": _cell_text(_get(row, col_map, "nomre")),
                    "veziyyet": _cell_text(_get(row, col_map, "veziyyet")),
                    "nov": _cell_text(_get(row, col_map, "nov")),
                    "esas": _to_float(_get(row, col_map, "esas")) or 0.0,
                    "edv": _to_float(_get(row, col_map, "edv")) or 0.0,
                    "avans_seriya": _cell_text(_get(row, col_map, "avans_seriya")),
                    "avans_nomre": _cell_text(_get(row, col_map, "avans_nomre")),
                    "avans_mebleg": _to_float(_get(row, col_map, "avans_mebleg")),
                }
            )
        return result
    raise ValueError(
        "GƏLƏN faylında başlıq sətri tapılmadı (VÖEN + Qaimə tarixi gözlənilir)."
    )


def _filter_invoices(
    rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """Qaimələri süzür: passiv və avans(AA) qaimələri hesablamadan çıxarır.

    Qaytarır: (saxlanan, çıxarılan, xəbərdarlıqlar). Çıxarılan hər sətirdə
    "sebeb" olur. Naməlum seriya / "Təsdiq gözləyən" statusu saxlanılır, lakin
    xəbərdarlıq yazılır.
    """
    kept: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []
    warnings: list[str] = []
    unknown_series: set[str] = set()
    pending: list[str] = []
    legve: list[str] = []
    for g in rows:
        veziyyet = normalize_text(g.get("veziyyet", ""))
        nov = normalize_text(g.get("nov", ""))
        seriya = str(g.get("seriya") or "").strip().upper()
        reason = None
        if veziyyet.startswith(_QAIME_PASSIV_ACAR):
            reason = "Passiv qaimə"
        elif _QAIME_AVANS_ACAR in nov or seriya.startswith(_QAIME_AA_PREFIX):
            reason = "Avans qaiməsi (AA)"
        if reason is not None:
            removed.append(
                {
                    "voen": g.get("voen", ""),
                    "ad": g.get("ad", ""),
                    "tarix": g.get("tarix", ""),
                    "tarix_dt": g.get("tarix_dt"),
                    "seriya": g.get("seriya", ""),
                    "nomre": g.get("nomre", ""),
                    "nov": g.get("nov", ""),
                    "veziyyet": g.get("veziyyet", ""),
                    "esas": g.get("esas", 0.0),
                    "edv": g.get("edv", 0.0),
                    "sebeb": reason,
                }
            )
            continue
        prefix = seriya[:2]
        if prefix and prefix not in _QAIME_MELUM_PREFIKSLER:
            unknown_series.add(seriya)
        if _QAIME_VEZIYYET_GOZLEYEN in veziyyet:
            pending.append(f"{seriya}/{g.get('nomre','')}")
        if _QAIME_VEZIYYET_LEGVE in veziyyet:
            legve.append(f"{seriya}/{g.get('nomre','')}")
        if prefix == _QAIME_MQ_PREFIX:
            # MQ (Malların qaytarılması) adi müsbət borc qaiməsi kimi saxlanılır,
            # qaimə sətrində "Qaytarma qaiməsi" kimi işarələnir.
            g = {**g, "qeyd": _QAIME_QAYTARMA_QEYD}
        kept.append(g)
    if unknown_series:
        warnings.append(
            "Naməlum qaimə seriyası: " + ", ".join(sorted(unknown_series))
        )
    if pending:
        warnings.append(
            "Təsdiq gözləyən qaimə (saxlanıldı): " + ", ".join(pending)
        )
    if legve:
        warnings.append(
            "Ləğvə göndərilmiş qaimə (saxlanıldı): " + ", ".join(legve)
        )
    if removed:
        warnings.append(f"{len(removed)} qaimə nəzərə alınmadı (passiv/avans).")
    return kept, removed, warnings


def _norm_refs_text(text: str) -> str:
    """İstinad mətnini normallaşdırır (tab/yeni sətir/çox boşluq → tək boşluq)."""
    return re.sub(r"\s+", " ", _cell_text(text)).strip()


# Parçalanmış istinad: prefiks (MT/AA) + aralarında tək boşluq/tab olan rəqəm
# qrupları. Hər qrup bütöv götürülür; cəm TAM 12 rəqəm toplayanda (seriya 4 +
# nömrə 8) namizəd ref yaranır.
_SPLIT_REF_RE = re.compile(
    r"(?<![A-Za-z0-9])(MT|AA)((?:\s?\d+)+)"
)


def _collapse_split_refs(
    norm: str,
    known_refs: set[tuple[str, str]],
) -> list[tuple[str, str]]:
    """Prefiksdən sonra parçalanmış rəqəm qruplarını 12 rəqəmə birləşdirir.

    Dar qayda: "MT"/"AA" prefiksindən sonra gələn rəqəm qrupları (qruplar arası
    TƏK boşluq/tab) bütöv götürülərək 12 rəqəm toplanana qədər birləşdirilir.
    Cəm TAM 12 olmalıdır; qrup 12-ni aşırsa → istinad sayılmır. Nəticə YALNIZ
    `known_refs` (həmin VÖEN-in qüvvədə/çıxarılmış qaimə seriya+nömrəsi) ilə TAM
    uyğun gələrsə qəbul olunur.
    """
    out: list[tuple[str, str]] = []
    for m in _SPLIT_REF_RE.finditer(norm):
        prefix = m.group(1)
        digits = re.findall(r"\d+", m.group(2))
        buf = ""
        for grp in digits:
            if len(buf) + len(grp) > 12:
                break  # 12-ni aşır → istinad sayılmır
            buf += grp
            if len(buf) == 12:
                seriya = prefix + buf[:4]
                nomre = buf[4:]
                if (seriya, nomre) in known_refs:
                    out.append((seriya, nomre))
                break  # hər prefiks üçün yalnız ilk 12 rəqəmli birləşmə
    return out


def _complete_partial_refs(
    norm: str,
    known_refs: set[tuple[str, str]],
) -> list[tuple[str, str]]:
    """Kəsilmiş istinadı (prefiks + seriya + nömrənin ilk 7 rəqəmi) tamamlayır.

    Yalnız həmin VÖEN-in həmin seriyalı qaimələri arasında nömrəsi bu prefiksdən
    başlayan TƏK qaimə varsa həll olunur. Birdən çox və ya heç biri → həll
    olunmur (mövcud məbləğ uyğunluğu/blok qaydaları davam edir).
    """
    if not known_refs:
        return []
    by_series: dict[str, list[str]] = {}
    for ser, nom in known_refs:
        by_series.setdefault(ser, []).append(nom)
    out: list[tuple[str, str]] = []
    for m in _PARTIAL_REF_RE.finditer(norm):
        seriya = m.group(1) + m.group(2)
        partial = m.group(3)
        cands = [n for n in by_series.get(seriya, []) if n.startswith(partial)]
        if len(cands) == 1:
            out.append((seriya, cands[0]))
    return out


def _extract_invoice_refs(
    text: str,
    qaime_by_nomre: dict[str, tuple[str, str]],
    known_refs: set[tuple[str, str]] | None = None,
) -> list[tuple[str, str]]:
    """Ödəniş mətnindən (seriya, nömrə) qaimə istinadlarını çıxarır.

    Sıra: (B1) əsas regex `MT/AA/MQ + 4 + 8 rəqəm`; (B1) parçalanmış yazılış
    `_collapse_split_refs`; (B1c) kəsilmiş nömrə `_complete_partial_refs`;
    fallback — yalnız həmin VÖEN-in məlum qaimə nömrələri arasından 8 rəqəmli
    uyğunluq.
    """
    norm = _norm_refs_text(text)
    if not norm:
        return []
    refs = [(g1 + g2, g3) for g1, g2, g3 in _INVOICE_REF_RE.findall(norm)]
    if known_refs:
        refs += _collapse_split_refs(norm, known_refs)
        refs += _complete_partial_refs(norm, known_refs)
    refs = list(dict.fromkeys(refs))
    if refs:
        return refs
    bare = _BARE_INVOICE_NO_RE.findall(norm)
    matched = [qaime_by_nomre[n] for n in bare if n in qaime_by_nomre]
    return list(dict.fromkeys(matched))


def _allocate_supplier(
    s: dict[str, Any],
    removed_refs: frozenset = frozenset(),
    direction: str = "kreditor",
    removed_meta: dict[tuple[str, str], dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Bir təsərrüfat subyektinin ödənişlərini qaimələrə bölüşdürür.

    `direction="kreditor"` (default) — mövcud (dəyişməz) davranış; tie-break
    ödəniş tarixinə ƏN YAXIN qaiməni seçir.
    `direction="debitor"` — müştəri ödənişi qaimədən SONRA gəldiyi üçün tie-break
    əvvəlcə qaimə tarixi ≤ ödəniş tarixi olanlardan ƏN KÖHNƏSİNİ, yoxdursa
    ödəniş tarixinə ən yaxınını seçir.

    Qaytarır: {
      "groups": [{"qaimə": q, "payments": [row]}],
      "unallocated": [row],
      "removed": [row],   # çıxarılmış (AA/passiv) qaiməyə istinad edən ödənişlər
    }
    hər row: {tarix, tarix_dt, esas, edv, bank, qeyd}.
    """
    # Sənəd tipi ayrılığı (müdafiə qatı): qaimələr INVOICE olmalıdır; qaimə
    # massivinə PAYMENT sızsa da ödəniş kimi emal edilməsin.
    qaimeler = [
        q for q in s["qaimeler"]
        if q.get("doc_type", DOC_INVOICE) == DOC_INVOICE
    ]
    qaimeler.sort(
        key=lambda q: _date_sort_key(
            q.get("tarix_dt"), q.get("tarix", ""),
            (str(q.get("seriya", "")), str(q.get("nomre", ""))),
        )
    )
    qaime_by_nomre: dict[str, tuple[str, str]] = {}
    idx_by_ref: dict[tuple[str, str], dict[str, Any]] = {}
    state: list[dict[str, Any]] = []
    for q in qaimeler:
        ref = (str(q.get("seriya") or "").strip().upper(), str(q.get("nomre") or "").strip())
        st = {
            "q": q,
            "esas": float(q.get("esas") or 0.0),
            "edv": float(q.get("edv") or 0.0),
            "esas_orig": float(q.get("esas") or 0.0),
            "edv_orig": float(q.get("edv") or 0.0),
            "payments": [],
        }
        state.append(st)
        idx_by_ref[ref] = st
        n = ref[1]
        if n and n not in qaime_by_nomre:
            qaime_by_nomre[n] = ref

    unallocated: list[dict[str, Any]] = []
    removed_rows: list[dict[str, Any]] = []

    def _emit(st: dict[str, Any], p: dict[str, Any], amount: float, qeyd: str) -> None:
        row = {
            "tarix": p["tarix"],
            "tarix_dt": p.get("tarix_dt"),
            "esas": amount if p["component"] == "esas" else None,
            "edv": amount if p["component"] == "edv" else None,
            "bank": p.get("bank", ""),
            "qeyd": qeyd,
        }
        st["payments"].append(row)

    def _row(tarix: str, tarix_dt: Any, esas: Any, edv: Any, bank: str, qeyd: str) -> dict[str, Any]:
        return {
            "tarix": tarix, "tarix_dt": tarix_dt, "esas": esas, "edv": edv,
            "bank": bank, "qeyd": qeyd,
        }

    def _pick_by_amount(comp: str, amt: float, pdt: Any) -> dict[str, Any] | None:
        """Tam məbləğ uyğunluğu: qalığı ±0.01 bərabər qaimələrdən seçim.

        Namizədlər TƏKDİRSƏ və ya `pdt` yoxdursa — tarixə görə ən köhnə.
        Kreditorlarda çox namizəd → ödəniş tarixinə ƏN YAXIN.
        Debitorlarda çox namizəd → əvvəlcə qaimə tarixi ≤ ödəniş tarixi
        olanların ƏN KÖHNƏSİ; beləsi yoxdursa → ödəniş tarixinə ən yaxın.
        Tolerans vermir.
        """
        cands = [
            st for st in state
            if abs(max(st[comp], 0.0) - amt) <= 0.01 and st[comp] > 0
        ]
        if not cands:
            return None
        if pdt is not None and len(cands) > 1:
            pdt_parsed = pdt if isinstance(pdt, date) else _parse_date(pdt)
            if pdt_parsed is not None:
                def _qdate(st):
                    return st["q"].get("tarix_dt") or _parse_date(
                        st["q"].get("tarix", "")
                    )

                if direction == "debitor":
                    not_after = [
                        st for st in cands
                        if _qdate(st) is not None and _qdate(st) <= pdt_parsed
                    ]
                    if not_after:
                        return min(not_after, key=_qdate)

                def _dist(st):
                    qd = _qdate(st)
                    if qd is None:
                        return 10**6
                    return abs((qd - pdt_parsed).days)
                return min(cands, key=_dist)
        cands.sort(
            key=lambda st: _date_sort_key(
                st["q"].get("tarix_dt"), st["q"].get("tarix", ""), None
            )
        )
        return cands[0]

    pays: list[dict[str, Any]] = []
    for p in s.get("odenis_esas", []):
        pays.append({
            "tarix": p["tarix"], "tarix_dt": p.get("tarix_dt"),
            "mebleg": float(p["mebleg"]), "component": "esas",
            "bank": p.get("bank", ""), "text": p.get("text", ""),
        })
    for p in s.get("odenis_edv", []):
        pays.append({
            "tarix": p["tarix"], "tarix_dt": p.get("tarix_dt"),
            "mebleg": float(p["mebleg"]), "component": "edv",
            "bank": p.get("bank", ""), "text": p.get("text", ""),
        })
    pays.sort(key=lambda p: _date_sort_key(p["tarix_dt"], p["tarix"], (p["component"],)))

    # Birləşdirmə üçün məlum ref-lər: qüvvədə olan + çıxarılmış (həmin VÖEN).
    known_refs: set[tuple[str, str]] = set(idx_by_ref) | set(removed_refs)

    # Avans zənciri: (AA seriya, AA nömrə) → (MT seriya, MT nömrə) — MT qaimənin
    # "Avans Seriası/Nömrəsi" sütunlarından (eyni VÖEN, MT qüvvədədir).
    chain: dict[tuple[str, str], tuple[str, str]] = {}
    if _AVANS_ZENCIRI_AKTIV.get(direction) and removed_meta:
        for st in state:
            q = st["q"]
            a_ser = str(q.get("avans_seriya") or "").strip().upper()
            a_nom = str(q.get("avans_nomre") or "").strip()
            if a_ser and a_nom:
                chain[(a_ser, a_nom)] = (
                    str(q.get("seriya") or "").strip().upper(),
                    str(q.get("nomre") or "").strip(),
                )

    def _chain_target(removed_hit_: list[tuple[str, str]]):
        """Çıxarılmış AA ref-lərindən zənciri olan ilkini qaytarır."""
        for r in removed_hit_:
            if r in chain and chain[r] in idx_by_ref:
                return r, chain[r]
        return None

    def _match_removed_aa_by_amount_date(comp_: str, amt_: float, pdt_: Any):
        """(2): istinadsız ödənişi AA komponenti + tarix pəncərəsi ilə bağlayır."""
        if not chain or not removed_meta:
            return None
        pdt_parsed = pdt_ if isinstance(pdt_, date) else _parse_date(pdt_)
        for aa_ref, mt_ref in chain.items():
            aa = removed_meta.get(aa_ref)
            if aa is None or mt_ref not in idx_by_ref:
                continue
            st = idx_by_ref[mt_ref]
            if st[comp_] <= 0:
                continue
            comp_amount = float(
                (aa.get("edv") if comp_ == "edv" else aa.get("esas")) or 0.0
            )
            if comp_amount <= 0 or abs(comp_amount - amt_) > 0.01:
                continue
            if pdt_parsed is not None:
                aa_dt = aa.get("tarix_dt")
                if aa_dt is None or abs((pdt_parsed - aa_dt).days) > _AVANS_TARIX_PENCERESI_GUN:
                    continue
            return aa_ref, mt_ref
        return None

    for p in pays:
        comp = p["component"]
        amt = p["mebleg"]
        refs = _extract_invoice_refs(p.get("text", ""), qaime_by_nomre, known_refs)
        removed_hit = [r for r in refs if r in removed_refs]
        resolved = [r for r in refs if r in idx_by_ref]
        unresolved = [r for r in refs if r not in idx_by_ref and r not in removed_refs]
        norm_text = _norm_refs_text(p.get("text", ""))
        if removed_hit and not resolved:
            # (1) Avans zənciri: çıxarılmış AA-ya istinad VAR və zəncir varsa →
            # həmin MT qaiməyə yaz.
            target = _chain_target(removed_hit)
            if target is not None:
                aa_ref, mt_ref = target
                st = idx_by_ref[mt_ref]
                st[comp] = round(st[comp] - amt, 4)
                _emit(st, p, amt, f"Avans zənciri: {aa_ref[0]}/{aa_ref[1]}")
                continue
            if _CIXARILMIS_QAIME_ODENIS_REJIMI == "avans":
                # Avans rejimi: istinadsız avans kimi işlənir → məbləğ/FIFO/blok.
                # refs təmizlənir ki, çıxarılmış qaimə istinad sayılmasın.
                refs = []
                norm_text_for_hint = ""
            else:
                removed_rows.append({
                    "voen": s.get("voen", ""), "ad": s.get("ad", ""),
                    "tarix": p["tarix"], "tarix_dt": p.get("tarix_dt"),
                    "menbe": p.get("bank", ""),
                    "bank": p.get("bank", ""),
                    "istiqamet": "Məxaric" if comp == "esas" else "Mədaxil",
                    "mebleg": amt, "teyinat": p.get("text", ""),
                    "sebeb": "Passiv/avans (AA) qaiməyə istinad edən ödəniş",
                })
                continue
        if resolved:
            # resolved-ləri tarixə artan sırala
            resolved_sorted = sorted(
                resolved,
                key=lambda r: _date_sort_key(
                    idx_by_ref[r]["q"].get("tarix_dt"),
                    idx_by_ref[r]["q"].get("tarix", "") or _DATE_MAX, None,
                ),
            )
            absorb_last = not unresolved
            n = len(resolved_sorted)
            remaining = amt
            total_orig = amt
            allocated_any = False
            for i, r in enumerate(resolved_sorted):
                st = idx_by_ref[r]
                rem = st[comp]
                if i == n - 1 and absorb_last:
                    take = remaining
                else:
                    take = min(max(rem, 0.0), remaining)
                if take > 0.0000001:
                    st[comp] = round(rem - take, 4)
                    remaining = round(remaining - take, 4)
                    qeyd = ""
                    if n > 1 or (unresolved and remaining > 0.0000001):
                        qeyd = f"Ümumi {total_orig:.2f}, bölündü"
                    _emit(st, p, take, qeyd)
                    allocated_any = True
                if remaining <= 0.0000001:
                    break
            if remaining > 0.0000001:
                note_bits = []
                if unresolved:
                    note_bits.append(
                        "Dövr xarici qaimə: "
                        + ", ".join(f"{r[0]}/{r[1]}" for r in unresolved)
                    )
                if removed_hit:
                    note_bits.append(
                        "çıxarılmış qaimə: "
                        + ", ".join(f"{r[0]}/{r[1]}" for r in removed_hit)
                    )
                note = "; ".join(note_bits)
                unallocated.append(_row(
                    p["tarix"], p.get("tarix_dt"),
                    remaining if comp == "esas" else None,
                    remaining if comp == "edv" else None,
                    p.get("bank", ""),
                    f"{note}; qaimələrə {total_orig - remaining:.2f}, "
                    f"qalan {remaining:.2f}",
                ))
            continue
        if unresolved:
            # Həll olunmayan ref varsa: yalnız dövr xarici → bloka (birləşdirmə/
            # məbləğ tətbiq olunmur).
            unallocated.append(_row(
                p["tarix"], p.get("tarix_dt"),
                amt if comp == "esas" else None,
                amt if comp == "edv" else None,
                p.get("bank", ""),
                "Dövr xarici/həll olunmayan qaimə: " + ", ".join(
                    f"{r[0]}/{r[1]}" for r in unresolved
                ),
            ))
            continue
        # (2) İstinadsız: çıxarılmış AA-nın komponentinə məbləğ + tarix uyğunluğu
        # (yalnız zəncir varsa) → həmin MT qaiməyə yaz.
        if not refs:
            aa_target = _match_removed_aa_by_amount_date(comp, amt, p.get("tarix_dt") or p.get("tarix"))
            if aa_target is not None:
                aa_ref, mt_ref = aa_target
                st = idx_by_ref[mt_ref]
                st[comp] = round(st[comp] - amt, 4)
                _emit(st, p, amt, "Avans zənciri (məbləğ+tarix)")
                continue
        if _INVOICE_HINT_RE.search(norm_text):
            # Natamam istinad nişanəsi: bloka atmazdan ƏVVƏL (ii) tam məbləğ
            # uyğunluğu — qalığı ödəniş məbləğinə bərabər TƏK qaimə.
            picked = _pick_by_amount(comp, amt, p.get("tarix_dt") or p.get("tarix"))
            if picked is not None:
                picked[comp] = round(picked[comp] - amt, 4)
                _emit(picked, p, amt, "Məbləğ uyğunluğu")
                continue
            unallocated.append(_row(
                p["tarix"], p.get("tarix_dt"),
                amt if comp == "esas" else None,
                amt if comp == "edv" else None,
                p.get("bank", ""), "Natamam qaimə istinadı",
            ))
            continue
        # İstinadsız: (2a) tam məbləğ uyğunluğu
        best = _pick_by_amount(comp, amt, p.get("tarix_dt") or p.get("tarix"))
        if best is not None:
            best[comp] = round(best[comp] - amt, 4)
            _emit(best, p, amt, "")
            continue
        # (2b) FIFO — ən köhnə qaimədən
        remaining = amt
        for st in state:
            if remaining <= 0.0000001:
                break
            rem = st[comp]
            if rem <= 0:
                continue
            take = min(rem, remaining)
            st[comp] = round(rem - take, 4)
            remaining = round(remaining - take, 4)
            _emit(st, p, take, f"FIFO (ümumi {amt:.2f})" if remaining > 0.0000001 or take < amt else "")
        if remaining > 0.0000001:
            unallocated.append(_row(
                p["tarix"], p.get("tarix_dt"),
                remaining if comp == "esas" else None,
                remaining if comp == "edv" else None,
                p.get("bank", ""), "Qaimə qalığı kifayət etmədi",
            ))

    return {"groups": state, "unallocated": unallocated, "removed": removed_rows}


def _qaime_ref_map(gelen_rows: list[dict[str, Any]]) -> dict[str, dict[str, tuple[str, str]]]:
    """VÖEN → {nömrə: (seriya, nömrə)} xəritəsi."""
    m: dict[str, dict[str, tuple[str, str]]] = {}
    for g in gelen_rows:
        voen = g.get("voen")
        if not voen:
            continue
        n = str(g.get("nomre") or "").strip()
        if n:
            m.setdefault(voen, {}).setdefault(
                n, (str(g.get("seriya") or "").strip().upper(), n)
            )
    return m


def _drop_removed_invoice_payments(
    rows: list[dict[str, Any]],
    ref_map: dict[str, dict[str, tuple[str, str]]],
    removed_refs: frozenset,
    menbe: str,
    istiqamet: str = "Məxaric",
    avans_chain: set[tuple[str, str, str]] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Yalnız çıxarılmış (AA/passiv) qaiməyə istinad edən ödənişləri ayırır.

    Həm çıxarılmış, həm qüvvədə olan qaiməyə istinad varsa ödəniş saxlanılır.
    `avans_chain` verilibsə və çıxarılmış AA-nın (VÖEN, seriya, nömrə) zənciri
    varsa, ödəniş buradan çıxarılmır (bölüşdürmə mərhələsində MT qaiməyə yazılır).
    """
    kept: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []
    for p in rows:
        byn = ref_map.get(p.get("voen") or "", {})
        text = p.get("text") or p.get("teyinat", "")
        # Çıxarılmış AA ref-ləri də məlum ref sayılır ki, bölünmüş ("AA260212 160090")
        # və kəsilmiş yazılışlar erkən mərhələdə tutulsun.
        known = set(byn.values()) | set(removed_refs)
        refs = _extract_invoice_refs(text, byn, known)
        removed_hit = [r for r in refs if r in removed_refs]
        resolved = [r for r in refs if r in set(byn.values())]
        if removed_hit and not resolved:
            if avans_chain is not None:
                voen = p.get("voen") or ""
                if any((voen, r[0], r[1]) in avans_chain for r in removed_hit):
                    kept.append(p)  # avans zənciri var → MT qaiməyə yazılacaq
                    continue
            removed.append(
                {
                    "voen": p.get("voen") or "",
                    "ad": p.get("ad", ""),
                    "tarix": p.get("tarix", ""),
                    "tarix_dt": p.get("tarix_dt"),
                    "menbe": menbe,
                    "bank": p.get("bank", ""),
                    "istiqamet": istiqamet,
                    "mebleg": p.get("mebleg", p.get("mexaric", 0.0)),
                    "teyinat": text,
                    "sebeb": "Passiv/avans (AA) qaiməyə istinad edən ödəniş",
                }
            )
        else:
            kept.append(p)
    return kept, removed


def parse_depozit(
    raw: bytes, direction: str = "expense"
) -> list[dict[str, Any]]:
    """DEPOZİT çıxarışını oxuyur.

    direction="expense" (default): "Məxaric" > 0 (Kreditorlar — VÖEN-ə ödənilən
    ƏDV). direction="income": "Mədaxil" > 0 (Debitorlar — müştərinin ƏDV-ni
    depozit hesabına ödəməsi).
    """
    column = "mexaric" if direction == "expense" else "medaxil"
    rows_list = _load_rows(raw)
    for rows in rows_list:
        found = _find_header(rows, _DEPOZIT_REQUIRED, _DEPOZIT_FIELDS)
        if found is None:
            continue
        header_idx, col_map = found
        # "Ödəniş olunmuş elektron qaimə fakturanın Seriyası/Nömrəsi" (L/M):
        # başlıq birləşmiş ola bilər, ona görə başlıq sətri və növbəti sətirdə
        # DƏQİQ "seriyasi"/"nomresi" axtarılır (real fayllarda boşdur).
        for hi in (header_idx, header_idx + 1):
            if hi >= len(rows):
                break
            for ci, value in enumerate(rows[hi]):
                key = normalize_text(_cell_text(value))
                if key == "seriyasi":
                    col_map.setdefault("seriya", ci)
                elif key == "nomresi":
                    col_map.setdefault("nomre", ci)
        result: list[dict[str, Any]] = []
        for row in rows[header_idx + 1 :]:
            voen = _voen_str(_get(row, col_map, "voen"))
            if not voen:
                continue
            amount = _to_float(_get(row, col_map, column)) or 0.0
            if amount <= 0:
                continue
            tarix_raw = _get(row, col_map, "tarix")
            result.append(
                {
                    "doc_type": DOC_PAYMENT,
                    "voen": voen,
                    "ad": _cell_text(_get(row, col_map, "ad")),
                    "tarix": _format_date(tarix_raw),
                    "tarix_dt": _parse_date(tarix_raw),
                    "mexaric": amount,
                    "teyinat": _cell_text(_get(row, col_map, "teyinat")),
                    "text": _cell_text(_get(row, col_map, "teyinat")),
                    "seriya": _cell_text(_get(row, col_map, "seriya")),
                    "nomre": _cell_text(_get(row, col_map, "nomre")),
                }
            )
        return result
    raise ValueError(
        "DEPOZİT faylında başlıq sətri tapılmadı (VÖEN + Məxaric/Mədaxil gözlənilir)."
    )


def parse_depozit_names(raw: bytes) -> list[tuple[str, str]]:
    """DEPOZİT-dən (ad, VÖEN) cütlərini oxuyur (ad lüğəti üçün; geriyə uyğun).

    `parse_depozit` çıxışına toxunmur; ad sütunu olan sətirlərdən ad→VÖEN
    lüğəti qurmaq üçün ayrıca oxuyur.
    """
    rows_list = _load_rows(raw)
    pairs: list[tuple[str, str]] = []
    for rows in rows_list:
        found = _find_header(rows, _DEPOZIT_REQUIRED, _DEPOZIT_FIELDS)
        if found is None:
            continue
        header_idx, col_map = found
        for row in rows[header_idx + 1 :]:
            voen = _voen_str(_get(row, col_map, "voen"))
            ad = _cell_text(_get(row, col_map, "ad"))
            if voen and ad:
                pairs.append((ad, voen))
        return pairs
    return pairs


def parse_bank_statement(
    raw: bytes,
    direction: str = "debit",
    self_voens: set[str] | None = None,
    self_names: set[str] | None = None,
) -> list[dict[str, Any]]:
    """Çoxbanklı bank çıxarışı parseri (dinamik auto-detection).

    direction="debit" (default): Kreditorlar — tədarükçüyə edilən ödənişlər
    (debit/məxaric). direction="credit": Debitorlar — müştəridən daxil olan
    ödənişlər (kredit/mədaxil); eyni col_map istifadə olunur, yalnız əks
    məbləğ sütunu oxunur.

    Kapital (VÖEN qarışıq mətndə + mənfi Məbləğ/AZN dəyəri), PAŞA
    (VÖEN ayrıca sütunda + müsbət Məxaric), Yapı Kredi "standard"
    (VÖEN yoxdur + mənfi Debit), ABB (VÖEN ayrıca sütunda + işarəsiz
    müsbət Debit məbləğ), Unibank (VÖEN ayrıca sütunda + müsbət Məxaric,
    məbləğlər string, boş xana "-"), TuranBank (VÖEN yoxdur + mənfi
    Məbləğ/AZN dəyəri), Rabitəbank (VÖEN ayrıca hesab sütununda
    `IBAN VOEN___<rəqəm>` + işarəsiz müsbət Debet/Kredit iki-sütun) və
    Access Bank (ayrıca `Tax ID` sütunu + işarəsiz müsbət Məxaric, ad
    sütunu çoxsətirli), Bank Respublika (ayrıca `VÖEN` sütunu + işarəsiz
    müsbət `Debet`, `Kredit` atlanır), Yelo Bank (`Ödəyən/alan`
    kompozit sütununda `AD<boşluq>VÖEN` + işarəsiz müsbət `Məxaric`) və
    Bank of Baku (`Credit Account` sütununda `AD / VÖEN / IBAN` +
    işarəsiz müsbət `Expenditure`), VTB Bank (ayrıca `VÖEN` sütunu +
    işarəsiz müsbət `Debet`), ATB Bank (`Müxbir` sütununda `AD + VÖEN` +
    işarəsiz müsbət `Debit`) və Ziraat Bank (`Müxbirin adı`/`Təyinat`-da
    VÖEN + işarəsiz müsbət `Debet`) formatlarını tanıyır; VTB/ATB/Ziraat-də
    bankın/müştərinin öz VÖEN-li sətirləri, ATB/Ziraat-da həmçinin VÖEN-siz
    sətirlər preamble-dən aşkarlanıb atılır. Balans/summary sətirləri
    atlanır. Yalnız debit
    sətirlər qaytarılır — bunlar tədarükçüyə edilən Əsas (principal)
    ödənişlərdir. VÖEN sütunu mövcud olub sətirdə VÖEN çıxmırsa (daxili
    əməliyyat: bank haqqı/komissiya), həmçinin tarixi boş olan footer/cəm
    sətirləri atlanır.
    """
    rows_list = _load_rows(raw)
    for rows in rows_list:
        found = _detect_bank_header(rows)
        if found is None:
            continue
        header_idx, col_map = found
        has_voen_col = "voen" in col_map
        preamble = rows[:header_idx]
        # VTB/ATB/Ziraat kimi fayllarda bank/müştəri öz VÖEN-lərini kreditor sayma
        self_voens = _detect_self_voens(preamble) | set(self_voens or ())
        self_names_norm = {_norm_ad(n) for n in (self_names or ()) if n}
        family = _detect_preamble_family(preamble)
        is_atb = family == "atb"
        is_ziraat = family == "ziraat"
        # Yapı Kredi "VÖEN-siz ad-only" rejimi (yalnız bu formatda aktiv)
        is_adonly = _is_adonly_format(preamble, col_map)
        qurum_ad_norm = _norm_ad(_extract_qurum_ad(preamble)) if is_adonly else ""
        qurum_ad_compact = _compact(qurum_ad_norm)
        result: list[dict[str, Any]] = []
        for row in rows[header_idx + 1 :]:
            name_cell = _get(row, col_map, "name")
            desc_cell = _get(row, col_map, "desc")
            voen_cell = _get(row, col_map, "voen")
            ref_cell = _get(row, col_map, "ref")
            name_text = _cell_text(name_cell)
            desc_text = _cell_text(desc_cell)
            voen_raw = _cell_text(voen_cell)
            ref_raw = _cell_text(ref_cell)
            # Sahələr AYRI normalize edilib boşluqla birləşir: sahə sərhədində
            # yalan "cem"/"yekun" tokeni yaranmasın (məs. "...bankASC"+"EMeliyyat").
            combined = " ".join(
                normalize_text(x)
                for x in (name_text, desc_text, voen_raw, ref_raw)
            )
            if (not is_ziraat) and _is_summary_row(combined):
                # Ziraat-da Təyinat "qalıq" sözünü ehtiva edə bilər → yalnış
                # summary; ora footer onsuz da VÖEN-siz skip ilə atılır.
                continue  # balans/yekun/cəm summary sətirləri
            if not _cell_text(_get(row, col_map, "date")):
                continue  # tarixi boş footer/cəm sətri (məs. Rabitə 289) → işlənmir

            if direction == "credit":
                # Debitorlar: əks istiqamət — kredit/mədaxil sütunu
                if "credit" in col_map:
                    amount = _to_float(_get(row, col_map, "credit"))
                    if amount is None or amount <= 0:
                        continue
                elif "signed" in col_map:
                    amount = _to_float(_get(row, col_map, "signed"))
                    if amount is None or amount <= 0:
                        continue
                else:
                    continue  # kredit sütunu yoxdursa (məs. yalnız debit) → atla
            elif "debit" in col_map:
                amount = _to_float(_get(row, col_map, "debit"))
                if amount is None or amount <= 0:
                    continue
            else:
                amount = _to_float(_get(row, col_map, "signed"))
                if amount is None or amount >= 0:
                    continue
                amount = abs(amount)

            if is_adonly:
                # Yapı Kredi: VÖEN sütunu yox, qarşı tərəf yalnız ad + IBAN ilə.
                # Desc-dən VÖEN fallback-ı SÖNDÜRÜLÜR (e-qaimə nömrəsi yanlış
                # VÖEN verərdi).
                ad = _cell_text(name_cell)
                if ad:
                    ad_norm = _norm_ad(ad)
                    # Öz şirkət köçürmələri (ad üzrə self-skip; norm + compact)
                    if qurum_ad_norm and (
                        ad_norm == qurum_ad_norm
                        or _compact(ad_norm) == qurum_ad_compact
                    ):
                        continue  # öz hesablar arası köçürmə → ödəniş deyil
                    result.append(
                        {
                            "doc_type": DOC_PAYMENT,
                            "voen": None,
                            "ad": ad,
                            "tarix": _format_date(_get(row, col_map, "date")),
                            "tarix_dt": _parse_date(_get(row, col_map, "date")),
                            "mebleg": amount,
                            "text": desc_text,
                            "ad_ile": True,
                        }
                    )
                else:
                    # Qarşı tərəf adı yox (ATM/kart/faiz/daxili) — FALLBACK olaraq
                    # "Əməliyyat Detalları" (desc) mətni ad kimi yazılır. Lakin
                    # "ad_yoxdur=True" bayrağı SAXLANILIR: bu sətir ad→VÖEN həllinə
                    # və Uyğunsuz sheet-ə girmir (təyinatdaki e-qaimə nömrəsi
                    # yanlış VÖEN əmələ gətirə bilər).
                    result.append(
                        {
                            "doc_type": DOC_PAYMENT,
                            "voen": None,
                            "ad": re.sub(r"_x000D_", " ", desc_text).strip(),
                            "tarix": _format_date(_get(row, col_map, "date")),
                            "tarix_dt": _parse_date(_get(row, col_map, "date")),
                            "mebleg": amount,
                            "text": desc_text,
                            "ad_yoxdur": True,
                        }
                    )
                continue

            voen, ad = _extract_voen_and_name(voen_cell, name_cell)
            if voen is None and is_ziraat:
                # Ziraat: VÖEN "Müxbirin adı" (D) yoxdursa "Təyinat"dan (H)
                voen = _extract_voen_from_text(desc_text)
            if voen is None:
                if is_atb or is_ziraat:
                    continue  # VÖEN-siz (komissiya/kassa/daxili) → kreditor deyil
                if has_voen_col and "/" not in voen_raw:
                    # VÖEN sütunu mövcud, sətirdə VÖEN yox → daxili əməliyyat.
                    # "/" olan qarışıq xana (Kapital `IBAN / AD / VÖEN`) real
                    # qarşı tərəfdir; VÖEN-siz olsa da Uyğunsuz kimi saxlanılır.
                    continue
                voen = _extract_voen_from_text(desc_text)
            if voen and voen in self_voens:
                continue  # bank/müştəri öz VÖEN-i → kreditor deyil (VTB/ATB/Ziraat)
            if not ad:
                ad = re.sub(r"_x000D_", " ", desc_text).strip()  # təyinatdan ad (Turan)
            if ad and self_names_norm and _norm_ad(ad) in self_names_norm:
                continue  # öz şirkət adı ilə köçürmə → ödəniş deyil
            result.append(
                {
                    "doc_type": DOC_PAYMENT,
                    "voen": voen,
                    "ad": ad,
                    "tarix": _format_date(_get(row, col_map, "date")),
                    "tarix_dt": _parse_date(_get(row, col_map, "date")),
                    "mebleg": amount,
                    "text": desc_text,
                }
            )
        return result
    raise ValueError(
        "Bank çıxarışında başlıq sətri tapılmadı "
        "(VÖEN + Tarix + Məbləğ/Məxaric sütunları gözlənilir)."
    )


def _desc_flag_level(detail: Any) -> str | None:
    """Təyinat mətnində səhv/geri-qaytarma açar sözünü aşkarlayır.

    Compact normallaşdırma (diakritiksiz, boşluqsuz) ilə: "səhv"→"sehv",
    "ƏDVgeri"→"edvgeri". Qaytarır: "guclu" | "zeif" | None.
    """
    compact = _compact(_norm_ad(_cell_text(detail)))
    if any(k in compact for k in _SEHV_ACARLARI_GUCLU):
        return "guclu"
    if any(k in compact for k in _SEHV_ACARLARI_ZEIF):
        return "zeif"
    return None


def _entry_direction(entry: dict[str, Any]) -> str | None:
    return entry.get("_istiqamet")


def _entry_amount(entry: dict[str, Any], direction: str) -> float | None:
    """Sətirin verilmiş istiqamətdəki məbləği (hər iki açar adını dəstəkləyir)."""
    for key in (
        ("_mexaric", "mexaric") if direction == "mexaric" else ("_medaxil", "medaxil")
    ):
        if entry.get(key) is not None:
            return entry[key]
    return None


def _cancel_reversals(
    entries: list[dict[str, Any]], target_direction: str
) -> dict[str, Any]:
    """Səhv köçürmə/geri qaytarma cütlərini ləğv edir.

    `entries`: HƏR İKİ istiqamətli sətirlər (`_istiqamet` = "medaxil"|"mexaric",
    `_menbe`, `_mexaric`/`_medaxil` məbləğ). `target_direction` — ödəniş sayılan
    istiqamət (kreditorlar: "mexaric"; debitorlar: "medaxil").

    Qaytarır: {"legvedilen": set(id(entry)), "nezere_alinmayan": [record]}.
    """
    opp = "medaxil" if target_direction == "mexaric" else "mexaric"
    used: set[int] = set()
    cancelled: set[int] = set()
    uncounted: list[dict[str, Any]] = []

    # Flag-lı sətirlər
    flagged = [e for e in entries if _desc_flag_level(e.get("teyinat"))]

    # Əvvəlcə cütləri tap (güclü+zəif): eyni VÖEN, əks istiqamət, bərabər məbləğ
    for flag in flagged:
        if id(flag) in used:
            continue
        voen = flag.get("voen")
        if not voen:
            # VÖEN-siz (ad-only) — təxmin etmə
            continue
        flag_dir = flag["_istiqamet"]
        flag_amt = _entry_amount(flag, flag_dir)
        if flag_amt is None:
            continue
        # Əks istiqamətli namizəd (mütləq flag-ın əksində)
        candidates = [
            e for e in entries
            if id(e) not in used
            and id(e) != id(flag)
            and e.get("voen") == voen
            and e["_istiqamet"] != flag_dir
        ]
        best = None
        best_key = None
        for e in candidates:
            amt = _entry_amount(e, e["_istiqamet"])
            if amt is None or abs(amt - flag_amt) > 0.01:
                continue
            # Flag tarixinə ən yaxın (əvvəlki üstün)
            fd = flag.get("tarix_dt")
            ed = e.get("tarix_dt")
            if fd is not None and ed is not None:
                delta = (ed - fd).days
                key = delta if delta <= 0 else delta + 1000  # əvvəlkilər üstün
            else:
                key = 999999
            if best_key is None or key < best_key:
                best_key, best = key, e
        if best is not None:
            used.add(id(flag))
            used.add(id(best))
            cancelled.add(id(flag))
            cancelled.add(id(best))
            # Cütün HƏR İKİ sətri "Nəzərə alınmayan"da göstərilir.
            for e in (flag, best):
                e_dir = e["_istiqamet"]
                uncounted.append(
                    {
                        "voen": voen,
                        "ad": e.get("ad", ""),
                        "tarix": e.get("tarix", ""),
                        "tarix_dt": e.get("tarix_dt"),
                        "menbe": e.get("_menbe", ""),
                        "istiqamet": "Mədaxil" if e_dir == "medaxil" else "Məxaric",
                        "mebleg": _entry_amount(e, e_dir),
                        "teyinat": e.get("teyinat", ""),
                        "sebeb": "Səhv köçürmə — qarşılıqlı ləğv olundu",
                    }
                )

    # Cüt tapılmayan GÜCLÜ flag-lı hədəf sətirləri → ödənişdən çıxar
    for flag in flagged:
        if id(flag) in used:
            continue
        level = _desc_flag_level(flag.get("teyinat"))
        flag_dir = flag["_istiqamet"]
        if flag_dir != target_direction and level != "guclu":
            # əks istiqamətli zəif flag → yalnız xəbərdarlıq (ödəniş sayılmır,
            # çünki hədəf deyil; sadəcə qeyd)
            continue
        amt = _entry_amount(flag, flag_dir)
        if level == "guclu":
            # Hədəf istiqamətindəsə ödənişdən çıxar; əks istiqamətdədirsə toxunma
            if flag_dir == target_direction:
                cancelled.add(id(flag))
                uncounted.append(
                    {
                        "voen": flag.get("voen") or "",
                        "ad": flag.get("ad", ""),
                        "tarix": flag.get("tarix", ""),
                        "tarix_dt": flag.get("tarix_dt"),
                        "menbe": flag.get("_menbe", ""),
                        "istiqamet": "Mədaxil" if flag_dir == "medaxil" else "Məxaric",
                        "mebleg": amt,
                        "teyinat": flag.get("teyinat", ""),
                        "sebeb": "Səhv/geri qaytarma — cüt tapılmadı (yoxlayın)",
                    }
                )

    return {"legvedilen": cancelled, "nezere_alinmayan": uncounted}


def parse_depozit_all(raw: bytes) -> list[dict[str, Any]]:
    """DEPOZİT-dən HƏR İKİ istiqamətli sətirləri oxuyur (yalnız ləğv axtarışı).

    Nəticə cəmlərə/uğunlaşmaya düşmür; `_istiqamet` sahəsi ilə işarələnir.
    """
    rows_list = _load_rows(raw)
    for rows in rows_list:
        found = _find_header(rows, _DEPOZIT_REQUIRED, _DEPOZIT_FIELDS)
        if found is None:
            continue
        header_idx, col_map = found
        out: list[dict[str, Any]] = []
        for row in rows[header_idx + 1 :]:
            voen = _voen_str(_get(row, col_map, "voen"))
            if not voen:
                continue
            tarix_raw = _get(row, col_map, "tarix")
            medaxil = _to_float(_get(row, col_map, "medaxil")) or 0.0
            mexaric = _to_float(_get(row, col_map, "mexaric")) or 0.0
            direction = "mexaric" if mexaric > 0 else ("medaxil" if medaxil > 0 else None)
            if direction is None:
                continue
            out.append(
                {
                    "voen": voen,
                    "ad": _cell_text(_get(row, col_map, "ad")),
                    "tarix": _format_date(tarix_raw),
                    "tarix_dt": _parse_date(tarix_raw),
                    "medaxil": medaxil,
                    "mexaric": mexaric,
                    "teyinat": _cell_text(_get(row, col_map, "teyinat")),
                    "_istiqamet": direction,
                    "_menbe": "Depozit",
                }
            )
        return out
    return []


def _apply_reversals(
    depozit_target_rows: list[dict[str, Any]],
    bank_target_rows: list[dict[str, Any]],
    depozit_raw: bytes,
    banks: Any,
    target_direction: str,
) -> list[dict[str, Any]]:
    """Səhv köçürmə/geri qaytarma cütlərini ləğv edir və nəticəni qaytarır.

    `depozit_target_rows`/`bank_target_rows` — hədəf istiqamətli (ödəniş sayılan)
    siyahılar; funksiya ləğv olunan sətirləri onlardan SİLİR (qarşılıqlı ləğv).
    Əks istiqamətli sətirlər yalnız ləğv axtarışı üçün ayrıca oxunur.
    Qaytarır: "Nəzərə alınmayan əməliyyatlar" üçün rekord siyahısı.
    """
    depozit_all = parse_depozit_all(depozit_raw)
    bank_all = _parse_banks_all(banks, direction=("debit" if target_direction == "mexaric" else "credit"))

    uncounted: list[dict[str, Any]] = []
    cancelled_ids: set[int] = set()

    for pool_all, target_rows in (
        (depozit_all, depozit_target_rows),
        (bank_all, bank_target_rows),
    ):
        res = _cancel_reversals(pool_all, target_direction)
        cancelled_ids |= res["legvedilen"]
        uncounted.extend(res["nezere_alinmayan"])
        # Hədəf siyahıdan ləğv olunanları çıxar (id-lə deyil, məzmunla uyğunlaşdır)
        removed = [e for e in pool_all if id(e) in res["legvedilen"]]
        to_drop = _match_entries(target_rows, removed)
        for r in sorted(to_drop, reverse=True):
            del target_rows[r]

    return uncounted


def _match_entries(
    target_rows: list[dict[str, Any]], removed: list[dict[str, Any]]
) -> list[int]:
    """`removed` sətirlərinə uyğun `target_rows` indekslərini qaytarır."""
    drops: list[int] = []
    used: set[int] = set()
    for rem in removed:
        for i, t in enumerate(target_rows):
            if i in used:
                continue
            if (
                t.get("voen") == rem.get("voen")
                and t.get("tarix") == rem.get("tarix")
                and abs((t.get("mebleg") or t.get("mexaric") or 0.0)
                        - (rem.get("mexaric") or rem.get("medaxil") or 0.0)) <= 0.01
            ):
                drops.append(i)
                used.add(i)
                break
    return drops


def _normalize_banks(
    banks: Any,
) -> list[tuple[str, bytes]]:
    """`banks` bayt (köhnə) və ya (ad, bayt) siyahısı ola bilər."""
    if isinstance(banks, (bytes, bytearray)):
        return [("bank", bytes(banks))]
    result: list[tuple[str, bytes]] = []
    for item in banks:
        if isinstance(item, (bytes, bytearray)):
            result.append(("bank", bytes(item)))
        else:
            name, raw = item
            result.append((name, bytes(raw)))
    return result


def _parse_banks(
    banks: Any,
    direction: str,
    self_voens: set[str] | None = None,
    self_names: set[str] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """Çoxlu bank faylını tək-tək parse edib birləşdirir.

    Qaytarır: (birləşmiş bank_rows [hər sətirdə "bank" etiketi], xülasə
    sətirləri, xəbərdarlıq siyahısı). Eyni SHA-256-lı təkrar fayl sayılmır.
    """
    normalized = _normalize_banks(banks)
    if not normalized:
        raise ValueError("Bank çıxarışı faylı tələb olunur.")

    used_labels: set[str] = set()
    seen_hashes: dict[str, str] = {}  # hash -> label
    combined: list[dict[str, Any]] = []
    summary: list[dict[str, Any]] = []
    warnings: list[str] = []

    for idx, (filename, raw) in enumerate(normalized, start=1):
        digest = _sha256(raw)
        if digest in seen_hashes:
            dup_of = seen_hashes[digest]
            label = _bank_source_label(filename, raw, used_labels)
            warnings.append(
                f"'{filename or label}' faylı '{dup_of}' ilə eynidir, sayılmadı."
            )
            summary.append(
                {
                    "bank": label,
                    "fayl": filename or label,
                    "say": 0,
                    "cem": 0.0,
                    "uygunlasan": 0.0,
                    "uygunsuz": 0.0,
                    "nezere_alinmayan": 0.0,
                    "adsiz_say": 0,
                    "adsiz_cem": 0.0,
                    "qeyd": "Dublikat, sayılmadı",
                }
            )
            continue
        seen_hashes[digest] = filename or f"bank_{idx}"

        label = _bank_source_label(filename, raw, used_labels)
        try:
            rows = parse_bank_statement(
                raw,
                direction=direction,
                self_voens=self_voens,
                self_names=self_names,
            )
        except ValueError as exc:
            raise ValueError(f"'{filename or label}' faylı oxunmadı: {exc}")

        # Adonly faylda öz şirkət köçürmələrinin xülasə xəbərdarlığı
        self_count, self_sum = count_adonly_self_skips(raw, direction)
        if self_count:
            warnings.append(
                f"{label}: {self_count} öz köçürmə ({self_sum:.2f} AZN) "
                f"ödəniş sayılmadı."
            )

        total = 0.0
        for r in rows:
            r["bank"] = label
            r["_istiqamet"] = "mexaric" if direction == "debit" else "medaxil"
            r["_menbe"] = "Bank"
            combined.append(r)
            total += r["mebleg"]
        summary.append(
            {
                "bank": label,
                "fayl": filename or label,
                "say": len(rows),
                "cem": round(total, 4),
                "uygunlasan": 0.0,
                "uygunsuz": 0.0,
                "nezere_alinmayan": 0.0,
                "adsiz_say": 0,
                "adsiz_cem": 0.0,
                "qeyd": "",
            }
        )

    return combined, summary, warnings


def _parse_banks_all(
    banks: Any, direction: str
) -> list[dict[str, Any]]:
    """Bank fayllarını hər iki istiqamətdə oxuyur (yalnız ləğv axtarışı üçün).

    Hədəf `direction` sətirləri + əks istiqamət sətirləri birlikdə qaytarılır,
    `_istiqamet`/`_menbe` sahası ilə işarələnir. Cəmlərə/uğunlaşmaya düşmür.
    SHA-256 dublikat qaydası qorunur.
    """
    normalized = _normalize_banks(banks)
    if not normalized:
        return []
    opp = "credit" if direction == "debit" else "debit"
    used_labels: set[str] = set()
    seen: set[str] = set()
    out: list[dict[str, Any]] = []
    for filename, raw in normalized:
        digest = _sha256(raw)
        if digest in seen:
            continue
        seen.add(digest)
        label = _bank_source_label(filename, raw, used_labels)
        for d in (direction, opp):
            try:
                rows = parse_bank_statement(raw, direction=d)
            except ValueError:
                rows = []
            for r in rows:
                r["bank"] = label
                r["_istiqamet"] = "mexaric" if d == "debit" else "medaxil"
                r["_menbe"] = "Bank"
                out.append(r)
    return out


def _finalize_bank_summary(
    summary: list[dict[str, Any]],
    matched: dict[str, Any],
    uncounted: list[dict[str, Any]] | None = None,
) -> None:
    """Uyğunlaşan/uyğunsuz/nəzərə alınmayan məbləğləri bank etiketi üzrə doldurur.

    `uygunlasan` YALNIZ vərəqlərə həqiqətən düşən bank (Əsas) ödənişlərindən
    hesablanır — çıxarılmış (AA) ödənişlər vərəqə düşmədiyi üçün buraya aid
    edilmir (əks halda xülase vərəqlərdən böyük görünür).
    """
    by_label = {r["bank"]: r for r in summary if r.get("qeyd") != "Dublikat, sayılmadı"}
    for s in matched["suppliers"]:
        alloc = s.get("_alloc")
        if alloc is not None:
            # Vərəqə düşən HƏR bank ödənişi: həm qaimə altına bölüşdürülənlər,
            # həm də "Qaimə ilə əlaqələndirilməyən ödənişlər" blokundakılar.
            pay_entries = [
                p for g in alloc["groups"] for p in g["payments"]
            ] + list(alloc.get("unallocated", []))
            for p in pay_entries:
                if p.get("esas") is None:  # yalnız bank (Əsas) ödənişləri
                    continue
                label = p.get("bank", "")
                if label in by_label:
                    by_label[label]["uygunlasan"] += p["esas"]
        else:
            # Ehtiyat: birbaşa qurulmuş match (alloc olmadan).
            for p in s.get("odenis_esas", []):
                label = p.get("bank", "")
                if label in by_label:
                    by_label[label]["uygunlasan"] += p["mebleg"]
    for u in matched["unmatched"]:
        if u.get("menbe") != "Bank":
            continue  # Depozit uyğunsuzları bank xülasəsinə düşmür
        label = u.get("bank", "")
        if label in by_label:
            by_label[label]["uygunsuz"] += u["mebleg"]
    for u in (uncounted or []):
        if u.get("menbe") != "Bank":
            continue
        label = u.get("bank", "")
        if label in by_label:
            by_label[label]["nezere_alinmayan"] = (
                by_label[label].get("nezere_alinmayan", 0.0)
                + float(u.get("mebleg") or 0.0)
            )
    for a in matched.get("ad_yoxdur", []):
        label = a.get("bank", "")
        if label in by_label:
            by_label[label]["adsiz_say"] += 1
            by_label[label]["adsiz_cem"] += a["mebleg"]
    for r in summary:
        r["uygunlasan"] = round(r["uygunlasan"], 4)
        r["uygunsuz"] = round(r["uygunsuz"], 4)
        r["nezere_alinmayan"] = round(r.get("nezere_alinmayan", 0.0), 4)
        r["adsiz_cem"] = round(r.get("adsiz_cem", 0.0), 4)


def _build_name_voen_dict(
    name_pairs: list[tuple[str, str]],
    translit: bool = False,
) -> dict[str, set[str]]:
    """(ad, VÖEN) cütlərindən normallaşdırılmış ad → {VÖEN} lüğəti qurur.

    `translit=True` olduqda ad transliterasiya variantları ilə də açarlanır
    (bank ad-only sətirlərinin ASCII adları ilə uyğunlaşma üçün).
    """
    name_to_voen: dict[str, set[str]] = {}
    for ad, voen in name_pairs:
        if not voen:
            continue
        if translit:
            norms = _norm_ad_translit_variants(ad)
        else:
            norm = _norm_ad(ad)
            norms = [norm] if norm else []
        for norm in norms:
            if norm:
                name_to_voen.setdefault(norm, set()).add(voen)
    return name_to_voen


def _resolve_voen_by_name(
    norm: str, name_to_voen: dict[str, set[str]]
) -> str | None:
    """Normallaşdırılmış adı VÖEN-ə çevirir (pilləli; qeyri-müəyyəndə None).

    Pillələr: (1) dəqiq norm ad; (2) token-sorted; (3) compact;
    (4) mühafizəkar fuzzy (ratio ≥ 0.92, unikal ən yaxşı namizəd və
    2-ci namizədlə fərq ≥ 0.05). Çox namizəd/qeyri-müəyyənlik → None.
    """
    if not norm:
        return None
    # 1) dəqiq
    exact = name_to_voen.get(norm)
    if exact and len(exact) == 1:
        return next(iter(exact))
    # 2) token-sorted
    ts = _token_sorted(norm)
    if ts != norm:
        cand = {
            v
            for k, vs in name_to_voen.items()
            if _token_sorted(k) == ts
            for v in vs
        }
        if cand and len(cand) == 1:
            return next(iter(cand))
    # 3) compact
    comp = _compact(norm)
    cand_c = {
        v
        for k, vs in name_to_voen.items()
        if _compact(k) == comp
        for v in vs
    }
    if cand_c and len(cand_c) == 1:
        return next(iter(cand_c))
    # 4) mühafizəkar fuzzy
    scored: list[tuple[float, str]] = []
    for k, vs in name_to_voen.items():
        ratio = difflib.SequenceMatcher(None, norm, k).ratio()
        if ratio >= 0.92:
            for v in vs:
                scored.append((ratio, v))
    if not scored:
        return None
    scored.sort(reverse=True)
    best_ratio, best_voen = scored[0]
    second_ratio = scored[1][0] if len(scored) > 1 else 0.0
    best_voens = {v for r, v in scored if r == best_ratio}
    if len(best_voens) == 1 and (best_ratio - second_ratio) >= 0.05:
        return best_voen
    return None


def _resolve_voen_by_name_translit(
    text: str, name_to_voen: dict[str, set[str]]
) -> str | None:
    """Transliterasiya variantları ilə ad→VÖEN həlli (unikal nəticə tələbi).

    Bütün variantlar sınanır; yalnız TƏK unikal VÖEN çıxarsa qaytarılır,
    əks halda None (çox namizəd / heç biri → Uyğunsuz).
    """
    resolved: set[str] = set()
    for norm in _norm_ad_translit_variants(text):
        v = _resolve_voen_by_name(norm, name_to_voen)
        if v:
            resolved.add(v)
    return next(iter(resolved)) if len(resolved) == 1 else None


def resolve_adonly_payments(
    bank_rows: list[dict[str, Any]],
    name_pairs: list[tuple[str, str]],
    translit: bool = False,
) -> dict[str, tuple[int, float, int, float]]:
    """`ad_ile` (VÖEN-siz, adlı) bank sətirlərini ad üzrə VÖEN-ə bağlayır.

    Yalnız `rel["ad_ile"]` sətirlərə toxunur (VÖEN-li sətirlər dəyişmir).
    `ad_yoxdur` (qarşı tərəfi olmayan) sətirlərə heç bir halda toxunulmur.
    `translit=True` olduqda transliterasiya ad-həlli tətbiq olunur (DEBİTORLAR).
    Qaytarır: bank etiketi → (uyğunlaşan say, cəm, uyğunlaşmayan say, cəm).
    """
    name_to_voen = _build_name_voen_dict(name_pairs, translit=translit)
    stats: dict[str, list[int | float]] = {}

    def _bump(label: str, ok: bool, amount: float) -> None:
        if label not in stats:
            stats[label] = [0, 0.0, 0, 0.0]
        if ok:
            stats[label][0] += 1
            stats[label][1] += amount
        else:
            stats[label][2] += 1
            stats[label][3] += amount

    for r in bank_rows:
        if not r.get("ad_ile") or r.get("ad_yoxdur"):
            continue  # adsız sətirlər ad→VÖEN həllinə girmir
        label = r.get("bank", "")
        if translit:
            matched = _resolve_voen_by_name_translit(r.get("ad", ""), name_to_voen)
        else:
            norm = _norm_ad(r.get("ad", ""))
            matched = _resolve_voen_by_name(norm, name_to_voen)
        if matched:
            r["voen"] = matched
            _bump(label, True, r["mebleg"])
        else:
            _bump(label, False, r["mebleg"])

    return {
        k: (int(v[0]), round(float(v[1]), 4), int(v[2]), round(float(v[3]), 4))
        for k, v in stats.items()
    }


def match_by_voen(
    gelen_rows: list[dict[str, Any]],
    depozit_rows: list[dict[str, Any]],
    bank_rows: list[dict[str, Any]],
) -> dict[str, Any]:
    suppliers: dict[str, dict[str, Any]] = {}
    for g in gelen_rows:
        voen = g["voen"]
        s = suppliers.setdefault(
            voen,
            {
                "voen": voen,
                "ad": g["ad"],
                "qaimeler": [],
                "odenis_esas": [],
                "odenis_edv": [],
            },
        )
        s["qaimeler"].append(g)

    unmatched: list[dict[str, Any]] = []
    ad_yoxdur: list[dict[str, Any]] = []
    for b in bank_rows:
        # "bank" YALNIZ etiketdir — uyğunlaşdırma açarına qatılmır.
        bank_label = b.get("bank", "")
        if b.get("ad_yoxdur"):
            # Qarşı tərəfi olmayan (ATM/kart/faiz) sətir — Uyğunsuz sheet-ə
            # YAZILMIR; yalnız xülasədə "Adsız" sütununda sayılır.
            ad_yoxdur.append(
                {"bank": bank_label, "mebleg": b["mebleg"]}
            )
            continue
        if b["voen"] and b["voen"] in suppliers:
            suppliers[b["voen"]]["odenis_esas"].append(
                {
                    "tarix": b["tarix"],
                    "tarix_dt": b.get("tarix_dt"),
                    "mebleg": b["mebleg"],
                    "bank": bank_label,
                    "text": b.get("text", ""),
                }
            )
        else:
            unmatched.append(
                {
                    "voen": b["voen"] or "",
                    "ad": b["ad"],
                    "tarix": b["tarix"],
                    "tarix_dt": b.get("tarix_dt"),
                    "menbe": "Bank",
                    "mebleg": b["mebleg"],
                    "bank": bank_label,
                    "text": b.get("text", ""),
                }
            )
    for d in depozit_rows:
        if d["voen"] and d["voen"] in suppliers:
            suppliers[d["voen"]]["odenis_edv"].append(
                {
                    "tarix": d["tarix"],
                    "tarix_dt": d.get("tarix_dt"),
                    "mebleg": d["mexaric"],
                    "bank": DEPOZIT_BANK_LABEL,
                    "text": d.get("text", d.get("teyinat", "")),
                }
            )
        else:
            unmatched.append(
                {
                    "voen": d["voen"] or "",
                    "ad": d["ad"],
                    "tarix": d["tarix"],
                    "tarix_dt": d.get("tarix_dt"),
                    "menbe": "Depozit",
                    "mebleg": d["mexaric"],
                    "bank": DEPOZIT_BANK_LABEL,
                    "text": d.get("text", d.get("teyinat", "")),
                }
            )

    supplier_list: list[dict[str, Any]] = []
    for s in suppliers.values():
        hes_esas = round(sum(q["esas"] for q in s["qaimeler"]), 4)
        hes_edv = round(sum(q["edv"] for q in s["qaimeler"]), 4)
        od_esas = round(sum(p["mebleg"] for p in s["odenis_esas"]), 4)
        od_edv = round(sum(p["mebleg"] for p in s["odenis_edv"]), 4)
        s["hesablanib_esas"] = hes_esas
        s["hesablanib_edv"] = hes_edv
        s["odenilib_esas"] = od_esas
        s["odenilib_edv"] = od_edv
        s["qaliq_esas"] = round(hes_esas - od_esas, 4)
        s["qaliq_edv"] = round(hes_edv - od_edv, 4)
        supplier_list.append(s)

    supplier_list.sort(key=lambda s: s["ad"].lower())
    unmatched.sort(
        key=lambda u: _date_sort_key(
            u.get("tarix_dt"), u.get("tarix", ""), (u.get("menbe", ""),)
        )
    )
    return {
        "suppliers": supplier_list,
        "unmatched": unmatched,
        "ad_yoxdur": ad_yoxdur,
    }


def sanitize_sheet_name(name: str) -> str:
    cleaned = "".join(ch for ch in name if ch not in _INVALID_SHEET_CHARS)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or "Tedarukcu"


def unique_sheet_name(name: str, voen: str, used: set[str]) -> str:
    base = sanitize_sheet_name(name)[:_SHEET_NAME_MAX]
    if not base:
        base = (voen or "Tedarukcu")[:_SHEET_NAME_MAX]
    candidate = base
    if candidate in used:
        suffix = "_" + (voen[-4:] if voen else "")
        candidate = base[: _SHEET_NAME_MAX - len(suffix)] + suffix
    idx = 2
    while candidate in used:
        n = f"_{idx}"
        candidate = base[: _SHEET_NAME_MAX - len(n)] + n
        idx += 1
    used.add(candidate)
    return candidate


def _quote_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def _qeyd_flag(s: dict[str, Any]) -> str:
    flags = []
    if s["qaliq_esas"] < -0.004:
        flags.append("Əsas")
    if s["qaliq_edv"] < -0.004:
        flags.append("ƏDV")
    return "Artıq ödəmə: " + ", ".join(flags) if flags else ""


def _style_header_row(ws, row_idx: int, last_col: int) -> None:
    for col in range(1, last_col + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER_ALIGNMENT


def _style_money(ws, start_row: int, end_row: int, cols: list[int]) -> None:
    for col in cols:
        for row in range(start_row, end_row + 1):
            cell = ws.cell(row=row, column=col)
            cell.number_format = _MONEY_FORMAT
            cell.alignment = _MONEY_ALIGNMENT


def _style_dates(ws, start_row: int, end_row: int, cols: list[int]) -> None:
    """Tarix sütunlarını real Excel tarixi formatı ilə göstərir."""
    for col in cols:
        for row in range(start_row, end_row + 1):
            cell = ws.cell(row=row, column=col)
            cell.number_format = _DATE_NUMBER_FORMAT
            cell.alignment = _TEXT_ALIGNMENT


def _set_link(cell, location: str, display: str | None = None) -> None:
    cell.hyperlink = Hyperlink(
        ref=cell.coordinate, location=location, display=display
    )


def _apply_borders(
    ws, start_row: int, end_row: int, start_col: int, end_col: int
) -> None:
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            ws.cell(row=row, column=col).border = _THIN_BORDER


def _apply_text_alignment(ws, start_row: int, end_row: int, cols: list[int]) -> None:
    for col in cols:
        for row in range(start_row, end_row + 1):
            ws.cell(row=row, column=col).alignment = _TEXT_ALIGNMENT


def _auto_fit_columns(ws, money_cols: tuple[int, ...] = ()) -> None:
    money_cols = set(money_cols)
    multi_col_merged: set[tuple[int, int]] = set()
    for rng in ws.merged_cells.ranges:
        if rng.max_col > rng.min_col:
            multi_col_merged.add((rng.min_row, rng.min_col))
    for col in range(1, ws.max_column + 1):
        max_len = _AUTO_FIT_MIN
        for row in range(1, ws.max_row + 1):
            if (row, col) in multi_col_merged:
                continue
            cell = ws.cell(row=row, column=col)
            value = cell.value
            if value is None:
                continue
            if isinstance(value, str) and value.startswith("="):
                if col in money_cols:
                    max_len = max(max_len, _MONEY_FIT_WIDTH)
                continue
            max_len = max(max_len, len(str(value)))
        width = min(max_len + _AUTO_FIT_PADDING, _AUTO_FIT_MAX)
        ws.column_dimensions[get_column_letter(col)].width = width


def _build_supplier_sheet(
    wb: Workbook,
    s: dict[str, Any],
    sheet_name: str,
    main_sheet_name: str = MAIN_SHEET_NAME,
    back_label: str = "← KREDİTORLAR",
) -> int:
    ws = wb.create_sheet(sheet_name)

    ws["A1"] = s["ad"]
    _set_link(ws["A1"], main_sheet_name + "!A1", display=s["ad"])
    ws.merge_cells("A1:J1")
    voen_val: Any = s["voen"]
    if voen_val.isdigit() and not voen_val.startswith("0"):
        voen_val = int(voen_val)
    ws["A2"] = voen_val
    ws.merge_cells("A2:J2")
    k2 = ws["K2"]
    k2.value = back_label
    _set_link(k2, main_sheet_name + "!A1", display=back_label)

    ws["A4"] = "Date"
    ws["B4"] = "Elektron Qaimə"
    ws["D4"] = "Ödəniş tarixi"
    ws["E4"] = "Hesablanıb"
    ws["G4"] = "Ödənilib"
    ws["I4"] = "Qalıq"
    ws["K4"] = "Bank"
    ws["L4"] = "Qeyd"
    ws["E5"] = "Əsas"
    ws["F5"] = "ƏDV"
    ws["G5"] = "Əsas"
    ws["H5"] = "ƏDV"
    ws["I5"] = "Əsas"
    ws["J5"] = "ƏDV"
    for rng in (
        "A4:A5", "B4:C5", "D4:D5", "E4:F4", "G4:H4", "I4:J4", "K4:K5", "L4:L5"
    ):
        ws.merge_cells(rng)

    alloc = s.get("_alloc")
    if alloc is None:
        alloc = _allocate_supplier(s)

    row = 6

    def _write_payment(p: dict[str, Any], negate_settle: bool = False) -> None:
        nonlocal row
        if p.get("tarix_dt") is not None or p.get("tarix"):
            ws.cell(row=row, column=4, value=_date_cell_value(p))
        has_g = bool(p.get("esas"))
        has_h = bool(p.get("edv"))
        if has_g:
            ws.cell(row=row, column=7, value=p["esas"])
        if has_h:
            ws.cell(row=row, column=8, value=p["edv"])
        if p.get("bank"):
            ws.cell(row=row, column=11, value=p["bank"])
        if p.get("qeyd"):
            ws.cell(row=row, column=12, value=p["qeyd"])
        # Blok sətri: Qalıq = −Ödənilib (mənfi, ödəniş çıxılır).
        if negate_settle:
            if has_g:
                ws.cell(row=row, column=9, value=f"=-G{row}")
            if has_h:
                ws.cell(row=row, column=10, value=f"=-H{row}")
        row += 1

    for st in alloc["groups"]:
        q = st["q"]
        invoice_row = row
        ws.cell(row=row, column=1, value=_date_cell_value(q))
        ws.cell(row=row, column=2, value=q.get("seriya", ""))
        ws.cell(row=row, column=3, value=q.get("nomre", ""))
        ws.cell(row=row, column=5, value=q.get("esas"))
        ws.cell(row=row, column=6, value=q.get("edv"))
        row += 1
        first_pay = row
        for p in st["payments"]:
            _write_payment(p)
        last_pay = row - 1
        # Ana qaimə sətrində Ödənilib (G/H) BOŞ qalır. Qalıq = Hesablanıb −
        # alt ödənişlərin cəmi (ödənişsiz qaimədə I=E, J=F).
        if last_pay >= first_pay:
            ws.cell(row=invoice_row, column=9,
                    value=f"=E{invoice_row}-SUM(G{first_pay}:G{last_pay})")
            ws.cell(row=invoice_row, column=10,
                    value=f"=F{invoice_row}-SUM(H{first_pay}:H{last_pay})")
        else:
            ws.cell(row=invoice_row, column=9, value=f"=E{invoice_row}")
            ws.cell(row=invoice_row, column=10, value=f"=F{invoice_row}")
        # Artıq ödəmə qeydi + qaimə qeydi (məs. "Qaytarma qaiməsi") L sütununda
        paid_esas = sum((p.get("esas") or 0.0) for p in st["payments"])
        paid_edv = sum((p.get("edv") or 0.0) for p in st["payments"])
        flags = []
        if paid_esas - (q.get("esas") or 0.0) > 0.004:
            flags.append("Əsas")
        if paid_edv - (q.get("edv") or 0.0) > 0.004:
            flags.append("ƏDV")
        invoice_notes = []
        if q.get("qeyd"):
            invoice_notes.append(str(q["qeyd"]))
        if flags:
            invoice_notes.append("Artıq ödəmə: " + ", ".join(flags))
        if invoice_notes:
            ws.cell(row=invoice_row, column=12, value="; ".join(invoice_notes))

    unallocated = [p for p in alloc.get("unallocated", []) if p.get("esas") or p.get("edv")]
    if unallocated:
        ws.cell(row=row, column=1, value="Qaimə ilə əlaqələndirilməyən ödənişlər")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
        ws.cell(row=row, column=1).font = _HEADER_FONT
        row += 1
        for p in unallocated:
            _write_payment(p, negate_settle=True)

    last = row - 1
    sum_row = row
    # Cəm: E–J hamısı sadə =SUM (G/H yalnız ödəniş/blok sətirlərindədir → ikiqat
    # sayılma strukturaca mümkün deyil).
    if last >= 6:
        for col in "EFGHIJ":
            ws[f"{col}{sum_row}"] = f"=SUM({col}6:{col}{last})"
    else:
        for col in "EFGHIJ":
            ws[f"{col}{sum_row}"] = f"=SUM({col}6:{col}6)"

    _style_header_row(ws, 4, 12)
    _style_header_row(ws, 5, 12)
    _style_money(ws, 6, sum_row, [5, 6, 7, 8, 9, 10])
    _style_dates(ws, 6, sum_row, [1, 4])
    _apply_text_alignment(ws, 6, sum_row - 1, [2, 3, 11, 12])
    _apply_borders(ws, 4, sum_row, 1, 12)
    _auto_fit_columns(ws, money_cols=(5, 6, 7, 8, 9, 10))
    return sum_row


def _build_unmatched_sheet(wb: Workbook, unmatched: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(UNMATCHED_SHEET_NAME)
    ws.append(["VÖEN", "Ad", "Tarix", "Mənbə", "Məbləğ", "Bank"])
    # Xronoloji sıralama (tarix_dt); tarixsizlər sona.
    ordered = sorted(
        unmatched,
        key=lambda u: _date_sort_key(
            u.get("tarix_dt"), u.get("tarix", ""), (u.get("menbe", ""),)
        ),
    )
    for u in ordered:
        ws.append(
            [
                u["voen"],
                u["ad"],
                u.get("tarix_dt") if u.get("tarix_dt") is not None else u["tarix"],
                u["menbe"],
                u["mebleg"],
                u.get("bank", ""),
            ]
        )
    _style_header_row(ws, 1, 6)
    _style_money(ws, 2, max(ws.max_row, 2), [5])
    _style_dates(ws, 2, max(ws.max_row, 2), [3])
    _apply_text_alignment(ws, 2, max(ws.max_row, 2), [1, 2, 4, 6])
    _apply_borders(ws, 1, max(ws.max_row, 2), 1, 6)
    _auto_fit_columns(ws, money_cols=(5,))


def _build_reversals_sheet(
    wb: Workbook, reversals: list[dict[str, Any]]
) -> None:
    """'Nəzərə alınmayan əməliyyatlar' — yalnız ≥1 sətir olduqda yaradılır."""
    ws = wb.create_sheet(REVERSALS_SHEET_NAME)
    ws.append(
        ["VÖEN", "Ad", "Tarix", "Mənbə", "İstiqamət", "Məbləğ", "Təyinat", "Səbəb"]
    )
    ordered = sorted(
        reversals,
        key=lambda u: _date_sort_key(
            u.get("tarix_dt"), u.get("tarix", ""), (u.get("menbe", ""),)
        ),
    )
    for u in ordered:
        ws.append(
            [
                u.get("voen", ""),
                u.get("ad", ""),
                u.get("tarix_dt") if u.get("tarix_dt") is not None else u.get("tarix", ""),
                u.get("menbe", ""),
                u.get("istiqamet", ""),
                u.get("mebleg", 0.0),
                u.get("teyinat", ""),
                u.get("sebeb", ""),
            ]
        )
    _style_header_row(ws, 1, 8)
    _style_money(ws, 2, max(ws.max_row, 2), [6])
    _style_dates(ws, 2, max(ws.max_row, 2), [3])
    _apply_text_alignment(ws, 2, max(ws.max_row, 2), [1, 2, 4, 5, 7, 8])
    _apply_borders(ws, 1, max(ws.max_row, 2), 1, 8)
    _auto_fit_columns(ws, money_cols=(6,))


def _build_removed_invoices_sheet(
    wb: Workbook, removed: list[dict[str, Any]]
) -> None:
    """'Nəzərə alınmayan qaimələr' — yalnız ≥1 sətir olduqda yaradılır."""
    ws = wb.create_sheet(REMOVED_INVOICES_SHEET_NAME)
    ws.append(
        [
            "VÖEN", "Ad", "Tarix", "Seriya", "Nömrə", "Növ",
            "Vəziyyət", "ƏDV-siz", "ƏDV", "Səbəb",
        ]
    )
    ordered = sorted(
        removed,
        key=lambda u: _date_sort_key(
            u.get("tarix_dt"), u.get("tarix", ""),
            (str(u.get("seriya", "")), str(u.get("nomre", ""))),
        ),
    )
    for u in ordered:
        ws.append(
            [
                u.get("voen", ""),
                u.get("ad", ""),
                u.get("tarix_dt") if u.get("tarix_dt") is not None else u.get("tarix", ""),
                u.get("seriya", ""),
                u.get("nomre", ""),
                u.get("nov", ""),
                u.get("veziyyet", ""),
                u.get("esas", 0.0),
                u.get("edv", 0.0),
                u.get("sebeb", ""),
            ]
        )
    _style_header_row(ws, 1, 10)
    _style_money(ws, 2, max(ws.max_row, 2), [8, 9])
    _style_dates(ws, 2, max(ws.max_row, 2), [3])
    _apply_text_alignment(ws, 2, max(ws.max_row, 2), [1, 2, 4, 5, 6, 7, 10])
    _apply_borders(ws, 1, max(ws.max_row, 2), 1, 10)
    _auto_fit_columns(ws, money_cols=(8, 9))


def _build_bank_summary_sheet(
    wb: Workbook,
    bank_summary: list[dict[str, Any]],
    main_sheet_name: str,
    warnings: list[str],
) -> None:
    """Hər bank faylı üzrə xülasə + (varsa) xəbərdarlıq sətirləri + ÜMUMİ."""
    ws = wb.create_sheet(BANK_SUMMARY_SHEET_NAME)
    ws.append(
        [
            "Bank",
            "Fayl adı",
            "Əməliyyat sayı",
            "Ümumi məbləğ",
            "Uyğunlaşan",
            "Uyğunsuz",
            "Adsız (qarşı tərəf yoxdur)",
            "Adsız sayı",
            "Qeyd",
            "Nəzərə alınmayan",
        ]
    )
    for r in bank_summary:
        ws.append(
            [
                r["bank"],
                r["fayl"],
                r["say"],
                r["cem"],
                r["uygunlasan"],
                r["uygunsuz"],
                r.get("adsiz_cem", 0.0),
                r.get("adsiz_say", 0),
                r.get("qeyd", ""),
                r.get("nezere_alinmayan", 0.0),
            ]
        )

    first_data = 2
    last_data = ws.max_row
    total_row = last_data + 1
    ws.cell(row=total_row, column=1, value="ÜMUMİ")
    if last_data >= first_data:
        for col in (3, 4, 5, 6, 7, 8, 10):
            letter = get_column_letter(col)
            ws[f"{letter}{total_row}"] = f"=SUM({letter}{first_data}:{letter}{last_data})"

    # Xəbərdarlıqlar ayrı sətirlər kimi (cədvəldən sonra)
    warn_row = total_row + 2
    if warnings:
        ws.cell(row=warn_row, column=1, value="Xəbərdarlıq")
        for i, msg in enumerate(warnings, start=1):
            ws.cell(row=warn_row + i, column=1, value=msg)

    # Geri-link master-ə
    a1 = ws["A1"]
    a1.value = "Banklar üzrə xülasə"
    _set_link(a1, main_sheet_name + "!A1", display="Banklar üzrə xülasə")

    _style_header_row(ws, 1, 10)
    _style_money(ws, first_data, total_row, [4, 5, 6, 7, 10])
    _apply_text_alignment(ws, first_data, total_row, [1, 2, 9])
    _apply_borders(ws, 1, total_row, 1, 10)
    _auto_fit_columns(ws, money_cols=(4, 5, 6, 7, 10))


def build_workbook(
    match_result: dict[str, Any],
    sheet_title: str = MAIN_SHEET_NAME,
    back_label: str = "← KREDİTORLAR",
    aciqlama: str = _ACIQLAMA,
    bank_summary: list[dict[str, Any]] | None = None,
    warnings: list[str] | None = None,
    reversals: list[dict[str, Any]] | None = None,
    removed_invoices: list[dict[str, Any]] | None = None,
) -> bytes:
    wb = Workbook()
    main = wb.active
    main.title = sheet_title

    main["A1"] = "№"
    main["B1"] = sheet_title
    main["C1"] = "Açıqlama"
    main["D1"] = "Hesablanıb"
    main["F1"] = "Ödənilib"
    main["H1"] = "Qalıq"
    main["J1"] = "Qeyd"
    main["D2"] = "Əsas"
    main["E2"] = "ƏDV"
    main["F2"] = "Əsas"
    main["G2"] = "ƏDV"
    main["H2"] = "Əsas"
    main["I2"] = "ƏDV"
    for rng in ("A1:A2", "B1:B2", "C1:C2", "D1:E1", "F1:G1", "H1:I1", "J1:J2"):
        main.merge_cells(rng)

    # Qorunan sheet adları: təchizatçı adı bunlarla üst-üstə düşməsin
    used: set[str] = {
        UNMATCHED_SHEET_NAME,
        BANK_SUMMARY_SHEET_NAME,
        REVERSALS_SHEET_NAME,
        REMOVED_INVOICES_SHEET_NAME,
    }
    row = 3
    for i, s in enumerate(match_result["suppliers"], start=1):
        sheet_name = unique_sheet_name(s["ad"], s["voen"], used)
        sum_row = _build_supplier_sheet(wb, s, sheet_name, sheet_title, back_label)

        main.cell(row=row, column=1, value=i)
        target = f"{_quote_sheet(sheet_name)}!A1"
        # Görünən mətn HƏMİŞƏ əsl addır; hyperlink.location arxada qalır
        b_cell = main.cell(row=row, column=2, value=s["ad"])
        _set_link(b_cell, target, display=s["ad"])
        main.cell(row=row, column=3, value=aciqlama)
        q = _quote_sheet(sheet_name)
        for col, letter in zip((4, 5, 6, 7, 8, 9), "EFGHIJ"):
            cell = main.cell(row=row, column=col, value=f"={q}!{letter}{sum_row}")
            _set_link(cell, target)
        main.cell(row=row, column=10, value=_qeyd_flag(s))
        row += 1

    last_data = row - 1
    main.cell(row=row, column=2, value="CƏMİ")
    for col in "DEFGHI":
        main[f"{col}{row}"] = f"=SUM({col}3:{col}{last_data})"

    _style_header_row(main, 1, 10)
    _style_header_row(main, 2, 10)
    _style_money(main, 3, row, [4, 5, 6, 7, 8, 9])
    _apply_text_alignment(main, 3, row, [2, 3, 10])
    for r in range(3, row + 1):
        main.cell(row=r, column=1).alignment = _MONEY_ALIGNMENT
    _apply_borders(main, 1, row, 1, 10)
    _auto_fit_columns(main, money_cols=(4, 5, 6, 7, 8, 9))

    _build_unmatched_sheet(wb, match_result["unmatched"])
    # "Nəzərə alınmayan əməliyyatlar" Uyğunsuz-dan sonra, xülasədən əvvəl;
    # yalnız ≥1 sətir olduqda yaradılır.
    if reversals:
        _build_reversals_sheet(wb, reversals)
    if removed_invoices:
        _build_removed_invoices_sheet(wb, removed_invoices)
    # "Banklar üzrə xülasə" ən SONA (index-ə bağlı köhnə testlər pozulmasın)
    _build_bank_summary_sheet(
        wb, bank_summary or [], sheet_title, warnings or []
    )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def generate_creditors_workbook(
    gelen_raw: bytes, depozit_raw: bytes, banks: Any
) -> bytes:
    """Kreditorlar workbook-u. `banks` tək bayt və ya (ad, bayt) siyahısı."""
    _clear_load_rows_cache()
    gelen_rows, removed_invoices, inv_warnings = _filter_invoices(parse_gelen(gelen_raw))
    removed_refs = frozenset(
        (str(r.get("seriya") or "").strip().upper(), str(r.get("nomre") or "").strip())
        for r in removed_invoices
    )
    depozit_rows = parse_depozit(depozit_raw, direction="expense")
    bank_rows, summary, bank_warnings = _parse_banks(banks, direction="debit")
    warnings = list(inv_warnings) + list(bank_warnings)
    # Ad-only (VÖEN-siz) sətirləri ad üzrə VÖEN-ə bağla (match_by_voen-dən əvvəl)
    name_pairs = [(r["ad"], r["voen"]) for r in gelen_rows]
    name_pairs += parse_depozit_names(depozit_raw)
    stat = resolve_adonly_payments(
        bank_rows, name_pairs, translit=_TRANSLIT_AD_HELLI_AKTIV.get("kreditor")
    )
    for label, (ok, ok_sum, miss, miss_sum) in stat.items():
        warnings.append(
            f"{label}: {ok} əməliyyat ad üzrə uyğunlaşdırıldı (VÖEN yoxdur), "
            f"{miss} əməliyyat uyğunlaşmadı."
        )
    # Səhv köçürmə / geri qaytarma ləğvi (uyğunlaşmadan ƏVVƏL).
    # Kreditorlarda ödəniş sayılan istiqamət MƏXARİC-dir.
    reversals = _apply_reversals(
        depozit_rows, bank_rows, depozit_raw, banks, target_direction="mexaric"
    )
    # Çıxarılmış (passiv/AA) qaiməyə istinad edən ödənişlər ödəniş sayılmır
    ref_map = _qaime_ref_map(gelen_rows)
    # Avans zənciri KREDİTORLAR üçün default BAĞLIDIR; yalnız açar açıqdırsa
    # (test/konfiq) tətbiq olunur. Bağlı halda nəticələr dəyişmir.
    avans_chain = None
    if _AVANS_ZENCIRI_AKTIV.get("kreditor"):
        avans_chain = set()
        for g in gelen_rows:
            a_ser = str(g.get("avans_seriya") or "").strip().upper()
            a_nom = str(g.get("avans_nomre") or "").strip()
            if a_ser and a_nom:
                avans_chain.add((g.get("voen") or "", a_ser, a_nom))
    depozit_rows, dz_removed = _drop_removed_invoice_payments(
        depozit_rows, ref_map, removed_refs, "Depozit", avans_chain=avans_chain
    )
    bank_rows, bk_removed = _drop_removed_invoice_payments(
        bank_rows, ref_map, removed_refs, "Bank", avans_chain=avans_chain
    )
    reversals = reversals + dz_removed + bk_removed
    matched = match_by_voen(gelen_rows, depozit_rows, bank_rows)
    removed_meta_by_voen: dict[str, dict[tuple[str, str], dict[str, Any]]] = {}
    if avans_chain is not None:
        for r in removed_invoices:
            v = r.get("voen")
            if not v:
                continue
            key = (
                str(r.get("seriya") or "").strip().upper(),
                str(r.get("nomre") or "").strip(),
            )
            removed_meta_by_voen.setdefault(v, {})[key] = r
    for s in matched["suppliers"]:
        s["_alloc"] = _allocate_supplier(
            s,
            removed_refs,
            direction="kreditor",
            removed_meta=removed_meta_by_voen.get(s["voen"], {}),
        )
        # `_allocate_supplier`-də tutulan çıxarılmış (AA) ödənişlər də məruzəyə
        # düşməlidir (əks halda xülasedə "Uyğunlaşan" vərəqlərdən böyük olur).
        reversals = reversals + s["_alloc"].get("removed", [])
    warnings.append(
        f"{len(reversals)} əməliyyat nəzərə alınmadı (səhv köçürmə/geri qaytarma)."
    )
    _finalize_bank_summary(summary, matched, reversals)
    return build_workbook(
        matched,
        bank_summary=summary,
        warnings=warnings,
        reversals=reversals,
        removed_invoices=removed_invoices,
    )


DEBITOR_MAIN_SHEET_NAME = "DEBİTORLAR"
_DEBITOR_ACIQLAMA = "Mal material satışı"


def generate_debitors_workbook(
    geden_raw: bytes, depozit_raw: bytes, banks: Any
) -> bytes:
    """Debitorlar (Gedən) workbook-u. `banks` tək bayt və ya (ad, bayt) siyahısı.

    GEDƏN (göndərilən qaimələr) − [bank kredit + depozit mədaxil] = Qalıq.
    KREDİTORLAR ilə eyni infrastruktur, istiqamət tərsinə.
    """
    _clear_load_rows_cache()
    geden_rows, removed_invoices, inv_warnings = _filter_invoices(parse_gelen(geden_raw))
    removed_refs = frozenset(
        (str(r.get("seriya") or "").strip().upper(), str(r.get("nomre") or "").strip())
        for r in removed_invoices
    )
    depozit_rows = parse_depozit(depozit_raw, direction="income")
    # Öz şirkət VÖEN-i/adı ilə daxil olan bank köçürmələri ödəniş sayılmır.
    self_voens, self_names = _extract_self_identity(geden_raw)
    bank_rows, summary, bank_warnings = _parse_banks(
        banks,
        direction="credit",
        self_voens=self_voens,
        self_names=self_names,
    )
    warnings = list(inv_warnings) + list(bank_warnings)
    # Ad-only (VÖEN-siz) sətirləri ad üzrə VÖEN-ə bağla (match_by_voen-dən əvvəl)
    name_pairs = [(r["ad"], r["voen"]) for r in geden_rows]
    name_pairs += parse_depozit_names(depozit_raw)
    stat = resolve_adonly_payments(
        bank_rows, name_pairs, translit=_TRANSLIT_AD_HELLI_AKTIV.get("debitor")
    )
    for label, (ok, ok_sum, miss, miss_sum) in stat.items():
        warnings.append(
            f"{label}: {ok} əməliyyat ad üzrə uyğunlaşdırıldı (VÖEN yoxdur), "
            f"{miss} əməliyyat uyğunlaşmadı."
        )
    # Səhv köçürmə / geri qaytarma ləğvi (simmetrik; ödəniş = MƏDAXİL).
    reversals = _apply_reversals(
        depozit_rows, bank_rows, depozit_raw, banks, target_direction="medaxil"
    )
    ref_map = _qaime_ref_map(geden_rows)
    # Avans zənciri olan çıxarılmış AA ödənişləri burada çıxarılmır (aşağıda
    # bölüşdürmə mərhələsində MT qaiməyə yazılır).
    avans_chain: set[tuple[str, str, str]] | None = None
    if _AVANS_ZENCIRI_AKTIV.get("debitor"):
        avans_chain = set()
        for g in geden_rows:
            a_ser = str(g.get("avans_seriya") or "").strip().upper()
            a_nom = str(g.get("avans_nomre") or "").strip()
            if a_ser and a_nom:
                avans_chain.add((g.get("voen") or "", a_ser, a_nom))
    depozit_rows, dz_removed = _drop_removed_invoice_payments(
        depozit_rows, ref_map, removed_refs, "Depozit", istiqamet="Mədaxil",
        avans_chain=avans_chain,
    )
    bank_rows, bk_removed = _drop_removed_invoice_payments(
        bank_rows, ref_map, removed_refs, "Bank", istiqamet="Mədaxil",
        avans_chain=avans_chain,
    )
    reversals = reversals + dz_removed + bk_removed
    matched = match_by_voen(geden_rows, depozit_rows, bank_rows)
    # Avans zənciri üçün çıxarılmış AA qaimələrinin VÖEN üzrə metadata-sı.
    removed_meta_by_voen: dict[str, dict[tuple[str, str], dict[str, Any]]] = {}
    for r in removed_invoices:
        v = r.get("voen")
        if not v:
            continue
        key = (str(r.get("seriya") or "").strip().upper(), str(r.get("nomre") or "").strip())
        removed_meta_by_voen.setdefault(v, {})[key] = r
    for s in matched["suppliers"]:
        s["_alloc"] = _allocate_supplier(
            s,
            removed_refs,
            direction="debitor",
            removed_meta=removed_meta_by_voen.get(s["voen"], {}),
        )
        reversals = reversals + s["_alloc"].get("removed", [])
    warnings.append(
        f"{len(reversals)} əməliyyat nəzərə alınmadı (səhv köçürmə/geri qaytarma)."
    )
    _finalize_bank_summary(summary, matched, reversals)
    return build_workbook(
        matched,
        sheet_title=DEBITOR_MAIN_SHEET_NAME,
        back_label="← DEBİTORLAR",
        aciqlama=_DEBITOR_ACIQLAMA,
        bank_summary=summary,
        warnings=warnings,
        reversals=reversals,
        removed_invoices=removed_invoices,
    )
