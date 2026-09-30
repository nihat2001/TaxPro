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

import hashlib
import io
import os
import re
from datetime import date, datetime
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from app.services.deterministic.excel_parser import normalize_text

# ===== Sheet / sütun sabitləri =====
MAIN_SHEET_NAME = "KREDİTORLAR"
BANK_SUMMARY_SHEET_NAME = "Banklar üzrə xülasə"
# Çoxlu bank çıxarışı: maksimum fayl sayı (endpoint səviyyəsində yoxlanılır)
MAX_BANK_FILES = 20
# Depozit ödənişlərinin "Bank" sütununda göstərilən mənbə etiketi
DEPOZIT_BANK_LABEL = "Depozit (ƏDV)"
# Bank preamble ailələri üçün oxunaqlı etiketlər
_FAMILY_LABELS = {"vtb": "VTB", "atb": "ATB", "ziraat": "Ziraat"}
UNMATCHED_SHEET_NAME = "Uyğunsuz ödənişlər"
_ACIQLAMA = "Mal material satışı"

# Excel sheet adı qadağanları (spec: : \ / ? * [ ]) + formula/hiperlink
# təhlükəsizliyi üçün əlavə " ' simvolları da atılır.
_INVALID_SHEET_CHARS = set(':\\/?*[]\'"')
_SHEET_NAME_MAX = 31

# ===== Formatlaşdırma =====
_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
_HEADER_FONT = Font(bold=True)
_MONEY_FORMAT = "#,##0.00"
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
    "malinedvsizumumideyeri": "esas",
    "malinedvmeblegi": "edv",
}
_GELEN_REQUIRED = ("voen", "qaimetarixi")

# DEPOZİT.xlsx: başlıq 10-11-ci sətirdə (VÖEN, Vergi ödəyicisinin adı,
# Əməliyyat tarixi, Mədaxil, Məxaric). "Məxaric" = VÖEN-ə ödənilən ƏDV.
_DEPOZIT_FIELDS = {
    "voen": "voen",
    "vergioddeyicisininadi": "ad",
    "emeliyyattarixi": "tarix",
    "medaxil": "medaxil",
    "mexaric": "mexaric",
}
_DEPOZIT_REQUIRED = ("voen", "mexaric")

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


def _load_rows(raw: bytes) -> list[list[list[Any]]]:
    try:
        wb = load_workbook(io.BytesIO(raw))
    except Exception as exc:
        raise ValueError(f"Fayl oxunmadı (Excel gözlənilir): {exc}")
    return [[[cell.value for cell in row] for row in ws.iter_rows()] for ws in wb.worksheets]


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
            result.append(
                {
                    "voen": voen,
                    "ad": ad,
                    "tarix": _format_date(_get(row, col_map, "tarix")),
                    "seriya": _cell_text(_get(row, col_map, "seriya")),
                    "nomre": _cell_text(_get(row, col_map, "nomre")),
                    "esas": _to_float(_get(row, col_map, "esas")) or 0.0,
                    "edv": _to_float(_get(row, col_map, "edv")) or 0.0,
                }
            )
        return result
    raise ValueError(
        "GƏLƏN faylında başlıq sətri tapılmadı (VÖEN + Qaimə tarixi gözlənilir)."
    )


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
        result: list[dict[str, Any]] = []
        for row in rows[header_idx + 1 :]:
            voen = _voen_str(_get(row, col_map, "voen"))
            if not voen:
                continue
            amount = _to_float(_get(row, col_map, column)) or 0.0
            if amount <= 0:
                continue
            result.append(
                {
                    "voen": voen,
                    "ad": _cell_text(_get(row, col_map, "ad")),
                    "tarix": _format_date(_get(row, col_map, "tarix")),
                    "mexaric": amount,
                }
            )
        return result
    raise ValueError(
        "DEPOZİT faylında başlıq sətri tapılmadı (VÖEN + Məxaric/Mədaxil gözlənilir)."
    )


def parse_bank_statement(
    raw: bytes, direction: str = "debit"
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
        self_voens = _detect_self_voens(preamble)
        family = _detect_preamble_family(preamble)
        is_atb = family == "atb"
        is_ziraat = family == "ziraat"
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
            result.append(
                {
                    "voen": voen,
                    "ad": ad,
                    "tarix": _format_date(_get(row, col_map, "date")),
                    "mebleg": amount,
                }
            )
        return result
    raise ValueError(
        "Bank çıxarışında başlıq sətri tapılmadı "
        "(VÖEN + Tarix + Məbləğ/Məxaric sütunları gözlənilir)."
    )


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
    banks: Any, direction: str
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
                    "qeyd": "Dublikat, sayılmadı",
                }
            )
            continue
        seen_hashes[digest] = filename or f"bank_{idx}"

        label = _bank_source_label(filename, raw, used_labels)
        try:
            rows = parse_bank_statement(raw, direction=direction)
        except ValueError as exc:
            raise ValueError(f"'{filename or label}' faylı oxunmadı: {exc}")

        total = 0.0
        for r in rows:
            r["bank"] = label
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
                "qeyd": "",
            }
        )

    return combined, summary, warnings


def _finalize_bank_summary(
    summary: list[dict[str, Any]],
    matched: dict[str, Any],
) -> None:
    """Uyğunlaşan/uyğunsuz məbləğləri bank etiketi üzrə doldurur (hesablama dəyişmir)."""
    by_label = {r["bank"]: r for r in summary if r.get("qeyd") != "Dublikat, sayılmadı"}
    for s in matched["suppliers"]:
        for p in s["odenis_esas"]:
            label = p.get("bank", "")
            if label in by_label:
                by_label[label]["uygunlasan"] += p["mebleg"]
    for u in matched["unmatched"]:
        if u.get("menbe") != "Bank":
            continue  # Depozit uyğunsuzları bank xülasəsinə düşmür
        label = u.get("bank", "")
        if label in by_label:
            by_label[label]["uygunsuz"] += u["mebleg"]
    for r in summary:
        r["uygunlasan"] = round(r["uygunlasan"], 4)
        r["uygunsuz"] = round(r["uygunsuz"], 4)


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
    for b in bank_rows:
        # "bank" YALNIZ etiketdir — uyğunlaşdırma açarına qatılmır.
        bank_label = b.get("bank", "")
        if b["voen"] and b["voen"] in suppliers:
            suppliers[b["voen"]]["odenis_esas"].append(
                {"tarix": b["tarix"], "mebleg": b["mebleg"], "bank": bank_label}
            )
        else:
            unmatched.append(
                {
                    "voen": b["voen"] or "",
                    "ad": b["ad"],
                    "tarix": b["tarix"],
                    "menbe": "Bank",
                    "mebleg": b["mebleg"],
                    "bank": bank_label,
                }
            )
    for d in depozit_rows:
        if d["voen"] and d["voen"] in suppliers:
            suppliers[d["voen"]]["odenis_edv"].append(
                {"tarix": d["tarix"], "mebleg": d["mexaric"], "bank": DEPOZIT_BANK_LABEL}
            )
        else:
            unmatched.append(
                {
                    "voen": d["voen"] or "",
                    "ad": d["ad"],
                    "tarix": d["tarix"],
                    "menbe": "Depozit",
                    "mebleg": d["mexaric"],
                    "bank": DEPOZIT_BANK_LABEL,
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
    unmatched.sort(key=lambda u: (u["menbe"], u["tarix"]))
    return {"suppliers": supplier_list, "unmatched": unmatched}


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
    ws["E5"] = "Əsas"
    ws["F5"] = "ƏDV"
    ws["G5"] = "Əsas"
    ws["H5"] = "ƏDV"
    ws["I5"] = "Əsas"
    ws["J5"] = "ƏDV"
    for rng in ("A4:A5", "B4:C5", "D4:D5", "E4:F4", "G4:H4", "I4:J4", "K4:K5"):
        ws.merge_cells(rng)

    qaimeler = sorted(s["qaimeler"], key=lambda q: q["tarix"])
    # Sıralama sabit (stable): eyni tarixdə əvvəl əsas, sonra ƏDV ödənişləri
    events = [
        (p["tarix"], p["mebleg"], 0.0, p.get("bank", "")) for p in s["odenis_esas"]
    ] + [
        (p["tarix"], 0.0, p["mebleg"], p.get("bank", "")) for p in s["odenis_edv"]
    ]
    events.sort(key=lambda e: e[0])

    row = 6
    for i, q in enumerate(qaimeler):
        ws.cell(row=row, column=1, value=q["tarix"])
        ws.cell(row=row, column=2, value=q["seriya"])
        ws.cell(row=row, column=3, value=q["nomre"])
        ws.cell(row=row, column=5, value=q["esas"])
        ws.cell(row=row, column=6, value=q["edv"])
        if i == 0 and events:
            tarix, esas, edv, bank = events.pop(0)
            ws.cell(row=row, column=4, value=tarix)
            if esas:
                ws.cell(row=row, column=7, value=esas)
            if edv:
                ws.cell(row=row, column=8, value=edv)
            if bank:
                ws.cell(row=row, column=11, value=bank)
        ws.cell(row=row, column=9, value=f"=E{row}-G{row}")
        ws.cell(row=row, column=10, value=f"=F{row}-H{row}")
        row += 1

    for tarix, esas, edv, bank in events:
        ws.cell(row=row, column=4, value=tarix)
        if esas:
            ws.cell(row=row, column=7, value=esas)
        if edv:
            ws.cell(row=row, column=8, value=edv)
        if bank:
            ws.cell(row=row, column=11, value=bank)
        ws.cell(row=row, column=9, value=f"=E{row}-G{row}")
        ws.cell(row=row, column=10, value=f"=F{row}-H{row}")
        row += 1

    last = row - 1
    sum_row = row
    if last >= 6:
        for col in "EFGHIJ":
            ws[f"{col}{sum_row}"] = f"=SUM({col}6:{col}{last})"
    else:
        for col in "EFGHIJ":
            ws[f"{col}{sum_row}"] = f"=SUM({col}6:{col}6)"

    _style_header_row(ws, 4, 11)
    _style_header_row(ws, 5, 11)
    _style_money(ws, 6, sum_row, [5, 6, 7, 8, 9, 10])
    _apply_text_alignment(ws, 6, sum_row, [1, 2, 3, 4, 11])
    _apply_borders(ws, 4, sum_row, 1, 11)
    _auto_fit_columns(ws, money_cols=(5, 6, 7, 8, 9, 10))
    return sum_row


def _build_unmatched_sheet(wb: Workbook, unmatched: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(UNMATCHED_SHEET_NAME)
    ws.append(["VÖEN", "Ad", "Tarix", "Mənbə", "Məbləğ", "Bank"])
    for u in unmatched:
        ws.append(
            [
                u["voen"],
                u["ad"],
                u["tarix"],
                u["menbe"],
                u["mebleg"],
                u.get("bank", ""),
            ]
        )
    _style_header_row(ws, 1, 6)
    _style_money(ws, 2, max(ws.max_row, 2), [5])
    _apply_text_alignment(ws, 2, max(ws.max_row, 2), [1, 2, 3, 4, 6])
    _apply_borders(ws, 1, max(ws.max_row, 2), 1, 6)
    _auto_fit_columns(ws, money_cols=(5,))


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
            "Qeyd",
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
                r.get("qeyd", ""),
            ]
        )

    first_data = 2
    last_data = ws.max_row
    total_row = last_data + 1
    ws.cell(row=total_row, column=1, value="ÜMUMİ")
    if last_data >= first_data:
        for col in (3, 4, 5, 6):
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

    _style_header_row(ws, 1, 7)
    _style_money(ws, first_data, total_row, [4, 5, 6])
    _apply_text_alignment(ws, first_data, total_row, [1, 2, 7])
    _apply_borders(ws, 1, total_row, 1, 7)
    _auto_fit_columns(ws, money_cols=(4, 5, 6))


def build_workbook(
    match_result: dict[str, Any],
    sheet_title: str = MAIN_SHEET_NAME,
    back_label: str = "← KREDİTORLAR",
    aciqlama: str = _ACIQLAMA,
    bank_summary: list[dict[str, Any]] | None = None,
    warnings: list[str] | None = None,
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
    used: set[str] = {UNMATCHED_SHEET_NAME, BANK_SUMMARY_SHEET_NAME}
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
    gelen_rows = parse_gelen(gelen_raw)
    depozit_rows = parse_depozit(depozit_raw, direction="expense")
    bank_rows, summary, warnings = _parse_banks(banks, direction="debit")
    matched = match_by_voen(gelen_rows, depozit_rows, bank_rows)
    _finalize_bank_summary(summary, matched)
    return build_workbook(matched, bank_summary=summary, warnings=warnings)


DEBITOR_MAIN_SHEET_NAME = "DEBİTORLAR"
_DEBITOR_ACIQLAMA = "Mal material satışı"


def generate_debitors_workbook(
    geden_raw: bytes, depozit_raw: bytes, banks: Any
) -> bytes:
    """Debitorlar (Gedən) workbook-u. `banks` tək bayt və ya (ad, bayt) siyahısı.

    GEDƏN (göndərilən qaimələr) − [bank kredit + depozit mədaxil] = Qalıq.
    KREDİTORLAR ilə eyni infrastruktur, istiqamət tərsinə.
    """
    geden_rows = parse_gelen(geden_raw)
    depozit_rows = parse_depozit(depozit_raw, direction="income")
    bank_rows, summary, warnings = _parse_banks(banks, direction="credit")
    matched = match_by_voen(geden_rows, depozit_rows, bank_rows)
    _finalize_bank_summary(summary, matched)
    return build_workbook(
        matched,
        sheet_title=DEBITOR_MAIN_SHEET_NAME,
        back_label="← DEBİTORLAR",
        aciqlama=_DEBITOR_ACIQLAMA,
        bank_summary=summary,
        warnings=warnings,
    )
