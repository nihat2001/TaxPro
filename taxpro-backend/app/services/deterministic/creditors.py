"""Kreditorlar/Debitorlar — "Gələn" (Kreditorlar) modulu.

3 mənbə faylı (GƏLƏN.xlsx, DEPOZİT.xlsx, BANK ÇIXARIŞI.xlsx) deterministik
alqoritmlə VÖEN üzrə uzlaşdırılır və TƏK nəticə workbook-u yaranır:

  - "KREDİTORLAR" əsas sheet (hər tədarükçü bir sətir, hiperlink + formula)
  - hər tədarükçü üçün ayrı sheet (qaimə/ödəniş sətirləri + SUM formulları)
  - "Uyğunsuz ödənişlər" sheet (GƏLƏN-də olmayan VÖEN-ə ödənişlər)

V1: stateless — DB-yə yazılmır; yalnız .xlsx generasiya olunur.
BANK ÇIXARIŞI parseri dinamik auto-detection ilə çoxbanklıdır (Kapital +
PAŞA; digər banklar alias siyahıları ilə genişləndirilə bilər).
AI/ML yoxdur.
"""

import io
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
_BANK_VOEN_ALIASES = (
    "voen",
    "terefinvoeni",
    "eksterefvoen",
    "ekshesab",
    "hesablasmahesabivoeen",
)
_BANK_DEBIT_ALIASES = ("mexaric", "debitmebleg")
_BANK_SIGNED_ALIASES = ("debit", "debet", "azndeyeri", "mebleg")
_BANK_DATE_ALIASES = ("emeliyyatintarixi", "emeliyyattarixi", "tarix")
_BANK_NAME_ALIASES = (
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
)
_BANK_BALANCE_TOKENS = ("balans", "dovrun", "movcud", "qaliq", "yekun", "cem")

_VOEN_RE = re.compile(r"^\d{7,12}$")
_IBAN_RE = re.compile(r"^[A-Za-z]{2}\d")
_VOEN_LABEL_RE = re.compile(r"(?:voen|vöen)\s*[:#]?\s*(\d{7,12})", re.IGNORECASE)
_BARE_VOEN_RE = re.compile(r"(?<!\d)(\d{10})(?!\d)")


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
    """Ödəyən/Benefisiar mətnindən adı təmizləyir (PAŞA `_x000D_` artefaktı).

    `/AZ03KABA..._x000D_CAUCASIAN SUPPLY..._x000D_/` → `CAUCASIAN SUPPLY...`.
    IBAN-a bənzər və boş seqmentlər atlanır; ilk mənalı seqment ad sayılır.
    """
    parts = [p.strip() for p in re.split(r"_x000D_|/", text) if p.strip()]
    for part in parts:
        if not _IBAN_RE.match(part):
            return part
    return ""


def _extract_voen_and_name(voen_cell: Any, name_cell: Any) -> tuple[str | None, str]:
    """Ağıllı VÖEN + ad çıxarışı: xalis rəqəm → birbaşa; qarışıq mətn → split."""
    voen_text = _cell_text(voen_cell)
    if _VOEN_RE.match(voen_text):
        return voen_text, _clean_counterparty_name(_cell_text(name_cell))
    voen, ad = _split_counter_account(voen_text)
    return voen, ad or _clean_counterparty_name(_cell_text(name_cell))


def _extract_voen_from_text(text: str) -> str | None:
    """Təyinat/əməliyyat mətnindən VÖEN-i regex-lə çıxarır.

    Əvvəlcə label forması (`VOEN: 1403610931`, `VÖEN 9900001231`), sonra boş
    dayanan 10 rəqəmli pattern. IBAN substring-i (bitişik rəqəmlər) tutulmur.
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
        name_col = _first_col(keys, key_to_col, _BANK_NAME_ALIASES)
        desc_col = _first_col(keys, key_to_col, _BANK_DESC_ALIASES)

        if date_col is None:
            continue
        if debit_col is None and signed_col is None:
            continue

        col_map: dict[str, int] = {"date": date_col}
        if voen_col is not None:
            col_map["voen"] = voen_col
        if debit_col is not None:
            col_map["debit"] = debit_col
        if signed_col is not None:
            col_map["signed"] = signed_col
        if name_col is not None:
            col_map["name"] = name_col
        if desc_col is not None:
            col_map["desc"] = desc_col
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


def parse_depozit(raw: bytes) -> list[dict[str, Any]]:
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
            mexaric = _to_float(_get(row, col_map, "mexaric")) or 0.0
            if mexaric <= 0:
                continue  # yalnız VÖEN-ə ödənilən ƏDV (Məxaric)
            result.append(
                {
                    "voen": voen,
                    "ad": _cell_text(_get(row, col_map, "ad")),
                    "tarix": _format_date(_get(row, col_map, "tarix")),
                    "mexaric": mexaric,
                }
            )
        return result
    raise ValueError(
        "DEPOZİT faylında başlıq sətri tapılmadı (VÖEN + Məxaric gözlənilir)."
    )


def parse_bank_statement(raw: bytes) -> list[dict[str, Any]]:
    """Çoxbanklı bank çıxarışı parseri (dinamik auto-detection).

    Kapital (VÖEN qarışıq mətndə + mənfi Məbləğ/AZN dəyəri), PAŞA
    (VÖEN ayrıca sütunda + müsbət Məxaric), Yapı Kredi "standard"
    (VÖEN yoxdur + mənfi Debit), ABB (VÖEN ayrıca sütunda + işarəsiz
    müsbət Debit məbləğ) və Unibank (VÖEN ayrıca sütunda + müsbət Məxaric,
    məbləğlər string, boş xana "-") formatlarını tanıyır. Balans/summary
    sətirləri atlanır. Yalnız debit sətirlər qaytarılır — bunlar tədarükçüyə
    edilən Əsas (principal) ödənişlərdir. VÖEN sütunu mövcud olub dəyəri boş
    olan debit sətirləri (bank haqqı/kart sifarişi — daxili əməliyyat) atlanır.
    """
    rows_list = _load_rows(raw)
    for rows in rows_list:
        found = _detect_bank_header(rows)
        if found is None:
            continue
        header_idx, col_map = found
        has_voen_col = "voen" in col_map
        result: list[dict[str, Any]] = []
        for row in rows[header_idx + 1 :]:
            name_cell = _get(row, col_map, "name")
            desc_cell = _get(row, col_map, "desc")
            voen_cell = _get(row, col_map, "voen")
            name_text = _cell_text(name_cell)
            desc_text = _cell_text(desc_cell)
            voen_raw = _cell_text(voen_cell)
            combined = normalize_text(name_text + " " + desc_text + " " + voen_raw)
            if _is_summary_row(combined):
                continue  # balans/yekun/cəm summary sətirləri

            if "debit" in col_map:
                amount = _to_float(_get(row, col_map, "debit"))
                if amount is None or amount <= 0:
                    continue
            else:
                amount = _to_float(_get(row, col_map, "signed"))
                if amount is None or amount >= 0:
                    continue
                amount = abs(amount)

            voen, ad = _extract_voen_and_name(voen_cell, name_cell)
            if voen is None and has_voen_col and voen_raw == "":
                continue  # VÖEN sütunu mövcud, dəyər boş → daxili əməliyyat
            if voen is None:
                voen = _extract_voen_from_text(desc_text)
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
        if b["voen"] and b["voen"] in suppliers:
            suppliers[b["voen"]]["odenis_esas"].append(
                {"tarix": b["tarix"], "mebleg": b["mebleg"]}
            )
        else:
            unmatched.append(
                {
                    "voen": b["voen"] or "",
                    "ad": b["ad"],
                    "tarix": b["tarix"],
                    "menbe": "Bank",
                    "mebleg": b["mebleg"],
                }
            )
    for d in depozit_rows:
        if d["voen"] and d["voen"] in suppliers:
            suppliers[d["voen"]]["odenis_edv"].append(
                {"tarix": d["tarix"], "mebleg": d["mexaric"]}
            )
        else:
            unmatched.append(
                {
                    "voen": d["voen"] or "",
                    "ad": d["ad"],
                    "tarix": d["tarix"],
                    "menbe": "Depozit",
                    "mebleg": d["mexaric"],
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


def _build_supplier_sheet(wb: Workbook, s: dict[str, Any], sheet_name: str) -> int:
    ws = wb.create_sheet(sheet_name)

    ws["A1"] = s["ad"]
    _set_link(ws["A1"], MAIN_SHEET_NAME + "!A1", display=s["ad"])
    ws.merge_cells("A1:J1")
    voen_val: Any = s["voen"]
    if voen_val.isdigit() and not voen_val.startswith("0"):
        voen_val = int(voen_val)
    ws["A2"] = voen_val
    ws.merge_cells("A2:J2")
    k2 = ws["K2"]
    k2.value = "← KREDİTORLAR"
    _set_link(k2, MAIN_SHEET_NAME + "!A1", display="← KREDİTORLAR")

    ws["A4"] = "Date"
    ws["B4"] = "Elektron Qaimə"
    ws["D4"] = "Ödəniş tarixi"
    ws["E4"] = "Hesablanıb"
    ws["G4"] = "Ödənilib"
    ws["I4"] = "Qalıq"
    ws["E5"] = "Əsas"
    ws["F5"] = "ƏDV"
    ws["G5"] = "Əsas"
    ws["H5"] = "ƏDV"
    ws["I5"] = "Əsas"
    ws["J5"] = "ƏDV"
    for rng in ("A4:A5", "B4:C5", "D4:D5", "E4:F4", "G4:H4", "I4:J4"):
        ws.merge_cells(rng)

    qaimeler = sorted(s["qaimeler"], key=lambda q: q["tarix"])
    events = [(p["tarix"], p["mebleg"], 0.0) for p in s["odenis_esas"]] + [
        (p["tarix"], 0.0, p["mebleg"]) for p in s["odenis_edv"]
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
            tarix, esas, edv = events.pop(0)
            ws.cell(row=row, column=4, value=tarix)
            if esas:
                ws.cell(row=row, column=7, value=esas)
            if edv:
                ws.cell(row=row, column=8, value=edv)
        ws.cell(row=row, column=9, value=f"=E{row}-G{row}")
        ws.cell(row=row, column=10, value=f"=F{row}-H{row}")
        row += 1

    for tarix, esas, edv in events:
        ws.cell(row=row, column=4, value=tarix)
        if esas:
            ws.cell(row=row, column=7, value=esas)
        if edv:
            ws.cell(row=row, column=8, value=edv)
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

    _style_header_row(ws, 4, 10)
    _style_header_row(ws, 5, 10)
    _style_money(ws, 6, sum_row, [5, 6, 7, 8, 9, 10])
    _apply_text_alignment(ws, 6, sum_row, [1, 2, 3, 4])
    _apply_borders(ws, 4, sum_row, 1, 10)
    _auto_fit_columns(ws, money_cols=(5, 6, 7, 8, 9, 10))
    return sum_row


def _build_unmatched_sheet(wb: Workbook, unmatched: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(UNMATCHED_SHEET_NAME)
    ws.append(["VÖEN", "Ad", "Tarix", "Mənbə", "Məbləğ"])
    for u in unmatched:
        ws.append([u["voen"], u["ad"], u["tarix"], u["menbe"], u["mebleg"]])
    _style_header_row(ws, 1, 5)
    _style_money(ws, 2, max(ws.max_row, 2), [5])
    _apply_text_alignment(ws, 2, max(ws.max_row, 2), [1, 2, 3, 4])
    _apply_borders(ws, 1, max(ws.max_row, 2), 1, 5)
    _auto_fit_columns(ws, money_cols=(5,))


def build_workbook(match_result: dict[str, Any]) -> bytes:
    wb = Workbook()
    main = wb.active
    main.title = MAIN_SHEET_NAME

    main["A1"] = "№"
    main["B1"] = "KREDİTORLAR"
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

    used: set[str] = set()
    row = 3
    for i, s in enumerate(match_result["suppliers"], start=1):
        sheet_name = unique_sheet_name(s["ad"], s["voen"], used)
        sum_row = _build_supplier_sheet(wb, s, sheet_name)

        main.cell(row=row, column=1, value=i)
        target = f"{_quote_sheet(sheet_name)}!A1"
        b_cell = main.cell(row=row, column=2, value=s["ad"])
        _set_link(b_cell, target, display=s["ad"])
        main.cell(row=row, column=3, value=_ACIQLAMA)
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

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def generate_creditors_workbook(
    gelen_raw: bytes, depozit_raw: bytes, bank_raw: bytes
) -> bytes:
    gelen_rows = parse_gelen(gelen_raw)
    depozit_rows = parse_depozit(depozit_raw)
    bank_rows = parse_bank_statement(bank_raw)
    matched = match_by_voen(gelen_rows, depozit_rows, bank_rows)
    return build_workbook(matched)
