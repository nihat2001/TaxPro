"""Deterministik bank çıxarışı kateqoriyalaşdırması + hesabat.

AI/ML yoxdur — yalnız sabit, konfiqurasiya edilə bilən pattern-matching.
Giriş: meta sətirləri + dinamik başlıq sətri olan .xlsx/.csv çıxarış.
Çıxış: kateqoriyalaşdırılmış .xlsx (orijinal hücrələr dəyişməz; Kateqoriya
sütunu doldurulur və onun dərhal sağında sətir-üzrə məbləğ (debit və ya
kredit, abs) ilə "İcmal" sütunu əlavə olunur) + ayrıca "Xülasə" sheet-i +
summary dict.

Kateqoriyalar (prioritet sırası ilə, yoxsa boş qalır):
  1. "Nağd çəkim komissiyası"  (öncəlik 1: KOMİSİYA-ya düşməməli)
  2. "KOMİSİYA"                (Com./COMM- prefixi, "KOM" və ya "COMMISSION")
  3. "BANK FAİZ XƏRCİ"
  4. "ƏDV ÖDƏNİŞİ"             (debit + EDV/ƏDV)
  5. "RÜSUM"                   (debit + RUSUM/RÜSUM)
  6. "VERGİ MƏXARİC"           (debit + dövlət/vergi açar sözü)
  7. "MAAŞ MƏXARİC"            (debit + maaş pattern-i)
  8. "TƏSİSÇİDƏN BORC"         (kredit + NAGD YATIRMA/CASH DEPOSIT + tesisciden)
  9. "DİGƏR ŞİRKƏT DAXİL OLMALAR" (kredit + şirkət adı)
  10. "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"   (debit + şirkət adı)
1-3 detail pattern-ləri CATEGORY_RULES ilə; 4-10 üçün VAT_PATTERNS,
DUTY_PATTERNS, TAX_AUTHORITY_KEYWORDS, SALARY_PATTERNS, _FOUNDER_*,
_COMPANY_SUFFIXES sabitləri istifadə olunur. Sıra önəmlidir — daha
spesifik olanlar əvvəl yoxlanır.
"""

import csv
import io
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.services.deterministic.excel_parser import detect_format, normalize_text

# Başlıq sətrinin tapılması üçün mütləq tələb olunan sütunlar.
_REQUIRED_HEADER_FIELDS = ("emeliyyatdetallari", "debit")

# Bilinən sütun başlıqları (normalize olunmuş) → sahə adı.
_STATEMENT_COLUMN_FIELDS = {
    "emeliyyattarixi": "emeliyyat_tarixi",
    "heyatakecmetarixi": "heyata_kechme_tarixi",
    "gonderenalicimelumatlari": "gonderen_alici",
    "emeliyyatdetallari": "emeliyyat_detallari",
    "emeliyyatno": "emeliyyat_no",
    "kredit": "kredit",
    "debit": "debit",
    "balans": "balans",
    "kateqoriya": "kateqoriya",
}

# Kateqoriya qaydaları — sıra vacibdir.
# Hər sətir: (kateqoriya, normalize edilmiş pattern siyahısı).
# Pattern-lər normalize edilmiş Əməliyyat Detalları daxilində axtarılır.
# Türk və Azərbaycan orfoqrafik variantları hər ikisi daxildir.
CATEGORY_RULES: list[tuple[str, list[str]]] = [
    (
        "Nağd çəkim komissiyası",
        [
            "nakitcekmekomisyonu",   # NAKIT CEKME KOMISYONU
            "nagdcekimkomisyonu",    # NAĞD ÇƏKİM KOMİSYONU
            "1faiznagdcekim",        # 1 FAIZ NAGD CEKIM / 1 FAİZ NAĞD ÇƏKİM
        ],
    ),
    (
        "KOMİSİYA",
        ["kom", "commission"],  # DK KOM, ..., PAYMENT ORDER COMMISSION
    ),
    (
        "BANK FAİZ XƏRCİ",
        [
            "taksitfaiz",   # Taksit Faiz
            "faizcerime",   # Faiz+Cərimə
        ],
    ),
]

# KOMİSİYA üçün başlanğıc (prefix) qaydaları — normalize etmədən yoxlanır.
_KOMISIYA_PREFIXES = ("com.", "comm-")

# ===== Kateqoriya konfiqurasiyaları (genişlənə bilən) =====

# Dövlət/vergi orqanı açar sözləri — VERGİ MƏXARİC üçün (sender və ya detail).
TAX_AUTHORITY_KEYWORDS: list[str] = [
    "İCBARI TIBBI SIGORTA",
    "İŞSİZLİK",
    "ISSIZLIK",
    "DSMF",
    "MDSS",
    "MƏRK.SOS.SIĞ",
    "MERK.SOS.SIG",
    "MƏRKƏZİ SOS",
    "VN MİLLİ GƏLİR",
    "MİLLİ GƏLİR DEPARTAMENTİ",
    "DÖVLƏT RÜSUMU",
    "DOVLET RUSUMU",
]

# MAAŞ MƏXARİC üçün normalize edilmiş detail pattern-ləri.
SALARY_PATTERNS: list[str] = ["maasodemesi", "maaspaket", "emekhaqq"]

# ƏDV ÖDƏNİŞİ üçün normalize pattern-ləri (EDV/ƏDV/EDV ODENILIR/ƏDV ÖDƏNİLİR).
VAT_PATTERNS: list[str] = ["edv"]

# RÜSUM üçün normalize pattern-ləri (RUSUM/RÜSUM/DOVLET RUSUMU/DÖVLƏT RÜSUMU).
DUTY_PATTERNS: list[str] = ["rusum"]

# TƏSİSÇİDƏN BORC: sender token-ləri + detail token-i (diakritiksiz).
_FOUNDER_SENDER_TOKENS: tuple[str, ...] = ("nagdyatirma", "cashdeposit")
_FOUNDER_DETAIL_TOKEN = "tesisciden"

# Şirkət adı suffiksləri (normalize) — DİGƏR ŞİRKƏT DAXİL OLMALAR/ÖDƏNİŞLƏRİ.
_COMPANY_SUFFIXES: tuple[str, ...] = (
    "mmc",
    "qsc",
    "asc",
    "ojsc",
    "sirketi",
    "ltd",
)

# Ümumi kateqoriya sırası (prioritet) — summary və Xülasə üçün.
_CATEGORY_ORDER: list[str] = [name for name, _ in CATEGORY_RULES] + [
    "ƏDV ÖDƏNİŞİ",
    "RÜSUM",
    "VERGİ MƏXARİC",
    "MAAŞ MƏXARİC",
    "TƏSİSÇİDƏN BORC",
    "DİGƏR ŞİRKƏT DAXİL OLMALAR",
    "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ",
]

_SUMMARY_SHEET_NAME = "Xülasə"

# Kateqoriya sütununun dərhal sağında əlavə olunan sətir-üzrə məbləğ sütunu.
_AMOUNT_COLUMN_TITLE = "İcmal"

# ===== Formatlaşdırma =====
_HEADER_FILL = PatternFill(
    start_color="E8E8E8", end_color="E8E8E8", fill_type="solid"
)
_HEADER_FONT = Font(bold=True)
_MONEY_FORMAT = "#,##0.00"
_MONEY_ALIGNMENT = Alignment(horizontal="right")
_TOP_BORDER = Border(top=Side(style="thin", color="000000"))
_AUTO_FIT_MIN = 8
_AUTO_FIT_MAX = 50
_AUTO_FIT_PADDING = 2


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if value != value:  # NaN
            return ""
    except Exception:
        pass
    return str(value).strip()


def _match_category(sender: Any, detail: Any, kredit: Any, debit: Any) -> str | None:
    """Sətir kontekstini prioritet sırası ilə kateqoriyalaşdırır.

    Sıra: 1) Com./COMM- → KOMİSİYA; 2) CATEGORY_RULES (Nağd çəkim →
    KOMİSİYA → BANK FAİZ); 3) VERGİ MƏXARİC; 4) MAAŞ MƏXARİC;
    5) TƏSİSÇİDƏN BORC; 6) DİGƏR ŞİRKƏT DAXİL OLMALAR;
    7) DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ; 8) None.
    """
    text = _cell_text(detail)
    lowered = text.lower()
    if lowered.startswith(_KOMISIYA_PREFIXES):
        return "KOMİSİYA"

    norm_detail = normalize_text(text)
    for category, patterns in CATEGORY_RULES:
        for pattern in patterns:
            if pattern in norm_detail:
                return category

    debit_filled = _has_amount(debit)
    kredit_filled = _has_amount(kredit)

    if debit_filled and any(p in norm_detail for p in VAT_PATTERNS):
        return "ƏDV ÖDƏNİŞİ"

    if debit_filled and any(p in norm_detail for p in DUTY_PATTERNS):
        return "RÜSUM"

    if debit_filled and _has_tax_keyword(sender, norm_detail):
        return "VERGİ MƏXARİC"

    if debit_filled and any(p in norm_detail for p in SALARY_PATTERNS):
        return "MAAŞ MƏXARİC"

    if (
        kredit_filled
        and _sender_has(sender, _FOUNDER_SENDER_TOKENS)
        and _FOUNDER_DETAIL_TOKEN in norm_detail
    ):
        return "TƏSİSÇİDƏN BORC"

    if kredit_filled and _is_company_name(sender):
        return "DİGƏR ŞİRKƏT DAXİL OLMALAR"

    if debit_filled and _is_company_name(sender):
        return "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"

    return None


def _has_amount(value: Any) -> bool:
    """Debit/Kredit sütununun 'dolu' olmasını (pul hərəkəti) yoxlayır."""
    amount = _parse_amount(value)
    return amount is not None and amount != 0


def _row_amount(debit: Any, kredit: Any) -> float | None:
    """Sətir-üzrə məbləğ: debit doludursa abs(debit), yoxsa abs(kredit)."""
    if _has_amount(debit):
        return abs(_parse_amount(debit))
    if _has_amount(kredit):
        return abs(_parse_amount(kredit))
    return None


def _is_company_name(sender: Any) -> bool:
    """Sender şirkət adı formatındadırmı? (case-insensitive, suffiks axtarışı)."""
    norm = normalize_text(_cell_text(sender))
    if not norm:
        return False
    return any(s in norm for s in _COMPANY_SUFFIXES)


def _has_tax_keyword(sender: Any, norm_detail: str) -> bool:
    """Dövlət/vergi açar sözü sender və ya detail-də keçirmi? (diakritiksiz)."""
    norm_sender = normalize_text(_cell_text(sender))
    for keyword in TAX_AUTHORITY_KEYWORDS:
        nkw = normalize_text(keyword)
        if nkw and (nkw in norm_detail or nkw in norm_sender):
            return True
    return False


def _sender_has(sender: Any, tokens) -> bool:
    norm_sender = normalize_text(_cell_text(sender))
    return any(t in norm_sender for t in tokens)


def _find_header_row(rows: list[list[Any]]) -> tuple[int, dict[str, int]] | None:
    """Dinamik başlıq sətrini tapır: həm sətir indeksini, həm də
    sütun xəritəsini qaytarır (normalize başlıq → 0-əsaslı sütun idx).
    """
    for idx, row in enumerate(rows):
        col_map: dict[str, int] = {}
        for col_idx, value in enumerate(row):
            key = normalize_text(_cell_text(value))
            if key:
                col_map[key] = col_idx
        if all(field in col_map for field in _REQUIRED_HEADER_FIELDS):
            return idx, col_map
    return None


def _parse_amount(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None
    text = (
        text.replace(" ", "")
        .replace("\u00a0", "")
        .replace("AZN", "")
        .replace("₼", "")
        .replace("man", "")
        .strip()
    )
    if not text:
        return None

    if "," in text:
        text = text.replace(".", "")  # minlik ayırıcı nöqtələri sil
        text = text.replace(",", ".")
    else:
        parts = text.split(".")
        if len(parts) > 2:  # çoxsaylı nöqtə → minlik ayırıcı
            text = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(text)
    except ValueError:
        return None


def _empty_summary() -> dict[str, dict[str, float | int]]:
    return {name: {"cem": 0.0, "say": 0} for name in _CATEGORY_ORDER}


def _style_header_row(ws, row_idx: int, last_col: int) -> None:
    """Başlıq sətrini bold + açıq boz fon ilə vurğulayır."""
    for col in range(1, last_col + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL


def _auto_fit_columns(ws) -> None:
    """Sütun enlərini ən uzun mətnə görə hesablayır (auto-fit)."""
    for col in range(1, ws.max_column + 1):
        max_len = _AUTO_FIT_MIN
        for row in range(1, ws.max_row + 1):
            value = ws.cell(row=row, column=col).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = min(max_len + _AUTO_FIT_PADDING, _AUTO_FIT_MAX)
        ws.column_dimensions[get_column_letter(col)].width = width


def _style_amount_column(ws, col: int, start_row: int, end_row: int) -> None:
    """Məbləğ sütununu #,##0.00 formatı ilə sağa düzləndirir."""
    for row in range(start_row, end_row + 1):
        cell = ws.cell(row=row, column=col)
        cell.number_format = _MONEY_FORMAT
        cell.alignment = _MONEY_ALIGNMENT


def _add_summary_sheet(wb: Workbook, summary: dict) -> None:
    ws = wb.create_sheet(_SUMMARY_SHEET_NAME)
    ws.append(["Kateqoriya", "Əməliyyat sayı", "Cəm (abs)"])
    total_say = 0
    total_cem = 0.0
    for category, entry in summary.items():
        say = int(entry["say"])
        if say == 0:
            continue  # Xülasə sheet-də yalnız faktiki mövcud kateqoriyalar
        cem = round(float(entry["cem"]), 2)
        ws.append([category, say, cem])
        total_say += say
        total_cem += cem
    ws.append(["CƏMİ", total_say, round(total_cem, 2)])

    # Başlıq bold + fon; Cəm sütunu formatla; CƏMİ sətri bold + yuxarı border.
    _style_header_row(ws, 1, 3)
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=3)
        cell.number_format = _MONEY_FORMAT
        cell.alignment = _MONEY_ALIGNMENT
    total_row = ws.max_row
    for col in range(1, 4):
        cell = ws.cell(row=total_row, column=col)
        cell.font = _HEADER_FONT
        cell.border = _TOP_BORDER
    _auto_fit_columns(ws)

    # Fayl açılanda Xülasə görünsün — defolt aktiv sheet.
    wb.active = wb.sheetnames.index(_SUMMARY_SHEET_NAME)


def _last_column(ws) -> int:
    """Vəziyyəti olan sonuncu sütunun 1-əsaslı indeksi."""
    max_col = 1
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                max_col = max(max_col, cell.column)
    return max_col


def _categorize_xlsx(raw_bytes: bytes) -> tuple[bytes, dict]:
    wb = load_workbook(io.BytesIO(raw_bytes))

    # Başlığı aktiv sheet ilə məhdudlaşdırmırıq: çıxışımız Xülasə-ni aktiv
    # edir — ilk başlığı olan sheet-i (sheet sırası ilə) işlədirik.
    ws = None
    header_row_idx = None
    col_map = None
    for sheet in wb.worksheets:
        rows = [[cell.value for cell in row] for row in sheet.iter_rows()]
        found = _find_header_row(rows)
        if found is not None:
            ws = sheet
            header_row_idx, col_map = found
            break
    if ws is None:
        raise ValueError(
            "Başlıq sətri tapılmadı. 'Əməliyyat Detalları' və 'Debit' "
            "sütunları olan sətri axtarın."
        )

    details_col = col_map["emeliyyatdetallari"]  # 0-əsaslı
    debit_col = col_map["debit"]  # 0-əsaslı
    sender_col = col_map.get("gonderenalicimelumatlari")
    kredit_col = col_map.get("kredit")
    if "kateqoriya" in col_map:
        cat_col = col_map["kateqoriya"] + 1  # 1-əsaslı
    else:
        cat_col = _last_column(ws) + 1  # 1-əsaslı; yeni sütun olaraq əlavə edilir
        ws.cell(row=header_row_idx + 1, column=cat_col, value="Kateqoriya")
    amount_col = cat_col + 1  # 1-əsaslı; Kateqoriya-nın dərhal sağı
    ws.cell(row=header_row_idx + 1, column=amount_col, value=_AMOUNT_COLUMN_TITLE)

    summary = _empty_summary()
    for row in ws.iter_rows(min_row=header_row_idx + 2):
        detail = row[details_col].value
        sender = row[sender_col].value if sender_col is not None else None
        kredit = row[kredit_col].value if kredit_col is not None else None
        debit = row[debit_col].value
        category = _match_category(sender, detail, kredit, debit)
        if category is None:
            continue
        ws.cell(row=row[0].row, column=cat_col).value = category
        entry = summary[category]
        entry["say"] = int(entry["say"]) + 1
        amount = _row_amount(debit, kredit)
        if amount is not None:
            entry["cem"] = float(entry["cem"]) + amount
            ws.cell(row=row[0].row, column=amount_col).value = amount

    _style_header_row(ws, header_row_idx + 1, amount_col)
    _style_amount_column(ws, amount_col, header_row_idx + 2, ws.max_row)
    _auto_fit_columns(ws)

    _add_summary_sheet(wb, summary)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), summary


def _decode_csv(raw_bytes: bytes) -> str:
    for enc in ("utf-8-sig", "cp1254", "latin-1"):
        try:
            return raw_bytes.decode(enc)
        except (UnicodeDecodeError, ValueError):
            continue
    return raw_bytes.decode("utf-8", errors="replace")


def _categorize_csv(raw_bytes: bytes) -> tuple[bytes, dict]:
    text = _decode_csv(raw_bytes)
    rows = [list(r) for r in csv.reader(io.StringIO(text))]
    if not rows:
        raise ValueError("Fayl boşdur.")

    found = _find_header_row(rows)
    if found is None:
        raise ValueError(
            "Başlıq sətri tapılmadı. 'Əməliyyat Detalları' və 'Debit' "
            "sütunları olan sətri axtarın."
        )
    header_row_idx, col_map = found

    details_col = col_map["emeliyyatdetallari"]
    debit_col = col_map["debit"]
    sender_col = col_map.get("gonderenalicimelumatlari")
    kredit_col = col_map.get("kredit")
    cat_col = col_map.get("kateqoriya")
    max_cols = max((len(r) for r in rows), default=0)
    if cat_col is None:
        cat_col = max_cols  # 0-əsaslı yeni sütun
    cat_col_1 = cat_col + 1  # 1-əsaslı
    amount_col_1 = cat_col_1 + 1  # 1-əsaslı; Kateqoriya-nın dərhal sağı
    target_cols = max(max_cols, amount_col_1)

    summary = _empty_summary()
    wb = Workbook()
    ws = wb.active
    ws.title = "Çıxarış"

    for r_idx, row in enumerate(rows):
        values = list(row)
        if len(values) < target_cols:
            values += [""] * (target_cols - len(values))
        ws.append(values)

        if r_idx == header_row_idx:
            if cat_col == max_cols:
                ws.cell(row=r_idx + 1, column=cat_col_1, value="Kateqoriya")
            ws.cell(row=r_idx + 1, column=amount_col_1, value=_AMOUNT_COLUMN_TITLE)

        if r_idx > header_row_idx:
            detail = values[details_col] if details_col < len(values) else ""
            sender = (
                values[sender_col]
                if sender_col is not None and sender_col < len(values)
                else ""
            )
            kredit = (
                values[kredit_col]
                if kredit_col is not None and kredit_col < len(values)
                else None
            )
            debit_value = values[debit_col] if debit_col < len(values) else None
            category = _match_category(sender, detail, kredit, debit_value)
            if category is None:
                continue
            ws.cell(row=r_idx + 1, column=cat_col_1, value=category)
            entry = summary[category]
            entry["say"] = int(entry["say"]) + 1
            debit_value = values[debit_col] if debit_col < len(values) else None
            amount = _row_amount(debit_value, kredit)
            if amount is not None:
                entry["cem"] = float(entry["cem"]) + amount
                ws.cell(row=r_idx + 1, column=amount_col_1, value=amount)

    _style_header_row(ws, header_row_idx + 1, amount_col_1)
    _style_amount_column(ws, amount_col_1, header_row_idx + 2, ws.max_row)
    _auto_fit_columns(ws)

    _add_summary_sheet(wb, summary)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), summary


def categorize_statement(raw_bytes: bytes) -> tuple[bytes, dict]:
    """Bank çıxarışını kateqoriyalaşdırır.

    Qaytarır: (kateqoriyalaşdırılmış .xlsx bytes, summary dict)
    Summary: {"KOMİSİYA": {"cem": float, "say": int}, ...} — hər 3 kateqoriya.
    Başlıq sətri tapılmazsa ValueError (→ endpoint 422).
    """
    fmt = detect_format(raw_bytes)
    if fmt == "xlsx":
        return _categorize_xlsx(raw_bytes)
    if fmt == "csv":
        return _categorize_csv(raw_bytes)
    raise ValueError(".xls faylları dəstəklənmir; yalnız .xlsx və .csv yükləyin.")