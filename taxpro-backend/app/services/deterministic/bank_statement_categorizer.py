"""Deterministik bank çıxarışı kateqoriyalaşdırması + hesabat.

AI/ML yoxdur — yalnız sabit, konfiqurasiya edilə bilən pattern-matching.
Giriş: meta sətirləri + dinamik başlıq sətri olan .xlsx/.csv çıxarış.
Çıxış: kateqoriyalaşdırılmış .xlsx (orijinal hücrələr dəyişməz; Kateqoriya
sütunu doldurulur və onun dərhal sağında sətir-üzrə abs(debit) dəyəri ilə
"Hesablanan Məbləğ" sütunu əlavə olunur) + ayrıca "Xülasə" sheet-i + summary dict.

Kateqoriyalar (yalnız 3, yoxsa boş qalır):
  - "Nağd çəkim komissiyası"  (öncəlik 1: KOMİSİYA-ya düşməməli)
  - "KOMİSİYA"                (Com./COMM- prefixi və ya "KOM" ehtivası)
  - "BANK FAİZ XƏRCİ"
Qaydalar CATEGORY_RULES ilə konfiqurasiya edilir; sıra önəmlidir.
"""

import csv
import io
from typing import Any

from openpyxl import Workbook, load_workbook

from app.services.deterministic.excel_parser import detect_format, normalize_text

# Başlıq sətrinin tapılması üçün mütləq tələb olunan sütunlar.
_REQUIRED_HEADER_FIELDS = ("emeliyyatdetallari", "debit")

# Bilinən sütun başlıqları (normalize olunmuş) → sahə adı.
_STATEMENT_COLUMN_FIELDS = {
    "emeliyyattarixi": "emeliyyat_tarixi",
    "heyatakecmetarixi": "heyata_kechme_tarixi",
    "gonderencialicimelumatlari": "gonderen_alici",
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
        ["kom"],  # DK KOM, HESAB ACILIS KOM, KREDIT TAHSIS KOM, KOMISSIYA...
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

_SUMMARY_SHEET_NAME = "Xülasə"

# Kateqoriya sütununun dərhal sağında əlavə olunan sətir-üzrə məbləğ sütunu.
_AMOUNT_COLUMN_TITLE = "Hesablanan Məbləğ"


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if value != value:  # NaN
            return ""
    except Exception:
        pass
    return str(value).strip()


def _match_category(detail: Any) -> str | None:
    """Əməliyyat Detalları mətnini qayda sırası ilə yoxlayır."""
    text = _cell_text(detail)
    lowered = text.lower()
    if lowered.startswith(_KOMISIYA_PREFIXES):
        return "KOMİSİYA"

    norm = normalize_text(text)
    if not norm:
        return None

    for category, patterns in CATEGORY_RULES:
        for pattern in patterns:
            if pattern in norm:
                return category
    return None


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
    return {name: {"cem": 0.0, "say": 0} for name, _ in CATEGORY_RULES}


def _add_summary_sheet(wb: Workbook, summary: dict) -> None:
    ws = wb.create_sheet(_SUMMARY_SHEET_NAME)
    ws.append(["Kateqoriya", "Əməliyyat sayı", "Cəm (Debit, abs)"])
    total_say = 0
    total_cem = 0.0
    for category, entry in summary.items():
        say = int(entry["say"])
        cem = round(float(entry["cem"]), 2)
        ws.append([category, say, cem])
        total_say += say
        total_cem += cem
    ws.append(["CƏMİ", total_say, round(total_cem, 2)])


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
    ws = wb.active

    rows = [[cell.value for cell in row] for row in ws.iter_rows()]
    found = _find_header_row(rows)
    if found is None:
        raise ValueError(
            "Başlıq sətri tapılmadı. 'Əməliyyat Detalları' və 'Debit' "
            "sütunları olan sətri axtarın."
        )
    header_row_idx, col_map = found

    details_col = col_map["emeliyyatdetallari"]  # 0-əsaslı
    debit_col = col_map["debit"]  # 0-əsaslı
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
        category = _match_category(detail)
        if category is None:
            continue
        ws.cell(row=row[0].row, column=cat_col).value = category
        entry = summary[category]
        entry["say"] = int(entry["say"]) + 1
        amount = _parse_amount(row[debit_col].value)
        if amount is not None:
            entry["cem"] = float(entry["cem"]) + abs(amount)
            ws.cell(row=row[0].row, column=amount_col).value = abs(amount)

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
            category = _match_category(detail)
            if category is None:
                continue
            ws.cell(row=r_idx + 1, column=cat_col_1, value=category)
            entry = summary[category]
            entry["say"] = int(entry["say"]) + 1
            debit_value = values[debit_col] if debit_col < len(values) else None
            amount = _parse_amount(debit_value)
            if amount is not None:
                entry["cem"] = float(entry["cem"]) + abs(amount)
                ws.cell(row=r_idx + 1, column=amount_col_1, value=abs(amount))

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