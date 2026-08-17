import re


def normalize_phone(value) -> str:
    """Azərbaycan telefonunu kanonik 9-rəqəmli formata çevirir.

    Portallar +994 prefiksini özləri əlavə etdiyi üçün saxlanan nömrə
    +994 və aparıcı 0 olmadan, yalnız milli nömrə (501234567, 703889246)
    şəklindədir. Tanınmayan format dəyişdirilmir.
    """
    if value is None:
        return ""
    digits = re.sub(r"\D", "", str(value))
    if digits.startswith("994") and len(digits) == 12:
        return digits[3:]
    if len(digits) == 10 and digits.startswith("0"):
        return digits[1:]
    if len(digits) == 9:
        return digits
    return str(value)
