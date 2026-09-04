"""Deterministik bank çıxarışı kateqoriyalaşdırması + hesabat.

AI/ML yoxdur — yalnız sabit, konfiqurasiya edilə bilən pattern-matching.
Giriş: meta sətirləri + dinamik başlıq sətri olan .xlsx/.csv çıxarış.
Dəstəklənən formatlar (avtomatik tanınır, _FORMAT_PROFILES):
  - standard: Əməliyyat Detalları + Debit/Kredit (YapıKredi/AccessBank)
  - kapital:  № | Tarix | ... | Məbləğ (işarəli) | ... | Təyinat
  - yelo:     çoxsətirli header, Mədaxil/Məxaric ayrı, Ödəyən/alan sender
  - access:   Header 2-ci sətirdə, Mədaxil/Məxaric ayrı,
              Göndərənin/Alanın Adı sender (ad + IBAN birlikdə)
  - vtb:      Header 11-ci sətirdə, Debet/Kredit ayrı, kiril
              'Комиссия за срочность' → KOMİSİYA (raw lowered yoxlanır)
  - turan:    Header 13-cü sətirdə, tək 'Məbləğ' (mənfi=məxaric,
              müsbət=mədaxil); sender yoxdur → catch-all sender tələb etmir
  - ziraat:   Header 13-cü sətirdə, Debet/Kredit ayrı, sender = Müxbirin adı;
              yekun (CƏMİ/Çıxış qalıq) sətirləri atlanır
  - atb:      Header 12-ci sətirdə, Debit/Kredit ayrı, Müxbir sender;
              'Xidmet Haqqi' yalnız debit-də KOMİSİYA (kredit → DAXİL);
              yekun (Müddət üzrə dövriyyə) sətirləri atlanır
  - unibank:  Header 13-cü sətirdə, Mədaxil/Məxaric ayrı, Tərəfdaş sender;
              sender-siz catch-all (KASSA, kredit ödənişləri);
              'konvertasiya/xidmet haqqi' yalnız debit-də KOMİSİYA;
              yekun (Yekun (Summary)) sətirləri atlanır
  - rabite:   Header 10-cu sətirdə, Debet/Kredit ayrı, sender = Qarşı tərəfin adı;
              sender 'Haqq və komissiya gəlirləri' debit-də KOMİSİYA
              (detail 'Emek haqqi' olanlar MAAŞ-a qalır); yekun sətirləri atlanır
  - bankofbaku: Header 12-ci sətirdə, Profit/Expenditure, sender = Debet Account;
              xüsusi pattern yoxdur — ümumi zəncir kifayətdir; boş təyinat/son meta
              sətirləri atlanır
  - pasha:    Header 13-cü sətirdə, Mədaxil/Məxaric ayrı, sender = Ödəyən/Benefisiar
              ('_x000D_' artefaktı ilə — 2-ci seqment); 'xidmet haqqi' debit-də
              KOMİSİYA, maaş-guard olmadan ('EMEK HAQQI SIFARISI UZRE XIDMET HAQQI');
              balans/POS yığım sətirləri atlanır
Qlobal: 'KOMİTƏ' ('KOMITESI') 'KOM' substring-i ilə KOMİSİYA-ya düşmür;
              'GOMRUK'/'GÖMRÜK' → RÜSUM (DUTY_PATTERNS)
  - bankrespublika: Header 13-cü sətirdə, Kredit/Debet ayrı, sender = Göndərənin adı/Alanın adı;
              'xidmet haqqi' debit-də KOMİSİYA; 'faiz VƏ kredit'/'cərimə VƏ kredit' → BANK FAİZ;
              detail 'tesisciden' → TƏSİSÇİDƏN BORC (sender tələbi yox);
              'əsas borc' (faiz/cəriməsiz) → kateqoriyasız; MASGULLUQ AGENTLIYI → VERGİ
  - abb:       Header 42-ci sətirdə, Debit/Kredit ayrı, sender = Əks hesab üzrə məlumatlar;
              'Standing payment order' (kredit, sender boş) → DAXİL OLMALAR;
              'sadeleshmis' + 'negdmexaric' → Nağd çəkim komissiyası
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
  8. "TƏSİSÇİDƏN BORC"         (sender NAGD YATIRMA/CASH DEPOSIT — kredit və ya debit)
  9. "Hesab Möhkəmlənməsi"     (hesablar arası köçürmə / hesabın möhkəmləndirilməsi)
  10. "Nağdlaşdırma"           (müqavilə üzrə kredit məxaric — Nağd ÇƏKİMDƏN AYRI)
  11. "Konvertasiya"           (alış əməliyyatı / valyuta / əskinaz / konvertasiya)
  12. "Bank xərci"             (ay ödənişi / aylıq ödəniş)
  13. "Divident ödəniş"        (divident / dividend)
  14. "DİGƏR ŞİRKƏT DAXİL OLMALAR" (catch-all: sender boş deyil + kredit)
  15. "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"   (catch-all: sender boş deyil + debit)
  16. "Terminal-Komissiya"       (kapital: Pos-Terminal ∧ Xidmət haqqı; İcmal = fee)
1-3 detail pattern-ləri CATEGORY_RULES ilə; 4-10 üçün VAT_PATTERNS,
DUTY_PATTERNS, TAX_AUTHORITY_KEYWORDS, SALARY_PATTERNS, _FOUNDER_*
sabitləri istifadə olunur. Sıra önəmlidir — daha spesifik olanlar əvvəl
yoxlanır. Sender tamamilə boş olan sətirlər kateqoriyasız qalır.
"""

import csv
import io
import logging
import re
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.services.deterministic.excel_parser import detect_format, normalize_text

logger = logging.getLogger(__name__)

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

# ===== Format profilləri =====
# Hər bank çıxarışı öz sütun strukturuna görə ayrıca emal olunur.
# "required": başlıq sətrini tanımaq üçün zəruri normalize başlıqlar.
# "fields": normalize başlıq → kanonik sahə adı (col_map bununla qurulur).
_FORMAT_PROFILES = [
    {
        "name": "standard",  # YapıKredi / AccessBank / ümumi
        "required": ("emeliyyatdetallari", "debit"),
        "fields": {
            "emeliyyattarixi": "emeliyyat_tarixi",
            "heyatakecmetarixi": "heyata_kechme_tarixi",
            "gonderenalicimelumatlari": "gonderen_alici",
            "emeliyyatdetallari": "emeliyyat_detallari",
            "emeliyyatno": "emeliyyat_no",
            "kredit": "kredit",
            "debit": "debit",
            "balans": "balans",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "abb",  # ABB (Header 42-ci sətirdə; Debit/Kredit ayrı, sender = Əks hesab üzrə məlumatlar)
        "required": ("emeliyyatintarixi", "emeliyyatinreferensi", "ekshesabuzremelumatlar",
                     "debitmebleg", "kreditmebleg", "odenisinteyinati", "qarsiterefinvoeni"),
        "fields": {
            "emeliyyatintarixi": "emeliyyat_tarixi",
            "emeliyyatinreferensi": "emeliyyat_no",
            "ekshesabuzremelumatlar": "gonderen_alici",
            "debitmebleg": "mexaric",
            "kreditmebleg": "medaxil",
            "odenisinteyinati": "teyinat",
            "qarsiterefinvoeni": "gonderen_tax",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "turan",  # Turan Bank (Header 13-cü sətirdə; tək "Məbləğ", mənfi=məxaric)
        "required": ("hesabnomresi", "mebleg", "mebleginazndeyeri", "qaliq", "valyuta"),
        "fields": {
            "tarix": "emeliyyat_tarixi",
            "hesabnomresi": "hesab_no",
            "mebleg": "mebleg",
            "qaliq": "balans",
            "valyuta": "valyuta",
            "teyinat": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "kapital",  # Kapital Bank (№ | Tarix | ... | Məbləğ | ... | Təyinat)
        "required": ("tarix", "mebleg", "teyinat"),
        "fields": {
            "tarix": "tarix",
            "senedinnomresi": "sened_no",
            "ekshesabinnomresiadivevoeni": "eks_hesab",
            "mebleg": "mebleg",
            "valyuta": "valyuta",
            "azndeyeri": "azn_deyeri",
            "balans": "balans",
            "teyinat": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "yelo",  # Yelo Bank (çoxsətirli header; Mədaxil/Məxaric ayrı)
        "required": ("teyinat", "medaxil", "mexaric", "odeyenalan"),
        "fields": {
            "emeltarixitrdate": "emeliyyat_tarixi",
            "icratarixirealdate": "icra_tarixi",
            "teyinatdescription": "teyinat",
            "odeyenalanpayerreciever": "odeyen",
            "qarsihesabcounteraccount": "qarsi_hesab",
            "medaxilincome": "medaxil",
            "mexaricoutcome": "mexaric",
            "balansbalance": "balans",
            "kursrate": "kurs",
            "aznekv": "azn_ekv",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "access",  # Access Bank (Header 2-ci sətirdə; Mədaxil/Məxaric ayrı)
        "required": ("teyinat", "medaxil", "mexaric", "gondereninalanin"),
        "fields": {
            "emeliyyattarixipostingdate": "emeliyyat_tarixi",
            "emeliyyatinsitransactionid": "emeliyyat_no",
            "medaxilin": "medaxil",
            "mexaricout": "mexaric",
            "balansbalance": "balans",
            "gondereninalaninadisenderbeneficiaryname": "gonderen_alici",
            "gonderenalanintaxidsenderbeneficiarytaxid": "gonderen_tax",
            "teyinatdescription": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "vtb",  # VTB Bank (Header 11-ci sətirdə; Debet/Kredit ayrı)
        "required": ("qarsihesab", "debet", "kredit", "qaliq"),
        "fields": {
            "tarix": "emeliyyat_tarixi",
            "qarsihesab": "qarsi_hesab",
            "adi": "gonderen_alici",
            "voen": "gonderen_tax",
            "debet": "mexaric",
            "kredit": "medaxil",
            "qaliq": "balans",
            "teyinat": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "ziraat",  # Ziraat Bank (Header 13-cü sətirdə; Debet/Kredit ayrı, sender = Müxbirin adı)
        "required": ("muxbirin", "debet", "kredit", "araqaliq"),
        "fields": {
            "tarix": "emeliyyat_tarixi",
            "senedinnsi": "emeliyyat_no",
            "muxbirinhesabi": "qarsi_hesab",
            "muxbirinadi": "gonderen_alici",
            "debet": "mexaric",
            "kredit": "medaxil",
            "araqaliq": "balans",
            "teyinat": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "atb",  # ATB Bank (Header 12-ci sətirdə; Debit/Kredit ayrı, Müxbir sender)
        "required": ("valyutalasmatarixi", "tranzaksiyatarixi", "muxbir", "tesvir", "balans"),
        "fields": {
            "tranzaksiyatarixi": "emeliyyat_tarixi",
            "senednomresi": "emeliyyat_no",
            "muxbir": "gonderen_alici",
            "debit": "mexaric",
            "kredit": "medaxil",
            "balans": "balans",
            "tesvir": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "unibank",  # Unibank (Header 13-cü sətirdə; Mədaxil/Məxaric ayrı, Tərəfdaş sender)
        "required": ("emeliyyattarixi", "medaxil", "mexaric", "emeliyyatterefdasi", "teyinat"),
        "fields": {
            "emeliyyattarixi": "emeliyyat_tarixi",
            "transactionid": "emeliyyat_no",
            "emeliyyatterefdasi": "gonderen_alici",
            "mexaric": "mexaric",
            "medaxil": "medaxil",
            "balans": "balans",
            "teyinat": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "rabite",  # Rabitə Bank (Header 10-cu sətirdə; Debet/Kredit ayrı, sender = Qarşı tərəfin adı)
        "required": ("emeliyyattarixi", "silinmetarixi", "benefisiarinbanki", "qarsiteref", "debet", "kredit"),
        "fields": {
            "emeliyyattarixi": "emeliyyat_tarixi",
            "silinmetarixi": "icra_tarixi",
            "benefisiarinbanki": "benefisiar_banki",
            "qarsiterefinnomresi": "qarsi_hesab",
            "qarsiterefinadi": "gonderen_alici",
            "debet": "mexaric",
            "kredit": "medaxil",
            "balans": "balans",
            "tesvir": "teyinat",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "bankofbaku",  # Bank of Baku (Header 12-ci sətirdə; Profit/Expenditure, sender = Debet Account)
        "required": ("date", "debetaccount", "creditaccount", "description", "profit", "expenditure", "balance", "comission"),
        "fields": {
            "date": "emeliyyat_tarixi",
            "debetaccount": "gonderen_alici",
            "creditaccount": "qarsi_hesab",
            "description": "teyinat",
            "profit": "medaxil",
            "expenditure": "mexaric",
            "balance": "balans",
            "comission": "komissiya_col",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "pasha",  # Pasha Bank (Header 13-cü sətirdə; Mədaxil/Məxaric ayrı, sender = Ödəyən/Benefisiar)
        "required": ("icratarixi", "odeyenbenefisiar", "teyinat", "istinadno", "medaxil", "mexaric", "balans"),
        "fields": {
            "icratarixi": "emeliyyat_tarixi",
            "odeyenbenefisiar": "gonderen_alici",
            "teyinat": "teyinat",
            "istinadno": "istinad_no",
            "voen": "gonderen_tax",
            "medaxil": "medaxil",
            "mexaric": "mexaric",
            "balans": "balans",
            "kateqoriya": "kateqoriya",
        },
    },
    {
        "name": "bankrespublika",  # Bank Respublika (Header 13-cü sətirdə; Kredit/Debet ayrı, sender = Göndərənin adı/Alanın adı)
        "required": ("tarix", "senedinn", "gondereninadialaninadi", "voen", "hesab", "odenisinteyinati", "kredit", "debet", "qaliq"),
        "fields": {
            "tarix": "emeliyyat_tarixi",
            "senedinn": "emeliyyat_no",
            "gondereninadialaninadi": "gonderen_alici",
            "voen": "gonderen_tax",
            "hesab": "hesab",
            "odenisinteyinati": "teyinat",
            "kredit": "medaxil",
            "debet": "mexaric",
            "qaliq": "balans",
            "kateqoriya": "kateqoriya",
        },
    },
]

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

# KOMİSİYA üçün başlanğıc (prefix) qaydaları — normalize etmədən yoxlanır.
_KOMISIYA_PREFIXES = ("com.", "comm-")

# Profilə-xas KOMİSİYA əlavə açar sözləri (normalize). Yalnız həmin profildə.
_KOMISIYA_EXTRA_PATTERNS: dict[str, tuple[str, ...]] = {
    "kapital": ("xidmethaqqi",),
    "yelo": ("debbeodeme", "emeliyyatmesref", "xidmethaqqi"),
    "turan": ("xidmethaqqi", "cekkitabcasi", "hesabacilmasi"),
    "atb": ("xidmethaqqi",),
    "unibank": ("xidmethaqqi",),
    "bankrespublika": ("xidmethaqqi",),
}

# ATB: 'xidmet haqqi' yalnız DEBIT (məxaric) sətirlərində KOMİSİYA-dır;
# kredit (mədaxil) 'xidmet haqqi' catch-all-a (DAXİL OLMALAR) qalmalıdır.
# Unibank: 'konvertasiya'/'xidmet haqqi' eyni istiqamət qaydası ilə.
# Rabitə: sender 'Haqq və komissiya gəlirləri' debit-də KOMİSİYA-dır (detail-a yox).
# Pasha: 'xidmet haqqi' yalnız debit-də KOMİSİYA (kredit → DAXİL).
# Bank Respublika: 'xidmet haqqi' yalnız debit-də KOMİSİYA.
_KOMISIYA_EXTRA_DEBIT_ONLY = ("atb", "unibank", "rabite", "pasha", "bankrespublika")

# Rabitə: sender token-i — 'Haqq və komissiya gəlirləri' (normalize: komissiyagelir).
_KOMISIYA_SENDER_TOKEN = "komissiyagelir"

# Yelo: Nağd çəkim komissiyası üçün əlavə token.
_NAGD_CEKIM_YELO = "nagdlasdirmavergisi"

# Access: Nağd çəkim komissiyası üçün əlavə token ('Nağdlaşdırma üçün S/V').
_NAGD_CEKIM_ACCESS = "nagdlasdirma"

# Turan: Nağd çəkim komissiyası üçün əlavə token ('Nəğd vəsaitə görə 1 % vergi').
_NAGD_CEKIM_TURAN = "negdvesaitegore"

# Sender sütunu olmayan profillər üçün catch-all sender tələb etmir
# (Turan: debit → ÖDƏNİŞLƏRİ, kredit → DAXİL OLMALAR).
# Unibank-da çoxlu boş-sender əməliyyat var (KASSA, kredit ödənişləri, konvertasiya).
_CATCHALL_NO_SENDER = ("turan", "unibank")

# VTB: kiril 'комиссия' normalize_text ilə boşaldılır (yalnız [a-z0-9] saxlanır);
# buna görə KOMİSİYA raw (lowered) mətndə yoxlanılır ('Комиссия за срочность').
_KOMISIYA_CYRILLIC = "комиссия"

# ===== Yeni qlobal kateqoriyalar (bütün profillərdə; catch-all-dan ƏVVƏL yoxlanır) =====

# Hesab Möhkəmlənməsi: hesablar arası köçürmə / hesabın möhkəmləndirilməsi.
_HESAB_MOHKEMLENME_TOKENS = (
    "hesablararasikocurme",
    "hesabinmohkemlend",
    "hesabinmohkemlenmesi",
)

# Nağdlaşdırma: 'müqavilə üzrə kredit məxaric' — Nağd ÇƏKİM KOMİSSİYASINDAN AYRI.
_NAGD_LASDIRMA_TRIGGER = "muqavileuzrekreditmexaric"

# Konvertasiya: alış əməliyyatı / valyuta / əskinaz / konvertasiya.
_KONVERTASIYA_TOKENS = ("alisemeliyyati", "valyuta", "eskinas", "konvertasiya")

# Bank xərci: aylıq ödəniş (ay ödənişi / aylıq ödəniş / aylıq ödənişi).
_BANK_XERCI_TRIGGERS = ("ayodenisi", "ayliqodenis", "ayliqodenisi")

# Divident ödəniş: divident/dividend.
_DIVIDENT_TOKENS = ("divident", "dividend")

# Qlobal istisna: 'KOMİTƏ'/'KOMITESI' (komitə) sözü 'KOM' substring-i ilə
# səhvən KOMİSİYA-ya düşməməlidir (məs. 'DOVLET GOMRUK KOMITESI' → RÜSUM).
_KOMITE_TOKENS = ("komite",)

# Pasha: 'xidmət haqqı' → KOMİSİYA, maaş-guard olmadan.
# '266340 NOMRELI EMEK HAQQI SIFARISI UZRE XIDMET HAQQI' maaş ödənişi deyil —
# bankın 'əmək haqqı sifarişi' xidmətinə görə tutduğu komissiyadır.
_PASHA_KOMISIYA_TOKEN = "xidmethaqqi"

# ABB: Nağd çəkim komissiyası — 'sadələşmiş vergi' (sadeleshmis) + 'nağd məxaric'.
_ABB_NAGD_CEKIM_TOKENS = ("sadeleshmis", "negdmexaric")

# ABB: 'Standing payment order' (mütəmadi daxilolan ödəniş, sender boş) →
# kredit olduqda birbaşa DİGƏR ŞİRKƏT DAXİL OLMALAR (sender tələbi yoxdur).
_ABB_SPO_TOKEN = "standingpaymentorder"

# Bank of Baku: 'əmək haqqı' mətnli məbləğ ≤ bu hədd → komissiya (MAAŞ yox, KOMİSİYA).
_BOB_MAAS_COMMISSION_MAX = 45.0

# Yeni qlobal kateqoriyaların _CATEGORY_ORDER-da yeri (catch-all-dan əvvəl).
_NEW_GLOBAL_CATEGORIES = (
    "Hesab Möhkəmlənməsi",
    "Nağdlaşdırma",
    "Konvertasiya",
    "Bank xərci",
    "Divident ödəniş",
)

# Yelo: BANK FAİZ XƏRCİ üçün çevik yoxlama (faiz ∧ ödəmə).
_YELO_FAIZ_TOKEN = "faiz"
_YELO_ODEME_TOKEN = "odeme"

# Yelo: əsas borc ödənişi → kateqoriyasız (borc əsas ödəmə).
_YELO_ASAS_ODEME = "esasodeme"

# Terminal-Komissiya (kapital): Təyinat-da 'Pos-Terminal' VƏ 'Xidmət haqqı'.
_TERMINAL_FEE_RE = re.compile(r"[Xx]idm[əe]t haqq[ıi]:\s*-?([\d.,]+)\s*AZN")


def _match_komisiya_extra(profile: str, norm_detail: str) -> bool:
    """Profilə-xas KOMİSİYA əlavəsi (maaş-xaric).

    Maaş pattern-i ('emekhaqq'/'maasodemesi'/'maaspaket') varsa KOMİSİYA
    deyil — belə sətirlər MAAŞ MƏXARİC-ə qalır (reqressiya qorunur).
    """
    patterns = _KOMISIYA_EXTRA_PATTERNS.get(profile)
    if not patterns:
        return False
    if not any(p in norm_detail for p in patterns):
        return False
    if any(p in norm_detail for p in SALARY_PATTERNS):
        return False
    return True


def _match_terminal_komissiya(norm_detail: str) -> bool:
    """Terminal-Komissiya (yalnız kapital): Pos-Terminal ∧ Xidmət haqqı."""
    return "posterminal" in norm_detail and "xidmethaqqi" in norm_detail


def _extract_terminal_fee(detail: Any) -> float | None:
    """Təyinat mətnindən 'Xidmət haqqı: N AZN' rəqəmini çıxarır (abs).

    Regex uyğun gəlməzsə None qaytarır (çağıran abs(Məbləğ)-ə düşür).
    """
    text = _cell_text(detail)
    match = _TERMINAL_FEE_RE.search(text)
    if not match:
        return None
    raw = match.group(1)
    normalized = raw.replace(" ", "")
    if "," in normalized:
        normalized = normalized.replace(".", "").replace(",", ".")
    else:
        parts = normalized.split(".")
        if len(parts) > 2:
            normalized = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return abs(float(normalized))
    except ValueError:
        return None

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
    "VERGİ XİDMƏTİ",
    "VERGI XIDMETI",
    "VERGI XIDMATI",
    "GƏLİR VERGİSİ",
    "MƏNFƏƏT VERGİSİ",
    "MASGULLUQ AGENTLIYI",
]

# MAAŞ MƏXARİC üçün normalize edilmiş detail pattern-ləri.
SALARY_PATTERNS: list[str] = ["maasodemesi", "maaspaket", "emekhaqq"]

# ƏDV ÖDƏNİŞİ üçün normalize pattern-ləri (EDV/ƏDV/EDV ODENILIR/ƏDV ÖDƏNİLİR).
VAT_PATTERNS: list[str] = ["edv"]

# RÜSUM üçün normalize pattern-ləri (RUSUM/RÜSUM/DOVLET RUSUMU/DÖVLƏT RÜSUMU/GÖMRÜK).
DUTY_PATTERNS: list[str] = ["rusum", "gomruk"]

# TƏSİSÇİDƏN BORC: sender token-ləri (NAGD YATIRMA / CASH DEPOSIT).
_FOUNDER_SENDER_TOKENS: tuple[str, ...] = ("nagdyatirma", "cashdeposit")

# Ümumi kateqoriya sırası (prioritet) — summary və Xülasə üçün.
_CATEGORY_ORDER: list[str] = [name for name, _ in CATEGORY_RULES] + [
    "ƏDV ÖDƏNİŞİ",
    "RÜSUM",
    "VERGİ MƏXARİC",
    "MAAŞ MƏXARİC",
    "TƏSİSÇİDƏN BORC",
    *_NEW_GLOBAL_CATEGORIES,
    "DİGƏR ŞİRKƏT DAXİL OLMALAR",
    "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ",
    "Terminal-Komissiya",
]

_SUMMARY_SHEET_NAME = "Xülasə"

# 3-cü sheet: kateqoriyaları Mədaxil/Məxaric (kredit/debit) istiqaməti üzrə ayırır.
_MADAXIL_MAXARIC_SHEET_NAME = "Mədaxil və Məxaric"

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


def _cell_display(value: Any) -> str | None:
    """Hücrə dəyəri: boş/NaN → None, əks halda mətn."""
    text = _cell_text(value)
    return text if text else None


def _match_category(
    sender: Any,
    detail: Any,
    kredit: Any,
    debit: Any,
    profile: str = "standard",
    amount: float | None = None,
) -> str | None:
    """Sətir kontekstini prioritet sırası ilə kateqoriyalaşdırır.

    Sıra: 1) Com./COMM- → KOMİSİYA; 2) kapital 'Terminal-Komissiya'
    (Pos-Terminal ∧ Xidmət haqqı); 3) CATEGORY_RULES (Nağd çəkim →
    KOMİSİYA → BANK FAİZ) + profil-xas əlavələr (kapital 'xidmət haqqı',
    yelo 'DƏBBƏ ÖDƏMƏ/ƏMƏLİYYAT MƏSRƏFİ', yelo BANK FAİZ 'faiz∧odeme',
    yelo Nağd çəkim 'Nağdlaşdırma vergisi'); 4) yelo 'Əsas ödəmə' →
    kateqoriyasız; 5) ƏDV ÖDƏNİŞİ; 6) RÜSUM; 7) VERGİ MƏXARİC;
    8) MAAŞ MƏXARİC; 9) TƏSİSÇİDƏN BORC; 10) catch-all (DAXİL/ÖDƏNİŞLƏRİ).
    """
    text = _cell_text(detail)
    lowered = text.lower()
    if lowered.startswith(_KOMISIYA_PREFIXES):
        return "KOMİSİYA"

    norm_detail = normalize_text(text)

    # Terminal-Komissiya (kapital): Pos-Terminal ∧ Xidmət haqqı —
    # ümumi KOMİSİYA-dan əvvəl yoxlanılır (daha spesifik).
    if profile == "kapital" and _match_terminal_komissiya(norm_detail):
        return "Terminal-Komissiya"

    for i, (category, patterns) in enumerate(CATEGORY_RULES):
        if category == "Nağd çəkim komissiyası":
            matched = _match_nagd_cekim(norm_detail) or (
                profile == "yelo" and _NAGD_CEKIM_YELO in norm_detail
            ) or (
                profile == "access" and _NAGD_CEKIM_ACCESS in norm_detail
            ) or (
                profile == "turan" and _NAGD_CEKIM_TURAN in norm_detail
            ) or (
                profile == "abb"
                and _ABB_NAGD_CEKIM_TOKENS[0] in norm_detail
                and _ABB_NAGD_CEKIM_TOKENS[1] in norm_detail
            )
        elif category == "KOMİSİYA":
            extra = _match_komisiya_extra(profile, norm_detail)
            # Pasha: 'xidmət haqqı' → KOMİSİYA, maaş-guard olmadan
            # ('EMEK HAQQI SIFARISI UZRE XIDMET HAQQI' maaş ödənişi deyil).
            if profile == "pasha":
                extra = extra or _match_pasha_komisiya(norm_detail)
            if profile in _KOMISIYA_EXTRA_DEBIT_ONLY:
                extra = extra and _has_amount(debit)  # yalnız debit → KOMİSİYA
            # Rabitə: sender-də 'Haqq və komissiya gəlirləri' + debit → KOMİSİYA.
            # Detail 'Emek haqqi' olsa da (məs. 'Emek haqqı İyun \ 19.13 AZN') bu,
            # maaş köçürməsinin bank komissiyasıdır — maaş-guard tətbiq edilmir.
            if profile == "rabite" and _has_amount(debit):
                norm_sender_here = normalize_text(_cell_text(sender))
                if _KOMISIYA_SENDER_TOKEN in norm_sender_here:
                    extra = True
            # Qlobal istisna: 'KOMİTƏ' ('KOMITESI') sözü 'KOM' substring-i ilə
            # səhvən KOMİSİYA-ya düşməməli (məs. 'DOVLET GOMRUK KOMITESI' → RÜSUM).
            has_komite = any(k in norm_detail for k in _KOMITE_TOKENS)
            generic_kom = (
                "commission" in norm_detail
                if has_komite
                else any(p in norm_detail for p in patterns)
            )
            matched = (
                generic_kom
                or extra
                or (profile == "vtb" and _KOMISIYA_CYRILLIC in lowered)
            )
        elif category == "BANK FAİZ XƏRCİ":
            matched = any(p in norm_detail for p in patterns) or (
                profile == "yelo"
                and _YELO_FAIZ_TOKEN in norm_detail
                and _YELO_ODEME_TOKEN in norm_detail
            ) or (
                # Bank Respublika: çevik yoxlama — 'faiz' VƏ 'kredit', VƏ YA
                # 'cərimə' VƏ 'kredit' (KREDİTƏ GÖRƏ FAİZ / CƏRİMƏNİN ÖDƏNİŞİ).
                profile == "bankrespublika"
                and ("faiz" in norm_detail or "cerime" in norm_detail)
                and "kredit" in norm_detail
            )
        else:
            matched = any(p in norm_detail for p in patterns)
        if matched:
            return category

    # Yelo: əsas borc ödənişi → kateqoriyasız (catch-all-dan əvvəl).
    if profile == "yelo" and _YELO_ASAS_ODEME in norm_detail:
        return None

    debit_filled = _has_amount(debit)
    kredit_filled = _has_amount(kredit)

    if debit_filled and any(p in norm_detail for p in VAT_PATTERNS):
        return "ƏDV ÖDƏNİŞİ"

    if debit_filled and any(p in norm_detail for p in DUTY_PATTERNS):
        return "RÜSUM"

    if debit_filled and _has_tax_keyword(sender, norm_detail):
        return "VERGİ MƏXARİC"

    if debit_filled and any(p in norm_detail for p in SALARY_PATTERNS):
        # Bank of Baku: 'əmək haqqı' mətnli kiçik məbləğ (≤ 45 AZN) əsl maaş deyil,
        # maaş köçürməsinin bank komissiyasıdır → KOMİSİYA (yalnız bankofbaku).
        if profile == "bankofbaku" and amount is not None and amount <= _BOB_MAAS_COMMISSION_MAX:
            return "KOMİSİYA"
        return "MAAŞ MƏXARİC"

    # TƏSİSÇİDƏN BORC: sender-də NAGD YATIRMA/CASH DEPOSIT keçən hər sətir
    # (Kredit və ya Debit — hər iki istiqamət, tesisciden şərti yoxdur).
    if _sender_has(sender, _FOUNDER_SENDER_TOKENS) and (kredit_filled or debit_filled):
        return "TƏSİSÇİDƏN BORC"
    # Bank Respublika: detail-də 'tesisciden' → TƏSİSÇİDƏN BORC, sender-in
    # NAGD YATIRMA olması tələb olunmadan ('TESİSCİDEN MADAXİL' + şəxs sender).
    if profile == "bankrespublika" and "tesisciden" in norm_detail and (
        kredit_filled or debit_filled
    ):
        return "TƏSİSÇİDƏN BORC"

    # Yeni qlobal kateqoriyalar — spesifik qaydalardan sonra, catch-all-dan əvvəl.
    if any(t in norm_detail for t in _HESAB_MOHKEMLENME_TOKENS):
        return "Hesab Möhkəmlənməsi"
    if _NAGD_LASDIRMA_TRIGGER in norm_detail:
        return "Nağdlaşdırma"
    if any(t in norm_detail for t in _KONVERTASIYA_TOKENS):
        return "Konvertasiya"
    if any(t in norm_detail for t in _BANK_XERCI_TRIGGERS):
        return "Bank xərci"
    if any(t in norm_detail for t in _DIVIDENT_TOKENS):
        return "Divident ödəniş"

    # Bank Respublika: 'əsas borc' (faiz/cərimə sözü olmadan) → kateqoriyasız
    # (kredit əsas məbləğinin ödənişi, catch-all-a düşməməlidir; Yelo 'Əsas ödəmə'
    # presedenti). Faiz/cərimə ehtiva edən əsas borc yuxarıda BANK FAİZ-ə düşür.
    if (
        profile == "bankrespublika"
        and "esasborc" in norm_detail
        and "faiz" not in norm_detail
        and "cerime" not in norm_detail
    ):
        return None

    # ABB: 'Standing payment order' (mütəmadi daxilolan ödəniş, sender boş) →
    # kredit olduqda birbaşa DAXİL OLMALAR (catch-all-ın sender şərtini keçərsiz edir).
    if profile == "abb" and _ABB_SPO_TOKEN in norm_detail and kredit_filled:
        return "DİGƏR ŞİRKƏT DAXİL OLMALAR"

    # Catch-all fallback: sender boş deyilsə, əvvəlki heç bir qayda düşməyibsə.
    # Kredit dolu → DİGƏR ŞİRKƏT DAXİL OLMALAR, Debit dolu → DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ.
    # Sender sütunu olmayan profillər (_CATCHALL_NO_SENDER) üçün sender tələb olunmur.
    if _sender_nonempty(sender) or profile in _CATCHALL_NO_SENDER:
        if kredit_filled:
            return "DİGƏR ŞİRKƏT DAXİL OLMALAR"
        if debit_filled:
            return "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"

    return None


_NAGD_CEKIM_PATTERNS = [
    "nakitcekmekomisyonu",
    "nagdcekimkomisyonu",
    "1faiznagdcekim",
]

# Çevik variantlar: "1 faiz" VƏ (naglashdirma/nagd/cekim/cektirme) birlikdə.
_NAGD_CEKIM_FLEX_TOKEN = "1faiz"
_NAGD_CEKIM_FLEX_WITH = ("naglashdirma", "nagd", "cekim", "cektirme")


def _match_nagd_cekim(norm_detail: str) -> bool:
    """Nağd çəkim komissiyası: sabit pattern-lər VƏ ya çevik birləşmə."""
    if any(p in norm_detail for p in _NAGD_CEKIM_PATTERNS):
        return True
    if _NAGD_CEKIM_FLEX_TOKEN in norm_detail and any(
        t in norm_detail for t in _NAGD_CEKIM_FLEX_WITH
    ):
        return True
    return False


def _extract_sender_name(raw_cell: Any) -> str:
    """'IBAN / AD / VÖEN' formatından ad hissəsini ayırır.

    "/" ilə split; 3+ hissə → orta (index 1), 2 hissə → ikinci,
    başqa halda tam mətn (IBAN və ya ad).
    """
    text = _cell_text(raw_cell)
    if not text:
        return ""
    parts = [p.strip() for p in text.split("/") if p.strip()]
    if len(parts) >= 2:
        return parts[1]
    return text


def _extract_sender_name_pasha(raw_cell: Any) -> str:
    """Pasha sender-i: '_x000D_' (carriage-return artefaktı) ilə split edir.

    Format: 'IBAN_x000D_AD_x000D_' → ad hissəsi (adətən 2-ci seqment).
    Sadə mətn ('Pasha Bank') olduğu kimi saxlanılır.
    """
    text = _cell_text(raw_cell)
    if not text:
        return ""
    parts = [p.strip() for p in text.split("_x000D_") if p.strip()]
    if len(parts) >= 2:
        return parts[1]
    return text


def _match_pasha_komisiya(norm_detail: str) -> bool:
    """Pasha: 'xidmət haqqı' → KOMİSİYA, maaş-guard OLMAYAN məntiq.

    Kapital-dan fərqlidir: 'EMEK HAQQI SIFARISI UZRE XIDMET HAQQI' maaş
    ödənişi deyil — bankın 'əmək haqqı sifarişi' xidmətinə görə komissiyadır,
    ona görə mətndə emekhaqq olsa belə KOMİSİYA-ya düşməlidir.
    """
    return _PASHA_KOMISIYA_TOKEN in norm_detail


def _split_signed_amount(value: Any) -> tuple[str | None, str | None]:
    """İşarəli Məbləği (Kapital) debit/kredit-ə ayırır.

    Mənfi → (abs, None) [debit/pul çıxır]; müsbət → (None, abs) [kredit].
    """
    amount = _parse_amount(value)
    if amount is None or amount == 0:
        return None, None
    text = str(int(amount)) if amount == int(amount) else str(amount)
    if amount < 0:
        return text.replace("-", ""), None
    return None, text


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


def _sender_nonempty(sender: Any) -> bool:
    """Göndərən/Alıcı Məlumatları sahəsi boş deyilmi (catch-all üçün)."""
    return bool(_cell_text(sender))


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


def _find_header_row(rows: list[list[Any]]) -> tuple[int, dict[str, int], str] | None:
    """Dinamik başlıq sətrini format profilinə görə tapır.

    Qaytarır: (sətir idx, kanonik sahə → sütun idx xəritəsi, profil adı).
    Hər profil öz required başlıqları ilə tanınır; ilk uyğun gələn qalib gəlir.
    Uyğunluq SUBSTRING əsaslıdır: çoxsətirli hüceyrələr normalize_text ilə
    birləşdiyindən 'Təyinat/Description' → 'teyinatdescription' kimi açarlar
    yaranır və 'teyinat' substring-i onu tutur.
    """
    for idx, row in enumerate(rows):
        raw_map: dict[str, int] = {}
        for col_idx, value in enumerate(row):
            key = normalize_text(_cell_text(value))
            if key:
                raw_map[key] = col_idx
        raw_keys = list(raw_map)
        for profile in _FORMAT_PROFILES:
            if not all(any(req in k for k in raw_keys) for req in profile["required"]):
                continue
            col_map: dict[str, int] = {}
            for norm, field in profile["fields"].items():
                for key in raw_keys:
                    if norm in key:
                        col_map[field] = raw_map[key]
                        break
            return idx, col_map, profile["name"]
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


def _build_madaxil_maxaric_sheet(
    wb: Workbook,
    categorized_rows: list[tuple[str, str, float]],
    opening_balance: float | None,
    closing_balance: float | None,
    total_madaxil: float = 0.0,
    total_mexaric: float = 0.0,
):
    """'Mədaxil və Məxaric' sheet-i — kateqoriyalar istiqamət üzrə ayrılır.

    categorized_rows: (kateqoriya, istiqamət, məbləğ) — istiqamət 'kredit'/'debit'.
    Eyni kateqoriya hər iki istiqamətdə ola bilər → hər iki tərəfdə ayrı sətir.
    Yalnız faylda faktiki mövcud kateqoriyalar göstərilir (sıfır sətir əlavə edilmir).
    total_madaxil/total_mexaric: bütün kredit/debit sətirlərinin cəmi (unclassified daxil);
    verilməzsə categorized_rows cəmi istifadə olunur.
    """
    credits: dict[str, float] = {}
    debits: dict[str, float] = {}
    for category, direction, amount in categorized_rows:
        if direction == "kredit":
            credits[category] = credits.get(category, 0.0) + amount
        else:
            debits[category] = debits.get(category, 0.0) + amount

    if not total_madaxil and not total_mexaric:
        total_madaxil = round(sum(credits.values()), 2)
        total_mexaric = round(sum(debits.values()), 2)

    ws = wb.create_sheet(_MADAXIL_MAXARIC_SHEET_NAME)
    ws["A1"] = "İlkin qalıq:"
    ws["B1"] = opening_balance
    ws["D1"] = "Son qalıq:"
    ws["E1"] = closing_balance
    ws["A2"] = "MƏDAXİL:"
    ws["B2"] = round(total_madaxil, 2)
    ws["D2"] = "MƏXARİC:"
    ws["E2"] = round(total_mexaric, 2)

    # Başlıq sətri (sətir 4): sol blok kredit, sağ blok debit.
    ws["A4"] = "Təyinat"
    ws["B4"] = "Məbləğ"
    ws["D4"] = "Təyinat"
    ws["E4"] = "Məbləğ"

    credit_order = [c for c in _CATEGORY_ORDER if c in credits]
    debit_order = [c for c in _CATEGORY_ORDER if c in debits]

    row = 5
    for cat in credit_order:
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2, value=round(credits[cat], 2))
        row += 1
    row = 5
    for cat in debit_order:
        ws.cell(row=row, column=4, value=cat)
        ws.cell(row=row, column=5, value=round(debits[cat], 2))
        row += 1

    # Başlıq sətri bold + fon.
    for col in (1, 2, 4, 5):
        cell = ws.cell(row=4, column=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
    for col in (1, 2, 4, 5):
        for r in (1, 2):
            cell = ws.cell(row=r, column=col)
            if cell.value is not None:
                cell.font = _HEADER_FONT

    # Məbləğ sütunları (#,##0.00, sağa düzləndirmə) + avto-fit.
    _style_amount_column(ws, 2, 2, max(ws.max_row, 2))
    _style_amount_column(ws, 5, 2, max(ws.max_row, 2))
    _auto_fit_columns(ws)
    # Xülasə aktiv qalsın (create_sheet yeni sheet-i aktiv edir).
    if _SUMMARY_SHEET_NAME in wb.sheetnames:
        wb.active = wb.sheetnames.index(_SUMMARY_SHEET_NAME)
    return ws


def _collect_madaxil_maxaric_from_sheet(
    ws,
    header_row: int,
    cat_col: int,
    amount_col: int,
    profile: str,
    debit_col: int | None,
    kredit_col: int | None,
    mebleg_col: int | None,
    balans_col: int | None,
) -> tuple[list[tuple[str, str, float]], float, float, float | None, float | None]:
    """Esas sheet-dən kateqoriya+istiqamət+məbləğ, cəmlər və ilk/son balansı toplayır.

    finalize override-larından sonra 3-cü sheet-i yenidən qurmaq üçün.
    MƏDAXİL/MƏXARİC cəmləri bütün sətirləri (unclassified daxil) əhatə edir.
    """
    categorized_rows: list[tuple[str, str, float]] = []
    first_balance: float | None = None
    last_balance: float | None = None
    total_madaxil = 0.0
    total_mexaric = 0.0

    for row in ws.iter_rows(min_row=header_row + 1):
        category = row[cat_col - 1].value
        if category and category not in _CATEGORY_ORDER:
            continue
        if balans_col is not None:
            bal = _parse_amount(row[balans_col - 1].value)
            if first_balance is None and bal is not None:
                first_balance = bal
            if bal is not None:
                last_balance = bal
        if profile in ("kapital", "turan"):
            mebleg = row[mebleg_col - 1].value if mebleg_col else None
            debit_s, kredit_s = _split_signed_amount(mebleg)
            debit_amt = abs(_parse_amount(debit_s)) if debit_s else None
            kredit_amt = abs(_parse_amount(kredit_s)) if kredit_s else None
        else:
            debit_amt = (
                abs(_parse_amount(row[debit_col - 1].value))
                if debit_col and _has_amount(row[debit_col - 1].value)
                else None
            )
            kredit_amt = (
                abs(_parse_amount(row[kredit_col - 1].value))
                if kredit_col and _has_amount(row[kredit_col - 1].value)
                else None
            )
        if kredit_amt is not None:
            total_madaxil += kredit_amt
        if debit_amt is not None:
            total_mexaric += debit_amt
        if category:
            if kredit_amt is not None:
                categorized_rows.append((category, "kredit", kredit_amt))
            if debit_amt is not None:
                categorized_rows.append((category, "debit", debit_amt))
    return categorized_rows, total_madaxil, total_mexaric, first_balance, last_balance


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
    profile = None
    for sheet in wb.worksheets:
        rows = [[cell.value for cell in row] for row in sheet.iter_rows()]
        found = _find_header_row(rows)
        if found is not None:
            ws = sheet
            header_row_idx, col_map, profile = found
            break
    if ws is None:
        raise ValueError(
            "Başlıq sətri tapılmadı. 'Əməliyyat Detalları'+'Debit' və ya "
            "'Tarix'+'Məbləğ'+'Təyinat' sütunları olan sətri axtarın."
        )

    if profile == "kapital":
        teyinat_col = col_map["teyinat"]
        mebleg_col = col_map["mebleg"]
        eks_hesab_col = col_map.get("eks_hesab")
    elif profile == "turan":
        teyinat_col = col_map["teyinat"]
        mebleg_col = col_map["mebleg"]
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "ziraat":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "atb":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "unibank":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "rabite":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "bankofbaku":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "pasha":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "bankrespublika":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "abb":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "yelo":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        odeyen_col = col_map.get("odeyen")
    elif profile in ("access", "vtb"):
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
    else:
        details_col = col_map["emeliyyat_detallari"]  # 0-əsaslı
        debit_col = col_map["debit"]  # 0-əsaslı
        sender_col = col_map.get("gonderen_alici")
        kredit_col = col_map.get("kredit")
    if "kateqoriya" in col_map:
        cat_col = col_map["kateqoriya"] + 1  # 1-əsaslı
    else:
        cat_col = _last_column(ws) + 1  # 1-əsaslı; yeni sütun olaraq əlavə edilir
        ws.cell(row=header_row_idx + 1, column=cat_col, value="Kateqoriya")
    amount_col = cat_col + 1  # 1-əsaslı; Kateqoriya-nın dərhal sağı
    ws.cell(row=header_row_idx + 1, column=amount_col, value=_AMOUNT_COLUMN_TITLE)

    summary = _empty_summary()
    unclassified: list[dict] = []
    categorized_rows: list[tuple[str, str, float]] = []
    balans_col = col_map.get("balans")
    first_balance: float | None = None
    last_balance: float | None = None
    total_madaxil = 0.0
    total_mexaric = 0.0
    for row in ws.iter_rows(min_row=header_row_idx + 2):
        if balans_col is not None:
            bal = _parse_amount(row[balans_col].value)
            if first_balance is None and bal is not None:
                first_balance = bal
            if bal is not None:
                last_balance = bal
        if profile == "kapital":
            sender = (
                _extract_sender_name(row[eks_hesab_col].value)
                if eks_hesab_col is not None
                else ""
            )
            detail = row[teyinat_col].value
            debit, kredit = _split_signed_amount(row[mebleg_col].value)
        elif profile == "turan":
            tarix = row[tarix_col].value if tarix_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail):
                continue  # yekun/summary sətri (yalnız qalıq) — atla
            sender = ""  # sender sütunu yoxdur
            debit, kredit = _split_signed_amount(row[mebleg_col].value)
        elif profile == "ziraat":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or (
                not _cell_text(detail) and not _sender_nonempty(sender)
            ):
                continue  # yekun/summary sətri (CƏMİ / Çıxış qalıq) — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "atb":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail) or not _sender_nonempty(sender):
                continue  # yekun/summary sətri (Müddət üzrə dövriyyə / sonunda qalıq) — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "unibank":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail):
                continue  # yekun/summary sətri (Yekun (Summary)) — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "rabite":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail):
                continue  # yekun/summary sətri (cəmi) — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "bankofbaku":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail):
                continue  # boş whitespace təyinat / son meta sətirləri — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "pasha":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = (
                _extract_sender_name_pasha(row[sender_col].value)
                if sender_col is not None
                else ""
            )
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail):
                continue  # yekun/balans/POS yığım sətirləri — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "bankrespublika":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail):
                continue  # yekun/summary sətri — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "abb":
            tarix = row[tarix_col].value if tarix_col is not None else None
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            if not _cell_text(tarix) or not _cell_text(detail):
                continue  # yekun/summary sətri — atla
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile == "yelo":
            sender = row[odeyen_col].value if odeyen_col is not None else None
            detail = row[teyinat_col].value
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        elif profile in ("access", "vtb"):
            sender = row[sender_col].value if sender_col is not None else None
            detail = row[teyinat_col].value
            kredit = row[medaxil_col].value
            debit = row[mexaric_col].value
        else:
            detail = row[details_col].value
            sender = row[sender_col].value if sender_col is not None else None
            kredit = row[kredit_col].value if kredit_col is not None else None
            debit = row[debit_col].value
        if _has_amount(kredit):
            total_madaxil += abs(_parse_amount(kredit))
        if _has_amount(debit):
            total_mexaric += abs(_parse_amount(debit))
        category = _match_category(
            sender, detail, kredit, debit, profile, amount=_row_amount(debit, kredit)
        )
        if category is None:
            unclassified.append(
                {
                    "row_index": row[0].row,
                    "sender": _cell_display(sender) or "",
                    "detail": _cell_display(detail) or "",
                    "kredit": _cell_display(kredit),
                    "debit": _cell_display(debit),
                }
            )
            continue
        ws.cell(row=row[0].row, column=cat_col).value = category
        entry = summary[category]
        entry["say"] = int(entry["say"]) + 1
        amount = _row_amount(debit, kredit)
        if category == "Terminal-Komissiya":
            fee = _extract_terminal_fee(detail)
            if fee is not None:
                amount = fee
            else:
                logger.warning(
                    "Terminal-Komissiya sətirində 'Xidmət haqqı' tapılmadı; "
                    "abs(Məbləğ) istifadə olunur",
                    extra={"row": row[0].row},
                )
        if amount is not None:
            entry["cem"] = float(entry["cem"]) + amount
            ws.cell(row=row[0].row, column=amount_col).value = amount

        # Mədaxil/Məxaric sheet-i üçün kateqoriya + istiqamət toplanması.
        if _has_amount(kredit):
            categorized_rows.append(
                (category, "kredit", abs(_parse_amount(kredit)))
            )
        if _has_amount(debit):
            categorized_rows.append(
                (category, "debit", abs(_parse_amount(debit)))
            )

    _style_header_row(ws, header_row_idx + 1, amount_col)
    _style_amount_column(ws, amount_col, header_row_idx + 2, ws.max_row)
    _auto_fit_columns(ws)

    _add_summary_sheet(wb, summary)
    _build_madaxil_maxaric_sheet(
        wb, categorized_rows, first_balance, last_balance, total_madaxil, total_mexaric
    )
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), summary, unclassified


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
            "Başlıq sətri tapılmadı. 'Əməliyyat Detalları'+'Debit' və ya "
            "'Tarix'+'Məbləğ'+'Təyinat' sütunları olan sətri axtarın."
        )
    header_row_idx, col_map, profile = found

    if profile == "kapital":
        teyinat_col = col_map["teyinat"]
        mebleg_col = col_map["mebleg"]
        eks_hesab_col = col_map.get("eks_hesab")
    elif profile == "turan":
        teyinat_col = col_map["teyinat"]
        mebleg_col = col_map["mebleg"]
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "ziraat":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "atb":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "unibank":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "rabite":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "bankofbaku":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "pasha":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "bankrespublika":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "abb":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
        tarix_col = col_map.get("emeliyyat_tarixi")
    elif profile == "yelo":
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        odeyen_col = col_map.get("odeyen")
    elif profile in ("access", "vtb"):
        teyinat_col = col_map["teyinat"]
        medaxil_col = col_map["medaxil"]
        mexaric_col = col_map["mexaric"]
        sender_col = col_map.get("gonderen_alici")
    else:
        details_col = col_map["emeliyyat_detallari"]
        debit_col = col_map["debit"]
        sender_col = col_map.get("gonderen_alici")
        kredit_col = col_map.get("kredit")
    cat_col = col_map.get("kateqoriya")
    max_cols = max((len(r) for r in rows), default=0)
    if cat_col is None:
        cat_col = max_cols  # 0-əsaslı yeni sütun
    cat_col_1 = cat_col + 1  # 1-əsaslı
    amount_col_1 = cat_col_1 + 1  # 1-əsaslı; Kateqoriya-nın dərhal sağı
    target_cols = max(max_cols, amount_col_1)

    summary = _empty_summary()
    unclassified: list[dict] = []
    categorized_rows: list[tuple[str, str, float]] = []
    balans_col = col_map.get("balans")
    first_balance: float | None = None
    last_balance: float | None = None
    total_madaxil = 0.0
    total_mexaric = 0.0
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
            if balans_col is not None and balans_col < len(values):
                bal = _parse_amount(values[balans_col])
                if first_balance is None and bal is not None:
                    first_balance = bal
                if bal is not None:
                    last_balance = bal
            if profile == "kapital":
                sender = (
                    _extract_sender_name(values[eks_hesab_col])
                    if eks_hesab_col is not None and eks_hesab_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                debit_value, kredit = _split_signed_amount(
                    values[mebleg_col] if mebleg_col < len(values) else None
                )
            elif profile == "turan":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail):
                    continue  # yekun/summary sətri — atla
                sender = ""  # sender sütunu yoxdur
                debit_value, kredit = _split_signed_amount(
                    values[mebleg_col] if mebleg_col < len(values) else None
                )
            elif profile == "ziraat":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or (
                    not _cell_text(detail) and not _sender_nonempty(sender)
                ):
                    continue  # yekun/summary sətri (CƏMİ / Çıxış qalıq) — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "atb":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail) or not _sender_nonempty(sender):
                    continue  # yekun/summary sətri — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "unibank":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail):
                    continue  # yekun/summary sətri (Yekun (Summary)) — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "rabite":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail):
                    continue  # yekun/summary sətri — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "bankofbaku":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail):
                    continue  # boş whitespace təyinat / son meta sətirləri — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "pasha":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    _extract_sender_name_pasha(values[sender_col])
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail):
                    continue  # yekun/balans/POS yığım sətirləri — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "bankrespublika":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail):
                    continue  # yekun/summary sətri — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "abb":
                tarix = (
                    values[tarix_col]
                    if tarix_col is not None and tarix_col < len(values)
                    else ""
                )
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                if not _cell_text(tarix) or not _cell_text(detail):
                    continue  # yekun/summary sətri — atla
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile == "yelo":
                sender = (
                    values[odeyen_col]
                    if odeyen_col is not None and odeyen_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            elif profile in ("access", "vtb"):
                sender = (
                    values[sender_col]
                    if sender_col is not None and sender_col < len(values)
                    else ""
                )
                detail = values[teyinat_col] if teyinat_col < len(values) else ""
                kredit = values[medaxil_col] if medaxil_col < len(values) else None
                debit_value = values[mexaric_col] if mexaric_col < len(values) else None
            else:
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
            if _has_amount(kredit):
                total_madaxil += abs(_parse_amount(kredit))
            if _has_amount(debit_value):
                total_mexaric += abs(_parse_amount(debit_value))
            category = _match_category(
                sender, detail, kredit, debit_value, profile, amount=_row_amount(debit_value, kredit)
            )
            if category is None:
                unclassified.append(
                    {
                        "row_index": r_idx + 1,
                        "sender": _cell_display(sender) or "",
                        "detail": _cell_display(detail) or "",
                        "kredit": _cell_display(kredit),
                        "debit": _cell_display(debit_value),
                    }
                )
                continue
            ws.cell(row=r_idx + 1, column=cat_col_1, value=category)
            entry = summary[category]
            entry["say"] = int(entry["say"]) + 1
            amount = _row_amount(debit_value, kredit)
            if category == "Terminal-Komissiya":
                fee = _extract_terminal_fee(detail)
                if fee is not None:
                    amount = fee
                else:
                    logger.warning(
                        "Terminal-Komissiya sətirində 'Xidmət haqqı' tapılmadı; "
                        "abs(Məbləğ) istifadə olunur",
                        extra={"row": r_idx + 1},
                    )
            if amount is not None:
                entry["cem"] = float(entry["cem"]) + amount
                ws.cell(row=r_idx + 1, column=amount_col_1, value=amount)

            # Mədaxil/Məxaric sheet-i üçün kateqoriya + istiqamət toplanması.
            if _has_amount(kredit):
                categorized_rows.append((category, "kredit", abs(_parse_amount(kredit))))
            if _has_amount(debit_value):
                categorized_rows.append(
                    (category, "debit", abs(_parse_amount(debit_value)))
                )

    _style_header_row(ws, header_row_idx + 1, amount_col_1)
    _style_amount_column(ws, amount_col_1, header_row_idx + 2, ws.max_row)
    _auto_fit_columns(ws)

    _add_summary_sheet(wb, summary)
    _build_madaxil_maxaric_sheet(
        wb, categorized_rows, first_balance, last_balance, total_madaxil, total_mexaric
    )
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), summary, unclassified


def categorize_statement_verbose(raw_bytes: bytes) -> tuple[bytes, dict, list[dict]]:
    """Bank çıxarışını kateqoriyalaşdırır.

    Qaytarır: (kateqoriyalaşdırılmış .xlsx bytes, summary dict, unclassified)
    Summary: 10 kateqoriya üzrə {"cem": float, "say": int}.
    Unclassified: kateqoriyasız sətirlərin siyahısı
        [{row_index, sender, detail, kredit, debit}].
    Başlıq sətri tapılmazsa ValueError (→ endpoint 422).
    """
    fmt = detect_format(raw_bytes)
    if fmt == "xlsx":
        return _categorize_xlsx(raw_bytes)
    if fmt == "csv":
        return _categorize_csv(raw_bytes)
    raise ValueError(".xls faylları dəstəklənmir; yalnız .xlsx və .csv yükləyin.")


def categorize_statement(raw_bytes: bytes) -> tuple[bytes, dict]:
    """categorize_statement_verbose-in 2-lik (bytes, summary) wrapper-i."""
    out, summary, _ = categorize_statement_verbose(raw_bytes)
    return out, summary


def _locate_columns(wb):
    """Kateqoriyalaşdırılmış çıxış faylında lazımi sütunları tapır.

    Qaytarır: (ws, header_row(1-əsaslı), cat_col, amount_col,
    debit_col(1-əsaslı|None), kredit_col(1-əsaslı|None)).
    """
    for sheet in wb.worksheets:
        rows = [[cell.value for cell in row] for row in sheet.iter_rows()]
        found = _find_header_row(rows)
        if found is not None:
            header_row_idx, col_map, profile = found
            header_row = header_row_idx + 1
            cat_col = col_map["kateqoriya"] + 1
            amount_col = cat_col + 1
            balans_col = col_map.get("balans")
            balans_col = balans_col + 1 if balans_col is not None else None
            debit_col = kredit_col = mebleg_col = None
            if profile in ("kapital", "turan"):
                mebleg_col = col_map["mebleg"] + 1
            elif profile in ("yelo", "access", "vtb", "ziraat", "atb", "unibank", "rabite", "bankofbaku", "pasha", "bankrespublika", "abb"):
                debit_col = col_map["mexaric"] + 1
                kredit_col = col_map["medaxil"] + 1
            else:
                debit_col = col_map["debit"] + 1
                kredit = col_map.get("kredit")
                kredit_col = kredit + 1 if kredit is not None else None
            return sheet, header_row, cat_col, amount_col, profile, debit_col, kredit_col, mebleg_col, balans_col
    raise ValueError("Kateqoriyalaşdırılmış faylda başlıq sətri tapılmadı.")


def _summarize_sheet(ws, header_row: int, cat_col: int, amount_col: int) -> dict:
    """Sheet-dəki Kateqoriya + İcmal sütunlarından summary-nı yenidən qurur."""
    summary = _empty_summary()
    for row in ws.iter_rows(min_row=header_row + 1):
        category = row[cat_col - 1].value
        if not category or category not in summary:
            continue
        entry = summary[category]
        entry["say"] = int(entry["say"]) + 1
        amount = row[amount_col - 1].value
        if isinstance(amount, (int, float)):
            entry["cem"] = float(entry["cem"]) + abs(amount)
    return summary


def finalize_statement(raw_bytes: bytes, selections: dict[int, str]) -> tuple[bytes, dict]:
    """Kateqoriyalaşdırılmış fayla istifadəçi override-larını tətbiq edir.

    selections: {row_index: kateqoriya} — boş string "kateqoriyasız qalsın"
    deməkdir. Hər row_index tanınmamış sətir olmalıdır; qeyri-boş kateqoriya
    _CATEGORY_ORDER-dan olmalıdır. Yekun fayl (Kateqoriya+İcmal dolmuş,
    Xülasə yenidən hesablanmış) və yeni summary qaytarılır.
    """
    out_bytes, _, unclassified = categorize_statement_verbose(raw_bytes)
    unclassified_indexes = {r["row_index"] for r in unclassified}

    parsed: dict[int, str] = {}
    for key, value in selections.items():
        row_index = int(key)
        category = (value or "").strip()
        if row_index not in unclassified_indexes:
            raise ValueError(f"Sətir {row_index} tanınmamış sətir deyil.")
        if category and category not in _CATEGORY_ORDER:
            raise ValueError(f"Keçərsiz kateqoriya: {category}")
        parsed[row_index] = category

    wb = load_workbook(io.BytesIO(out_bytes))
    (
        ws,
        header_row,
        cat_col,
        amount_col,
        profile,
        debit_col,
        kredit_col,
        mebleg_col,
        balans_col,
    ) = _locate_columns(wb)

    for row_index, category in parsed.items():
        if not category:
            continue  # "kateqoriyasız qalsın"
        ws.cell(row=row_index, column=cat_col).value = category
        if profile in ("kapital", "turan"):
            mebleg = ws.cell(row=row_index, column=mebleg_col).value
            debit, kredit = _split_signed_amount(mebleg)
        else:
            debit = ws.cell(row=row_index, column=debit_col).value if debit_col else None
            kredit = ws.cell(row=row_index, column=kredit_col).value if kredit_col else None
        amount = _row_amount(debit, kredit)
        if amount is not None:
            ws.cell(row=row_index, column=amount_col).value = amount

    new_summary = _summarize_sheet(ws, header_row, cat_col, amount_col)
    del wb[_SUMMARY_SHEET_NAME]
    _add_summary_sheet(wb, new_summary)

    # 3-cü sheet-i override-lardan sonra yenidən qur (override kateqoriyaları əks olunsun).
    if _MADAXIL_MAXARIC_SHEET_NAME in wb.sheetnames:
        del wb[_MADAXIL_MAXARIC_SHEET_NAME]
    categorized_rows, total_madaxil, total_mexaric, first_balance, last_balance = (
        _collect_madaxil_maxaric_from_sheet(
            ws, header_row, cat_col, amount_col, profile, debit_col, kredit_col, mebleg_col, balans_col
        )
    )
    _build_madaxil_maxaric_sheet(
        wb, categorized_rows, first_balance, last_balance, total_madaxil, total_mexaric
    )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), new_summary