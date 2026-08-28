from pydantic import BaseModel, ConfigDict, Field


class ImportExcelResponse(BaseModel):
    status: str = "done"
    message: str
    inserted: int = 0
    sources: list[str] = []


class BankStatementCategorySummary(BaseModel):
    cem: float = 0.0
    say: int = Field(default=0, ge=0)


class BankStatementSummary(BaseModel):
    """11 kateqoriya üzrə xülasə (bank çıxarışı kateqoriyalaşdırması).

    Alias-lar kateqoriya adlarının orijinal (Azərbaycan) yazılışıdır.
    """

    model_config = ConfigDict(populate_by_name=True)

    komisiya: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="KOMİSİYA"
    )
    nagd_cekim_komissiyasi: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="Nağd çəkim komissiyası"
    )
    bank_faiz_xerci: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="BANK FAİZ XƏRCİ"
    )
    edv_odenisi: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="ƏDV ÖDƏNİŞİ"
    )
    rusum: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="RÜSUM"
    )
    vergi_mexaric: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="VERGİ MƏXARİC"
    )
    maas_mexaric: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="MAAŞ MƏXARİC"
    )
    tesisciden_borc: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="TƏSİSÇİDƏN BORC"
    )
    diger_sirket_daxil_olmalar: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="DİGƏR ŞİRKƏT DAXİL OLMALAR"
    )
    diger_sirket_odenisleri: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ"
    )
    terminal_komissiya: BankStatementCategorySummary = Field(
        default_factory=BankStatementCategorySummary, alias="Terminal-Komissiya"
    )


class BankStatementUnclassifiedRow(BaseModel):
    """Kateqoriyasız qalmış sətir (manual təsnifat üçün)."""

    row_index: int
    sender: str = ""
    detail: str = ""
    kredit: str | None = None
    debit: str | None = None


class BankStatementCategorizeResult(BaseModel):
    """Mərhələ 1 cavabı: summary + tanınmamış sətirlər + fayl (base64)."""

    summary: BankStatementSummary
    unclassified: list[BankStatementUnclassifiedRow] = []
    file_base64: str