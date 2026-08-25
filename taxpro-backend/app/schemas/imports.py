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
    """3 kateqoriya üzrə xülasə (bank çıxarışı kateqoriyalaşdırması).

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
