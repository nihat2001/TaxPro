from pydantic import BaseModel


class ImportExcelResponse(BaseModel):
    status: str = "done"
    message: str
    inserted: int = 0
    sources: list[str] = []
