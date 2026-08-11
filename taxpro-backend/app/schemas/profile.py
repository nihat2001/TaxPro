from datetime import date, datetime
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field, field_validator


class ProfileBase(BaseModel):
    ad_soyad: str = Field(min_length=2, max_length=255)
    voen: str = Field(min_length=7, max_length=50)
    nomre: str = Field(min_length=7, max_length=50)
    asan_id: str = Field(min_length=1, max_length=50)
    bitme_tarixi: date | None = None
    qeyd: str | None = None
    reng: str = Field(default="#3b82f6", max_length=20)
    ikon: str = Field(default="user", max_length=50)


class ProfileCreate(ProfileBase):
    pin1: str = Field(min_length=4, max_length=100)
    pin2: str = Field(min_length=4, max_length=100)
    puk: str = Field(min_length=4, max_length=100)


class ProfileUpdate(BaseModel):
    ad_soyad: str | None = None
    voen: str | None = None
    nomre: str | None = None
    asan_id: str | None = None
    bitme_tarixi: date | None = None
    qeyd: str | None = None
    pin1: str | None = None
    pin2: str | None = None
    puk: str | None = None
    reng: str | None = None
    ikon: str | None = None


class ProfileRead(ProfileBase):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    pin1: str = "****"
    pin2: str = "****"
    puk: str = "****"
    created_at: datetime
    updated_at: datetime

    @field_validator("pin1", "pin2", "puk")
    @classmethod
    def mask_secret(cls, value: str) -> str:
        return "****" if value else ""


class ProfileAutoFill(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    ad_soyad: str
    voen: str
    nomre: str
    asan_id: str
    bitme_tarixi: date | None
    qeyd: str | None
    pin1: str
    pin2: str
    puk: str
