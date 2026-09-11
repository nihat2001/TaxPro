from datetime import datetime
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field


class ProfileBase(BaseModel):
    ad_soyad: str = Field(min_length=2, max_length=255)
    nomre: str = Field(min_length=7, max_length=50)
    asan_id: str = Field(min_length=1, max_length=50)


class ProfileCreate(ProfileBase):
    pass


class ProfileUpdate(BaseModel):
    ad_soyad: str | None = None
    nomre: str | None = None
    asan_id: str | None = None


class ProfileRead(ProfileBase):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    created_at: datetime
    updated_at: datetime


class ProfileAutoFill(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    ad_soyad: str
    nomre: str
    asan_id: str
