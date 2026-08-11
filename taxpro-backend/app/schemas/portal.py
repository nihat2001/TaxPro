from typing import Literal

from uuid import UUID

from pydantic import BaseModel, ConfigDict


class PortalRead(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    name: str
    url: str
    category: str
    logo_url: str | None
    selector_mapping: dict
    auth_type: str
    sort_order: int


class SelectorMappingUpdate(BaseModel):
    selector_mapping: dict
    auth_type: Literal["asan_id_form", "none", "unknown"]


class AutoFillPayload(BaseModel):
    profile_id: UUID
    portal_id: UUID
    portal_name: str
    fields: dict[str, str]
    field_map: dict[str, str] = {}
