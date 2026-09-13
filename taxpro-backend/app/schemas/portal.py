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
    # Selector-u qeyd olunmuş sahələr: {selector: value}
    fields: dict[str, str]
    # Selector -> məntiqi sahə adı (fields-in tərsi)
    field_map: dict[str, str] = {}
    # Selector-u olmayan, lakin content-script heuristikası ilə tapılmalı sahələr:
    # {məntiqi_sahə: value}
    fallback_fields: dict[str, str] = {}
    # Selector-u olmayan məntiqi sahələrin siyahısı (UI-də xəbərdarlıq üçün)
    missing_selectors: list[str] = []
    phone_prefix_selector: str | None = None
