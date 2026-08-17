from fastapi import APIRouter

from app.api.v1.endpoints import import_excel, portals, profiles

api_router = APIRouter(prefix="/api/v1")
api_router.include_router(profiles.router)
api_router.include_router(portals.router)
api_router.include_router(import_excel.router)
