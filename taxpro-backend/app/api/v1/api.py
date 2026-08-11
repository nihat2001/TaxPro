from fastapi import APIRouter

from app.api.v1.endpoints import (
    bank_statements,
    chat,
    import_excel,
    portals,
    profiles,
    resume,
)

api_router = APIRouter(prefix="/api/v1")
api_router.include_router(profiles.router)
api_router.include_router(portals.router)
api_router.include_router(import_excel.router)
api_router.include_router(bank_statements.router)
api_router.include_router(chat.router)
api_router.include_router(resume.router)
