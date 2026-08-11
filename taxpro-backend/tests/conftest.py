import asyncio
import os
from pathlib import Path

os.environ.setdefault("DB_URL", "postgresql+asyncpg://taxpro:taxpro@localhost:5432/taxpro_test")
os.environ.setdefault("APP_ENV", "development")
os.environ.setdefault(
    "ENCRYPTION_KEY",
    "zQYo10eooLLAaz2FRLKmdVzoXydus7K_tL2ngIL1PbM=",
)

import pytest
import pytest_asyncio
from langgraph.checkpoint.memory import MemorySaver
from sqlalchemy import text

from app.services.ai.graph import build_graph
from tests.fakes import ScriptedModel

ADMIN_DB_URL = "postgresql+asyncpg://taxpro:taxpro@localhost:5432/taxpro"
TEST_DB_NAME = "taxpro_test"
BACKEND_DIR = Path(__file__).resolve().parent.parent


async def _recreate_test_db() -> None:
    from sqlalchemy.ext.asyncio import create_async_engine

    engine = create_async_engine(ADMIN_DB_URL, isolation_level="AUTOCOMMIT")
    async with engine.connect() as conn:
        await conn.execute(
            text(
                "SELECT pg_terminate_backend(pid) FROM pg_stat_activity "
                "WHERE datname = :name AND pid <> pg_backend_pid()"
            ),
            {"name": TEST_DB_NAME},
        )
        await conn.execute(text(f'DROP DATABASE IF EXISTS "{TEST_DB_NAME}"'))
        await conn.execute(text(f'CREATE DATABASE "{TEST_DB_NAME}"'))
    await engine.dispose()


@pytest.fixture(scope="session", autouse=True)
def _setup_database():
    asyncio.run(_recreate_test_db())

    from alembic import command
    from alembic.config import Config

    config = Config(str(BACKEND_DIR / "alembic.ini"))
    config.set_main_option("script_location", str(BACKEND_DIR / "alembic"))
    command.upgrade(config, "head")

    yield



@pytest_asyncio.fixture(autouse=True)
async def _clean_tables():
    from app.core.database import async_session_maker

    yield

    async with async_session_maker() as session:
        await session.execute(
            text("TRUNCATE vector_documents, profiles, portals CASCADE")
        )
        await session.commit()


@pytest.fixture
def fake_model():
    return ScriptedModel()


@pytest.fixture
def fake_graph(fake_model):
    return build_graph(checkpointer=MemorySaver(), model=fake_model)


@pytest_asyncio.fixture
async def client(fake_graph):
    import httpx

    import app.services.ai.graph as graph_mod
    from app.core.database import async_session_maker, get_db
    from app.main import app

    graph_mod._graph = fake_graph

    async def _override_db():
        async with async_session_maker() as session:
            yield session

    app.dependency_overrides[get_db] = _override_db

    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="http://test"
    ) as c:
        yield c

    app.dependency_overrides.clear()
    graph_mod._graph = None
