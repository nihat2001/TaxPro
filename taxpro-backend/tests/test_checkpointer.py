import pytest

from app.core.config import get_settings


@pytest.fixture(autouse=True)
def _reset_settings():
    get_settings.cache_clear()
    yield
    get_settings.cache_clear()


def _prod_settings():
    settings = get_settings()
    settings.app_env = "production"
    return settings


def _dev_settings():
    settings = get_settings()
    settings.app_env = "development"
    return settings


def test_build_graph_production_without_checkpointer_raises(monkeypatch):
    monkeypatch.setattr("app.services.ai.graph.get_settings", _prod_settings)
    monkeypatch.setattr("app.services.ai.checkpointer._checkpointer", None)
    from app.services.ai.graph import build_graph

    from tests.fakes import ScriptedModel

    with pytest.raises(RuntimeError, match="MemorySaver fallback-i qadağandır"):
        build_graph(model=ScriptedModel())


def test_build_graph_development_without_checkpointer_falls_back(monkeypatch):
    monkeypatch.setattr("app.services.ai.graph.get_settings", _dev_settings)
    monkeypatch.setattr("app.services.ai.checkpointer._checkpointer", None)
    from app.services.ai.graph import build_graph

    from tests.fakes import ScriptedModel

    graph = build_graph(model=ScriptedModel())
    assert graph is not None


def test_lifespan_production_fails_fast(monkeypatch):
    monkeypatch.setattr("app.main.get_settings", _prod_settings)
    monkeypatch.setattr(
        "app.main.init_checkpointer",
        _fail_init,
    )
    from fastapi.testclient import TestClient

    from app.main import app

    with pytest.raises(RuntimeError):
        with TestClient(app):
            pass


def test_lifespan_development_starts_with_fallback(monkeypatch):
    monkeypatch.setattr("app.main.get_settings", _dev_settings)
    monkeypatch.setattr(
        "app.main.init_checkpointer",
        _fail_init,
    )
    from fastapi.testclient import TestClient

    from app.main import app

    with TestClient(app) as c:
        resp = c.get("/health")
        assert resp.status_code == 200


async def _fail_init():
    raise RuntimeError("db down")
