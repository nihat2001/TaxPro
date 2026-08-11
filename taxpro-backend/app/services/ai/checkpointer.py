from langgraph.checkpoint.postgres.aio import AsyncPostgresSaver
from psycopg_pool import AsyncConnectionPool

from app.core.config import get_settings

_pool: AsyncConnectionPool | None = None
_checkpointer: AsyncPostgresSaver | None = None


def _saver_dsn() -> str:
    return get_settings().db_url.replace("postgresql+asyncpg://", "postgresql://", 1)


async def init_checkpointer() -> AsyncPostgresSaver:
    global _checkpointer, _pool
    if _checkpointer is not None:
        return _checkpointer
    _pool = AsyncConnectionPool(conninfo=_saver_dsn(), open=False, max_size=10)
    await _pool.open()
    _checkpointer = AsyncPostgresSaver(_pool)
    await _checkpointer.setup()
    return _checkpointer


def get_checkpointer(testing: bool = False) -> AsyncPostgresSaver | None:
    if testing:
        return None
    return _checkpointer


async def close_checkpointer() -> None:
    global _checkpointer, _pool
    if _pool is not None:
        await _pool.close()
    _checkpointer = None
    _pool = None
