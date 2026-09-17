"""資料庫連線與 session 管理。

整份程式碼裡唯一知道連線字串的地方。要從 SQLite 換成 PostgreSQL，
改的是這裡讀取的那個環境變數，其他一行都不用動——
因為所有存取都經過 repository 介面，而且沒有任何原生 SQL。

    SQLite      sqlite:///./reconciliation.db
    PostgreSQL  postgresql+psycopg://user:pass@localhost/reconciliation
"""

from __future__ import annotations

import os
from collections.abc import Iterator
from contextlib import contextmanager

from sqlalchemy import Engine, create_engine, event
from sqlalchemy.orm import Session, sessionmaker
from sqlalchemy.pool import StaticPool

from .tables import Base

__all__ = ["DEFAULT_DATABASE_URL", "create_db_engine", "init_schema", "session_scope"]

DEFAULT_DATABASE_URL = "sqlite:///./reconciliation.db"


def create_db_engine(url: str | None = None, *, echo: bool = False) -> Engine:
    url = url or os.environ.get("RECONCILIATION_DATABASE_URL", DEFAULT_DATABASE_URL)
    connect_args: dict[str, object] = {}
    kwargs: dict[str, object] = {}

    if url.startswith("sqlite"):
        # FastAPI 的同一個請求可能跨執行緒，SQLite 預設會擋
        connect_args["check_same_thread"] = False

        if ":memory:" in url or "mode=memory" in url:
            # 記憶體資料庫是「每條連線一個」的：預設的連線池每次都開新連線，
            # 於是建好資料表的那個連線關掉之後，下一個 session 看到的是
            # 一個全新的空資料庫。StaticPool 讓整個 engine 共用同一條連線，
            # 測試才能在記憶體裡跑完整的請求流程。
            kwargs["poolclass"] = StaticPool

    engine = create_engine(url, echo=echo, future=True, connect_args=connect_args, **kwargs)

    if url.startswith("sqlite"):
        # SQLite 預設不強制外鍵，必須每條連線各自開啟。
        # 不開的話 ForeignKey 只是註解，不是約束——那會讓
        # 「資料庫層面保證完整性」這句話變成空話。
        @event.listens_for(engine, "connect")
        def _enable_foreign_keys(dbapi_connection: object, _record: object) -> None:
            cursor = dbapi_connection.cursor()  # type: ignore[attr-defined]
            cursor.execute("PRAGMA foreign_keys=ON")
            cursor.close()

    return engine


def init_schema(engine: Engine) -> None:
    """建立資料表。

    開發與測試用。正式環境應該跑 Alembic migration——
    ``create_all`` 不會處理既有資料庫的結構變更。
    """
    Base.metadata.create_all(engine)


@contextmanager
def session_scope(engine: Engine) -> Iterator[Session]:
    """一個交易範圍。

    離開時沒有例外就 commit，有例外就 rollback。這是冪等匯入的第三道
    防線：整批寫入要嘛全部成功，要嘛完全沒發生，不會留下半批資料。
    """
    factory = sessionmaker(bind=engine, expire_on_commit=False)
    session = factory()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
