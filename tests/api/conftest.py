"""API 測試的共用 fixture。

預設用記憶體 SQLite，跑一輪不到一秒、不需要任何外部服務。

但整份測試也可以指向真的 PostgreSQL::

    RECONCILIATION_TEST_DATABASE_URL="postgresql+psycopg://user@localhost/recon_test" \\
        python -m pytest tests/api

這件事的意義不只是「多一種測試環境」。README 說「SQLite 換 PostgreSQL
只要改連線字串」——這個 fixture 讓那句話變成**可執行的驗證**，
而不是一句沒人查證過的宣稱。W6 的 CI 會兩種都跑一遍。
"""

from __future__ import annotations

import os
from collections.abc import Iterator

import pytest
from fastapi.testclient import TestClient

from reconciliation.api import create_app
from reconciliation.repositories.database import create_db_engine
from reconciliation.repositories.sqlalchemy_repo import SqlAlchemyOrderRepository
from reconciliation.repositories.tables import Base

DEFAULT_TEST_URL = "sqlite:///:memory:"


def test_database_url() -> str:
    return os.environ.get("RECONCILIATION_TEST_DATABASE_URL", DEFAULT_TEST_URL)


@pytest.fixture
def client() -> Iterator[TestClient]:
    """每個測試一個乾淨的資料庫。

    記憶體 SQLite 天生就是乾淨的（每個 engine 一個新資料庫）；
    指向 PostgreSQL 時則在每個測試前後把資料表整組重建，
    讓兩種後端的測試語意完全一致。
    """
    url = test_database_url()
    is_sqlite_memory = url.startswith("sqlite") and ":memory:" in url

    if not is_sqlite_memory:
        engine = create_db_engine(url)
        Base.metadata.drop_all(engine)
        engine.dispose()

    app = create_app(url)
    with TestClient(app) as test_client:
        yield test_client

    # 明確關掉連線池。不關的話 psycopg 會在物件被回收時才關，
    # 而那個時間點在 pytest 的掌控之外，會冒出 ResourceWarning。
    app.state.engine.dispose()

    if not is_sqlite_memory:
        engine = create_db_engine(url)
        Base.metadata.drop_all(engine)
        engine.dispose()


@pytest.fixture
def seeded_client(client: TestClient) -> TestClient:
    """先塞好訂單的 client。

    訂單目前還沒有匯入端點（我方訂單的來源在 W5 之後才會接上），
    所以直接用 repository 寫進去。這是測試的權宜做法，不是產品行為。
    """
    from tests.domain.test_matching import make_order

    factory = client.app.state.session_factory  # type: ignore[attr-defined]
    session = factory()
    SqlAlchemyOrderRepository(session).add_many(
        [
            make_order(1, external="AAAA0001", gross="1000", fee="55.30"),
            make_order(2, external="AAAA0002", gross="2500", fee="138.25"),
            make_order(3, external="AAAA0003", gross="680", fee="37.60"),
            make_order(4, external="AAAA0004", gross="3000", fee="165.90"),
        ]
    )
    session.commit()
    session.close()
    return client
