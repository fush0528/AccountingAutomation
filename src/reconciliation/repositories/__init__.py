"""持久層：資料表、Repository 介面與 SQLAlchemy 實作。

服務層只該 import 這裡的介面與工廠函式，不該 import ``tables`` 或
``sqlalchemy_repo`` 裡的具體類別。
"""

from .database import (
    DEFAULT_DATABASE_URL,
    create_db_engine,
    init_schema,
    session_scope,
)
from .interfaces import (
    ImportBatchRepository,
    OrderRepository,
    ReconciliationRepository,
    SettlementRepository,
)
from .mappers import row_fingerprint
from .sqlalchemy_repo import (
    SqlAlchemyImportBatchRepository,
    SqlAlchemyOrderRepository,
    SqlAlchemyReconciliationRepository,
    SqlAlchemySettlementRepository,
)
from .tables import Base

__all__ = [
    "DEFAULT_DATABASE_URL",
    "Base",
    "ImportBatchRepository",
    "OrderRepository",
    "ReconciliationRepository",
    "SettlementRepository",
    "SqlAlchemyImportBatchRepository",
    "SqlAlchemyOrderRepository",
    "SqlAlchemyReconciliationRepository",
    "SqlAlchemySettlementRepository",
    "create_db_engine",
    "init_schema",
    "row_fingerprint",
    "session_scope",
]
