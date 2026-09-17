"""Repository 介面。

用 :class:`typing.Protocol` 而不是抽象基底類別，因為實作不需要繼承——
只要方法簽章對得上就算符合。這讓測試可以用一個五十行的記憶體版本
取代 SQLAlchemy 實作，而不必為了測試去繼承一個真實的基底類別。

這一層存在的唯一理由，是讓 ``services/`` 依賴**介面**而不是 SQLAlchemy。
因此「SQLite 換成 PostgreSQL」對服務層來說完全不可見，
而服務層的測試也不需要資料庫。

所有方法都收發領域物件（``Order``、``SettlementRecord``），不收發
SQLAlchemy 的 row 物件。ORM 的型別不該洩漏到這個介面之外。
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol

from ..domain.matching import ReconciliationReport
from ..domain.models import ImportBatch, Order, SettlementRecord

__all__ = [
    "ImportBatchRepository",
    "OrderRepository",
    "ReconciliationRepository",
    "SettlementRepository",
    "StoredRun",
]


class OrderRepository(Protocol):
    def add_many(self, orders: Sequence[Order]) -> int:
        """寫入訂單，回傳實際新增的筆數（已存在的會被跳過）。"""
        ...

    def list_all(self, platform: str | None = None) -> list[Order]: ...

    def count(self) -> int: ...


class ImportBatchRepository(Protocol):
    def find_by_content_hash(self, content_hash: str) -> ImportBatch | None:
        """冪等匯入的第一道防線：同樣的檔案內容是否已經匯入過。"""
        ...

    def get(self, batch_id: str) -> ImportBatch | None: ...

    def list_all(self) -> list[ImportBatch]: ...


class SettlementRepository(Protocol):
    def list_all(self, platform: str | None = None) -> list[SettlementRecord]: ...

    def list_by_batch(self, batch_id: str) -> list[SettlementRecord]: ...

    def count(self) -> int: ...


class StoredRun(Protocol):
    """一次已保存的對帳執行。"""

    run_id: str
    report: ReconciliationReport


class ReconciliationRepository(Protocol):
    def save(self, run_id: str, report: ReconciliationReport) -> None: ...

    def get(self, run_id: str) -> ReconciliationReport | None: ...

    def list_run_ids(self) -> list[str]: ...
