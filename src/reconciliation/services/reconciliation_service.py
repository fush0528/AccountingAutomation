"""對帳服務：把資料庫裡的訂單與結算記錄跑一次對帳並保存結果。

這一層很薄，而且應該很薄。所有的演算法都在 ``domain/matching.py``，
所有的存取都在 ``repositories/``。服務層只做三件事：
取資料、呼叫領域邏輯、保存結果。

這個「薄」是可驗證的：整個檔案不到一百行，而且沒有任何一個 if 是在
處理業務規則——業務規則全部在領域層。
"""

from __future__ import annotations

import logging
import uuid
from dataclasses import dataclass

from sqlalchemy.orm import Session

from ..domain.matching import MatchingConfig, ReconciliationReport, reconcile
from ..repositories.sqlalchemy_repo import (
    SqlAlchemyOrderRepository,
    SqlAlchemyReconciliationRepository,
    SqlAlchemySettlementRepository,
)

__all__ = ["NothingToReconcile", "ReconciliationOutcome", "ReconciliationService"]

logger = logging.getLogger(__name__)


class NothingToReconcile(Exception):
    """資料庫裡沒有訂單或沒有結算記錄。"""


@dataclass(frozen=True, slots=True)
class ReconciliationOutcome:
    run_id: str
    report: ReconciliationReport


class ReconciliationService:
    def __init__(self, session: Session) -> None:
        self._session = session
        self._orders = SqlAlchemyOrderRepository(session)
        self._settlements = SqlAlchemySettlementRepository(session)
        self._runs = SqlAlchemyReconciliationRepository(session)

    def run(
        self, *, platform: str | None = None, config: MatchingConfig | None = None
    ) -> ReconciliationOutcome:
        orders = self._orders.list_all(platform)
        records = self._settlements.list_all(platform)

        if not orders and not records:
            raise NothingToReconcile("資料庫裡既沒有訂單也沒有結算記錄。請先上傳結算單並匯入訂單。")

        report = reconcile(orders, records, config)
        run_id = f"RUN-{uuid.uuid4().hex[:12].upper()}"
        self._runs.save(run_id, report)

        logger.info(
            "對帳完成",
            extra={
                "run_id": run_id,
                "orders": len(orders),
                "records": len(records),
                "automation_rate": report.metrics.automation_rate,
            },
        )
        return ReconciliationOutcome(run_id=run_id, report=report)

    def get(self, run_id: str) -> ReconciliationReport | None:
        return self._runs.get(run_id)

    def list_run_ids(self) -> list[str]:
        return self._runs.list_run_ids()
