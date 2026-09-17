"""Repository 的 SQLAlchemy 實作。

這些類別是整個專案裡唯一寫 SQLAlchemy 查詢的地方。它們符合
``interfaces.py`` 的 Protocol，所以服務層拿到的是介面，不是這些類別。

沒有任何一行原生 SQL——這是「SQLite 換 PostgreSQL 只改連線字串」
那句話的依據。
"""

from __future__ import annotations

import json
from collections.abc import Sequence
from decimal import Decimal

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from ..domain.matching import (
    MatchingConfig,
    ReconciliationMetrics,
    ReconciliationReport,
)
from ..domain.models import (
    Candidate,
    ImportBatch,
    MatchOutcome,
    MatchResult,
    MatchStage,
    Order,
    SettlementRecord,
    VarianceReason,
)
from ..domain.money import Money
from .mappers import (
    order_to_row,
    row_to_batch,
    row_to_order,
    row_to_record,
)
from .tables import (
    ImportBatchRow,
    MatchResultRow,
    OrderRow,
    ReconciliationRunRow,
    SettlementRecordRow,
)

__all__ = [
    "SqlAlchemyImportBatchRepository",
    "SqlAlchemyOrderRepository",
    "SqlAlchemyReconciliationRepository",
    "SqlAlchemySettlementRepository",
]


class SqlAlchemyOrderRepository:
    def __init__(self, session: Session) -> None:
        self._session = session

    def add_many(self, orders: Sequence[Order]) -> int:
        """寫入訂單，已存在的（同 order_id）跳過。

        訂單的來源是我方系統，重複匯入同一份訂單檔是常見操作，
        所以這裡採「跳過既有」而不是報錯。
        """
        if not orders:
            return 0
        existing = set(
            self._session.scalars(
                select(OrderRow.order_id).where(OrderRow.order_id.in_([o.order_id for o in orders]))
            ).all()
        )
        added = 0
        for order in orders:
            if order.order_id in existing:
                continue
            self._session.add(order_to_row(order))
            existing.add(order.order_id)
            added += 1
        self._session.flush()
        return added

    def list_all(self, platform: str | None = None) -> list[Order]:
        stmt = select(OrderRow).order_by(OrderRow.order_id)
        if platform:
            stmt = stmt.where(OrderRow.platform == platform)
        return [row_to_order(row) for row in self._session.scalars(stmt)]

    def count(self) -> int:
        return self._session.scalar(select(func.count()).select_from(OrderRow)) or 0


class SqlAlchemyImportBatchRepository:
    def __init__(self, session: Session) -> None:
        self._session = session

    def find_by_content_hash(self, content_hash: str) -> ImportBatch | None:
        row = self._session.scalar(
            select(ImportBatchRow).where(ImportBatchRow.content_hash == content_hash)
        )
        return row_to_batch(row) if row else None

    def get(self, batch_id: str) -> ImportBatch | None:
        row = self._session.scalar(
            select(ImportBatchRow).where(ImportBatchRow.batch_id == batch_id)
        )
        return row_to_batch(row) if row else None

    def list_all(self) -> list[ImportBatch]:
        stmt = select(ImportBatchRow).order_by(ImportBatchRow.imported_at.desc())
        return [row_to_batch(row) for row in self._session.scalars(stmt)]


class SqlAlchemySettlementRepository:
    def __init__(self, session: Session) -> None:
        self._session = session

    def list_all(self, platform: str | None = None) -> list[SettlementRecord]:
        stmt = select(SettlementRecordRow).order_by(SettlementRecordRow.id)
        if platform:
            stmt = stmt.where(SettlementRecordRow.platform == platform)
        return [row_to_record(row) for row in self._session.scalars(stmt)]

    def list_by_batch(self, batch_id: str) -> list[SettlementRecord]:
        stmt = (
            select(SettlementRecordRow)
            .join(ImportBatchRow)
            .where(ImportBatchRow.batch_id == batch_id)
            .order_by(SettlementRecordRow.source_row)
        )
        return [row_to_record(row) for row in self._session.scalars(stmt)]

    def count(self) -> int:
        return self._session.scalar(select(func.count()).select_from(SettlementRecordRow)) or 0


class SqlAlchemyReconciliationRepository:
    """對帳結果的保存與讀回。

    結算列不重複存——``MatchResultRow.record_ids_json`` 存的是
    ``settlement_records`` 的主鍵。這樣同一列被不同次對帳引用時，
    資料庫裡仍然只有一份，而且之後修正某一列的資料，
    歷史報告讀回來會跟著更新（這是刻意的：報告是視圖，不是快照）。
    """

    def __init__(self, session: Session) -> None:
        self._session = session

    def save(self, run_id: str, report: ReconciliationReport) -> None:
        run = ReconciliationRunRow(
            run_id=run_id,
            created_at=__import__("datetime").datetime.now(),
            config_json=json.dumps(_config_to_dict(report.config), ensure_ascii=False),
            metrics_json=json.dumps(_metrics_to_dict(report.metrics), ensure_ascii=False),
        )
        self._session.add(run)
        self._session.flush()

        fingerprint_to_pk = self._fingerprint_index()
        for result in report.results:
            record_pks = [
                pk
                for rec in result.records
                if (pk := fingerprint_to_pk.get(_fingerprint(rec))) is not None
            ]
            self._session.add(
                MatchResultRow(
                    run_pk=run.id,
                    outcome=result.outcome.value,
                    stage=result.stage.value if result.stage else None,
                    order_id=result.order.order_id if result.order else None,
                    record_ids_json=json.dumps(record_pks),
                    variance_minor=(result.variance.minor_units if result.variance else None),
                    variance_reason=(
                        result.variance_reason.value if result.variance_reason else None
                    ),
                    candidates_json=json.dumps(
                        [
                            {
                                "order_id": c.order.order_id,
                                "score": str(c.score),
                                "reasons": list(c.reasons),
                            }
                            for c in result.candidates
                        ],
                        ensure_ascii=False,
                    ),
                )
            )
        self._session.flush()

    def get(self, run_id: str) -> ReconciliationReport | None:
        run = self._session.scalar(
            select(ReconciliationRunRow).where(ReconciliationRunRow.run_id == run_id)
        )
        if run is None:
            return None

        orders = {o.order_id: o for o in SqlAlchemyOrderRepository(self._session).list_all()}
        records_by_pk = {
            row.id: row_to_record(row) for row in self._session.scalars(select(SettlementRecordRow))
        }

        results: list[MatchResult] = []
        for row in run.results:
            record_pks = json.loads(row.record_ids_json)
            candidates = tuple(
                Candidate(
                    order=orders[c["order_id"]],
                    score=Decimal(c["score"]),
                    reasons=tuple(c["reasons"]),
                )
                for c in json.loads(row.candidates_json)
                if c["order_id"] in orders
            )
            currency = "TWD"
            results.append(
                MatchResult(
                    outcome=MatchOutcome(row.outcome),
                    stage=MatchStage(row.stage) if row.stage else None,
                    order=orders.get(row.order_id) if row.order_id else None,
                    records=tuple(records_by_pk[pk] for pk in record_pks if pk in records_by_pk),
                    variance=(
                        Money(row.variance_minor, currency)
                        if row.variance_minor is not None
                        else None
                    ),
                    variance_reason=(
                        VarianceReason(row.variance_reason) if row.variance_reason else None
                    ),
                    candidates=candidates,
                )
            )

        return ReconciliationReport(
            results=tuple(results),
            metrics=_metrics_from_dict(json.loads(run.metrics_json)),
            config=_config_from_dict(json.loads(run.config_json)),
        )

    def list_run_ids(self) -> list[str]:
        stmt = select(ReconciliationRunRow.run_id).order_by(ReconciliationRunRow.created_at.desc())
        return list(self._session.scalars(stmt))

    # ------------------------------------------------------------------
    def _fingerprint_index(self) -> dict[str, int]:
        return {
            row.row_fingerprint: row.id
            for row in self._session.scalars(select(SettlementRecordRow))
        }


def _fingerprint(record: SettlementRecord) -> str:
    from .mappers import row_fingerprint

    return row_fingerprint(record)


# ----------------------------------------------------------------------
def _config_to_dict(config: MatchingConfig) -> dict[str, object]:
    return {
        "settlement_lag": list(config.settlement_lag),
        "amount_tolerance_minor": config.amount_tolerance.minor_units,
        "bucket_width_minor": config.bucket_width.minor_units,
        "accept_threshold": str(config.accept_threshold),
        "max_candidates": config.max_candidates,
        "weight_amount": str(config.weight_amount),
        "weight_date": str(config.weight_date),
        "weight_product": str(config.weight_product),
    }


def _config_from_dict(data: dict[str, object]) -> MatchingConfig:
    lag = data["settlement_lag"]
    assert isinstance(lag, list)
    return MatchingConfig(
        settlement_lag=(int(lag[0]), int(lag[1])),
        amount_tolerance=Money(int(str(data["amount_tolerance_minor"]))),
        bucket_width=Money(int(str(data["bucket_width_minor"]))),
        accept_threshold=Decimal(str(data["accept_threshold"])),
        max_candidates=int(str(data["max_candidates"])),
        weight_amount=Decimal(str(data["weight_amount"])),
        weight_date=Decimal(str(data["weight_date"])),
        weight_product=Decimal(str(data["weight_product"])),
    )


def _metrics_to_dict(metrics: ReconciliationMetrics) -> dict[str, object]:
    return {
        "order_count": metrics.order_count,
        "record_count": metrics.record_count,
        "duplicates_dropped": metrics.duplicates_dropped,
        "matched": metrics.matched,
        "amount_variance": metrics.amount_variance,
        "missing_in_ledger": metrics.missing_in_ledger,
        "missing_in_settlement": metrics.missing_in_settlement,
        "stage_exact": metrics.stage_exact,
        "stage_tolerant": metrics.stage_tolerant,
        "stage_fuzzy": metrics.stage_fuzzy,
        "needs_review": metrics.needs_review,
        "elapsed_seconds": metrics.elapsed_seconds,
        "candidate_comparisons": metrics.candidate_comparisons,
    }


def _metrics_from_dict(data: dict[str, object]) -> ReconciliationMetrics:
    return ReconciliationMetrics(
        order_count=int(str(data["order_count"])),
        record_count=int(str(data["record_count"])),
        duplicates_dropped=int(str(data["duplicates_dropped"])),
        matched=int(str(data["matched"])),
        amount_variance=int(str(data["amount_variance"])),
        missing_in_ledger=int(str(data["missing_in_ledger"])),
        missing_in_settlement=int(str(data["missing_in_settlement"])),
        stage_exact=int(str(data["stage_exact"])),
        stage_tolerant=int(str(data["stage_tolerant"])),
        stage_fuzzy=int(str(data["stage_fuzzy"])),
        # 舊的報告沒有這個欄位，用 0 當預設值而不是讓整份報告讀不回來。
        # 欄位增加是常見的，序列化的讀取端要能容忍。
        needs_review=int(str(data.get("needs_review", 0))),
        elapsed_seconds=float(str(data["elapsed_seconds"])),
        candidate_comparisons=int(str(data["candidate_comparisons"])),
    )
