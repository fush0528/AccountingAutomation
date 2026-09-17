"""領域物件與資料表列之間的轉換。

集中放在這裡，而不是散在 repository 的各個方法裡，是因為轉換規則本身
就是設計決策的所在：金額怎麼拆成整數與幣別、``raw`` 怎麼序列化、
日期怎麼存。把它們放在一起，改動時才不會漏掉某一邊。
"""

from __future__ import annotations

import hashlib
import json
from datetime import date, datetime

from ..domain.models import (
    ImportBatch,
    Order,
    SettlementRecord,
    SettlementRecordType,
)
from ..domain.money import Money
from .tables import ImportBatchRow, OrderRow, SettlementRecordRow

__all__ = [
    "order_to_row",
    "record_to_row",
    "row_fingerprint",
    "row_to_batch",
    "row_to_order",
    "row_to_record",
]


def row_fingerprint(record: SettlementRecord) -> str:
    """一列結算記錄的身分雜湊。

    這是冪等匯入的第二道防線。輸入必須把 ``None`` 明確編碼成一個字串，
    因為 SQL 的 NULL 彼此不相等——如果直接對欄位下複合唯一約束，
    兩列都沒有訂單編號時會雙雙插入成功。

    ``source_row`` **不**納入：同一筆交易在不同檔案裡的列號當然不同，
    把它算進去就等於沒有去重。
    """
    parts = [
        record.platform,
        record.external_order_id if record.external_order_id is not None else "\x00NULL",
        record.settled_at.isoformat(),
        str(record.gross_amount.minor_units),
        str(record.fee_amount.minor_units),
        str(record.net_amount.minor_units),
        record.gross_amount.currency,
        record.record_type.value,
    ]
    return hashlib.sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


# ----------------------------------------------------------------------
def order_to_row(order: Order) -> OrderRow:
    return OrderRow(
        order_id=order.order_id,
        platform=order.platform,
        external_order_id=order.external_order_id,
        ordered_at=order.ordered_at,
        product_name=order.product_name,
        quantity=order.quantity,
        gross_minor=order.gross_amount.minor_units,
        expected_fee_minor=order.expected_fee.minor_units,
        currency=order.gross_amount.currency,
    )


def row_to_order(row: OrderRow) -> Order:
    return Order(
        order_id=row.order_id,
        platform=row.platform,
        external_order_id=row.external_order_id,
        ordered_at=row.ordered_at,
        product_name=row.product_name,
        quantity=row.quantity,
        gross_amount=Money(row.gross_minor, row.currency),
        expected_fee=Money(row.expected_fee_minor, row.currency),
    )


def record_to_row(record: SettlementRecord, batch_pk: int) -> SettlementRecordRow:
    return SettlementRecordRow(
        batch_pk=batch_pk,
        row_fingerprint=row_fingerprint(record),
        platform=record.platform,
        external_order_id=record.external_order_id,
        settled_at=datetime.combine(record.settled_at, datetime.min.time()),
        gross_minor=record.gross_amount.minor_units,
        fee_minor=record.fee_amount.minor_units,
        net_minor=record.net_amount.minor_units,
        currency=record.gross_amount.currency,
        record_type=record.record_type.value,
        product_name=record.product_name,
        source_row=record.source_row,
        raw_json=json.dumps(dict(record.raw), ensure_ascii=False),
    )


def row_to_record(row: SettlementRecordRow) -> SettlementRecord:
    settled: date = (
        row.settled_at.date() if isinstance(row.settled_at, datetime) else row.settled_at
    )
    return SettlementRecord(
        platform=row.platform,
        settled_at=settled,
        gross_amount=Money(row.gross_minor, row.currency),
        fee_amount=Money(row.fee_minor, row.currency),
        net_amount=Money(row.net_minor, row.currency),
        record_type=SettlementRecordType(row.record_type),
        external_order_id=row.external_order_id,
        product_name=row.product_name,
        source_row=row.source_row,
        raw=json.loads(row.raw_json or "{}"),
    )


def row_to_batch(row: ImportBatchRow) -> ImportBatch:
    return ImportBatch(
        batch_id=row.batch_id,
        platform=row.platform,
        filename=row.filename,
        content_hash=row.content_hash,
        imported_at=row.imported_at,
        row_count=row.row_count,
        accepted_count=row.accepted_count,
        rejected_count=row.rejected_count,
    )
