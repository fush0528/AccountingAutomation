"""資料表定義。

這一層是唯一知道「資料長什麼樣子存進資料庫」的地方。領域層完全不知道
它的存在——``domain/`` 底下沒有任何一行 import SQLAlchemy。

三個貫穿全部資料表的決定：

**金額存整數分，不存浮點數。** 每個金額欄位是一對
``*_minor``（``BigInteger``）與 ``*_currency``（``String(3)``）。
資料庫層面也可以用 ``NUMERIC``，但整數在所有資料庫上的行為完全一致，
而且與 :class:`~reconciliation.domain.money.Money` 的內部表示直接對應，
存取時不需要任何轉換與捨入。理由見 ADR 0001。

**冪等性靠唯一約束，不靠應用層檢查。** 「先查有沒有、沒有再插入」在併發
下會失效——兩個請求可能同時查到「沒有」。把約束下放到資料庫，讓它成為
真正的不變量，應用層只負責處理被拒絕的情況。

**不用資料庫特有的型別。** 沒有 PostgreSQL 的 ``JSONB``、沒有
``ARRAY``、沒有 ``SERIAL``。這是「SQLite 換 PostgreSQL 只改連線字串」
這句話能成立的前提。
"""

from __future__ import annotations

from datetime import date, datetime

from sqlalchemy import (
    BigInteger,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship

__all__ = [
    "Base",
    "ImportBatchRow",
    "MatchResultRow",
    "OrderRow",
    "ReconciliationRunRow",
    "SettlementRecordRow",
]


class Base(DeclarativeBase):
    pass


class OrderRow(Base):
    """我方訂單。"""

    __tablename__ = "orders"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    order_id: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    platform: Mapped[str] = mapped_column(String(32), index=True)
    external_order_id: Mapped[str] = mapped_column(String(128), index=True)
    ordered_at: Mapped[datetime] = mapped_column(DateTime)
    product_name: Mapped[str] = mapped_column(String(256))
    quantity: Mapped[int] = mapped_column(Integer)
    gross_minor: Mapped[int] = mapped_column(BigInteger)
    expected_fee_minor: Mapped[int] = mapped_column(BigInteger)
    currency: Mapped[str] = mapped_column(String(3), default="TWD")

    __table_args__ = (Index("ix_orders_platform_external", "platform", "external_order_id"),)


class ImportBatchRow(Base):
    """一次帳單匯入。

    ``content_hash`` 的唯一約束是冪等性的第一道防線：同一份檔案再上傳一次，
    INSERT 會被資料庫拒絕，服務層據此回傳既有批次而不是重新匯入。
    """

    __tablename__ = "import_batches"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_id: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    platform: Mapped[str] = mapped_column(String(32), index=True)
    filename: Mapped[str] = mapped_column(String(256))
    content_hash: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    imported_at: Mapped[datetime] = mapped_column(DateTime)
    row_count: Mapped[int] = mapped_column(Integer, default=0)
    accepted_count: Mapped[int] = mapped_column(Integer, default=0)
    rejected_count: Mapped[int] = mapped_column(Integer, default=0)

    records: Mapped[list[SettlementRecordRow]] = relationship(
        back_populates="batch", cascade="all, delete-orphan"
    )


class SettlementRecordRow(Base):
    """平台結算單的一列。

    ``row_fingerprint`` 是第二道防線。它是「這一列的身分」的雜湊：
    平台、訂單編號、結算日、三個金額、交易類型。

    為什麼用雜湊欄位而不是直接對這幾個欄位下複合唯一約束？因為
    ``external_order_id`` 可以是 NULL，而 SQL 標準規定 NULL 彼此不相等——
    兩列都沒有訂單編號時，複合唯一約束擋不住它們。把 NULL 明確編碼進
    雜湊的輸入字串，就沒有這個問題。
    """

    __tablename__ = "settlement_records"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    batch_pk: Mapped[int] = mapped_column(ForeignKey("import_batches.id"), index=True)
    row_fingerprint: Mapped[str] = mapped_column(String(64), unique=True, index=True)

    platform: Mapped[str] = mapped_column(String(32), index=True)
    external_order_id: Mapped[str | None] = mapped_column(String(128), nullable=True, index=True)
    settled_at: Mapped[date] = mapped_column(DateTime().with_variant(DateTime, "sqlite"))
    gross_minor: Mapped[int] = mapped_column(BigInteger)
    fee_minor: Mapped[int] = mapped_column(BigInteger)
    net_minor: Mapped[int] = mapped_column(BigInteger)
    currency: Mapped[str] = mapped_column(String(3), default="TWD")
    record_type: Mapped[str] = mapped_column(String(16))
    product_name: Mapped[str | None] = mapped_column(String(256), nullable=True)
    source_row: Mapped[int] = mapped_column(Integer, default=0)
    raw_json: Mapped[str] = mapped_column(Text, default="{}")

    batch: Mapped[ImportBatchRow] = relationship(back_populates="records")


class ReconciliationRunRow(Base):
    """一次對帳執行。

    設定與指標存成 JSON 字串。用 ``Text`` 而不是 PostgreSQL 的 ``JSONB``，
    是為了保持資料庫可替換——這兩個欄位只會被整包讀寫，不需要在資料庫裡
    對它們做查詢。真的需要查詢時，該做的是把欄位拉出來正規化，
    而不是改用資料庫特有的型別。
    """

    __tablename__ = "reconciliation_runs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    run_id: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime)
    config_json: Mapped[str] = mapped_column(Text)
    metrics_json: Mapped[str] = mapped_column(Text)

    results: Mapped[list[MatchResultRow]] = relationship(
        back_populates="run", cascade="all, delete-orphan"
    )


class MatchResultRow(Base):
    """一筆對帳結論。

    ``outcome`` 與 ``stage`` 拉成獨立欄位（而不是塞進 JSON），因為前端要
    依象限篩選與計數——那是真正會被查詢的維度。候選清單則整包存 JSON，
    因為它只會跟著單筆結論一起顯示。
    """

    __tablename__ = "match_results"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    run_pk: Mapped[int] = mapped_column(ForeignKey("reconciliation_runs.id"), index=True)
    outcome: Mapped[str] = mapped_column(String(32), index=True)
    stage: Mapped[int | None] = mapped_column(Integer, nullable=True, index=True)
    order_id: Mapped[str | None] = mapped_column(String(64), nullable=True, index=True)
    record_ids_json: Mapped[str] = mapped_column(Text, default="[]")
    variance_minor: Mapped[int | None] = mapped_column(BigInteger, nullable=True)
    variance_reason: Mapped[str | None] = mapped_column(String(32), nullable=True)
    candidates_json: Mapped[str] = mapped_column(Text, default="[]")

    run: Mapped[ReconciliationRunRow] = relationship(back_populates="results")

    __table_args__ = (
        UniqueConstraint("run_pk", "id", name="uq_match_result_run"),
        Index("ix_match_results_run_outcome", "run_pk", "outcome"),
    )
