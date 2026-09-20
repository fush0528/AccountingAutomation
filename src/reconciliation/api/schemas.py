"""API 的資料傳輸物件（DTO）。

**DTO 不等於領域模型。** 這是一條刻意的界線，理由有三個：

1. **領域模型不該被 API 的需求牽著走。** ``Money`` 內部存整數分，
   那是為了算術正確；但 API 回傳 ``{"amount": "944.70", "currency": "TWD"}``
   對前端才友善。如果讓領域模型直接序列化，遲早會有人為了讓 JSON 好看
   而去改 ``Money``——那是本末倒置。
2. **API 是對外承諾，領域模型是內部設計。** 領域模型重構時，
   API 不該跟著變動；反過來，API 要加一個欄位也不該逼領域模型長胖。
   兩者的變動頻率與變動原因完全不同。
3. **這一層是 OpenAPI 的來源。** W5 的前端會用 ``openapi-typescript``
   從這些 schema 自動產生 TypeScript 型別，做到前後端型別的單一真實來源。
   所以這裡的每一個欄位說明，最後都會變成前端開發者看到的文件。

金額一律以**字串**傳輸，不用浮點數。JSON 的 number 是 IEEE 754 雙精度，
把 ``944.70`` 放進去再拿出來不保證還是 ``944.70``——整個專案花了這麼多
力氣避開浮點數，不該在最後一哩前功盡棄。
"""

from __future__ import annotations

from datetime import date, datetime

from pydantic import BaseModel, ConfigDict, Field

from ..domain.matching import ReconciliationMetrics, ReconciliationReport
from ..domain.models import ImportBatch, MatchResult, Order, SettlementRecord
from ..domain.money import Money
from ..parsers.base import RowError
from ..services.import_service import ImportOutcome

__all__ = [
    "ErrorResponse",
    "ImportResponse",
    "MatchResultOut",
    "MetricsOut",
    "MoneyOut",
    "OrderOut",
    "PlatformOut",
    "ReconciliationRequest",
    "ReconciliationResponse",
    "ReportResponse",
    "SettlementRecordOut",
]


class MoneyOut(BaseModel):
    """金額。

    ``amount`` 是字串而不是數字，避免 JSON 的浮點數表示破壞精度。
    ``minor_units`` 一併回傳，讓需要精確計算的前端可以直接用整數。
    """

    model_config = ConfigDict(frozen=True)

    amount: str = Field(description='以元為單位的十進位字串，例如 "944.70"')
    currency: str = Field(description="ISO 4217 幣別代碼", examples=["TWD"])
    minor_units: int = Field(description="最小貨幣單位的整數值，例如 94470")

    @classmethod
    def of(cls, money: Money) -> MoneyOut:
        return cls(
            amount=str(money.to_decimal()),
            currency=money.currency,
            minor_units=money.minor_units,
        )

    @classmethod
    def maybe(cls, money: Money | None) -> MoneyOut | None:
        return cls.of(money) if money is not None else None


class PlatformOut(BaseModel):
    code: str
    display_name: str


class RowErrorOut(BaseModel):
    row_number: int
    message: str

    @classmethod
    def of(cls, error: RowError) -> RowErrorOut:
        return cls(row_number=error.row_number, message=error.message)


class ImportResponse(BaseModel):
    """匯入結果。

    ``was_duplicate`` 為真時代表這份檔案先前已經匯入過，
    本次沒有寫入任何資料——呼叫端必須能分辨這兩種「成功」。
    """

    batch_id: str
    platform: str
    filename: str
    content_hash: str = Field(description="檔案內容的 SHA-256，冪等匯入的依據")
    imported_at: datetime
    row_count: int
    accepted_count: int
    rejected_count: int
    skipped_duplicate_rows: int = Field(
        default=0, description="檔案指紋不同、但個別列先前已存在而被略過的筆數"
    )
    was_duplicate: bool = Field(description="這份檔案先前是否已匯入過。為真時本次未寫入任何資料")
    message: str
    row_errors: list[RowErrorOut] = Field(default_factory=list)

    @classmethod
    def of(cls, outcome: ImportOutcome) -> ImportResponse:
        b = outcome.batch
        return cls(
            batch_id=b.batch_id,
            platform=b.platform,
            filename=b.filename,
            content_hash=b.content_hash,
            imported_at=b.imported_at,
            row_count=b.row_count,
            accepted_count=b.accepted_count,
            rejected_count=b.rejected_count,
            skipped_duplicate_rows=outcome.skipped_duplicate_rows,
            was_duplicate=outcome.was_duplicate,
            message=outcome.message,
            row_errors=[RowErrorOut.of(e) for e in outcome.row_errors[:20]],
        )


class BatchOut(BaseModel):
    batch_id: str
    platform: str
    filename: str
    imported_at: datetime
    row_count: int
    accepted_count: int
    rejected_count: int

    @classmethod
    def of(cls, batch: ImportBatch) -> BatchOut:
        return cls(
            batch_id=batch.batch_id,
            platform=batch.platform,
            filename=batch.filename,
            imported_at=batch.imported_at,
            row_count=batch.row_count,
            accepted_count=batch.accepted_count,
            rejected_count=batch.rejected_count,
        )


class OrderOut(BaseModel):
    order_id: str
    platform: str
    external_order_id: str
    ordered_at: datetime
    product_name: str
    quantity: int
    gross_amount: MoneyOut
    expected_fee: MoneyOut
    expected_net: MoneyOut

    @classmethod
    def of(cls, order: Order) -> OrderOut:
        return cls(
            order_id=order.order_id,
            platform=order.platform,
            external_order_id=order.external_order_id,
            ordered_at=order.ordered_at,
            product_name=order.product_name,
            quantity=order.quantity,
            gross_amount=MoneyOut.of(order.gross_amount),
            expected_fee=MoneyOut.of(order.expected_fee),
            expected_net=MoneyOut.of(order.expected_net),
        )


class SettlementRecordOut(BaseModel):
    platform: str
    settled_at: date
    external_order_id: str | None
    product_name: str | None
    source_row: int
    record_type: str
    gross_amount: MoneyOut
    fee_amount: MoneyOut
    net_amount: MoneyOut

    @classmethod
    def of(cls, record: SettlementRecord) -> SettlementRecordOut:
        return cls(
            platform=record.platform,
            settled_at=record.settled_at,
            external_order_id=record.external_order_id,
            product_name=record.product_name,
            source_row=record.source_row,
            record_type=record.record_type.value,
            gross_amount=MoneyOut.of(record.gross_amount),
            fee_amount=MoneyOut.of(record.fee_amount),
            net_amount=MoneyOut.of(record.net_amount),
        )


class CandidateOut(BaseModel):
    """Stage 3 的候選。``reasons`` 是給人看的，不是給程式判斷的。"""

    order_id: str
    external_order_id: str
    product_name: str
    score: str
    reasons: list[str]


class MatchResultOut(BaseModel):
    outcome: str = Field(
        description="四象限之一：matched / amount_variance / "
        "missing_in_ledger / missing_in_settlement"
    )
    stage: int | None = Field(default=None, description="1 精確、2 容差、3 模糊；未匹配時為 null")
    needs_review: bool = Field(description="是否需要人工確認。只有 Stage 1 的 matched 為 false")
    order: OrderOut | None = None
    # 這兩個刻意不給預設值。有預設值的 Pydantic 欄位在 OpenAPI 裡是
    # optional，產生出來的 TypeScript 型別就會是 `T[] | undefined`，
    # 逼前端去處理一個實際上不可能發生的狀態。`of()` 永遠會給值，
    # 那就讓型別誠實地說出來。
    records: list[SettlementRecordOut]
    settled_amount: MoneyOut | None = Field(
        default=None, description="平台實際撥款合計（多列時已抵銷退款）"
    )
    variance: MoneyOut | None = None
    variance_reason: str | None = None
    candidates: list[CandidateOut]

    @classmethod
    def of(cls, result: MatchResult) -> MatchResultOut:
        return cls(
            outcome=result.outcome.value,
            stage=result.stage.value if result.stage else None,
            needs_review=result.needs_review,
            order=OrderOut.of(result.order) if result.order else None,
            records=[SettlementRecordOut.of(r) for r in result.records],
            settled_amount=MoneyOut.maybe(result.settled_amount),
            variance=MoneyOut.maybe(result.variance),
            variance_reason=(result.variance_reason.value if result.variance_reason else None),
            candidates=[
                CandidateOut(
                    order_id=c.order.order_id,
                    external_order_id=c.order.external_order_id,
                    product_name=c.order.product_name,
                    score=str(c.score),
                    reasons=list(c.reasons),
                )
                for c in result.candidates
            ],
        )


class MetricsOut(BaseModel):
    order_count: int
    record_count: int
    duplicates_dropped: int
    matched: int
    amount_variance: int
    missing_in_ledger: int
    missing_in_settlement: int
    stage_exact: int
    stage_tolerant: int
    stage_fuzzy: int
    needs_review: int = Field(description="需人工確認的筆數。只有 Stage 1 的 matched 不需要")
    automation_rate: float = Field(
        description="完全不需人工介入的比例。只計 Stage 1——那是唯一可以直接入帳的情況"
    )
    review_rate: float
    elapsed_seconds: float
    candidate_comparisons: int = Field(
        description="Stage 3 建了候選索引之後，實際做了幾次評分比對"
    )
    fuzzy_record_count: int = Field(
        description="有幾列結算資料進到 Stage 3（前兩階段靠編號配不掉的）"
    )
    naive_comparisons: int = Field(
        description=(
            "同樣的工作不建索引要做幾次比對（進入 Stage 3 的結算列 × 尚未被認領的訂單）。"
            "不是全部結算列 × 全部訂單——那會把索引的功勞誇大一個數量級"
        )
    )

    @classmethod
    def of(cls, metrics: ReconciliationMetrics) -> MetricsOut:
        return cls(
            order_count=metrics.order_count,
            record_count=metrics.record_count,
            duplicates_dropped=metrics.duplicates_dropped,
            matched=metrics.matched,
            amount_variance=metrics.amount_variance,
            missing_in_ledger=metrics.missing_in_ledger,
            missing_in_settlement=metrics.missing_in_settlement,
            stage_exact=metrics.stage_exact,
            stage_tolerant=metrics.stage_tolerant,
            stage_fuzzy=metrics.stage_fuzzy,
            needs_review=metrics.needs_review,
            automation_rate=metrics.automation_rate,
            review_rate=metrics.review_rate,
            elapsed_seconds=metrics.elapsed_seconds,
            candidate_comparisons=metrics.candidate_comparisons,
            fuzzy_record_count=metrics.fuzzy_record_count,
            naive_comparisons=metrics.naive_comparisons,
        )


class ConfigOut(BaseModel):
    """對帳當下使用的設定。

    報告一定要附上設定。「自動匹配率 79.6%」在門檻 0.5 與 0.9 下
    是完全不同的兩件事，沒有設定的數字沒有意義。
    """

    accept_threshold: str
    amount_tolerance: MoneyOut
    settlement_lag_days: list[int]
    weight_amount: str
    weight_date: str
    weight_product: str


class ReconciliationRequest(BaseModel):
    platform: str | None = Field(default=None, description="留空則對所有平台")
    accept_threshold: str | None = Field(
        default=None, description='Stage 3 採納門檻，0 到 1 之間，例如 "0.82"'
    )


class ReconciliationResponse(BaseModel):
    run_id: str
    metrics: MetricsOut
    config: ConfigOut


class ReportResponse(BaseModel):
    run_id: str
    metrics: MetricsOut
    config: ConfigOut
    results: list[MatchResultOut]
    total: int = Field(description="符合篩選條件的總筆數（未分頁前）")

    @classmethod
    def of(
        cls,
        run_id: str,
        report: ReconciliationReport,
        results: list[MatchResult],
        total: int,
    ) -> ReportResponse:
        return cls(
            run_id=run_id,
            metrics=MetricsOut.of(report.metrics),
            config=ConfigOut(
                accept_threshold=str(report.config.accept_threshold),
                amount_tolerance=MoneyOut.of(report.config.amount_tolerance),
                settlement_lag_days=list(report.config.settlement_lag),
                weight_amount=str(report.config.weight_amount),
                weight_date=str(report.config.weight_date),
                weight_product=str(report.config.weight_product),
            ),
            results=[MatchResultOut.of(r) for r in results],
            total=total,
        )


class ErrorResponse(BaseModel):
    """統一的錯誤格式。

    ``code`` 給程式判斷，``message`` 給人看。兩者分開，前端才不需要
    比對中文字串來決定行為。
    """

    code: str
    message: str
