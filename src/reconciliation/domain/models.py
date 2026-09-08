"""對帳領域模型。

這一層是整個系統的核心，規則只有一條：**不 import 任何框架**。
沒有 FastAPI、沒有 SQLAlchemy、沒有 pandas、沒有 openpyxl。
因此這裡的每一個型別都能在沒有資料庫、沒有網路的情況下完整測試，
而 ``services/`` 與 ``repositories/`` 是依賴這一層，不是反過來。

領域的核心概念是「兩份事實」：

* :class:`Order` —— **我方**系統認為發生了什麼（銷售端的事實）
* :class:`SettlementRecord` —— **平台**結算單說發生了什麼（撥款端的事實）

對帳就是為這兩份事實建立對應關係，並且對「對不上」的部分給出解釋。
:class:`MatchResult` 就是那個解釋。
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from types import MappingProxyType

from .money import Money

__all__ = [
    "Candidate",
    "DomainError",
    "ImportBatch",
    "MatchOutcome",
    "MatchResult",
    "MatchStage",
    "Order",
    "SettlementRecord",
    "SettlementRecordType",
    "VarianceReason",
]


class DomainError(ValueError):
    """領域不變量被違反時拋出。"""


# ----------------------------------------------------------------------
# 列舉
# ----------------------------------------------------------------------
class SettlementRecordType(Enum):
    """結算單上一列的性質。

    退款與調整之所以要獨立出來，是因為它們讓「一筆訂單對一列結算」的
    天真假設失效：同一個訂單編號可能同時出現一列銷售與一列退款。
    """

    SALE = "sale"
    REFUND = "refund"
    ADJUSTMENT = "adjustment"


class MatchStage(Enum):
    """這筆對應是在哪一階段建立的。

    階段本身就是信心水準的粗略代理：EXACT 可直接入帳，
    FUZZY 一律進人工確認佇列。
    """

    EXACT = 1
    TOLERANT = 2
    FUZZY = 3


class MatchOutcome(Enum):
    """對帳報告的四個象限。"""

    MATCHED = "matched"
    """兩邊都有、金額相符。"""

    AMOUNT_VARIANCE = "amount_variance"
    """兩邊都有，但金額不符——差額已計算，並嘗試歸因。"""

    MISSING_IN_LEDGER = "missing_in_ledger"
    """平台有、我方無。通常是漏記單，或是平台自行產生的調整列。"""

    MISSING_IN_SETTLEMENT = "missing_in_settlement"
    """我方有、平台無。通常是尚未撥款，也可能是真的沒收到錢。"""


class VarianceReason(Enum):
    """金額差異的歸因。

    歸因規則寫在 ``domain/variance.py``（W3 實作）。留 UNKNOWN 是刻意的：
    無法解釋的差異必須顯眼，不能被塞進某個看起來合理的分類裡。
    """

    FEE_RATE_CHANGE = "fee_rate_change"
    SHIPPING_SUBSIDY = "shipping_subsidy"
    PARTIAL_REFUND = "partial_refund"
    ROUNDING = "rounding"
    UNKNOWN = "unknown"


_EMPTY_RAW: Mapping[str, str] = MappingProxyType({})


# ----------------------------------------------------------------------
# 我方事實
# ----------------------------------------------------------------------
@dataclass(frozen=True, slots=True)
class Order:
    """我方系統中的一筆訂單。

    ``external_order_id`` 保留平台上的**原始**編號，不做任何清洗。
    正規化只在匹配時進行（見 ``domain/normalization.py``），
    這樣人工查核時才看得到平台實際印出來的字串。
    """

    order_id: str
    platform: str
    external_order_id: str
    ordered_at: datetime
    product_name: str
    quantity: int
    gross_amount: Money
    expected_fee: Money

    def __post_init__(self) -> None:
        if not self.order_id.strip():
            raise DomainError("order_id 不可為空")
        if not self.platform.strip():
            raise DomainError("platform 不可為空")
        if self.quantity <= 0:
            raise DomainError(f"quantity 必須為正整數，收到 {self.quantity}")
        if self.gross_amount.currency != self.expected_fee.currency:
            raise DomainError("gross_amount 與 expected_fee 幣別必須一致")
        if self.gross_amount.is_negative:
            raise DomainError("gross_amount 不可為負；退貨請以結算單的 REFUND 列表示")

    @property
    def expected_net(self) -> Money:
        """我方預期會收到的金額。對帳時拿它跟平台的 net_amount 比。"""
        return self.gross_amount - self.expected_fee


# ----------------------------------------------------------------------
# 平台事實
# ----------------------------------------------------------------------
@dataclass(frozen=True, slots=True)
class SettlementRecord:
    """平台結算單中的一列，已由 parser 正規化成統一結構。

    ``external_order_id`` 允許為 ``None``：實務上結算單確實有整列沒有訂單
    編號的情況（平台端的調整、補貼、跨期沖銷）。這種列會直接進入
    Stage 3 模糊匹配，而不是被丟掉——丟掉會讓總額對不起來。
    """

    platform: str
    settled_at: date
    gross_amount: Money
    fee_amount: Money
    net_amount: Money
    record_type: SettlementRecordType = SettlementRecordType.SALE
    external_order_id: str | None = None
    product_name: str | None = None
    source_row: int = 0
    raw: Mapping[str, str] = field(default_factory=lambda: _EMPTY_RAW)

    def __post_init__(self) -> None:
        if not self.platform.strip():
            raise DomainError("platform 不可為空")
        currencies = {
            self.gross_amount.currency,
            self.fee_amount.currency,
            self.net_amount.currency,
        }
        if len(currencies) != 1:
            raise DomainError(f"同一列的金額幣別必須一致，收到 {sorted(currencies)}")
        if self.external_order_id is not None and not self.external_order_id.strip():
            # 空字串與 None 是不同的意思，但對下游沒有差別，統一收斂成 None
            object.__setattr__(self, "external_order_id", None)
        object.__setattr__(self, "raw", MappingProxyType(dict(self.raw)))

    @property
    def residual(self) -> Money:
        """gross - fee - net。

        理論上應為零。不為零代表結算單上還有本模型沒有涵蓋的扣項
        （例如廣告費、金流費）。把它顯性化，而不是假設它是零——
        這個值不為零時，parser 就該被檢討。
        """
        return self.gross_amount - self.fee_amount - self.net_amount

    @property
    def has_order_id(self) -> bool:
        return self.external_order_id is not None


# ----------------------------------------------------------------------
# 匹配結果
# ----------------------------------------------------------------------
@dataclass(frozen=True, slots=True)
class Candidate:
    """Stage 3 模糊匹配產生的一個候選，附評分與可讀的理由。

    ``reasons`` 存在的原因很實際：人工確認時，操作者需要知道系統為什麼
    覺得這兩筆是同一件事。只給一個 0.87 的分數沒有人敢按確認。
    """

    order: Order
    score: Decimal
    reasons: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not (Decimal(0) <= self.score <= Decimal(1)):
            raise DomainError(f"score 必須落在 0 到 1 之間，收到 {self.score}")


@dataclass(frozen=True, slots=True)
class MatchResult:
    """一筆對帳結論。

    這個型別的 ``__post_init__`` 是整個領域層最重要的地方：它讓
    「MATCHED 卻沒有 order」這種不可能的狀態根本無法被建構出來。
    與其在下游到處寫防禦性檢查，不如讓非法狀態不可表示。
    """

    outcome: MatchOutcome
    stage: MatchStage | None = None
    order: Order | None = None
    record: SettlementRecord | None = None
    variance: Money | None = None
    variance_reason: VarianceReason | None = None
    candidates: tuple[Candidate, ...] = ()

    def __post_init__(self) -> None:
        if self.outcome in (MatchOutcome.MATCHED, MatchOutcome.AMOUNT_VARIANCE):
            if self.order is None or self.record is None:
                raise DomainError(f"{self.outcome.value} 必須同時具備 order 與 record")
            if self.stage is None:
                raise DomainError(f"{self.outcome.value} 必須標明是在哪一階段匹配的")
        if self.outcome is MatchOutcome.MISSING_IN_LEDGER and (
            self.record is None or self.order is not None
        ):
            raise DomainError("missing_in_ledger 必須只有 record，沒有 order")
        if self.outcome is MatchOutcome.MISSING_IN_SETTLEMENT and (
            self.order is None or self.record is not None
        ):
            raise DomainError("missing_in_settlement 必須只有 order，沒有 record")
        if self.outcome is MatchOutcome.AMOUNT_VARIANCE and (
            self.variance is None or self.variance.is_zero
        ):
            raise DomainError("amount_variance 的 variance 必須存在且不為零")
        if (
            self.outcome is MatchOutcome.MATCHED
            and self.variance is not None
            and not self.variance.is_zero
        ):
            raise DomainError("matched 的 variance 若存在必須為零")

    @property
    def needs_review(self) -> bool:
        """是否需要人工確認。"""
        return self.stage is MatchStage.FUZZY or self.outcome in (
            MatchOutcome.AMOUNT_VARIANCE,
            MatchOutcome.MISSING_IN_LEDGER,
            MatchOutcome.MISSING_IN_SETTLEMENT,
        )


# ----------------------------------------------------------------------
# 匯入批次
# ----------------------------------------------------------------------
@dataclass(frozen=True, slots=True)
class ImportBatch:
    """一次帳單匯入。

    ``content_hash`` 是冪等性的第一道防線：同一份檔案的 SHA-256 相同，
    重複上傳時直接回傳既有批次，不再寫入任何一列。第二道防線是資料庫上
    ``(platform, external_order_id, settled_at)`` 的唯一約束，第三道是
    整批寫入包在單一交易裡。三道合起來，才在「至少一次投遞」的環境下
    得到「恰好一次」的效果。
    """

    batch_id: str
    platform: str
    filename: str
    content_hash: str
    imported_at: datetime
    row_count: int = 0
    accepted_count: int = 0
    rejected_count: int = 0

    def __post_init__(self) -> None:
        if len(self.content_hash) != 64 or not all(
            c in "0123456789abcdef" for c in self.content_hash
        ):
            raise DomainError("content_hash 必須是 64 字元的小寫 SHA-256 十六進位字串")
        if self.accepted_count + self.rejected_count > self.row_count:
            raise DomainError("accepted + rejected 不可超過 row_count")
        for name in ("row_count", "accepted_count", "rejected_count"):
            if getattr(self, name) < 0:
                raise DomainError(f"{name} 不可為負")
