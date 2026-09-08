"""領域層：純 Python，零框架依賴。"""

from .models import (
    Candidate,
    DomainError,
    ImportBatch,
    MatchOutcome,
    MatchResult,
    MatchStage,
    Order,
    SettlementRecord,
    SettlementRecordType,
    VarianceReason,
)
from .money import CurrencyMismatchError, Money, MoneyError, money_sum
from .normalization import normalize_order_id, normalize_product_name

__all__ = [
    "Candidate",
    "CurrencyMismatchError",
    "DomainError",
    "ImportBatch",
    "MatchOutcome",
    "MatchResult",
    "MatchStage",
    "Money",
    "MoneyError",
    "Order",
    "SettlementRecord",
    "SettlementRecordType",
    "VarianceReason",
    "money_sum",
    "normalize_order_id",
    "normalize_product_name",
]
