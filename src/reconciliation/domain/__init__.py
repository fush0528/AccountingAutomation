"""領域層：純 Python，零框架依賴。"""

from .matching import (
    MatchingConfig,
    ReconciliationMetrics,
    ReconciliationReport,
    reconcile,
)
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
from .variance import VarianceAnalysis, analyse_variance

__all__ = [
    "Candidate",
    "CurrencyMismatchError",
    "DomainError",
    "ImportBatch",
    "MatchOutcome",
    "MatchResult",
    "MatchStage",
    "MatchingConfig",
    "Money",
    "MoneyError",
    "Order",
    "ReconciliationMetrics",
    "ReconciliationReport",
    "SettlementRecord",
    "SettlementRecordType",
    "VarianceAnalysis",
    "VarianceReason",
    "analyse_variance",
    "money_sum",
    "normalize_order_id",
    "normalize_product_name",
    "reconcile",
]
