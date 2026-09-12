"""差異歸因的測試。

歸因的核心是把總差異拆成 gross 與 fee 兩個分量。因為 net = gross - fee，
所以 variance = gross_delta - fee_delta 這個恆等式必須永遠成立——
第一組測試守的就是它。
"""

from __future__ import annotations

from datetime import date, datetime

from reconciliation.domain.models import (
    Order,
    SettlementRecord,
    SettlementRecordType,
    VarianceReason,
)
from reconciliation.domain.money import Money
from reconciliation.domain.variance import ROUNDING_TOLERANCE, analyse_variance


def order(gross: str = "1000", fee: str = "55.30") -> Order:
    return Order(
        order_id="ORD-1",
        platform="platform_a",
        external_order_id="AAAA0001",
        ordered_at=datetime(2026, 3, 1),
        product_name="無線滑鼠",
        quantity=1,
        gross_amount=Money.from_str(gross),
        expected_fee=Money.from_str(fee),
    )


def record(
    gross: str = "1000",
    fee: str = "55.30",
    kind: SettlementRecordType = SettlementRecordType.SALE,
) -> SettlementRecord:
    g, f = Money.from_str(gross), Money.from_str(fee)
    return SettlementRecord(
        platform="platform_a",
        settled_at=date(2026, 3, 15),
        gross_amount=g,
        fee_amount=f,
        net_amount=g - f,
        record_type=kind,
        external_order_id="SP-AAAA0001",
    )


class TestDecomposition:
    def test_variance_equals_gross_delta_minus_fee_delta(self) -> None:
        """恆等式：net = gross - fee，所以差異也必須這樣拆得開。"""
        analysis = analyse_variance(order(), [record(gross="1080", fee="75.30")])
        assert analysis.gross_delta == Money.from_str("80")
        assert analysis.fee_delta == Money.from_str("20")
        assert analysis.variance == analysis.gross_delta - analysis.fee_delta
        assert analysis.variance == Money.from_str("60")

    def test_no_difference_means_zero_everywhere(self) -> None:
        analysis = analyse_variance(order(), [record()])
        assert analysis.variance.is_zero
        assert analysis.gross_delta.is_zero
        assert analysis.fee_delta.is_zero
        assert not analysis.is_significant


class TestAttribution:
    def test_fee_only_is_rate_change(self) -> None:
        analysis = analyse_variance(order(), [record(fee="85.30")])
        assert analysis.reason is VarianceReason.FEE_RATE_CHANGE
        assert "費率" in analysis.explanation

    def test_gross_only_is_subsidy(self) -> None:
        analysis = analyse_variance(order(), [record(gross="1080")])
        assert analysis.reason is VarianceReason.SHIPPING_SUBSIDY
        assert "運費" in analysis.explanation

    def test_refund_row_takes_priority(self) -> None:
        """有退款列時金額本來就該對不上，那不是異常。"""
        refund = record(gross="-300", fee="-16.59", kind=SettlementRecordType.REFUND)
        analysis = analyse_variance(order(), [record(), refund])
        assert analysis.reason is VarianceReason.PARTIAL_REFUND

    def test_sub_dollar_is_rounding(self) -> None:
        analysis = analyse_variance(order(), [record(fee="55.80")])
        assert analysis.reason is VarianceReason.ROUNDING
        assert not analysis.is_significant

    def test_both_moved_stays_unknown(self) -> None:
        """兩邊都動了就不要硬猜。錯誤的解釋比沒有解釋更危險——
        使用者會以為系統懂了，於是不去追查。"""
        analysis = analyse_variance(order(), [record(gross="1200", fee="90")])
        assert analysis.reason is VarianceReason.UNKNOWN
        assert "人工" in analysis.explanation

    def test_inconsistent_row_is_flagged_as_parser_suspect(self) -> None:
        """gross 與 fee 都對，net 卻不對 → 該列自己的加減法不成立。"""
        broken = SettlementRecord(
            platform="platform_a",
            settled_at=date(2026, 3, 15),
            gross_amount=Money.from_str("1000"),
            fee_amount=Money.from_str("55.30"),
            net_amount=Money.from_str("900"),
            external_order_id="SP-AAAA0001",
        )
        analysis = analyse_variance(order(), [broken])
        assert analysis.reason is VarianceReason.UNKNOWN
        assert "尚未建模" in analysis.explanation

    def test_explanation_is_ready_to_display(self) -> None:
        """explanation 要能直接顯示在介面上，不需要前端再翻譯 enum。"""
        analysis = analyse_variance(order(), [record(fee="85.30")])
        assert len(analysis.explanation) > 10
        assert "30.00 TWD" in analysis.explanation


class TestSignificance:
    def test_tolerance_boundary(self) -> None:
        exactly_one = analyse_variance(order(), [record(fee="56.30")])
        assert exactly_one.variance == -ROUNDING_TOLERANCE
        assert not exactly_one.is_significant

        just_over = analyse_variance(order(), [record(fee="56.31")])
        assert just_over.is_significant
