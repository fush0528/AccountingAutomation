"""三階段匹配引擎的測試。

每一組測試對應一個真實會發生的對帳情境。測試名稱刻意寫成敘述句，
`pytest -v` 列出來就是一份規格書。
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import Decimal

import pytest

from reconciliation.domain.matching import MatchingConfig, reconcile
from reconciliation.domain.models import (
    MatchOutcome,
    MatchStage,
    Order,
    SettlementRecord,
    SettlementRecordType,
    VarianceReason,
)
from reconciliation.domain.money import Money

BASE_DATE = datetime(2026, 3, 1, 10, 0)


def make_order(
    n: int = 1,
    *,
    external: str | None = None,
    gross: str = "1000",
    fee: str = "55.30",
    product: str = "無線滑鼠",
    platform: str = "platform_a",
    days: int = 0,
) -> Order:
    return Order(
        order_id=f"ORD-{n:04d}",
        platform=platform,
        external_order_id=external if external is not None else f"AAAA{n:04d}",
        ordered_at=BASE_DATE + timedelta(days=days),
        product_name=product,
        quantity=1,
        gross_amount=Money.from_str(gross),
        expected_fee=Money.from_str(fee),
    )


def make_record(
    *,
    external: str | None = "SP-AAAA0001",
    gross: str = "1000",
    fee: str = "55.30",
    net: str | None = None,
    kind: SettlementRecordType = SettlementRecordType.SALE,
    platform: str = "platform_a",
    settled: date = date(2026, 3, 15),
    product: str | None = "無線滑鼠",
    row: int = 2,
) -> SettlementRecord:
    g = Money.from_str(gross)
    f = Money.from_str(fee)
    return SettlementRecord(
        platform=platform,
        settled_at=settled,
        gross_amount=g,
        fee_amount=f,
        net_amount=Money.from_str(net) if net is not None else g - f,
        record_type=kind,
        external_order_id=external,
        product_name=product,
        source_row=row,
    )


class TestStage1Exact:
    def test_matching_ids_and_amounts_are_auto_cleared(self) -> None:
        report = reconcile([make_order()], [make_record()])
        (result,) = report.results
        assert result.outcome is MatchOutcome.MATCHED
        assert result.stage is MatchStage.EXACT
        assert not result.needs_review
        assert report.metrics.automation_rate == 1.0

    def test_platform_prefix_and_case_do_not_break_matching(self) -> None:
        """結算單寫 'sp-aaaa0001'，我方存 'AAAA0001'——正規化後應該對得上。"""
        report = reconcile([make_order()], [make_record(external="sp-aaaa0001")])
        assert report.results[0].stage is MatchStage.EXACT

    def test_fullwidth_digits_still_match(self) -> None:
        order = make_order(external="24A7")
        report = reconcile([order], [make_record(external="＃２４Ａ７")])
        assert report.results[0].outcome is MatchOutcome.MATCHED


class TestStage2Tolerant:
    def test_fee_difference_is_attributed_to_rate_change(self) -> None:
        """訂單金額一致、手續費不同 → 費率調整。"""
        report = reconcile([make_order()], [make_record(fee="85.30")])
        (result,) = report.results
        assert result.outcome is MatchOutcome.AMOUNT_VARIANCE
        assert result.stage is MatchStage.TOLERANT
        assert result.variance_reason is VarianceReason.FEE_RATE_CHANGE
        assert result.variance == Money.from_str("-30")
        assert result.needs_review

    def test_gross_difference_is_attributed_to_subsidy(self) -> None:
        """手續費一致、訂單金額不同 → 運費補貼之類。"""
        report = reconcile([make_order()], [make_record(gross="1080")])
        (result,) = report.results
        assert result.variance_reason is VarianceReason.SHIPPING_SUBSIDY
        assert result.variance == Money.from_str("80")

    def test_sub_dollar_difference_is_rounding(self) -> None:
        report = reconcile([make_order()], [make_record(fee="55.60")])
        (result,) = report.results
        assert result.variance_reason is VarianceReason.ROUNDING

    def test_unexplainable_difference_stays_unknown(self) -> None:
        """兩邊都動了就不要硬猜——UNKNOWN 比錯誤的解釋安全。"""
        report = reconcile([make_order()], [make_record(gross="1200", fee="90")])
        (result,) = report.results
        assert result.variance_reason is VarianceReason.UNKNOWN


class TestMultipleRecordsPerOrder:
    def test_sale_plus_refund_are_grouped_into_one_result(self) -> None:
        """同一訂單的銷售列與退款列必須合併成一筆結論，不是兩筆。"""
        order = make_order()
        sale = make_record()
        refund = make_record(
            gross="-300",
            fee="-16.59",
            kind=SettlementRecordType.REFUND,
            settled=date(2026, 3, 20),
            row=9,
        )
        report = reconcile([order], [sale, refund])
        assert len(report.results) == 1
        (result,) = report.results
        assert len(result.records) == 2
        assert result.outcome is MatchOutcome.AMOUNT_VARIANCE
        assert result.variance_reason is VarianceReason.PARTIAL_REFUND

    def test_settled_amount_nets_the_refund_out(self) -> None:
        order = make_order()
        full_refund = make_record(
            gross="-1000",
            fee="-55.30",
            kind=SettlementRecordType.REFUND,
            settled=date(2026, 3, 20),
        )
        report = reconcile([order], [make_record(), full_refund])
        assert report.results[0].settled_amount == Money.zero()


class TestDeduplication:
    def test_identical_rows_are_dropped_and_reported(self) -> None:
        """平台重複匯出同一列，不能讓它被算兩次。"""
        record = make_record()
        report = reconcile([make_order()], [record, record, record])
        assert report.metrics.duplicates_dropped == 2
        assert report.metrics.record_count == 1
        assert report.results[0].outcome is MatchOutcome.MATCHED

    def test_legitimate_multiple_rows_are_not_deduplicated(self) -> None:
        """同一訂單的銷售與退款長得不一樣，不該被當成重複。"""
        report = reconcile(
            [make_order()],
            [
                make_record(),
                make_record(
                    gross="-300",
                    fee="-16.59",
                    kind=SettlementRecordType.REFUND,
                    settled=date(2026, 3, 20),
                ),
            ],
        )
        assert report.metrics.duplicates_dropped == 0
        assert len(report.results[0].records) == 2


class TestStage3Fuzzy:
    def test_missing_order_id_is_recovered_by_amount_and_product(self) -> None:
        """整欄沒有訂單編號，靠金額、日期、商品名稱救回來。"""
        order = make_order()
        report = reconcile([order], [make_record(external=None)])
        (result,) = report.results
        assert result.outcome is MatchOutcome.MATCHED
        assert result.stage is MatchStage.FUZZY
        assert result.order is not None
        assert result.order.order_id == order.order_id

    def test_fuzzy_match_always_needs_review(self) -> None:
        """模糊匹配是猜的，再高分也要人看過。"""
        report = reconcile([make_order()], [make_record(external=None)])
        assert report.results[0].needs_review
        assert report.metrics.automation_rate == 0.0

    def test_low_confidence_produces_candidates_not_a_guess(self) -> None:
        """分數不夠就不要猜——列為漏記單，附上候選給人工判斷。"""
        order = make_order(product="無線滑鼠", gross="1000", fee="55.30")
        record = make_record(external=None, product="完全不同的商品", gross="1090")
        report = reconcile([order], [record])
        missing = report.by_outcome(MatchOutcome.MISSING_IN_LEDGER)
        assert len(missing) == 1
        assert missing[0].candidates, "應該要附上候選清單"
        assert missing[0].candidates[0].order.order_id == order.order_id

    def test_candidates_carry_human_readable_reasons(self) -> None:
        record = make_record(external=None, product="完全不同的商品", gross="1090")
        report = reconcile([make_order()], [record])
        candidate = report.by_outcome(MatchOutcome.MISSING_IN_LEDGER)[0].candidates[0]
        assert candidate.reasons
        assert any("金額" in r for r in candidate.reasons)

    def test_one_order_cannot_be_claimed_twice(self) -> None:
        """兩筆無編號的結算列不能同時認領同一筆訂單。"""
        order = make_order()
        records = [
            make_record(external=None, row=2),
            make_record(external=None, row=3, settled=date(2026, 3, 16)),
        ]
        report = reconcile([order], records)
        claimed = [r for r in report.results if r.order is not None and r.records]
        assert len(claimed) == 1

    def test_different_platforms_never_match(self) -> None:
        """金額完全相同也不能跨平台配對。"""
        order = make_order(platform="platform_a")
        record = make_record(external=None, platform="platform_b")
        report = reconcile([order], [record])
        assert report.by_outcome(MatchOutcome.MISSING_IN_LEDGER)
        assert report.by_outcome(MatchOutcome.MISSING_IN_SETTLEMENT)

    def test_settlement_outside_the_lag_window_is_not_a_candidate(self) -> None:
        """訂單成立前就撥款不合理，一年後才撥款也不合理。"""
        order = make_order(days=0)
        record = make_record(external=None, settled=date(2026, 8, 1))
        report = reconcile([order], [record])
        assert report.by_outcome(MatchOutcome.MISSING_IN_LEDGER)


class TestFourQuadrants:
    def test_missing_in_ledger_when_platform_has_an_unknown_order(self) -> None:
        report = reconcile([], [make_record(external="SP-NOTOURS")])
        (result,) = report.results
        assert result.outcome is MatchOutcome.MISSING_IN_LEDGER
        assert result.order is None

    def test_missing_in_settlement_when_platform_never_paid(self) -> None:
        report = reconcile([make_order()], [])
        (result,) = report.results
        assert result.outcome is MatchOutcome.MISSING_IN_SETTLEMENT
        assert result.records == ()

    def test_all_four_quadrants_in_one_run(self) -> None:
        orders = [
            make_order(1),  # 完全相符
            make_order(2, external="AAAA0002", gross="2000", fee="110"),  # 金額有差
            make_order(3, external="AAAA0003", gross="3000", fee="165"),  # 平台沒撥款
        ]
        records = [
            make_record(external="SP-AAAA0001"),
            make_record(external="SP-AAAA0002", gross="2000", fee="150"),
            make_record(external="SP-UNKNOWN", gross="777", fee="42"),
        ]
        report = reconcile(orders, records)
        assert len(report.by_outcome(MatchOutcome.MATCHED)) == 1
        assert len(report.by_outcome(MatchOutcome.AMOUNT_VARIANCE)) == 1
        assert len(report.by_outcome(MatchOutcome.MISSING_IN_SETTLEMENT)) == 1
        assert len(report.by_outcome(MatchOutcome.MISSING_IN_LEDGER)) == 1


class TestMetrics:
    def test_automation_rate_counts_only_stage_one(self) -> None:
        """Stage 2 與 Stage 3 都要人看，不能算進自動化率。

        這個測試守住的是誠實性：把它們算進去數字會從 33% 跳到 100%，
        但那個數字是假的。
        """
        orders = [make_order(1), make_order(2, external="AAAA0002")]
        records = [
            make_record(external="SP-AAAA0001"),  # Stage 1
            make_record(external="SP-AAAA0002", fee="95.30"),  # Stage 2
        ]
        report = reconcile(orders, records)
        assert report.metrics.stage_exact == 1
        assert report.metrics.stage_tolerant == 1
        assert report.metrics.automation_rate == 0.5
        assert report.metrics.review_rate == 0.5

    def test_every_record_and_order_is_accounted_for(self) -> None:
        """沒有任何一筆資料可以憑空消失——這是對帳系統的基本要求。"""
        orders = [make_order(i, external=f"AAAA{i:04d}") for i in range(1, 6)]
        records = [make_record(external=f"SP-AAAA{i:04d}") for i in range(1, 4)]
        records.append(make_record(external="SP-GHOST", gross="500", fee="27"))
        report = reconcile(orders, records)

        seen_orders = {r.order.order_id for r in report.results if r.order}
        seen_rows = [rec for r in report.results for rec in r.records]
        assert len(seen_orders) == len(orders)
        assert len(seen_rows) == len(records)


class TestConfig:
    def test_bucket_width_must_exceed_tolerance(self) -> None:
        """設定錯了要當場爆，不要安靜地漏掉候選。"""
        with pytest.raises(ValueError, match="bucket_width"):
            MatchingConfig(
                amount_tolerance=Money.from_str("500"),
                bucket_width=Money.from_str("100"),
            )

    def test_weights_must_sum_to_one(self) -> None:
        with pytest.raises(ValueError, match="權重"):
            MatchingConfig(
                weight_amount=Decimal("0.5"),
                weight_date=Decimal("0.5"),
                weight_product=Decimal("0.5"),
            )

    def test_threshold_changes_the_outcome(self) -> None:
        """同一份資料、不同門檻，結果不同——所以報告一定要附上設定。"""
        order = make_order()
        record = make_record(external=None, product="無線滑鼠 黑", gross="1040")

        strict = reconcile([order], [record], MatchingConfig(accept_threshold=Decimal("0.99")))
        loose = reconcile([order], [record], MatchingConfig(accept_threshold=Decimal("0.60")))

        # 嚴格門檻下不敢認，只列出候選
        assert strict.by_outcome(MatchOutcome.MISSING_IN_LEDGER)
        assert strict.by_outcome(MatchOutcome.MISSING_IN_SETTLEMENT)

        # 寬鬆門檻下認了這筆訂單。金額仍有 40 元差異，所以落在
        # AMOUNT_VARIANCE 而不是 MATCHED——找到人跟金額對得上是兩件事。
        (found,) = loose.by_outcome(MatchOutcome.AMOUNT_VARIANCE)
        assert found.stage is MatchStage.FUZZY
        assert found.order is not None
        assert found.order.order_id == order.order_id

    def test_report_carries_the_config_used(self) -> None:
        config = MatchingConfig(accept_threshold=Decimal("0.9"))
        report = reconcile([make_order()], [make_record()], config)
        assert report.config.accept_threshold == Decimal("0.9")
