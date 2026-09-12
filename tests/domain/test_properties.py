"""以性質為基礎的測試（property-based testing）。

前面的測試都是「給定這個輸入，應該得到那個輸出」——我挑的案例，
只能證明我想到的情況是對的。這裡不一樣：由 Hypothesis 自動產生成千上萬
組隨機輸入，驗證的是**不管輸入是什麼都必須成立的性質**。

對帳系統有四條這樣的性質，違反任何一條都代表系統不能信：

1. **守恆**：每一筆訂單、每一列結算記錄都必須出現在結果裡，不多不少。
   對帳系統最不可原諒的錯誤就是讓資料憑空消失。
2. **不重複認領**：一筆訂單不能同時被兩個結論認領。
3. **金額一致**：``variance`` 必須等於實際撥款減去預期撥款，
   永遠。這條守住的是 Money 的算術與匯總邏輯。
4. **狀態合法**：任何輸入都不能讓引擎建構出不合法的 MatchResult。

Hypothesis 找到反例時會自動縮小到最小的失敗案例，通常一兩行就能看懂
哪裡壞了。這比隨機亂測有用太多。
"""

from __future__ import annotations

from datetime import date, datetime, timedelta

from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st

from reconciliation.domain.matching import reconcile
from reconciliation.domain.models import (
    MatchOutcome,
    Order,
    SettlementRecord,
    SettlementRecordType,
)
from reconciliation.domain.money import Money, money_sum

PLATFORMS = ["platform_a", "platform_b", "platform_c"]
PRODUCTS = ["無線滑鼠", "機械鍵盤", "USB 集線器", "螢幕支架", "行動電源"]


@st.composite
def orders(draw: st.DrawFn) -> Order:
    n = draw(st.integers(min_value=1, max_value=9999))
    gross_units = draw(st.integers(min_value=100, max_value=1_000_000))
    fee_units = draw(st.integers(min_value=0, max_value=gross_units))
    return Order(
        order_id=f"ORD-{n:05d}",
        platform=draw(st.sampled_from(PLATFORMS)),
        external_order_id=draw(
            st.text(
                alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789",
                min_size=4,
                max_size=12,
            )
        ),
        ordered_at=datetime(2026, 3, 1) + timedelta(days=draw(st.integers(0, 27))),
        product_name=draw(st.sampled_from(PRODUCTS)),
        quantity=draw(st.integers(min_value=1, max_value=10)),
        gross_amount=Money(gross_units),
        expected_fee=Money(fee_units),
    )


@st.composite
def settlement_records(draw: st.DrawFn) -> SettlementRecord:
    gross_units = draw(st.integers(min_value=-1_000_000, max_value=1_000_000))
    fee_units = draw(st.integers(min_value=-100_000, max_value=100_000))
    net_units = draw(st.integers(min_value=-1_000_000, max_value=1_000_000))
    order_id = draw(
        st.one_of(
            st.none(),
            st.text(
                alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-",
                min_size=1,
                max_size=16,
            ),
        )
    )
    return SettlementRecord(
        platform=draw(st.sampled_from(PLATFORMS)),
        settled_at=date(2026, 3, 1) + timedelta(days=draw(st.integers(0, 60))),
        gross_amount=Money(gross_units),
        fee_amount=Money(fee_units),
        net_amount=Money(net_units),
        record_type=draw(st.sampled_from(list(SettlementRecordType))),
        external_order_id=order_id,
        product_name=draw(st.one_of(st.none(), st.sampled_from(PRODUCTS))),
        source_row=draw(st.integers(min_value=2, max_value=9999)),
    )


SETTINGS = settings(
    max_examples=200,
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow],
)


class TestConservation:
    """守恆：資料不能憑空消失，也不能憑空出現。"""

    @given(
        order_list=st.lists(orders(), max_size=20, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=20),
    )
    @SETTINGS
    def test_every_order_appears_exactly_once(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        report = reconcile(order_list, record_list)
        appearances = [r.order.order_id for r in report.results if r.order is not None]
        assert sorted(appearances) == sorted(o.order_id for o in order_list)
        assert len(appearances) == len(set(appearances)), "同一筆訂單被認領了兩次"

    @given(
        order_list=st.lists(orders(), max_size=20, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=20),
    )
    @SETTINGS
    def test_every_deduplicated_record_appears_exactly_once(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        report = reconcile(order_list, record_list)
        seen = [rec for r in report.results for rec in r.records]
        assert len(seen) == report.metrics.record_count
        assert len(seen) + report.metrics.duplicates_dropped == len(record_list)

    @given(
        order_list=st.lists(orders(), max_size=15, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=15),
    )
    @SETTINGS
    def test_metric_counts_add_up(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        report = reconcile(order_list, record_list)
        assert report.metrics.total_results == len(report.results)
        assert 0.0 <= report.metrics.automation_rate <= 1.0


class TestAmountIntegrity:
    """金額一致：variance 必須永遠等於實收減預期。"""

    @given(
        order_list=st.lists(orders(), max_size=15, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=15),
    )
    @SETTINGS
    def test_variance_equals_actual_minus_expected(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        report = reconcile(order_list, record_list)
        for result in report.results:
            if result.order is None or not result.records:
                continue
            actual = money_sum(r.net_amount for r in result.records)
            expected = result.order.expected_net
            if result.variance is not None:
                assert result.variance == actual - expected

    @given(
        order_list=st.lists(orders(), max_size=15, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=15),
    )
    @SETTINGS
    def test_matched_results_have_no_variance(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        """被判定為「已勾稽」的，金額一定分毫不差。"""
        report = reconcile(order_list, record_list)
        for result in report.by_outcome(MatchOutcome.MATCHED):
            assert result.order is not None
            actual = money_sum(r.net_amount for r in result.records)
            assert actual == result.order.expected_net

    @given(
        order_list=st.lists(orders(), max_size=15, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=15),
    )
    @SETTINGS
    def test_total_settled_amount_is_preserved(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        """結果裡所有結算列的撥款總額，必須等於輸入去重後的撥款總額。

        這條性質守住的是「對帳過程不會改動金額」——引擎只做分類，不做算術。
        """
        report = reconcile(order_list, record_list)
        in_results = money_sum(rec.net_amount for r in report.results for rec in r.records)
        seen: set[tuple[object, ...]] = set()
        deduped_total = Money.zero()
        for record in record_list:
            key = (
                record.platform,
                record.external_order_id,
                record.settled_at,
                record.gross_amount,
                record.fee_amount,
                record.net_amount,
                record.record_type,
            )
            if key in seen:
                continue
            seen.add(key)
            deduped_total = deduped_total + record.net_amount
        assert in_results == deduped_total


class TestRobustness:
    """任何輸入都不能讓引擎崩潰或產生不合法的狀態。"""

    @given(
        order_list=st.lists(orders(), max_size=25, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=25),
    )
    @SETTINGS
    def test_never_crashes_and_never_builds_illegal_state(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        """MatchResult 的 __post_init__ 會擋非法狀態。

        引擎能跑完就代表它產生的每一個結論都通過了那些不變量檢查——
        這正是把驗證放在建構子裡的價值：不需要另外寫檢查，
        只要讓引擎跑過隨機輸入，就等於驗證了全部的狀態合法性。
        """
        report = reconcile(order_list, record_list)
        for result in report.results:
            if result.outcome in (MatchOutcome.MATCHED, MatchOutcome.AMOUNT_VARIANCE):
                assert result.order is not None
                assert result.records
                assert result.stage is not None
            elif result.outcome is MatchOutcome.MISSING_IN_LEDGER:
                assert result.order is None
                assert result.records
            else:
                assert result.order is not None
                assert not result.records

    @given(record_list=st.lists(settlement_records(), max_size=20))
    @SETTINGS
    def test_no_orders_means_everything_is_missing_in_ledger(
        self, record_list: list[SettlementRecord]
    ) -> None:
        report = reconcile([], record_list)
        assert all(r.outcome is MatchOutcome.MISSING_IN_LEDGER for r in report.results)

    @given(order_list=st.lists(orders(), max_size=20, unique_by=lambda o: o.order_id))
    @SETTINGS
    def test_no_records_means_everything_is_unpaid(self, order_list: list[Order]) -> None:
        report = reconcile(order_list, [])
        assert len(report.results) == len(order_list)
        assert all(r.outcome is MatchOutcome.MISSING_IN_SETTLEMENT for r in report.results)


class TestDeterminism:
    @given(
        order_list=st.lists(orders(), max_size=15, unique_by=lambda o: o.order_id),
        record_list=st.lists(settlement_records(), max_size=15),
    )
    @SETTINGS
    def test_same_input_gives_same_outcome_counts(
        self, order_list: list[Order], record_list: list[SettlementRecord]
    ) -> None:
        """同樣的輸入必須得到同樣的結果——對帳不能有隨機性。"""
        first = reconcile(order_list, record_list)
        second = reconcile(order_list, record_list)
        assert [r.outcome for r in first.results] == [r.outcome for r in second.results]
        assert [r.stage for r in first.results] == [r.stage for r in second.results]
