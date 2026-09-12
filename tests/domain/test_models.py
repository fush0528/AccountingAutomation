"""領域模型的測試。

重點放在 ``__post_init__`` 的不變量：這一層的價值不在於能存資料，
而在於**讓非法狀態無法被建構出來**。每個 raises 測試都對應一個
「如果沒擋，下游就得到處寫防禦性 if」的情境。
"""

from __future__ import annotations

from dataclasses import FrozenInstanceError
from datetime import date, datetime
from decimal import Decimal

import pytest

from reconciliation.domain.models import (
    Candidate,
    DomainError,
    ImportBatch,
    MatchOutcome,
    MatchResult,
    MatchStage,
    Order,
    SettlementRecord,
    SettlementRecordType,
)
from reconciliation.domain.money import Money

HASH = "a" * 64


def make_order(**overrides: object) -> Order:
    defaults: dict[str, object] = {
        "order_id": "ORD-1",
        "platform": "platform_a",
        "external_order_id": "24A7X9K2",
        "ordered_at": datetime(2026, 3, 1, 14, 30),
        "product_name": "無線滑鼠",
        "quantity": 1,
        "gross_amount": Money.from_str("1000"),
        "expected_fee": Money.from_str("55.30"),
    }
    defaults.update(overrides)
    return Order(**defaults)  # type: ignore[arg-type]


def make_record(**overrides: object) -> SettlementRecord:
    defaults: dict[str, object] = {
        "platform": "platform_a",
        "settled_at": date(2026, 3, 15),
        "gross_amount": Money.from_str("1000"),
        "fee_amount": Money.from_str("55.30"),
        "net_amount": Money.from_str("944.70"),
        "external_order_id": "SP-24A7X9K2",
    }
    defaults.update(overrides)
    return SettlementRecord(**defaults)  # type: ignore[arg-type]


class TestOrder:
    def test_expected_net_is_gross_minus_fee(self) -> None:
        assert make_order().expected_net == Money.from_str("944.70")

    def test_quantity_must_be_positive(self) -> None:
        with pytest.raises(DomainError, match="quantity"):
            make_order(quantity=0)

    def test_blank_identifiers_rejected(self) -> None:
        with pytest.raises(DomainError, match="order_id"):
            make_order(order_id="   ")

    def test_mixed_currency_rejected(self) -> None:
        with pytest.raises(DomainError, match="幣別"):
            make_order(expected_fee=Money(100, "USD"))

    def test_negative_gross_rejected(self) -> None:
        """退貨要用結算單的 REFUND 列表示，不是把訂單金額改成負數。"""
        with pytest.raises(DomainError, match="gross_amount"):
            make_order(gross_amount=Money.from_str("-100"))

    def test_is_frozen(self) -> None:
        with pytest.raises(FrozenInstanceError):
            make_order().quantity = 5  # type: ignore[misc]


class TestSettlementRecord:
    def test_residual_is_zero_for_consistent_row(self) -> None:
        assert make_record().residual.is_zero

    def test_residual_surfaces_unmodelled_deductions(self) -> None:
        """淨額少了 20 元，residual 就會是 20——不會被默默吸收掉。"""
        record = make_record(net_amount=Money.from_str("924.70"))
        assert record.residual == Money.from_str("20")

    def test_missing_order_id_is_allowed(self) -> None:
        record = make_record(external_order_id=None)
        assert record.external_order_id is None
        assert not record.has_order_id

    def test_blank_order_id_collapses_to_none(self) -> None:
        """空字串與 None 對下游是同一件事，在邊界就收斂。"""
        assert make_record(external_order_id="   ").external_order_id is None

    def test_raw_is_read_only(self) -> None:
        record = make_record(raw={"訂單編號": "SP-1"})
        assert record.raw["訂單編號"] == "SP-1"
        with pytest.raises(TypeError):
            record.raw["訂單編號"] = "tampered"  # type: ignore[index]

    def test_mixed_currency_rejected(self) -> None:
        with pytest.raises(DomainError, match="幣別"):
            make_record(fee_amount=Money(100, "USD"))

    def test_refund_row(self) -> None:
        record = make_record(
            record_type=SettlementRecordType.REFUND,
            gross_amount=Money.from_str("-1000"),
            fee_amount=Money.from_str("-55.30"),
            net_amount=Money.from_str("-944.70"),
        )
        assert record.record_type is SettlementRecordType.REFUND
        assert record.residual.is_zero


class TestMatchResult:
    def test_matched_requires_both_sides(self) -> None:
        result = MatchResult(
            outcome=MatchOutcome.MATCHED,
            stage=MatchStage.EXACT,
            order=make_order(),
            records=(make_record(),),
        )
        assert not result.needs_review
        assert result.record is result.records[0]

    def test_one_order_can_have_several_records(self) -> None:
        """同一筆訂單可能有一列銷售加一列退款。"""
        sale = make_record()
        refund = make_record(
            record_type=SettlementRecordType.REFUND,
            gross_amount=Money.from_str("-1000"),
            fee_amount=Money.from_str("-55.30"),
            net_amount=Money.from_str("-944.70"),
        )
        result = MatchResult(
            outcome=MatchOutcome.MATCHED,
            stage=MatchStage.EXACT,
            order=make_order(),
            records=(sale, refund),
        )
        assert len(result.records) == 2
        assert result.record is sale  # record 是便利存取，取第一列
        assert result.settled_amount == Money.zero()

    def test_matched_without_order_is_unconstructable(self) -> None:
        with pytest.raises(DomainError, match="必須同時具備"):
            MatchResult(
                outcome=MatchOutcome.MATCHED,
                stage=MatchStage.EXACT,
                records=(make_record(),),
            )

    def test_matched_without_stage_is_unconstructable(self) -> None:
        with pytest.raises(DomainError, match="哪一階段"):
            MatchResult(
                outcome=MatchOutcome.MATCHED,
                order=make_order(),
                records=(make_record(),),
            )

    def test_variance_must_be_non_zero(self) -> None:
        with pytest.raises(DomainError, match="variance"):
            MatchResult(
                outcome=MatchOutcome.AMOUNT_VARIANCE,
                stage=MatchStage.TOLERANT,
                order=make_order(),
                records=(make_record(),),
                variance=Money.zero(),
            )

    def test_missing_in_ledger_must_not_carry_an_order(self) -> None:
        with pytest.raises(DomainError, match="missing_in_ledger"):
            MatchResult(
                outcome=MatchOutcome.MISSING_IN_LEDGER,
                order=make_order(),
                records=(make_record(),),
            )

    def test_missing_in_settlement_is_valid_with_order_only(self) -> None:
        result = MatchResult(
            outcome=MatchOutcome.MISSING_IN_SETTLEMENT,
            order=make_order(),
        )
        assert result.needs_review

    def test_fuzzy_match_always_needs_review(self) -> None:
        result = MatchResult(
            outcome=MatchOutcome.MATCHED,
            stage=MatchStage.FUZZY,
            order=make_order(),
            records=(make_record(),),
        )
        assert result.needs_review

    def test_candidates_carry_human_readable_reasons(self) -> None:
        candidate = Candidate(
            order=make_order(),
            score=Decimal("0.87"),
            reasons=("金額完全相符", "日期相差 2 天", "商品名稱相似度 0.91"),
        )
        assert len(candidate.reasons) == 3

    @pytest.mark.parametrize("score", ["-0.1", "1.1"])
    def test_candidate_score_must_be_a_probability(self, score: str) -> None:
        with pytest.raises(DomainError, match="score"):
            Candidate(order=make_order(), score=Decimal(score))


class TestImportBatch:
    def test_valid_batch(self) -> None:
        batch = ImportBatch(
            batch_id="B1",
            platform="platform_a",
            filename="2026-03.csv",
            content_hash=HASH,
            imported_at=datetime(2026, 4, 1),
            row_count=100,
            accepted_count=98,
            rejected_count=2,
        )
        assert batch.row_count == 100

    @pytest.mark.parametrize("bad_hash", ["", "abc", "A" * 64, "z" * 64])
    def test_content_hash_must_be_lowercase_sha256(self, bad_hash: str) -> None:
        with pytest.raises(DomainError, match="content_hash"):
            ImportBatch(
                batch_id="B1",
                platform="p",
                filename="f.csv",
                content_hash=bad_hash,
                imported_at=datetime(2026, 4, 1),
            )

    def test_counts_must_not_exceed_row_count(self) -> None:
        with pytest.raises(DomainError, match="row_count"):
            ImportBatch(
                batch_id="B1",
                platform="p",
                filename="f.csv",
                content_hash=HASH,
                imported_at=datetime(2026, 4, 1),
                row_count=10,
                accepted_count=9,
                rejected_count=5,
            )
