"""Money 值物件的測試。

第一組測試（TestFloatIsWrong）不是在測我們的程式碼，而是把**問題本身**
釘在測試檔裡：它證明 v1 的 float 實作為什麼一定會出錯。這種測試在
重構專案裡很有價值——它讓「為什麼要改」變成可執行的文件，而不是
README 裡的一句宣稱。
"""

from __future__ import annotations

from dataclasses import FrozenInstanceError
from decimal import Decimal

import pytest

from reconciliation.domain.money import (
    CurrencyMismatchError,
    Money,
    MoneyError,
    money_sum,
)


class TestFloatIsWrong:
    """記錄 v1 的缺陷：為什麼金額不能用 float。"""

    def test_float_addition_is_not_exact(self) -> None:
        assert 0.1 + 0.2 != 0.3

    def test_float_accumulation_drifts(self) -> None:
        """一千筆 0.1 元用 float 累加，結果不等於 100。"""
        total = 0.0
        for _ in range(1000):
            total += 0.1
        assert total != 100.0

    def test_money_accumulation_is_exact(self) -> None:
        """同樣一千筆，Money 精確等於 100 元。"""
        total = money_sum(Money.from_str("0.1") for _ in range(1000))
        assert total == Money.from_str("100")
        assert total.minor_units == 10_000


class TestConstruction:
    def test_direct_construction_takes_minor_units(self) -> None:
        assert Money(12345).to_decimal() == Decimal("123.45")

    def test_from_str_handles_plain_decimal(self) -> None:
        assert Money.from_str("123.45") == Money(12345)

    @pytest.mark.parametrize(
        ("text", "expected_minor"),
        [
            ("1,234.50", 123450),
            ("NT$1,234", 123400),
            ("  1234.5  ", 123450),
            ("＄1234", 123400),
            ("(1,234)", -123400),
            ("-56.78", -5678),
            ("0", 0),
        ],
    )
    def test_from_str_tolerates_settlement_sheet_formats(
        self, text: str, expected_minor: int
    ) -> None:
        """結算單上的金額寫法五花八門，統一在 from_str 處理。"""
        assert Money.from_str(text).minor_units == expected_minor

    def test_from_str_rounds_half_up_not_bankers(self) -> None:
        """財務慣例是四捨五入，不是 Python 預設的銀行家捨入。"""
        assert Money.from_str("0.125").minor_units == 13  # banker's 會給 12
        assert Money.from_str("0.135").minor_units == 14

    def test_float_is_rejected_with_a_useful_message(self) -> None:
        with pytest.raises(MoneyError, match="minor_units 必須是 int"):
            Money(123.45)  # type: ignore[arg-type]

    def test_from_decimal_rejects_float(self) -> None:
        with pytest.raises(MoneyError, match="需要 Decimal"):
            Money.from_decimal(1.5)  # type: ignore[arg-type]

    def test_bool_is_not_an_acceptable_minor_units(self) -> None:
        with pytest.raises(MoneyError):
            Money(True)  # type: ignore[arg-type]

    @pytest.mark.parametrize("text", ["", "   ", "-", "abc", "1.2.3"])
    def test_unparseable_strings_raise(self, text: str) -> None:
        with pytest.raises(MoneyError):
            Money.from_str(text)

    def test_currency_must_be_iso_code(self) -> None:
        with pytest.raises(MoneyError):
            Money(100, "NTD$")

    def test_currency_is_upper_cased(self) -> None:
        assert Money(100, "twd").currency == "TWD"

    def test_zero_decimal_currency(self) -> None:
        """JPY 沒有小數位，最小單位就是 1 圓。"""
        assert Money.from_str("1234", "JPY").minor_units == 1234
        assert str(Money.from_str("1234", "JPY")) == "1234 JPY"


class TestArithmetic:
    def test_add_and_sub(self) -> None:
        assert Money.from_str("100") + Money.from_str("23.45") == Money.from_str("123.45")
        assert Money.from_str("100") - Money.from_str("23.45") == Money.from_str("76.55")

    def test_negation_and_abs(self) -> None:
        assert -Money.from_str("100") == Money.from_str("-100")
        assert abs(Money.from_str("-100")) == Money.from_str("100")

    def test_multiply_by_int(self) -> None:
        assert Money.from_str("19.99") * 3 == Money.from_str("59.97")

    def test_multiply_by_decimal_rate(self) -> None:
        """平台費率是 Decimal，不是 float。"""
        gross = Money.from_str("1000")
        fee = gross * Decimal("0.0553")
        assert fee == Money.from_str("55.30")

    def test_multiply_by_float_is_rejected(self) -> None:
        with pytest.raises(MoneyError, match="不接受 float 乘數"):
            Money.from_str("100") * 0.05  # type: ignore[operator]

    def test_currency_mismatch_raises(self) -> None:
        with pytest.raises(CurrencyMismatchError):
            Money(100, "TWD") + Money(100, "USD")

    def test_comparison(self) -> None:
        assert Money.from_str("1") < Money.from_str("2")
        assert Money.from_str("2") >= Money.from_str("2")
        assert max(Money.from_str("1"), Money.from_str("3")) == Money.from_str("3")

    def test_comparing_different_currencies_raises(self) -> None:
        with pytest.raises(CurrencyMismatchError):
            _ = Money(1, "TWD") < Money(1, "USD")

    def test_truthiness_follows_zero(self) -> None:
        assert not Money.zero()
        assert Money.from_str("0.01")
        assert Money.zero().is_zero

    def test_money_sum_of_empty_returns_typed_zero(self) -> None:
        """空序列要回傳 Money.zero()，不是 int 0——否則下游型別會崩。"""
        total = money_sum([], "USD")
        assert total == Money.zero("USD")
        assert isinstance(total, Money)


class TestImmutability:
    def test_is_frozen(self) -> None:
        m = Money.from_str("100")
        with pytest.raises(FrozenInstanceError):
            m.minor_units = 200  # type: ignore[misc]

    def test_is_hashable_and_usable_as_dict_key(self) -> None:
        buckets = {Money.from_str("100"): "a", Money.from_str("200"): "b"}
        assert buckets[Money(10000)] == "a"


class TestAllocate:
    def test_allocation_preserves_total(self) -> None:
        """100 元分三份，總和必須還是 100——不能變成 99.99。"""
        parts = Money.from_str("100").allocate([1, 1, 1])
        assert money_sum(parts) == Money.from_str("100")
        assert [p.minor_units for p in parts] == [3334, 3333, 3333]

    def test_allocation_by_weight(self) -> None:
        """手續費依售價比例攤回品項。"""
        parts = Money.from_str("55.30").allocate([700, 300])
        assert money_sum(parts) == Money.from_str("55.30")
        assert parts[0] > parts[1]

    def test_allocation_of_negative_amount(self) -> None:
        """退款分攤同樣不能有尾差。"""
        parts = Money.from_str("-100").allocate([1, 1, 1])
        assert money_sum(parts) == Money.from_str("-100")

    @pytest.mark.parametrize(
        ("amount", "weights"),
        [
            ("0.01", [1, 1, 1]),
            ("0.02", [1, 1, 1]),
            ("99999.99", [3, 5, 7, 11]),
            ("1", [1]),
        ],
    )
    def test_allocation_always_preserves_total(self, amount: str, weights: list[int]) -> None:
        original = Money.from_str(amount)
        assert money_sum(original.allocate(weights)) == original

    def test_invalid_weights(self) -> None:
        with pytest.raises(MoneyError):
            Money.from_str("100").allocate([])
        with pytest.raises(MoneyError):
            Money.from_str("100").allocate([0, 0])
        with pytest.raises(MoneyError):
            Money.from_str("100").allocate([1, -1])


class TestRepresentation:
    def test_str_includes_currency(self) -> None:
        assert str(Money.from_str("1234.5")) == "1234.50 TWD"

    def test_format_supports_alignment(self) -> None:
        """報表要對齊金額欄位。"""
        assert f"{Money.from_str('12.5'):>12}" == "   12.50 TWD"
        assert f"{Money.from_str('12.5'):<12}" == "12.50 TWD   "

    def test_format_supports_numeric_spec(self) -> None:
        """有時只要數值，不要幣別。"""
        assert f"{Money.from_str('1234.5'):.2f}" == "1234.50"

    def test_repr_round_trips(self) -> None:
        m = Money.from_str("1234.56")
        assert eval(repr(m)) == m
