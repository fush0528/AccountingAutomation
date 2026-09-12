"""金額值物件。

設計決策（對應 docs/adr/0001-money-as-integer-minor-units.md）
------------------------------------------------------------
本專案第一版（見 ``legacy/``）以 ``float`` 儲存金額，在對帳總額累加時
會出現無法解釋的尾差。原因是二進位浮點數無法精確表示 0.1 這類十進位小數::

    >>> 0.1 + 0.2 == 0.3
    False

金額是離散量，不是連續量。因此本模組將金額表示為**最小貨幣單位的整數**
（TWD 為「分」，即 1 元 = 100），所有算術都在整數上進行，只有在需要
顯示或除法時才涉及十進位捨入，且捨入策略明確標示為 ROUND_HALF_UP
（財務慣例，非 Python 預設的 banker's rounding）。

``Money`` 是不可變的值物件：相等性由值決定、沒有身分、任何運算都回傳
新的實例。它不依賴任何框架，可以在沒有資料庫的情況下完整測試。
"""

from __future__ import annotations

import unicodedata
from collections.abc import Iterable
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Final

__all__ = ["CurrencyMismatchError", "Money", "MoneyError"]


class MoneyError(ValueError):
    """金額操作違反不變量時拋出。"""


class CurrencyMismatchError(MoneyError):
    """對兩個不同幣別的金額做算術時拋出。"""

    def __init__(self, left: str, right: str) -> None:
        super().__init__(
            f"不能對不同幣別做算術運算：{left} 與 {right}。請先明確換匯，本型別不會替你猜匯率。"
        )
        self.left = left
        self.right = right


# 各幣別的小數位數。未列出的幣別預設 2 位。
# TWD 在日常交易雖以元為單位，但平台手續費（費率 × 售價）會產生小數，
# 因此仍保留 2 位小數的精度，避免在中間計算階段就提早捨入。
_MINOR_DIGITS: Final[dict[str, int]] = {
    "TWD": 2,
    "USD": 2,
    "JPY": 0,
    "KRW": 0,
}
_DEFAULT_MINOR_DIGITS: Final[int] = 2

Numeric = int | Decimal


def _minor_digits(currency: str) -> int:
    return _MINOR_DIGITS.get(currency, _DEFAULT_MINOR_DIGITS)


@dataclass(frozen=True, slots=True, order=False)
class Money:
    """以最小貨幣單位（整數）表示的金額。

    直接建構時傳入的是**最小單位**，不是元::

        Money(12345)              # 123.45 元
        Money.from_str("123.45")  # 同上，日常使用請用這個

    刻意不提供 ``from_float``：把浮點數轉成金額這件事沒有安全的預設行為，
    呼叫端應該先決定要如何捨入，再透過 :meth:`from_decimal` 傳入。
    """

    minor_units: int
    currency: str = "TWD"

    def __post_init__(self) -> None:
        if isinstance(self.minor_units, bool) or not isinstance(self.minor_units, int):
            raise MoneyError(
                f"minor_units 必須是 int，收到 {type(self.minor_units).__name__}。"
                f"若你手上是小數金額，請改用 Money.from_str() 或 Money.from_decimal()。"
            )
        if not self.currency or not self.currency.isalpha() or len(self.currency) != 3:
            raise MoneyError(f"currency 必須是三個字母的 ISO 4217 代碼，收到 {self.currency!r}")
        object.__setattr__(self, "currency", self.currency.upper())

    # ------------------------------------------------------------------
    # 建構
    # ------------------------------------------------------------------
    @classmethod
    def zero(cls, currency: str = "TWD") -> Money:
        return cls(0, currency)

    @classmethod
    def from_str(cls, value: str, currency: str = "TWD") -> Money:
        """從字串建構，容忍千分位逗號、全形空白與貨幣符號。

        平台結算單常見 ``"1,234.50"``、``"NT$1,234"``、``" 1234.5 "`` 等寫法，
        統一在這裡處理，parser 就不必各自重複實作。
        """
        if not isinstance(value, str):
            raise MoneyError(f"from_str 需要字串，收到 {type(value).__name__}")
        # 先做 NFKC：全形數字、全形貨幣符號、全形逗號與空白一次收斂成半形，
        # 比逐個 replace 可靠得多——結算單的全形寫法遠比想像的多。
        cleaned = unicodedata.normalize("NFKC", value)
        for token in ("NT$", "TWD", "US$", "USD", "$", "¥", "￥", ",", " ", "　"):
            cleaned = cleaned.replace(token, "")
        cleaned = cleaned.strip()
        if cleaned in ("", "-", "--"):
            raise MoneyError(f"無法從空白或佔位字串解析金額：{value!r}")
        # 會計常見以括號表示負數：(1,234) 即 -1234
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = "-" + cleaned[1:-1]
        try:
            return cls.from_decimal(Decimal(cleaned), currency)
        except InvalidOperation as exc:
            raise MoneyError(f"無法解析金額字串：{value!r}") from exc

    @classmethod
    def from_decimal(cls, value: Decimal, currency: str = "TWD") -> Money:
        """從 :class:`~decimal.Decimal` 建構，以 ROUND_HALF_UP 捨入到最小單位。"""
        if not isinstance(value, Decimal):
            raise MoneyError(
                f"from_decimal 需要 Decimal，收到 {type(value).__name__}。"
                f"若來源是 float，請先自行決定捨入方式，例如 "
                f"Decimal(str(x))，並理解 float 本身已經帶有誤差。"
            )
        if not value.is_finite():
            raise MoneyError(f"金額必須是有限數值，收到 {value}")
        scale = Decimal(10) ** _minor_digits(currency)
        minor = (value * scale).quantize(Decimal(1), rounding=ROUND_HALF_UP)
        return cls(int(minor), currency)

    # ------------------------------------------------------------------
    # 轉換
    # ------------------------------------------------------------------
    def to_decimal(self) -> Decimal:
        """回傳以「元」為單位的 Decimal，位數與幣別一致。"""
        digits = _minor_digits(self.currency)
        return (Decimal(self.minor_units) / (Decimal(10) ** digits)).quantize(
            Decimal(1).scaleb(-digits)
        )

    def __str__(self) -> str:
        return f"{self.to_decimal()} {self.currency}"

    def __repr__(self) -> str:
        return f"Money.from_str({str(self.to_decimal())!r}, {self.currency!r})"

    def __format__(self, spec: str) -> str:
        """支援 f-string 的格式化。

        報表要對齊金額欄位，``f"{amount:>14}"`` 必須能用；偶爾也會需要
        只取數值而不要幣別，``f"{amount:.2f}"`` 也該work。沒有這個方法的話
        兩者都會拋 TypeError，而那個錯誤訊息完全看不出問題在哪。
        """
        if not spec:
            return str(self)
        if spec[-1] in "eEfFgGn%":
            return format(self.to_decimal(), spec)
        return format(str(self), spec)

    # ------------------------------------------------------------------
    # 算術
    # ------------------------------------------------------------------
    def _check(self, other: Money) -> None:
        if self.currency != other.currency:
            raise CurrencyMismatchError(self.currency, other.currency)

    def __add__(self, other: Money) -> Money:
        if not isinstance(other, Money):
            return NotImplemented
        self._check(other)
        return Money(self.minor_units + other.minor_units, self.currency)

    def __sub__(self, other: Money) -> Money:
        if not isinstance(other, Money):
            return NotImplemented
        self._check(other)
        return Money(self.minor_units - other.minor_units, self.currency)

    def __neg__(self) -> Money:
        return Money(-self.minor_units, self.currency)

    def __abs__(self) -> Money:
        return Money(abs(self.minor_units), self.currency)

    def __mul__(self, factor: Numeric) -> Money:
        """乘上倍數或費率。float 會被明確拒絕。"""
        if isinstance(factor, bool | float):
            raise MoneyError(
                "不接受 float 乘數。費率請用 Decimal，例如 "
                'Money.from_str("100") * Decimal("0.0553")。'
            )
        if isinstance(factor, int):
            return Money(self.minor_units * factor, self.currency)
        if isinstance(factor, Decimal):
            product = (Decimal(self.minor_units) * factor).quantize(
                Decimal(1), rounding=ROUND_HALF_UP
            )
            return Money(int(product), self.currency)
        return NotImplemented

    __rmul__ = __mul__

    # ------------------------------------------------------------------
    # 比較
    # ------------------------------------------------------------------
    def __lt__(self, other: Money) -> bool:
        if not isinstance(other, Money):
            return NotImplemented
        self._check(other)
        return self.minor_units < other.minor_units

    def __le__(self, other: Money) -> bool:
        if not isinstance(other, Money):
            return NotImplemented
        self._check(other)
        return self.minor_units <= other.minor_units

    def __gt__(self, other: Money) -> bool:
        if not isinstance(other, Money):
            return NotImplemented
        self._check(other)
        return self.minor_units > other.minor_units

    def __ge__(self, other: Money) -> bool:
        if not isinstance(other, Money):
            return NotImplemented
        self._check(other)
        return self.minor_units >= other.minor_units

    def __bool__(self) -> bool:
        return self.minor_units != 0

    @property
    def is_zero(self) -> bool:
        return self.minor_units == 0

    @property
    def is_negative(self) -> bool:
        return self.minor_units < 0

    # ------------------------------------------------------------------
    # 分攤
    # ------------------------------------------------------------------
    def allocate(self, weights: Iterable[int]) -> list[Money]:
        """依權重把金額分攤成數份，保證**分攤後總和等於原金額**。

        對帳時常要把一筆平台手續費依售價比例攤回各個品項。單純用比例相乘
        再各自捨入會產生尾差（例如 100 元分三份會變成 33.33 × 3 = 99.99）。
        這裡採 largest remainder method：先取整數商，再把餘下的最小單位
        依餘數大小逐一分配出去。

        >>> [str(m) for m in Money.from_str("100").allocate([1, 1, 1])]
        ['33.34 TWD', '33.33 TWD', '33.33 TWD']
        """
        weight_list = list(weights)
        if not weight_list:
            raise MoneyError("allocate 需要至少一個權重")
        if any(w < 0 for w in weight_list):
            raise MoneyError("權重不可為負")
        total_weight = sum(weight_list)
        if total_weight == 0:
            raise MoneyError("權重總和不可為零")

        base = [self.minor_units * w // total_weight for w in weight_list]
        remainder = self.minor_units - sum(base)
        # 餘數大者優先取得多出來的最小單位；同餘數時以原順序為準（穩定排序）
        order = sorted(
            range(len(weight_list)),
            key=lambda i: (-(self.minor_units * weight_list[i] % total_weight), i),
        )
        for i in range(abs(remainder)):
            idx = order[i % len(order)]
            base[idx] += 1 if remainder > 0 else -1
        return [Money(units, self.currency) for units in base]


def money_sum(items: Iterable[Money], currency: str = "TWD") -> Money:
    """加總一串金額；空序列回傳該幣別的零，而不是 int 0。"""
    total = Money.zero(currency)
    for item in items:
        total = total + item
    return total
