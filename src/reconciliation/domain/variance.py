"""金額差異的歸因。

對帳系統真正的價值不在於指出「這筆差了 50 元」——那用 Excel 也做得到。
價值在於**回答「為什麼差」**，因為使用者要據此決定行動：
費率被調整了要去跟平台確認合約，運費補貼是正常的可以放行，
部分退款要去核對出貨紀錄，而「不知道為什麼」的必須人工追查。

歸因的方法是把總差異拆成兩個獨立的分量再看誰動了：

* ``gross_delta`` —— 平台認列的訂單金額 減 我方的訂單金額
* ``fee_delta``   —— 平台實收的手續費   減 我方預期的手續費

因為 ``net = gross - fee``，所以 ``variance = gross_delta - fee_delta``。
這個恆等式讓歸因有數學基礎，而不是憑感覺猜。

設計上最重要的一條：**無法解釋的差異必須落到 UNKNOWN**。
把它硬塞進某個看起來合理的分類，會讓使用者以為系統理解了這筆差異，
於是不去追查——那比完全不歸因還危險。
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from .models import Order, SettlementRecord, SettlementRecordType, VarianceReason
from .money import Money

__all__ = ["ROUNDING_TOLERANCE", "VarianceAnalysis", "analyse_variance"]

#: 小於這個金額的差異視為捨入誤差。
#: 一元是刻意選的：手續費是費率乘售價，各平台捨入到分或到元的規則不同，
#: 累積起來的尾差不會超過一元。再大就不是捨入，是真的有事情發生。
ROUNDING_TOLERANCE: Money = Money.from_str("1")


@dataclass(frozen=True, slots=True)
class VarianceAnalysis:
    """一筆差異的拆解結果。

    三個金額欄位讓使用者看得到「差在哪一段」，``explanation`` 則是
    直接可以顯示在介面上的一句話——不需要前端再去翻譯 enum。
    """

    variance: Money
    gross_delta: Money
    fee_delta: Money
    reason: VarianceReason
    explanation: str

    @property
    def is_significant(self) -> bool:
        """超出捨入容差、需要人看的差異。"""
        return abs(self.variance) > ROUNDING_TOLERANCE


def analyse_variance(order: Order, records: Sequence[SettlementRecord]) -> VarianceAnalysis:
    """把一筆訂單與它對應的結算列拿來比對，並解釋差在哪裡。

    ``records`` 是複數：同一筆訂單可能有一列銷售加一列退款。
    退款列的金額是負數，所以直接加總就是平台的淨認列。
    """
    if not records:
        raise ValueError("analyse_variance 需要至少一列結算記錄")

    currency = order.gross_amount.currency
    actual_gross = _total(r.gross_amount for r in records)
    actual_fee = _total(r.fee_amount for r in records)
    actual_net = _total(r.net_amount for r in records)

    gross_delta = actual_gross - order.gross_amount
    fee_delta = actual_fee - order.expected_fee
    variance = actual_net - order.expected_net

    has_refund = any(r.record_type is SettlementRecordType.REFUND for r in records)
    reason, explanation = _classify(
        variance=variance,
        gross_delta=gross_delta,
        fee_delta=fee_delta,
        has_refund=has_refund,
        currency=currency,
    )
    return VarianceAnalysis(
        variance=variance,
        gross_delta=gross_delta,
        fee_delta=fee_delta,
        reason=reason,
        explanation=explanation,
    )


def _classify(
    *,
    variance: Money,
    gross_delta: Money,
    fee_delta: Money,
    has_refund: bool,
    currency: str,
) -> tuple[VarianceReason, str]:
    """判斷差異的成因。

    規則依序套用，先判斷的優先。順序本身就是設計決策：
    退款排在最前面，因為只要有退款列，金額本來就「應該」對不上我方
    原始訂單，這不是異常。捨入排在其後，因為一元以內的差異不值得
    使用者花時間。剩下的才依 gross 與 fee 誰動了來分類。
    """
    zero = Money.zero(currency)

    if has_refund:
        return (
            VarianceReason.PARTIAL_REFUND,
            f"結算單含退款列，平台淨認列與原始訂單相差 {variance}。請核對退貨紀錄是否相符。",
        )

    if abs(variance) <= ROUNDING_TOLERANCE:
        return (
            VarianceReason.ROUNDING,
            f"差異 {variance} 在一元以內，屬於手續費計算的捨入誤差，可忽略。",
        )

    gross_moved = gross_delta != zero
    fee_moved = fee_delta != zero

    if fee_moved and not gross_moved:
        direction = "多收" if fee_delta.minor_units > 0 else "少收"
        return (
            VarianceReason.FEE_RATE_CHANGE,
            f"訂單金額一致，但平台{direction}手續費 {abs(fee_delta)}。請確認費率方案是否有變動。",
        )

    if gross_moved and not fee_moved:
        direction = "高" if gross_delta.minor_units > 0 else "低"
        return (
            VarianceReason.SHIPPING_SUBSIDY,
            f"手續費一致，但平台認列的訂單金額比我方{direction} {abs(gross_delta)}。"
            f"常見於運費補貼、平台折扣或加購品未入帳。",
        )

    if gross_moved and fee_moved:
        return (
            VarianceReason.UNKNOWN,
            f"訂單金額與手續費同時與我方不符"
            f"（金額差 {gross_delta}、手續費差 {fee_delta}），"
            f"無法歸因，需人工追查。",
        )

    # gross 與 fee 都沒動，net 卻對不上 —— 代表該列自己的加減法不成立。
    # 這幾乎一定是 parser 漏了某個扣項，不是業務上的差異。
    return (
        VarianceReason.UNKNOWN,
        f"訂單金額與手續費都與我方一致，撥款金額卻差 {variance}。"
        f"結算單上可能有本系統尚未建模的扣項（例如廣告費、金流費），"
        f"應檢視該列的原始內容。",
    )


def _total(amounts: object) -> Money:
    """加總一串金額。空序列在呼叫端已被擋掉，這裡不需要處理。"""
    total: Money | None = None
    for amount in amounts:  # type: ignore[attr-defined]
        total = amount if total is None else total + amount
    assert total is not None
    return total
