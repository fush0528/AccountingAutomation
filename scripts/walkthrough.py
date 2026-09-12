"""領域層導覽：把 W1 建好的東西實際跑一遍。

這支腳本存在的目的是「看得見」。目前專案還沒有 CLI 也還沒有 API
（那是 W4、W5 的事），但領域層已經能動了。這裡用一組手寫的假資料，
把每個元件實際跑過，並解釋它在整個對帳流程裡的位置。

執行方式（在專案根目錄）::

    python scripts/walkthrough.py

讀完之後你應該能回答三件事：
1. 為什麼金額不能用 float
2. 為什麼 Stage 1 不能直接用字串相等去比對訂單編號
3. 為什麼領域模型要在建構時就擋掉非法狀態
"""

from __future__ import annotations

import sys
from collections.abc import Callable
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

# 讓腳本不必安裝套件也能跑：把 src/ 加進搜尋路徑
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from reconciliation.domain import (
    DomainError,
    MatchOutcome,
    MatchResult,
    MatchStage,
    Money,
    MoneyError,
    Order,
    SettlementRecord,
    money_sum,
    normalize_order_id,
)


def section(number: int, title: str) -> None:
    print()
    print("=" * 72)
    print(f"  {number}. {title}")
    print("=" * 72)


def note(text: str) -> None:
    print(f"\n  → {text}")


# ======================================================================
def demo_money() -> None:
    section(1, "Money —— 為什麼金額不能用 float")

    print("\n  先看問題。用 float 累加一千筆 0.1 元：")
    total_float = 0.0
    for _ in range(1000):
        total_float += 0.1
    print(f"    預期 100.0，實際得到 {total_float!r}")
    print(f"    total == 100.0 ? {total_float == 100.0}")

    note("差 1.4e-11 看起來很小，但對帳的判斷是「兩邊金額相不相等」。")
    print("     一個連相等性都不可靠的型別，不能拿來當對帳的基礎。")

    print("\n  同樣一千筆，改用 Money：")
    total_money = money_sum(Money.from_str("0.1") for _ in range(1000))
    print(f"    {total_money}")
    print(f"    total == 100 元 ? {total_money == Money.from_str('100')}")
    print(f"    內部儲存的是整數分：{total_money.minor_units}")

    print("\n  結算單上的金額寫法五花八門，from_str 統一處理：")
    for raw in ['"1,234.50"', '"NT$1,234"', '"＄1234"', '"(1,234)"', '"  56.78  "']:
        value = Money.from_str(raw.strip('"'))
        print(f"    {raw:<14} → {value}")
    note("(1,234) 是會計用括號表示負數的寫法，代表 -1234。")

    print("\n  平台手續費是費率乘售價，費率必須是 Decimal 不能是 float：")
    gross = Money.from_str("1000")
    fee = gross * Decimal("0.0553")
    print(f"    售價 {gross} × 費率 5.53% = 手續費 {fee}")
    print(f"    實收 = {gross - fee}")

    print("\n  傳 float 進來會被明確擋下，而且訊息會告訴你該怎麼做：")
    try:
        gross * 0.0553  # type: ignore[operator]
    except MoneyError as exc:
        print(f"    {type(exc).__name__}: {exc}")

    print("\n  分攤：把一筆手續費依售價比例攤回三個品項。")
    print("  天真的做法（各自算完再捨入）會產生尾差，這裡不會：")
    parts = Money.from_str("100").allocate([1, 1, 1])
    print(f"    100 元分三份 → {[str(p) for p in parts]}")
    print(f"    加回來 = {money_sum(parts)}   ← 一分不差")


# ======================================================================
def demo_normalization() -> None:
    section(2, "normalization —— 為什麼不能直接用字串相等")

    print("\n  同一筆訂單，在我方系統與六份不同來源上的寫法：")
    variants = [
        "24A7X9K2",
        "24a7x9k2",
        "SP-24A7X9K2",
        "  sp-24A7X9K2  ",
        "#24A7X9K2",
        "２４Ａ７Ｘ９Ｋ２",
    ]
    for raw in variants:
        print(f"    {raw!r:<22} → {normalize_order_id(raw)!r}")

    keys = {normalize_order_id(v) for v in variants}
    note(f"六種寫法正規化後落在 {len(keys)} 個鍵上：{keys}")
    print("     如果 Stage 1 直接用字串相等，這六筆會被判定成六筆不同的訂單。")

    print("\n  無法作為匹配鍵時回傳 None，不是空字串：")
    unusable: list[str | None] = [None, "", "   ", "---"]
    for blank in unusable:
        print(f"    {blank!r:<10} → {normalize_order_id(blank)!r}")
    note("這一點很關鍵。如果都轉成 ''，所有缺編號的記錄會在 Stage 1")
    print("     互相匹配上，產生一堆假的對帳結果。None 逼呼叫端明確處理。")


# ======================================================================
def demo_models() -> None:
    section(3, "領域模型 —— 讓非法狀態無法被建構出來")

    order = Order(
        order_id="ORD-0001",
        platform="platform_a",
        external_order_id="24A7X9K2",
        ordered_at=datetime(2026, 3, 1, 14, 30),
        product_name="無線滑鼠（黑色）",
        quantity=1,
        gross_amount=Money.from_str("1000"),
        expected_fee=Money.from_str("55.30"),
    )
    print("\n  我方的訂單（Order）——「我們認為發生了什麼」：")
    print(f"    訂單編號   {order.external_order_id}")
    print(f"    商品       {order.product_name}")
    print(f"    售價       {order.gross_amount}")
    print(f"    預期手續費 {order.expected_fee}")
    print(f"    預期實收   {order.expected_net}   ← 這是要拿去跟平台比的數字")

    record = SettlementRecord(
        platform="platform_a",
        settled_at=date(2026, 3, 15),
        gross_amount=Money.from_str("1000"),
        fee_amount=Money.from_str("55.30"),
        net_amount=Money.from_str("944.70"),
        external_order_id="SP-24A7X9K2",
        source_row=42,
        raw={"訂單編號": "SP-24A7X9K2", "撥款金額": "944.70"},
    )
    print("\n  平台的結算列（SettlementRecord）——「平台說發生了什麼」：")
    print(f"    結算單上的編號 {record.external_order_id}   ← 保留原文，沒有清洗")
    print(f"    撥款金額       {record.net_amount}")
    print(f"    residual       {record.residual}   ← gross - fee - net，應為零")

    print("\n  residual 不為零代表結算單上還有我們沒建模的扣項：")
    odd = SettlementRecord(
        platform="platform_a",
        settled_at=date(2026, 3, 15),
        gross_amount=Money.from_str("1000"),
        fee_amount=Money.from_str("55.30"),
        net_amount=Money.from_str("924.70"),
        external_order_id="SP-99999999",
    )
    print(f"    撥款少了 20 元 → residual = {odd.residual}")
    note("把它顯性化，而不是假設它是零。這個值不為零時，parser 就該被檢討。")

    print("\n  非法狀態在建構時就被擋下，下游不需要寫防禦性檢查：")
    bad_cases: list[tuple[str, Callable[[], MatchResult]]] = [
        (
            "MATCHED 卻沒有 order",
            lambda: MatchResult(
                outcome=MatchOutcome.MATCHED,
                stage=MatchStage.EXACT,
                record=record,
            ),
        ),
        (
            "MATCHED 卻沒說是哪一階段匹配的",
            lambda: MatchResult(outcome=MatchOutcome.MATCHED, order=order, record=record),
        ),
        (
            "金額差異的 variance 卻是零",
            lambda: MatchResult(
                outcome=MatchOutcome.AMOUNT_VARIANCE,
                stage=MatchStage.TOLERANT,
                order=order,
                record=record,
                variance=Money.zero(),
            ),
        ),
        (
            "「平台有我方無」卻帶著一筆 order",
            lambda: MatchResult(
                outcome=MatchOutcome.MISSING_IN_LEDGER,
                order=order,
                record=record,
            ),
        ),
    ]
    for label, build in bad_cases:
        try:
            build()
        except DomainError as exc:
            print(f"    ✗ {label}")
            print(f"        {exc}")


# ======================================================================
def demo_matching_preview() -> None:
    section(4, "對帳預覽 —— W3 的引擎要解決的問題長什麼樣")

    print("\n  注意：真正的三階段匹配引擎是 W3 的工作，這裡只是用")
    print("  已經做好的零件手寫一個最陽春的 Stage 1，讓你看見形狀。")

    orders = [
        Order(
            order_id=f"ORD-{i:04d}",
            platform="platform_a",
            external_order_id=ext,
            ordered_at=datetime(2026, 3, 1 + i),
            product_name=name,
            quantity=1,
            gross_amount=Money.from_str(gross),
            expected_fee=Money.from_str(fee),
        )
        for i, (ext, name, gross, fee) in enumerate(
            [
                ("24A7X9K2", "無線滑鼠", "1000", "55.30"),
                ("24B8Y1M3", "機械鍵盤", "2500", "138.25"),
                ("24C9Z2N4", "USB 集線器", "680", "37.60"),
                ("24D1A3P5", "螢幕支架", "1200", "66.36"),
            ],
            start=1,
        )
    ]

    records = [
        # 編號有前綴、金額相符 → Stage 1 應該要抓到
        SettlementRecord(
            platform="platform_a",
            settled_at=date(2026, 3, 15),
            gross_amount=Money.from_str("1000"),
            fee_amount=Money.from_str("55.30"),
            net_amount=Money.from_str("944.70"),
            external_order_id="SP-24a7x9k2",
        ),
        # 金額少了 50（運費補貼？部分退款？）→ Stage 2 的工作
        SettlementRecord(
            platform="platform_a",
            settled_at=date(2026, 3, 15),
            gross_amount=Money.from_str("2500"),
            fee_amount=Money.from_str("138.25"),
            net_amount=Money.from_str("2311.75"),
            external_order_id="SP-24B8Y1M3",
        ),
        # 完全沒有訂單編號 → Stage 3 的工作
        SettlementRecord(
            platform="platform_a",
            settled_at=date(2026, 3, 16),
            gross_amount=Money.from_str("680"),
            fee_amount=Money.from_str("37.60"),
            net_amount=Money.from_str("642.40"),
            external_order_id=None,
        ),
        # 我方根本沒有這筆 → 漏記單
        SettlementRecord(
            platform="platform_a",
            settled_at=date(2026, 3, 17),
            gross_amount=Money.from_str("399"),
            fee_amount=Money.from_str("22.06"),
            net_amount=Money.from_str("376.94"),
            external_order_id="SP-24E5B7Q9",
        ),
    ]

    print(f"\n  我方訂單 {len(orders)} 筆，平台結算列 {len(records)} 筆。開始比對：\n")

    # 建索引：正規化後的編號 → 訂單
    index = {key: o for o in orders if (key := normalize_order_id(o.external_order_id))}

    results: list[MatchResult] = []
    matched_orders: set[str] = set()

    for rec in records:
        key = normalize_order_id(rec.external_order_id)
        order = index.get(key) if key else None

        if order is None:
            reason = "結算列沒有訂單編號" if key is None else "我方查無此訂單編號"
            print(f"    {rec.external_order_id or '(無編號)':<14} {reason}")
            results.append(MatchResult(outcome=MatchOutcome.MISSING_IN_LEDGER, record=rec))
            continue

        matched_orders.add(order.order_id)
        variance = rec.net_amount - order.expected_net

        if variance.is_zero:
            print(f"    {rec.external_order_id:<14} 金額相符，自動勾稽")
            results.append(
                MatchResult(
                    outcome=MatchOutcome.MATCHED,
                    stage=MatchStage.EXACT,
                    order=order,
                    record=rec,
                )
            )
        else:
            print(f"    {rec.external_order_id:<14} 編號相符但金額差 {variance}")
            results.append(
                MatchResult(
                    outcome=MatchOutcome.AMOUNT_VARIANCE,
                    stage=MatchStage.TOLERANT,
                    order=order,
                    record=rec,
                    variance=variance,
                )
            )

    for order in orders:
        if order.order_id not in matched_orders:
            print(f"    {order.external_order_id:<14} 我方有此訂單，平台未撥款")
            results.append(MatchResult(outcome=MatchOutcome.MISSING_IN_SETTLEMENT, order=order))

    print("\n  四象限結果：\n")
    labels = {
        MatchOutcome.MATCHED: "已勾稽",
        MatchOutcome.AMOUNT_VARIANCE: "金額差異",
        MatchOutcome.MISSING_IN_LEDGER: "漏記單（平台有我方無）",
        MatchOutcome.MISSING_IN_SETTLEMENT: "未撥款（我方有平台無）",
    }
    for outcome, label in labels.items():
        count = sum(1 for r in results if r.outcome is outcome)
        print(f"    {label:<24} {count} 筆")

    review = sum(1 for r in results if r.needs_review)
    print(f"\n    需人工確認 {review} / {len(results)} 筆")

    note("看到那筆「沒有訂單編號」被丟進漏記單了嗎？那其實是誤判——")
    print("     它是 USB 集線器那筆，金額完全對得上，只是編號欄是空的。")
    print("     這正是 Stage 3 模糊匹配要救回來的東西：用日期窗口加金額")
    print("     分桶建候選集，就能找回它。這就是 W3 的價值所在。")


# ======================================================================
def main() -> None:
    print()
    print("多平台電商對帳引擎 —— 領域層導覽")
    print("目前完成度：W1（領域層）。W2 解析層、W3 對帳引擎尚未實作。")
    demo_money()
    demo_normalization()
    demo_models()
    demo_matching_preview()
    print()
    print("=" * 72)
    print("  導覽結束。想動手改的話，先從 tests/domain/ 裡挑一個測試改壞，")
    print("  跑 python -m pytest 看它怎麼失敗——那是最快的理解方式。")
    print("=" * 72)
    print()


if __name__ == "__main__":
    main()
