"""對帳引擎實測：把三個平台的結算單跟我方訂單對一遍。

這是整個專案目前為止最完整的一條路徑：

    合成資料 → 三個 parser → 統一的領域模型 → 三階段匹配 → 四象限報告

執行（先跑 gen_fixtures.py）::

    python scripts/gen_fixtures.py
    python scripts/reconcile_demo.py

看更多細節::

    python scripts/reconcile_demo.py --show 8        # 每個象限印 8 筆
    python scripts/reconcile_demo.py --threshold 0.7 # 調鬆 Stage 3 門檻
"""

from __future__ import annotations

import argparse
from decimal import Decimal

from _data import load_orders, load_settlements, require_data

from reconciliation.domain import (
    MatchingConfig,
    MatchOutcome,
    MatchStage,
    ReconciliationReport,
    money_sum,
    reconcile,
)

QUADRANT_LABELS = {
    MatchOutcome.MATCHED: "已勾稽",
    MatchOutcome.AMOUNT_VARIANCE: "金額差異",
    MatchOutcome.MISSING_IN_LEDGER: "漏記單（平台有、我方無）",
    MatchOutcome.MISSING_IN_SETTLEMENT: "未撥款（我方有、平台無）",
}

STAGE_LABELS = {
    MatchStage.EXACT: "Stage 1 精確",
    MatchStage.TOLERANT: "Stage 2 容差",
    MatchStage.FUZZY: "Stage 3 模糊",
}


def rule(char: str = "─", width: int = 74) -> None:
    print(char * width)


def show_metrics(report: ReconciliationReport) -> None:
    m = report.metrics
    rule("═")
    print("  對帳結果")
    rule("═")
    print(f"\n  輸入：我方訂單 {m.order_count} 筆、平台結算列 {m.record_count} 筆")
    if m.duplicates_dropped:
        print(f"        （另有 {m.duplicates_dropped} 列是平台重複匯出，已去重）")

    print("\n  四象限：\n")
    for outcome, label in QUADRANT_LABELS.items():
        count = len(report.by_outcome(outcome))
        bar = "█" * round(count / max(m.total_results, 1) * 40)
        print(f"    {label:<24} {count:>4} 筆  {bar}")

    print("\n  各階段：\n")
    for stage, label in STAGE_LABELS.items():
        count = sum(1 for r in report.results if r.stage is stage)
        print(f"    {label:<16} {count:>4} 筆")

    print(f"\n  自動化率 {m.automation_rate:.1%}　需人工確認 {m.review_rate:.1%}")
    print(f"  處理耗時 {m.elapsed_seconds * 1000:.1f} ms")
    print(f"  Stage 3 實際候選比對次數 {m.candidate_comparisons:,} 次")
    if m.order_count and m.record_count:
        naive = m.order_count * m.record_count
        print(f"    （若不建索引、每筆都比過所有訂單，會是 {naive:,} 次）")

    print("\n  自動化率只計 Stage 1——那是唯一可以直接入帳、不需要人看的情況。")
    print("  把 Stage 2、Stage 3 算進去數字會好看很多，但那是自我欺騙。")


def show_variances(report: ReconciliationReport, show: int) -> None:
    variances = report.by_outcome(MatchOutcome.AMOUNT_VARIANCE)
    if not variances:
        return
    print()
    rule()
    print(f"金額差異明細（共 {len(variances)} 筆，顯示前 {min(show, len(variances))} 筆）")
    print("\n  差異不只被指出來，還被歸因——使用者要據此決定行動。\n")

    by_reason: dict[str, int] = {}
    for result in variances:
        reason = result.variance_reason.value if result.variance_reason else "unknown"
        by_reason[reason] = by_reason.get(reason, 0) + 1

    for reason, count in sorted(by_reason.items(), key=lambda kv: -kv[1]):
        print(f"    {reason:<20} {count:>4} 筆")

    print()
    for result in variances[:show]:
        assert result.order is not None
        print(
            f"    {result.order.external_order_id:<14} "
            f"預期 {result.order.expected_net:>14}  "
            f"實收 {result.settled_amount!s:>14}  "
            f"差 {result.variance!s:>12}"
        )
        if result.variance_reason:
            print(f"        → {result.variance_reason.value}")

    total_variance = money_sum(r.variance for r in variances if r.variance is not None)
    print(f"\n    差異合計 {total_variance}")


def show_candidates(report: ReconciliationReport, show: int) -> None:
    with_candidates = [r for r in report.by_outcome(MatchOutcome.MISSING_IN_LEDGER) if r.candidates]
    if not with_candidates:
        return
    print()
    rule()
    print(f"人工確認佇列（{len(with_candidates)} 筆有候選）")
    print("\n  Stage 3 分數不夠高就不猜，但會把候選與理由列出來讓人判斷。")
    print("  只給一個 0.87 的分數，沒有人敢按下確認。\n")

    for result in with_candidates[:show]:
        record = result.record
        assert record is not None
        oid = record.external_order_id or "(無編號)"
        print(f"    結算列 第{record.source_row}列 {oid}  撥款 {record.net_amount}")
        for candidate in result.candidates:
            print(
                f"        候選 {candidate.order.external_order_id}  "
                f"分數 {candidate.score:.2f}  "
                f"（{'、'.join(candidate.reasons)}）"
            )
        print()


def show_fuzzy_wins(report: ReconciliationReport, show: int) -> None:
    wins = [r for r in report.results if r.stage is MatchStage.FUZZY and r.order]
    if not wins:
        return
    print()
    rule()
    print(f"Stage 3 救回來的（{len(wins)} 筆）")
    print("\n  這些結算列的訂單編號欄是空的。Stage 1 完全找不到，")
    print("  靠日期窗口與金額分桶建候選集、加權評分才找回對應的訂單。\n")
    for result in wins[:show]:
        assert result.order is not None
        record = result.record
        assert record is not None
        print(
            f"    第{record.source_row:>4}列 撥款 {record.net_amount:<14} "
            f"→ {result.order.order_id} {result.order.external_order_id} "
            f"（{result.order.product_name}）"
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="對帳引擎實測")
    parser.add_argument("--show", type=int, default=5, help="每個區塊顯示幾筆明細")
    parser.add_argument("--threshold", type=str, default=None, help="Stage 3 採納門檻，例如 0.7")
    args = parser.parse_args()

    if not require_data():
        return 1

    print("\n載入資料⋯")
    orders = load_orders()
    records = load_settlements()
    print(f"  我方訂單 {len(orders)} 筆")
    print(f"  平台結算列 {len(records)} 筆（三個平台、三種格式）\n")

    config = MatchingConfig(
        accept_threshold=Decimal(args.threshold)
        if args.threshold
        else MatchingConfig().accept_threshold
    )
    report = reconcile(orders, records, config)

    show_metrics(report)
    show_variances(report, args.show)
    show_fuzzy_wins(report, args.show)
    show_candidates(report, args.show)

    print()
    rule("═")
    print(
        f"  設定：Stage 3 門檻 {report.config.accept_threshold}、"
        f"金額容差 {report.config.amount_tolerance}"
    )
    print("  換個門檻結果就不同，所以報告一定要附上設定——")
    print("  「自動匹配率 92%」沒有附設定的話，是一個沒有意義的數字。")
    print()
    print("  接著跑 python scripts/benchmark.py，")
    print("  它會拿 manifest.json 的 ground truth 驗證上面這些分類是不是真的對。")
    rule("═")
    print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
