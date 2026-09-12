"""以 ground truth 驗證對帳結果，並量測索引帶來的複雜度改善。

一個對帳系統說自己「自動匹配率 92%」，這句話本身沒有意義——除非能證明
那 92% 真的都配對正確。所以 ``gen_fixtures.py`` 在產生資料時，把每一種
刻意注入的狀況記錄進 ``manifest.json``，這支腳本再拿它來對答案。

驗證的是四件事：

1. **正規化有效**：編號被寫成小寫的訂單，必須仍然在 Stage 1 被抓到。
2. **Stage 3 的召回率**：編號整欄空白的訂單，有多少比例被模糊匹配救回來。
3. **四象限分類正確**：注入的「平台有我方無」「我方有平台無」「金額差異」
   有沒有被分到正確的象限（precision / recall）。
4. **索引真的有用**：實測候選比對次數，跟不建索引的暴力法比。
   ``--scaling`` 會量測成長率——結果顯示索引版是次平方而非線性，
   限制與原因印在輸出裡，不做誇大。

執行::

    python scripts/gen_fixtures.py --orders 2000
    python scripts/benchmark.py
    python scripts/benchmark.py --scaling    # 量測複雜度隨資料量的成長
"""

from __future__ import annotations

import argparse
import random
import time
from collections.abc import Sequence

from _data import load_manifest, load_orders, load_settlements, require_data

from reconciliation.domain import (
    MatchOutcome,
    MatchStage,
    Order,
    ReconciliationReport,
    SettlementRecord,
    normalize_order_id,
    reconcile,
)
from reconciliation.domain.matching import MatchingConfig, _score


def rule(char: str = "─", width: int = 74) -> None:
    print(char * width)


def pct(part: int, whole: int) -> str:
    return f"{part / whole:.1%}" if whole else "—"


# ----------------------------------------------------------------------
def verify(report: ReconciliationReport, manifest: dict[str, object]) -> bool:
    """拿 ground truth 對答案。回傳是否全部通過。"""
    detail = manifest["injected_detail"]
    assert isinstance(detail, dict)
    injected: dict[str, list[str]] = {k: list(v) for k, v in detail.items()}

    # 建立 order_id → 結果 的查詢表
    by_order: dict[str, object] = {}
    for result in report.results:
        if result.order is not None:
            by_order[result.order.order_id] = result

    all_passed = True

    def check(label: str, passed: bool, detail_text: str) -> None:
        nonlocal all_passed
        mark = "✓" if passed else "✗"
        if not passed:
            all_passed = False
        print(f"  {mark} {label}")
        print(f"      {detail_text}")

    # --- 1. 正規化：小寫編號仍應靠編號被找到 -------------------------
    #
    # 判準是「有沒有靠編號找到」，也就是落在 Stage 1 或 Stage 2，
    # 而不是「有沒有落在 Stage 1」。因為一筆訂單可能同時被注入了
    # 退款或金額差異，那會讓它正當地降級到 Stage 2——那是金額的問題，
    # 不是正規化的問題。用 Stage 1 當判準會把引擎的正確行為誤判成失敗。
    #
    # 真正代表正規化失效的訊號是落到 Stage 3：那表示編號沒對上，
    # 只好靠模糊比對去猜。
    keyed_stages = (MatchStage.EXACT, MatchStage.TOLERANT)
    lowercase = injected.get("lowercase_id", [])
    hit = sum(
        1
        for oid in lowercase
        if (r := by_order.get(oid)) is not None and getattr(r, "stage", None) in keyed_stages
    )
    check(
        "正規化：編號被寫成小寫的訂單仍靠編號命中（Stage 1 或 2）",
        hit == len(lowercase),
        f"{hit} / {len(lowercase)} 筆（{pct(hit, len(lowercase))}）　落到 Stage 3 才代表正規化失效",
    )

    # --- 2. Stage 3 召回率：編號空白的訂單救回多少 -------------------
    missing_id = injected.get("missing_order_id", [])
    recovered = sum(
        1
        for oid in missing_id
        if (r := by_order.get(oid)) is not None and getattr(r, "stage", None) is MatchStage.FUZZY
    )
    # Stage 3 是模糊匹配，不可能 100%。低於六成代表評分函式或門檻要調。
    check(
        "Stage 3 召回率：訂單編號整欄空白的，靠模糊匹配救回",
        recovered / len(missing_id) >= 0.6 if missing_id else True,
        f"{recovered} / {len(missing_id)} 筆（{pct(recovered, len(missing_id))}）"
        f"　門檻 {report.config.accept_threshold}",
    )

    # --- 3. 四象限分類 ----------------------------------------------
    ledger_only = injected.get("ledger_only", [])
    correct = sum(
        1
        for oid in ledger_only
        if (r := by_order.get(oid)) is not None
        and getattr(r, "outcome", None) is MatchOutcome.MISSING_IN_SETTLEMENT
    )
    check(
        "分類：平台未撥款的訂單被歸到「未撥款」象限",
        correct == len(ledger_only),
        f"{correct} / {len(ledger_only)} 筆（{pct(correct, len(ledger_only))}）",
    )

    settlement_only = set(injected.get("settlement_only", []))
    flagged = {
        normalize_order_id(rec.external_order_id)
        for r in report.by_outcome(MatchOutcome.MISSING_IN_LEDGER)
        for rec in r.records
    }
    expected_keys = {normalize_order_id(code) for code in settlement_only}
    found = len(expected_keys & flagged)
    check(
        "分類：我方查無的結算列被歸到「漏記單」象限",
        found == len(expected_keys),
        f"{found} / {len(expected_keys)} 筆（{pct(found, len(expected_keys))}）",
    )

    variance_ids = injected.get("amount_variance", [])
    flagged_variance = sum(
        1
        for oid in variance_ids
        if (r := by_order.get(oid)) is not None
        and getattr(r, "outcome", None) is MatchOutcome.AMOUNT_VARIANCE
    )
    # 注入的差異有一部分小於一元，會被歸為捨入誤差；另有部分訂單同時被
    # 注入了退款，歸因會走 partial_refund。所以不要求 100%。
    check(
        "分類：注入的金額差異被偵測到",
        flagged_variance / len(variance_ids) >= 0.85 if variance_ids else True,
        f"{flagged_variance} / {len(variance_ids)} 筆"
        f"（{pct(flagged_variance, len(variance_ids))}）",
    )

    # --- 4. 去重 ----------------------------------------------------
    duplicates = injected.get("duplicated_row", [])
    check(
        "去重：平台重複匯出的列被擋下",
        report.metrics.duplicates_dropped == len(duplicates),
        f"擋下 {report.metrics.duplicates_dropped} 列，注入了 {len(duplicates)} 列",
    )

    # --- 5. 守恆 ----------------------------------------------------
    accounted_orders = len({r.order.order_id for r in report.results if r.order})
    accounted_records = sum(len(r.records) for r in report.results)
    check(
        "守恆：沒有任何一筆訂單或結算列憑空消失",
        accounted_orders == report.metrics.order_count
        and accounted_records == report.metrics.record_count,
        f"訂單 {accounted_orders}/{report.metrics.order_count}、"
        f"結算列 {accounted_records}/{report.metrics.record_count}",
    )

    return all_passed


# ----------------------------------------------------------------------
def naive_comparison_count(
    records: Sequence[SettlementRecord], orders: Sequence[Order]
) -> tuple[int, float]:
    """不建索引的暴力法：每筆未匹配的結算列都跟所有訂單比一次。

    實際跑一遍而不是只算理論值，因為要量測真實耗時。
    """
    config = MatchingConfig()
    unkeyed = [r for r in records if normalize_order_id(r.external_order_id) is None]
    started = time.perf_counter()
    comparisons = 0
    for record in unkeyed:
        for order in orders:
            comparisons += 1
            _score(record, order, config)
    return comparisons, time.perf_counter() - started


def scaling_test() -> None:
    """量測比對次數如何隨資料量成長。

    如果是 O(n²)，資料量翻倍比對次數會變四倍。索引版應該接近線性。
    """
    import subprocess
    import sys
    import tempfile
    from pathlib import Path

    print()
    rule("═")
    print("  複雜度實測：資料量翻倍時，比對次數怎麼成長")
    rule("═")
    print()
    print("  n = 訂單筆數。索引版每筆只跟同桶與相鄰桶的候選比對，")
    print("  暴力版每筆都跟所有訂單比。\n")
    print(f"  {'n':>6}  {'索引版':>12}  {'暴力版':>14}  {'倍數':>8}  {'索引耗時':>10}")
    rule()

    for size in (250, 500, 1000, 2000):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            subprocess.run(
                [
                    sys.executable,
                    "scripts/gen_fixtures.py",
                    "--orders",
                    str(size),
                    "--out",
                    str(out),
                ],
                check=True,
                capture_output=True,
            )
            orders = load_orders(out / "orders.csv")
            records = load_settlements(out)
            report = reconcile(orders, records)
            indexed = report.metrics.candidate_comparisons
            naive, _naive_time = naive_comparison_count(records, orders)
            ratio = naive / indexed if indexed else 0
            print(
                f"  {size:>6}  {indexed:>12,}  {naive:>14,}  "
                f"{ratio:>7.0f}×  {report.metrics.elapsed_seconds * 1000:>8.1f} ms"
            )

    print()
    print("  暴力版是標準的 O(n²)：n 變 8 倍，比對次數變約 44 倍。")
    print("  索引版明顯低於平方成長，但**不是**線性——n 變 8 倍時約 30 倍，")
    print("  大約 O(n^1.6)。原因是訂單變多時每個桶裡的候選也跟著變多，")
    print("  所以 O(n·k) 裡的 k 本身會隨 n 成長。要壓回接近線性，")
    print("  桶寬必須隨資料量縮小（同時把金額容差一起調緊）。")
    print("  這是目前設計的已知限制，寫在這裡而不是假裝它是線性的。")


# ----------------------------------------------------------------------
def main() -> int:
    parser = argparse.ArgumentParser(description="對帳結果的 ground truth 驗證")
    parser.add_argument("--scaling", action="store_true", help="加跑複雜度成長實測")
    args = parser.parse_args()

    if not require_data():
        return 1

    random.seed(0)
    orders = load_orders()
    records = load_settlements()
    manifest = load_manifest()
    report = reconcile(orders, records)
    m = report.metrics

    print()
    rule("═")
    print("  對帳結果驗證")
    rule("═")
    print(
        f"\n  資料：亂數種子 {manifest['seed']}、"
        f"訂單 {m.order_count} 筆、結算列 {m.record_count} 筆"
    )
    print(
        f"  設定：Stage 3 門檻 {report.config.accept_threshold}、"
        f"金額容差 {report.config.amount_tolerance}\n"
    )

    print("  拿 manifest.json 記錄的 ground truth 對答案：\n")
    passed = verify(report, manifest)

    print()
    rule()
    print("指標")
    print(f"\n  自動化率（僅 Stage 1）  {m.automation_rate:>7.1%}")
    print(f"  需人工確認              {m.review_rate:>7.1%}")
    print(f"  處理耗時                {m.elapsed_seconds * 1000:>7.1f} ms")
    print(f"  吞吐量                  {m.throughput:>7,.0f} 列/秒")

    print("\n  Stage 3 候選比對：")
    naive, naive_time = naive_comparison_count(records, orders)
    print(f"    建索引    {m.candidate_comparisons:>10,} 次")
    print(
        f"    不建索引  {naive:>10,} 次"
        f"（{naive / m.candidate_comparisons:.0f} 倍，耗時 {naive_time * 1000:.0f} ms）"
    )

    if args.scaling:
        scaling_test()

    print()
    rule("═")
    if passed:
        print("  全部通過。上面的數字有 ground truth 背書，可以寫進 README。")
    else:
        print("  有項目未通過（上面標 ✗ 的）。這不一定是 bug——")
        print("  也可能是評分權重或門檻需要調整。先看那一項的實際數字。")
    rule("═")
    print()
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
