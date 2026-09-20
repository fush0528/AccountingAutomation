"""三階段對帳匹配引擎。

這是整個專案的核心。對帳不是一句 SQL JOIN，因為兩邊的資料對不上的方式
太多種：編號格式不同、金額有差、整欄沒有編號、同一訂單拆成多列撥款、
平台重複匯出。每一種都需要不同的處理策略，而且**必須分階段**，
因為信心水準不同——精確匹配可以自動放行，模糊匹配一定要人看過。

流程
----

::

    去重 ─→ 依正規化編號分組 ─┬─→ Stage 1 精確：編號＋金額皆符 ─→ 自動勾稽
                              └─→ Stage 2 容差：編號符、金額差 ─→ 歸因後送審
                                        ↓（編號缺失或查無此單）
                              Stage 3 模糊：候選集評分 ─┬─→ 高分：暫定匹配，送審
                                                        └─→ 低分：漏記單＋候選清單
                                        ↓
                              剩餘未匹配的訂單 ─→ 未撥款

為什麼 Stage 3 不會是 O(n²)
---------------------------

天真的做法是每一筆未匹配的結算列都跟所有訂單比一次，1000 筆對 1000 筆
就是一百萬次比對。這裡先用 ``(平台, 金額分桶)`` 建索引，每筆只需要跟
同桶與相鄰桶的少數訂單比對，複雜度從 O(n²) 降到近似 O(n·k)，
k 是平均候選數。``scripts/benchmark.py`` 有實測數字。

金額分桶的桶寬必須**大於**允許的金額容差，否則會漏掉落在桶邊界外的
正確候選——所以查詢時一律看相鄰的三個桶。
"""

from __future__ import annotations

import time
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from decimal import Decimal
from difflib import SequenceMatcher

from .models import (
    Candidate,
    MatchOutcome,
    MatchResult,
    MatchStage,
    Order,
    SettlementRecord,
)
from .money import Money
from .normalization import normalize_order_id, normalize_product_name
from .variance import analyse_variance

__all__ = ["MatchingConfig", "ReconciliationReport", "reconcile"]


@dataclass(frozen=True, slots=True)
class MatchingConfig:
    """匹配引擎的可調參數。

    全部集中在一個物件而不是散落成函式參數，是為了讓「這次是用什麼設定
    跑出來的」可以被記錄下來——對帳結果沒有設定就沒有意義，
    92% 的匹配率在門檻 0.5 與 0.9 下是完全不同的兩件事。
    """

    #: 結算日與訂單日之間合理的間隔，超出就不當候選
    settlement_lag: tuple[int, int] = (0, 45)
    #: Stage 3 允許的金額差異上限
    amount_tolerance: Money = field(default_factory=lambda: Money.from_str("100"))
    #: 金額分桶的桶寬，必須大於 amount_tolerance
    bucket_width: Money = field(default_factory=lambda: Money.from_str("500"))
    #: Stage 3 自動採納的分數門檻；低於此值只產生候選清單交人工判斷
    accept_threshold: Decimal = Decimal("0.82")
    #: 每筆最多保留幾個候選給人工參考
    max_candidates: int = 3
    #: 評分權重，三者相加應為 1
    weight_amount: Decimal = Decimal("0.55")
    weight_date: Decimal = Decimal("0.15")
    weight_product: Decimal = Decimal("0.30")

    def __post_init__(self) -> None:
        if self.bucket_width <= self.amount_tolerance:
            raise ValueError(
                "bucket_width 必須大於 amount_tolerance，否則正確的候選會落在查詢範圍之外而被漏掉"
            )
        total = self.weight_amount + self.weight_date + self.weight_product
        if total != Decimal(1):
            raise ValueError(f"評分權重相加必須等於 1，目前是 {total}")


@dataclass(frozen=True, slots=True)
class ReconciliationMetrics:
    """對帳的量化結果。README 裡放的就是這幾個數字。"""

    order_count: int
    record_count: int
    duplicates_dropped: int
    matched: int
    amount_variance: int
    missing_in_ledger: int
    missing_in_settlement: int
    stage_exact: int
    stage_tolerant: int
    stage_fuzzy: int
    needs_review: int
    elapsed_seconds: float

    candidate_comparisons: int
    """Stage 3 實際做了幾次評分比對（建了候選索引之後）。"""

    fuzzy_record_count: int
    """有幾列結算資料進到 Stage 3（前兩階段靠編號配不掉的）。

    有了這個數字才算得出平均候選數 k = candidate_comparisons / fuzzy_record_count，
    而 k 隨 n 怎麼成長，才是索引到底有沒有漸近效果的真正判準。
    """

    naive_comparisons: int
    """同樣的工作不建索引要做幾次比對，作為對照基準。

    定義是「進入 Stage 3 的結算列 × Stage 1／2 之後還沒被認領的訂單」，
    **不是**「全部結算列 × 全部訂單」。這個區別很重要：不建索引的實作
    一樣知道哪些訂單已經在前兩階段配掉了，跳過它們不需要任何索引。
    拿全部訂單當分母會把索引的功勞誇大一個數量級——那是在跟一個
    沒有人會寫的爛實作比較，不是誠實的對照。
    """

    @property
    def mean_candidates(self) -> float:
        """Stage 3 每一列平均要跟幾個候選比對，也就是 O(n·k) 裡的 k。"""
        return (
            self.candidate_comparisons / self.fuzzy_record_count
            if self.fuzzy_record_count
            else 0.0
        )

    @property
    def total_results(self) -> int:
        return (
            self.matched
            + self.amount_variance
            + self.missing_in_ledger
            + self.missing_in_settlement
        )

    @property
    def automation_rate(self) -> float:
        """完全不需要人工介入的比例。

        只算 Stage 1，因為那是唯一可以直接入帳的情況。把 Stage 2、Stage 3
        算進來會讓數字好看很多，但那是在自我欺騙——那兩種都得有人看過。
        """
        return self.stage_exact / self.total_results if self.total_results else 0.0

    @property
    def review_rate(self) -> float:
        return 1.0 - self.automation_rate

    @property
    def throughput(self) -> float:
        """每秒處理幾列結算記錄。"""
        return self.record_count / self.elapsed_seconds if self.elapsed_seconds else 0.0


@dataclass(frozen=True, slots=True)
class ReconciliationReport:
    """一次對帳的完整輸出。"""

    results: tuple[MatchResult, ...]
    metrics: ReconciliationMetrics
    config: MatchingConfig = field(default_factory=MatchingConfig)

    def by_outcome(self, outcome: MatchOutcome) -> tuple[MatchResult, ...]:
        return tuple(r for r in self.results if r.outcome is outcome)

    def needing_review(self) -> tuple[MatchResult, ...]:
        return tuple(r for r in self.results if r.needs_review)


# ----------------------------------------------------------------------
def reconcile(
    orders: Iterable[Order],
    records: Iterable[SettlementRecord],
    config: MatchingConfig | None = None,
) -> ReconciliationReport:
    """對帳主流程。"""
    config = config or MatchingConfig()
    started = time.perf_counter()

    order_list = list(orders)
    deduped, duplicates_dropped = _deduplicate(records)

    keyed, unkeyed = _partition_by_key(deduped)
    order_index = _index_orders_by_key(order_list)

    results: list[MatchResult] = []
    matched_order_ids: set[str] = set()
    leftover_records: list[SettlementRecord] = list(unkeyed)

    # --- Stage 1 與 Stage 2：靠正規化後的訂單編號 --------------------
    for key, group in keyed.items():
        order = order_index.get(key)
        if order is None:
            # 編號存在但我方查無此單 —— 先留給 Stage 3，說不定是編號打錯
            leftover_records.extend(group)
            continue
        matched_order_ids.add(order.order_id)
        results.append(_compare(order, group, stage_hint=MatchStage.EXACT))

    # --- Stage 3：模糊匹配 -------------------------------------------
    remaining_orders = [o for o in order_list if o.order_id not in matched_order_ids]
    # 對照基準：不建索引時這一階段要做的比對次數。在進迴圈前算，
    # 因為 _fuzzy_match 過程中 claimed 會長大，事後就算不回來了。
    naive_comparisons = len(leftover_records) * len(remaining_orders)
    fuzzy_results, fuzzy_matched_ids, comparisons = _fuzzy_match(
        leftover_records, remaining_orders, config
    )
    results.extend(fuzzy_results)
    matched_order_ids |= fuzzy_matched_ids

    # --- 剩下沒對到的訂單：平台未撥款 --------------------------------
    for order in order_list:
        if order.order_id not in matched_order_ids:
            results.append(MatchResult(outcome=MatchOutcome.MISSING_IN_SETTLEMENT, order=order))

    elapsed = time.perf_counter() - started
    metrics = _build_metrics(
        results=results,
        order_count=len(order_list),
        record_count=len(deduped),
        duplicates_dropped=duplicates_dropped,
        elapsed=elapsed,
        comparisons=comparisons,
        fuzzy_record_count=len(leftover_records),
        naive_comparisons=naive_comparisons,
    )
    return ReconciliationReport(results=tuple(results), metrics=metrics, config=config)


# ----------------------------------------------------------------------
# 前處理
# ----------------------------------------------------------------------
def _dedup_key(record: SettlementRecord) -> tuple[object, ...]:
    """一列結算記錄的身分。

    平台重複匯出時，同一筆交易會一模一樣地出現兩次。用整列內容當指紋，
    比只看訂單編號安全——同一個訂單本來就可能有多列（銷售加退款），
    只看編號會把合法的多列誤判成重複。

    注意這裡**不**包含 ``source_row``：同一筆交易在檔案裡的列號當然不同，
    把它算進指紋就等於沒有去重。
    """
    return (
        record.platform,
        record.external_order_id,
        record.settled_at,
        record.gross_amount,
        record.fee_amount,
        record.net_amount,
        record.record_type,
    )


def _deduplicate(
    records: Iterable[SettlementRecord],
) -> tuple[list[SettlementRecord], int]:
    """移除完全相同的重複列，回傳去重後的列表與被丟掉的筆數。

    W4 會在資料庫層用唯一約束做同樣的事。這裡先做一次，是因為對帳
    本身就不該被重複列影響——而且被丟掉的筆數要報告出來，
    使用者有權知道系統替他們拿掉了什麼。
    """
    seen: set[tuple[object, ...]] = set()
    kept: list[SettlementRecord] = []
    dropped = 0
    for record in records:
        key = _dedup_key(record)
        if key in seen:
            dropped += 1
            continue
        seen.add(key)
        kept.append(record)
    return kept, dropped


def _partition_by_key(
    records: Sequence[SettlementRecord],
) -> tuple[dict[str, list[SettlementRecord]], list[SettlementRecord]]:
    """把結算列分成「有可用編號」與「沒有」兩堆。

    有編號的依正規化後的鍵分組——同一筆訂單的銷售列與退款列會落在同一組。
    """
    keyed: dict[str, list[SettlementRecord]] = {}
    unkeyed: list[SettlementRecord] = []
    for record in records:
        key = normalize_order_id(record.external_order_id)
        if key is None:
            unkeyed.append(record)
        else:
            keyed.setdefault(key, []).append(record)
    return keyed, unkeyed


def _index_orders_by_key(orders: Sequence[Order]) -> dict[str, Order]:
    index: dict[str, Order] = {}
    for order in orders:
        key = normalize_order_id(order.external_order_id)
        if key is not None:
            index[key] = order
    return index


# ----------------------------------------------------------------------
# Stage 1 / Stage 2
# ----------------------------------------------------------------------
def _compare(
    order: Order, records: Sequence[SettlementRecord], *, stage_hint: MatchStage
) -> MatchResult:
    """訂單與它的結算列已經配對上了，剩下的是金額對不對。"""
    analysis = analyse_variance(order, records)

    if analysis.variance.is_zero:
        return MatchResult(
            outcome=MatchOutcome.MATCHED,
            stage=stage_hint,
            order=order,
            records=tuple(records),
            variance=analysis.variance,
        )

    # 金額有差 —— 不管是靠編號還是靠模糊比對找到的，都降級為容差匹配。
    # 因為「編號對上但金額不對」這件事本身就需要人看過。
    stage = MatchStage.TOLERANT if stage_hint is MatchStage.EXACT else stage_hint
    return MatchResult(
        outcome=MatchOutcome.AMOUNT_VARIANCE,
        stage=stage,
        order=order,
        records=tuple(records),
        variance=analysis.variance,
        variance_reason=analysis.reason,
    )


# ----------------------------------------------------------------------
# Stage 3
# ----------------------------------------------------------------------
def _bucket(amount: Money, width: Money) -> int:
    """把金額映射到一個桶號。"""
    return amount.minor_units // width.minor_units


def _build_candidate_index(
    orders: Sequence[Order], config: MatchingConfig
) -> dict[tuple[str, int], list[Order]]:
    """依 ``(平台, 金額分桶)`` 建索引。

    這一步是 Stage 3 不會退化成 O(n²) 的原因。建索引本身是 O(n)，
    之後每次查詢只碰到少數幾個桶。
    """
    index: dict[tuple[str, int], list[Order]] = {}
    for order in orders:
        key = (order.platform, _bucket(order.expected_net, config.bucket_width))
        index.setdefault(key, []).append(order)
    return index


def _lookup_candidates(
    record: SettlementRecord,
    index: dict[tuple[str, int], list[Order]],
    config: MatchingConfig,
) -> list[Order]:
    """取出可能匹配這筆結算列的訂單。

    查相鄰三個桶，因為容差可能讓正確的候選落在隔壁桶。這也是為什麼
    ``bucket_width`` 必須大於 ``amount_tolerance``——否則要查的桶數
    會隨容差無限增加，索引就失去意義了。
    """
    centre = _bucket(record.net_amount, config.bucket_width)
    found: list[Order] = []
    for offset in (-1, 0, 1):
        found.extend(index.get((record.platform, centre + offset), []))
    return found


def _score(
    record: SettlementRecord, order: Order, config: MatchingConfig
) -> tuple[Decimal, tuple[str, ...]]:
    """替一個候選評分，並產生人看得懂的理由。

    理由不是裝飾。人工確認時操作者必須知道系統為什麼覺得這兩筆是同一件事，
    只給一個 0.87 沒有人敢按下確認。
    """
    reasons: list[str] = []

    # --- 金額接近度 -------------------------------------------------
    gap = abs(record.net_amount - order.expected_net)
    if gap > config.amount_tolerance:
        return Decimal(0), ()
    if gap.is_zero:
        amount_score = Decimal(1)
        reasons.append("金額完全相符")
    else:
        amount_score = Decimal(1) - (
            Decimal(gap.minor_units) / Decimal(config.amount_tolerance.minor_units)
        )
        reasons.append(f"金額相差 {gap}")

    # --- 日期合理性 -------------------------------------------------
    lag_days = (record.settled_at - order.ordered_at.date()).days
    low, high = config.settlement_lag
    if not (low <= lag_days <= high):
        return Decimal(0), ()
    span = max(high - low, 1)
    # 越接近典型的撥款週期中位數越可信
    midpoint = (low + high) / 2
    date_score = Decimal(1) - Decimal(abs(lag_days - midpoint) / span).quantize(Decimal("0.0001"))
    date_score = max(Decimal(0), min(Decimal(1), date_score))
    reasons.append(f"訂單後 {lag_days} 天結算")

    # --- 商品名稱相似度 ---------------------------------------------
    record_name = normalize_product_name(record.product_name)
    order_name = normalize_product_name(order.product_name)
    if record_name and order_name:
        ratio = SequenceMatcher(None, record_name, order_name).ratio()
        product_score = Decimal(str(round(ratio, 4)))
        if ratio >= 0.99:
            reasons.append("商品名稱相同")
        else:
            reasons.append(f"商品名稱相似度 {ratio:.0%}")
    else:
        # 沒有商品名稱可比時給中性分數，而不是 0——
        # 給 0 會讓「資料不足」被當成「證據不符」，那是兩回事。
        product_score = Decimal("0.5")
        reasons.append("無商品名稱可比對")

    total = (
        amount_score * config.weight_amount
        + date_score * config.weight_date
        + product_score * config.weight_product
    )
    return min(Decimal(1), max(Decimal(0), total)), tuple(reasons)


def _fuzzy_match(
    records: Sequence[SettlementRecord],
    orders: Sequence[Order],
    config: MatchingConfig,
) -> tuple[list[MatchResult], set[str], int]:
    """對沒有可用編號的結算列做模糊匹配。

    採「每筆各自取最高分」的貪婪策略，並且一筆訂單只能被認領一次。
    這不是全域最佳解（那是指派問題，可以用匈牙利演算法求解），
    但在對帳的情境下夠用：分數夠高的配對通常也是唯一合理的配對，
    而分數不夠高的本來就要人工確認。這個取捨寫在這裡是刻意的——
    如果之後發現誤配率偏高，這裡就是要改的地方。
    """
    index = _build_candidate_index(orders, config)
    claimed: set[str] = set()
    results: list[MatchResult] = []
    comparisons = 0

    for record in records:
        scored: list[Candidate] = []
        for order in _lookup_candidates(record, index, config):
            if order.order_id in claimed:
                continue
            comparisons += 1
            score, reasons = _score(record, order, config)
            if score > 0:
                scored.append(Candidate(order=order, score=score, reasons=reasons))

        scored.sort(key=lambda c: -c.score)
        best = scored[0] if scored else None

        if best is not None and best.score >= config.accept_threshold:
            claimed.add(best.order.order_id)
            results.append(_compare(best.order, [record], stage_hint=MatchStage.FUZZY))
            continue

        # 分數不夠——不猜。列為漏記單，但把候選附上去讓人工判斷。
        results.append(
            MatchResult(
                outcome=MatchOutcome.MISSING_IN_LEDGER,
                records=(record,),
                candidates=tuple(scored[: config.max_candidates]),
            )
        )

    return results, claimed, comparisons


# ----------------------------------------------------------------------
def _build_metrics(
    *,
    results: Sequence[MatchResult],
    order_count: int,
    record_count: int,
    duplicates_dropped: int,
    elapsed: float,
    comparisons: int,
    fuzzy_record_count: int,
    naive_comparisons: int,
) -> ReconciliationMetrics:
    def count_outcome(outcome: MatchOutcome) -> int:
        return sum(1 for r in results if r.outcome is outcome)

    def count_stage(stage: MatchStage) -> int:
        return sum(1 for r in results if r.stage is stage)

    return ReconciliationMetrics(
        order_count=order_count,
        record_count=record_count,
        duplicates_dropped=duplicates_dropped,
        matched=count_outcome(MatchOutcome.MATCHED),
        amount_variance=count_outcome(MatchOutcome.AMOUNT_VARIANCE),
        missing_in_ledger=count_outcome(MatchOutcome.MISSING_IN_LEDGER),
        missing_in_settlement=count_outcome(MatchOutcome.MISSING_IN_SETTLEMENT),
        stage_exact=count_stage(MatchStage.EXACT),
        stage_tolerant=count_stage(MatchStage.TOLERANT),
        stage_fuzzy=count_stage(MatchStage.FUZZY),
        needs_review=sum(1 for r in results if r.needs_review),
        elapsed_seconds=elapsed,
        candidate_comparisons=comparisons,
        fuzzy_record_count=fuzzy_record_count,
        naive_comparisons=naive_comparisons,
    )
