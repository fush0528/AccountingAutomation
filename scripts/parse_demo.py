"""解析層實測：把產生的三份結算單各自解析一遍，印出結果。

這支腳本回答一個問題：**「可插拔的解析層」實際上長什麼樣？**

它不指定平台，只把檔案丟給 registry 讓它自己認。三份檔案格式天差地遠
（UTF-8 CSV、Big5 CSV、多層表頭 Excel），但解析完得到的都是同一種
:class:`~reconciliation.domain.models.SettlementRecord`——下游的對帳引擎
完全不知道這些差異存在。

執行（先跑 gen_fixtures.py 產生資料）::

    python scripts/gen_fixtures.py
    python scripts/parse_demo.py

想看單一檔案的細節::

    python scripts/parse_demo.py --file data/generated/platform_c_2026-03.csv --show 5
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from reconciliation.domain import SettlementRecordType, money_sum  # noqa: E402
from reconciliation.parsers import (  # noqa: E402
    ParseError,
    ParseResult,
    SettlementSource,
    available_platforms,
    detect,
)
from reconciliation.parsers.registry import _REGISTRY  # noqa: E402

DATA_DIR = ROOT / "data" / "generated"


def rule(char: str = "─", width: int = 74) -> None:
    print(char * width)


def show_registry() -> None:
    print("\n已註冊的平台（registry 的內容）：")
    for code, name in available_platforms():
        print(f"  {code:<14} {name}")
    print("\n新增一個平台只要：新增一個檔案、加上 @register、在 __init__ 補一行 import。")
    print("核心的任何一行程式碼都不用改。")


def show_sniffing(source: SettlementSource) -> str:
    """把每個 parser 的信心分數攤開來看——偵測不是黑箱。"""
    scores = sorted(
        ((cls.sniff(source), cls) for cls in _REGISTRY.values()),
        key=lambda item: -item[0],
    )
    parts = [f"{cls.display_name.split('（')[0]} {score:.0%}" for score, cls in scores]
    return "　".join(parts)


def report(result: ParseResult, show: int) -> None:
    print(f"  {result.summary()}")

    sales = [r for r in result.records if r.record_type is SettlementRecordType.SALE]
    refunds = [r for r in result.records if r.record_type is SettlementRecordType.REFUND]
    no_id = [r for r in result.records if not r.has_order_id]

    print(f"    銷售 {len(sales)} 列、退款 {len(refunds)} 列、無訂單編號 {len(no_id)} 列")

    if result.records:
        total_net = money_sum(r.net_amount for r in result.records)
        total_fee = money_sum(r.fee_amount for r in result.records)
        print(f"    手續費合計 {total_fee}、撥款合計 {total_net}")

    bad_residual = [r for r in result.records if not r.residual.is_zero]
    if bad_residual:
        print(f"    ⚠ {len(bad_residual)} 列的 gross - fee - net 不為零，parser 可能漏了扣項")

    for err in result.errors[:3]:
        print(f"    ✗ {err}")
    if result.rejected_count > 3:
        print(f"    ✗ ⋯⋯ 另有 {result.rejected_count - 3} 列失敗")

    for record in result.records[:show]:
        oid = record.external_order_id or "(無編號)"
        print(
            f"      第{record.source_row:>4}列  {oid:<14} "
            f"{record.settled_at}  {record.net_amount}  {record.record_type.value}"
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="解析層實測")
    parser.add_argument("--file", type=Path, help="只解析這一個檔案")
    parser.add_argument("--show", type=int, default=0, help="每份檔案印出前 N 筆明細")
    args = parser.parse_args()

    if args.file:
        targets = [args.file]
    else:
        if not DATA_DIR.exists():
            print("找不到測試資料。請先執行：python scripts/gen_fixtures.py")
            return 1
        targets = sorted(
            p for p in DATA_DIR.iterdir() if p.name != "orders.csv" and p.suffix != ".json"
        )

    if not targets:
        print("沒有可解析的檔案。請先執行：python scripts/gen_fixtures.py")
        return 1

    print()
    rule("═")
    print("  解析層實測：三種格式，一個介面")
    rule("═")
    show_registry()

    all_records = 0
    all_errors = 0

    for path in targets:
        print()
        rule()
        print(f"檔案：{path.name}（{path.stat().st_size:,} bytes）")

        source = SettlementSource.from_path(path)
        print(f"  SHA-256：{source.sha256[:32]}⋯   ← W4 的冪等匯入用這個當指紋")
        print(f"  各平台信心分數：{show_sniffing(source)}")

        try:
            selected = detect(source)
        except ParseError as exc:
            print(f"  ✗ {exc}")
            continue

        print(f"  → 選中 {selected.display_name}")

        try:
            result = selected.parse(source)
        except ParseError as exc:
            print(f"  ✗ 解析失敗：{exc}")
            continue

        report(result, args.show)
        all_records += result.accepted_count
        all_errors += result.rejected_count

    print()
    rule("═")
    print(f"  合計解析出 {all_records} 筆 SettlementRecord，{all_errors} 列失敗。")
    print()
    print("  重點不是數字，是這件事：上面三份檔案的編碼、日期格式、表頭結構")
    print("  完全不同，但它們最後都變成同一種領域模型。W3 的對帳引擎只需要")
    print("  認識 SettlementRecord，永遠不會知道 Big5 或合併儲存格的存在。")
    rule("═")
    print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
