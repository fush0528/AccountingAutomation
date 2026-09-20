"""載入 gen_fixtures.py 產生的資料。

這是示範用的臨時管路。W4 之後訂單會存在資料庫裡，由 repository 讀取，
這支檔案就會被取代掉。放在 scripts/ 而不是 src/ 就是這個意思：
它不是產品的一部分。
"""

from __future__ import annotations

import csv
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from reconciliation.domain import Money, Order, SettlementRecord  # noqa: E402
from reconciliation.parsers import ParseError, SettlementSource, detect  # noqa: E402

DATA_DIR = ROOT / "data" / "generated"


def load_orders(path: Path | None = None) -> list[Order]:
    """讀取我方訂單。"""
    path = path or DATA_DIR / "orders.csv"
    orders: list[Order] = []
    # utf-8-sig：檔案有沒有 BOM 都能讀。訂單檔會被人用 Excel 打開，
    # 而 Excel 沒有 BOM 就會把中文顯示成亂碼，所以產生端會寫 BOM——
    # 讀取端就必須容忍它，否則第一個欄位名稱會多一個看不見的字元。
    with path.open(encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            orders.append(
                Order(
                    order_id=row["訂單編號"],
                    platform=row["平台"],
                    external_order_id=row["平台訂單編號"],
                    ordered_at=datetime.fromisoformat(row["訂單時間"]),
                    product_name=row["商品名稱"],
                    quantity=int(row["數量"]),
                    gross_amount=Money.from_str(row["訂單金額"]),
                    expected_fee=Money.from_str(row["預期手續費"]),
                )
            )
    return orders


def load_settlements(directory: Path | None = None) -> list[SettlementRecord]:
    """把目錄裡所有結算單解析成領域模型。

    注意這裡不指定平台——每一份都丟給 registry 自己認。
    """
    directory = directory or DATA_DIR
    records: list[SettlementRecord] = []
    # 白名單而不是黑名單：目錄裡會有 manifest.json、SAMPLES.md 這類
    # 附屬檔案，一個個排除遲早會漏掉。只收看起來像結算單的副檔名。
    settlement_suffixes = {".csv", ".xlsx", ".xlsm", ".txt"}
    for path in sorted(directory.iterdir()):
        if path.name == "orders.csv" or path.suffix.lower() not in settlement_suffixes:
            continue
        source = SettlementSource.from_path(path)
        try:
            records.extend(detect(source).parse(source).records)
        except ParseError as exc:
            print(f"  ⚠ {path.name} 解析失敗：{exc}")
    return records


def load_manifest(directory: Path | None = None) -> dict[str, object]:
    """讀取 ground truth。"""
    directory = directory or DATA_DIR
    data: dict[str, object] = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
    return data


def require_data(directory: Path | None = None) -> bool:
    directory = directory or DATA_DIR
    if not (directory / "manifest.json").exists():
        print("找不到測試資料。請先執行：python scripts/gen_fixtures.py")
        return False
    return True
