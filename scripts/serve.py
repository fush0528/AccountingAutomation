"""啟動 API 伺服器，順便把資料庫準備好。

為什麼需要這支腳本
------------------

直接下 ``uvicorn reconciliation.api.app:app`` 會遇到兩個問題：

1. **套件找不到。** 專案用 src layout，沒有 ``pip install -e .`` 的話
   Python 找不到 ``reconciliation``。這支腳本會自己把 ``src/`` 加進路徑。
2. **資料庫是空的。** 就算上傳了結算單，裡面沒有我方訂單，
   對帳結果會全部落在「漏記單」象限——看起來像壞掉，其實只是沒資料。

所以這支腳本會依序：建立資料表 → 匯入我方訂單 → 啟動伺服器。

執行::

    python scripts/gen_fixtures.py    # 先產生測試資料
    python scripts/serve.py

然後開瀏覽器到 http://127.0.0.1:8000/docs

想換埠號或資料庫::

    python scripts/serve.py --port 9000
    python scripts/serve.py --database-url "postgresql+psycopg://user@localhost/recon"
    python scripts/serve.py --fresh        # 砍掉重練

注意這裡**沒有開 --reload**。自動重載需要 uvicorn 用字串路徑在子行程裡
重新 import，而子行程拿不到這支腳本設定的 sys.path。要邊改邊重載的話，
先 ``pip install -e .`` 再用 ``uvicorn reconciliation.api.app:app --reload``。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from _data import DATA_DIR, load_orders  # noqa: E402

from reconciliation.api import create_app  # noqa: E402
from reconciliation.repositories.sqlalchemy_repo import (  # noqa: E402
    SqlAlchemyOrderRepository,
    SqlAlchemySettlementRepository,
)

DEFAULT_DB = f"sqlite:///{ROOT / 'reconciliation.db'}"


def seed_orders(app: object, orders_csv: Path | None = None) -> tuple[int, int]:
    """把我方訂單寫進資料庫。

    回傳 ``(本次新增, 目前總數)``。重複執行是安全的——
    repository 會跳過已存在的 order_id。

    訂單來源可以換：``data/generated/orders.csv`` 是統計用的大批資料，
    ``data/samples/orders.csv`` 是逐列講解用的小樣本。兩邊的訂單編號
    不重疊，所以同時匯入也不會互相污染。
    """
    orders_csv = orders_csv or DATA_DIR / "orders.csv"
    if not orders_csv.exists():
        return 0, 0

    factory = app.state.session_factory  # type: ignore[attr-defined]
    session = factory()
    try:
        repo = SqlAlchemyOrderRepository(session)
        added = repo.add_many(load_orders(orders_csv))
        session.commit()
        return added, repo.count()
    finally:
        session.close()


def count_records(app: object) -> int:
    factory = app.state.session_factory  # type: ignore[attr-defined]
    session = factory()
    try:
        return SqlAlchemySettlementRepository(session).count()
    finally:
        session.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="啟動對帳引擎 API")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--database-url", default=DEFAULT_DB)
    parser.add_argument("--fresh", action="store_true", help="先刪掉現有的 SQLite 檔案再啟動")
    parser.add_argument("--no-seed", action="store_true", help="不要自動匯入訂單")
    parser.add_argument(
        "--samples",
        action="store_true",
        help="改用 data/samples/orders.csv 的 12 筆小樣本訂單（搭配 data/samples 裡的結算單）",
    )
    parser.add_argument(
        "--orders",
        type=Path,
        default=None,
        help="指定訂單 CSV 的路徑（覆蓋 --samples）",
    )
    args = parser.parse_args()

    orders_csv = args.orders or (
        ROOT / "data" / "samples" / "orders.csv" if args.samples else None
    )
    if orders_csv is not None and not orders_csv.exists():
        print(f"找不到訂單檔 {orders_csv}")
        print("先執行 python scripts/gen_samples.py（或 gen_fixtures.py）產生資料。")
        return 1

    if args.fresh and args.database_url.startswith("sqlite:///"):
        db_file = Path(args.database_url.removeprefix("sqlite:///"))
        if db_file.exists():
            db_file.unlink()
            print(f"已刪除 {db_file.name}")

    # create_app 會順便建立資料表（開發用；正式環境請跑 alembic upgrade head）
    app = create_app(args.database_url)

    print()
    print("═" * 74)
    print("  多平台電商對帳引擎")
    print("═" * 74)
    print(f"\n  資料庫　{args.database_url}")

    if not args.no_seed:
        added, total = seed_orders(app, orders_csv)
        if total == 0:
            print("\n  ⚠ 資料庫裡沒有我方訂單。")
            print("    先執行 python scripts/gen_fixtures.py 產生測試資料，")
            print("    否則對帳結果會全部落在「漏記單」象限——那不是程式壞了，")
            print("    是真的沒有東西可以對。")
        else:
            print(f"  訂單　　{total} 筆（本次新增 {added}）")

    records = count_records(app)
    print(f"  結算列　{records} 筆")
    if records == 0:
        print("\n  還沒有結算單。等一下在網頁上用 POST /api/imports 上傳，")
        print(f"  檔案在 {orders_csv.parent if orders_csv else DATA_DIR}")

    print()
    print("─" * 74)
    print(f"  API 文件　http://{args.host}:{args.port}/docs")
    print(f"  OpenAPI　 http://{args.host}:{args.port}/openapi.json")
    print("─" * 74)
    print("\n  在 /docs 頁面上可以直接點按鈕呼叫 API。建議順序：")
    print("    1. POST /api/imports              上傳結算單（三份都傳）")
    print("    2. POST /api/reconciliations      執行對帳，記下回傳的 run_id")
    print("    3. GET  /api/reconciliations/{run_id}/report   看四象限報告")
    print("\n  按 Ctrl+C 停止伺服器。\n")

    import uvicorn

    uvicorn.run(app, host=args.host, port=args.port, log_level="info")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
