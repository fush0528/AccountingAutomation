"""API 實測：走一次完整的 HTTP 流程，並證明冪等匯入真的有效。

這支腳本不需要你另外開伺服器——它用 FastAPI 的 TestClient 在同一個
程序裡打真實的 HTTP 請求，走的是跟正式環境完全一樣的路徑：
路由、Pydantic 驗證、服務層、SQLAlchemy、SQLite。

執行::

    python scripts/gen_fixtures.py
    python scripts/api_demo.py

想自己用瀏覽器點::

    uvicorn reconciliation.api.app:app --reload
    # 然後開 http://127.0.0.1:8000/docs
"""

from __future__ import annotations

import json
import sys
from collections.abc import Callable
from pathlib import Path

from _data import DATA_DIR, load_orders, require_data

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import httpx  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

from reconciliation.api import create_app  # noqa: E402
from reconciliation.repositories.sqlalchemy_repo import (  # noqa: E402
    SqlAlchemyOrderRepository,
)


def rule(char: str = "─", width: int = 74) -> None:
    print(char * width)


def show(method: str, path: str, response: object, body: object = None) -> None:
    status = getattr(response, "status_code", "")
    print(f"  {method:<5} {path:<46} → {status}")
    if body is not None:
        text = json.dumps(body, ensure_ascii=False, indent=2)
        for line in text.splitlines()[:14]:
            print(f"        {line}")


def main() -> int:
    if not require_data():
        return 1

    print()
    rule("═")
    print("  API 實測：上傳 → 對帳 → 報告")
    rule("═")

    app = create_app("sqlite:///:memory:")
    with TestClient(app) as client:
        # --- 先把我方訂單塞進去（W5 之後會有正式的匯入端點）----------
        factory = app.state.session_factory
        session = factory()
        added = SqlAlchemyOrderRepository(session).add_many(load_orders())
        session.commit()
        session.close()
        print(f"\n  （前置）寫入我方訂單 {added} 筆")

        print()
        rule()
        print("1. 探索 API")
        response = client.get("/api/platforms")
        show("GET", "/api/platforms", response, response.json())
        print("\n  這個端點讀的是 parser registry。新增一個平台不需要改 API。")

        # --- 上傳三份結算單 -----------------------------------------
        print()
        rule()
        print("2. 上傳結算單（平台自動偵測）\n")
        files = sorted(
            p for p in DATA_DIR.iterdir() if p.name != "orders.csv" and p.suffix != ".json"
        )
        first_batch = None
        for path in files:
            response = client.post(
                "/api/imports",
                files={"file": (path.name, path.read_bytes(), "application/octet-stream")},
            )
            body = response.json()
            first_batch = first_batch or body
            print(
                f"  POST  /api/imports  {path.name:<28} → {response.status_code}  "
                f"{body['platform']}　{body['message']}"
            )

        # --- 冪等性：重上傳一次 --------------------------------------
        print()
        rule()
        print("3. 冪等匯入：把同一份檔案再上傳四次\n")
        target = files[0]
        before = client.get("/api/health").json()["settlement_records"]
        for attempt in range(1, 5):
            response = client.post(
                "/api/imports",
                files={
                    "file": (
                        f"{target.stem}-副本{attempt}.csv",
                        target.read_bytes(),
                        "application/octet-stream",
                    )
                },
            )
            body = response.json()
            mark = "重複，未寫入" if body["was_duplicate"] else "！！寫入了新資料"
            print(f"  第 {attempt} 次（改了檔名）→ batch {body['batch_id']}　{mark}")
        after = client.get("/api/health").json()["settlement_records"]

        print(f"\n  結算記錄筆數：上傳前 {before}　上傳後 {after}")
        print(f"  批次總數：{len(client.get('/api/imports').json())}（三份檔案，不是七份）")
        print("\n  指紋看的是**內容**不是檔名，所以改檔名也沒用。")
        print("  這就是「至少一次投遞下的恰好一次語意」——呼叫端可以放心重試。")
        assert before == after, "冪等性被破壞了"

        # --- 對帳 ----------------------------------------------------
        print()
        rule()
        print("4. 執行對帳\n")
        response = client.post("/api/reconciliations", json={})
        run = response.json()
        m = run["metrics"]
        print(f"  POST  /api/reconciliations  → {response.status_code}  {run['run_id']}")
        print(f"\n    訂單 {m['order_count']} 筆、結算列 {m['record_count']} 筆")
        print(
            f"    已勾稽 {m['matched']}　金額差異 {m['amount_variance']}　"
            f"漏記單 {m['missing_in_ledger']}　未撥款 {m['missing_in_settlement']}"
        )
        print(f"    自動化率 {m['automation_rate']:.1%}　耗時 {m['elapsed_seconds'] * 1000:.1f} ms")
        print(
            f"    設定：門檻 {run['config']['accept_threshold']}"
            f"（報告一定要附設定，否則數字沒有意義）"
        )

        # --- 報告與篩選 ----------------------------------------------
        print()
        rule()
        print("5. 查看報告，依象限篩選\n")
        run_id = run["run_id"]
        for outcome in (
            "matched",
            "amount_variance",
            "missing_in_ledger",
            "missing_in_settlement",
        ):
            body = client.get(
                f"/api/reconciliations/{run_id}/report?outcome={outcome}&limit=1"
            ).json()
            print(f"  GET   ?outcome={outcome:<24} → {body['total']:>4} 筆")

        review = client.get(
            f"/api/reconciliations/{run_id}/report?needs_review=true&limit=1"
        ).json()
        print(f"  GET   ?needs_review=true{' ':<17} → {review['total']:>4} 筆")

        # --- 一筆明細 ------------------------------------------------
        print()
        rule()
        print("6. 一筆金額差異的完整回應\n")
        body = client.get(
            f"/api/reconciliations/{run_id}/report?outcome=amount_variance&limit=1"
        ).json()
        if body["results"]:
            item = body["results"][0]
            print(
                json.dumps(
                    {
                        "outcome": item["outcome"],
                        "stage": item["stage"],
                        "needs_review": item["needs_review"],
                        "order": {
                            "order_id": item["order"]["order_id"],
                            "expected_net": item["order"]["expected_net"],
                        },
                        "settled_amount": item["settled_amount"],
                        "variance": item["variance"],
                        "variance_reason": item["variance_reason"],
                    },
                    ensure_ascii=False,
                    indent=2,
                ).replace("\n", "\n    ")
            )
            print("\n    金額全部是字串。JSON 的 number 是 IEEE 754 雙精度浮點數，")
            print("    944.70 放進去再拿出來不保證還是 944.70——整個專案花了這麼多")
            print("    力氣避開浮點數，不該在最後一哩前功盡棄。")

        # --- 錯誤處理 ------------------------------------------------
        print()
        rule()
        print("7. 錯誤處理：所有錯誤長同一個樣子\n")
        cases: list[tuple[str, Callable[[], httpx.Response]]] = [
            (
                "上傳無法辨識的檔案",
                lambda: client.post(
                    "/api/imports",
                    files={"file": ("x.csv", b"\xe5\xa7\x93\xe5\x90\x8d,x\n1,2", "text/csv")},
                ),
            ),
            ("查不存在的批次", lambda: client.get("/api/imports/BATCH-NOPE")),
            ("查不存在的對帳", lambda: client.get("/api/reconciliations/RUN-NOPE/report")),
            (
                "用不合法的象限篩選",
                lambda: client.get(f"/api/reconciliations/{run_id}/report?outcome=nonsense"),
            ),
        ]
        for label, call in cases:
            response = call()
            body = response.json()
            print(f"  {label:<22} → {response.status_code}  {body['code']}")
            print(f"        {body['message'][:62]}⋯")

        print("\n  code 給程式判斷，message 給人看。分開之後前端就不必")
        print("  比對中文字串來決定行為。")

    print()
    rule("═")
    print("  OpenAPI 文件在 http://127.0.0.1:8000/docs（uvicorn 啟動後）")
    print("  W5 的前端會用 openapi-typescript 從那份文件產生 TypeScript 型別，")
    print("  做到前後端型別的單一真實來源。")
    rule("═")
    print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
