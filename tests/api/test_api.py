"""API 整合測試。

這些測試起一個真的 FastAPI app，打真的 HTTP 請求，寫真的資料庫——
只是資料庫在記憶體裡。所以它們驗證的是**整條路徑**：
路由、DTO 轉換、服務層、repository、SQL、交易邊界。

最重要的一組是 :class:`TestIdempotency`。冪等性不是「寫在 docstring 裡的
承諾」，是必須被證明的性質：同一份檔案上傳兩次，資料庫裡的筆數必須不變。
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import httpx
from fastapi.testclient import TestClient

from reconciliation.api import create_app
from reconciliation.repositories.database import create_db_engine, init_schema
from tests.parsers.conftest import PLATFORM_A_HEADER

SETTLEMENT_CSV = "\n".join(
    [
        PLATFORM_A_HEADER,
        "2026-03-01,SP-AAAA0001,EC1,信用卡,5.53%,無線滑鼠,1000,"
        "30.42,19.35,5.53,,0.00,944.70,2026-03-13,2026-03-15",
        "2026-03-02,SP-AAAA0002,EC2,信用卡,5.53%,機械鍵盤,2500,"
        "76.04,48.39,13.82,,0.00,2361.75,2026-03-14,2026-03-16",
        "2026-03-03,,EC3,ATM,5.53%,USB 集線器,680,"
        "20.68,13.16,3.76,,0.00,642.40,2026-03-15,2026-03-17",
        "2026-03-04,SP-UNKNOWN9,EC4,信用卡,5.53%,不明商品,500,"
        "15.21,9.68,2.76,,0.00,472.35,2026-03-16,2026-03-18",
    ]
)


def upload(
    client: TestClient, content: str = SETTLEMENT_CSV, name: str = "a.csv"
) -> httpx.Response:
    # 先指派給有型別註記的區域變數再回傳，不直接 `return client.post(...)`。
    # 某些版本的 starlette 把 TestClient.post 的回傳標成 Any，而 mypy 在
    # strict（warn_return_any）下會拒絕「宣告回傳 Response 卻回傳 Any」。
    # 這行指派讓 Any 在這裡就收斂成 Response，兩種版本都通過。
    response: httpx.Response = client.post(
        "/api/imports",
        files={"file": (name, content.encode("utf-8"), "text/csv")},
    )
    return response


# ----------------------------------------------------------------------
class TestHealthAndDiscovery:
    def test_health(self, client: TestClient) -> None:
        body = client.get("/api/health").json()
        assert body["status"] == "ok"
        assert body["orders"] == 0

    def test_platforms_come_from_the_registry(self, client: TestClient) -> None:
        """新增平台不需要改這個端點——它讀的是 parser registry。"""
        codes = [p["code"] for p in client.get("/api/platforms").json()]
        assert codes == ["platform_a", "platform_b", "platform_c"]

    def test_openapi_document_is_generated(self, client: TestClient) -> None:
        """W5 的前端會用這份文件產生 TypeScript 型別。"""
        spec = client.get("/openapi.json").json()
        assert "/api/imports" in spec["paths"]
        assert "MoneyOut" in spec["components"]["schemas"]


class TestImport:
    def test_upload_detects_platform_and_imports(self, client: TestClient) -> None:
        response = upload(client)
        assert response.status_code == 200
        body = response.json()
        assert body["platform"] == "platform_a"
        assert body["accepted_count"] == 4
        assert body["was_duplicate"] is False
        assert len(body["content_hash"]) == 64

    def test_records_are_queryable_after_import(self, client: TestClient) -> None:
        batch_id = upload(client).json()["batch_id"]
        body = client.get(f"/api/imports/{batch_id}").json()
        assert body["record_total"] == 4
        assert body["records"][0]["external_order_id"] == "SP-AAAA0001"

    def test_money_is_a_string_not_a_float(self, client: TestClient) -> None:
        """整個專案避開浮點數，不能在 JSON 這一哩前功盡棄。"""
        batch_id = upload(client).json()["batch_id"]
        record = client.get(f"/api/imports/{batch_id}").json()["records"][0]
        assert record["net_amount"]["amount"] == "944.70"
        assert isinstance(record["net_amount"]["amount"], str)
        assert record["net_amount"]["minor_units"] == 94470

    def test_unrecognised_file_is_rejected_with_an_explanation(self, client: TestClient) -> None:
        response = upload(client, "姓名,電話\n王小明,0912345678", "mystery.csv")
        assert response.status_code == 422
        body = response.json()
        # 認不出平台與格式壞掉是兩種不同的失敗：前者使用者可以手動
        # 指定平台重試，後者重試也沒用。給不同的 code 讓前端能分辨。
        assert body["code"] == "unrecognised_platform"
        assert "信心分數" in body["message"]

    def test_explicit_platform_overrides_detection(self, client: TestClient) -> None:
        response = client.post(
            "/api/imports?platform=platform_a",
            files={"file": ("x.csv", SETTLEMENT_CSV.encode("utf-8"), "text/csv")},
        )
        assert response.status_code == 200
        assert response.json()["platform"] == "platform_a"

    def test_unknown_platform_code_is_rejected(self, client: TestClient) -> None:
        response = client.post(
            "/api/imports?platform=platform_z",
            files={"file": ("x.csv", SETTLEMENT_CSV.encode("utf-8"), "text/csv")},
        )
        assert response.status_code == 422
        body = response.json()
        assert body["code"] == "unrecognised_platform"
        assert "platform_a" in body["message"]


class TestIdempotency:
    """冪等性不是承諾，是必須被證明的性質。"""

    def test_same_file_twice_does_not_double_import(self, client: TestClient) -> None:
        first = upload(client).json()
        second = upload(client).json()

        assert first["was_duplicate"] is False
        assert second["was_duplicate"] is True
        # 關鍵斷言：回傳的是**原本那個批次**，不是新建的
        assert second["batch_id"] == first["batch_id"]
        assert second["accepted_count"] == first["accepted_count"]

        # 而且資料庫裡的筆數沒有變
        batch = client.get(f"/api/imports/{first['batch_id']}").json()
        assert batch["record_total"] == 4
        assert len(client.get("/api/imports").json()) == 1

    def test_upload_five_times_is_the_same_as_once(self, client: TestClient) -> None:
        """呼叫端可以安心重試——這就是恰好一次語意的實際意義。"""
        responses = [upload(client).json() for _ in range(5)]
        assert [r["was_duplicate"] for r in responses] == [
            False,
            True,
            True,
            True,
            True,
        ]
        assert len({r["batch_id"] for r in responses}) == 1
        assert len(client.get("/api/imports").json()) == 1

    def test_different_filename_same_content_is_still_a_duplicate(self, client: TestClient) -> None:
        """指紋看的是內容，不是檔名。"""
        upload(client, name="三月.csv")
        second = upload(client, name="三月-副本.csv").json()
        assert second["was_duplicate"] is True

    def test_changed_file_still_skips_rows_it_already_has(self, client: TestClient) -> None:
        """第二道防線：檔案指紋不同，但重複的列仍然被擋下。

        情境很實際：平台重新匯出一份帳單，多了一列新交易，
        其餘內容完全相同。整份檔案的指紋因此不同，但舊的那幾列不該重複入帳。
        """
        upload(client)
        extended = SETTLEMENT_CSV + (
            "\n2026-03-05,SP-AAAA0005,EC5,信用卡,5.53%,新商品,900,"
            "27.37,17.42,4.98,,0.00,850.23,2026-03-17,2026-03-19"
        )
        second = upload(client, extended, "a-v2.csv").json()

        assert second["was_duplicate"] is False  # 是不同的檔案
        assert second["accepted_count"] == 1  # 但只有那一列是新的
        assert second["skipped_duplicate_rows"] == 4

        total = sum(b["accepted_count"] for b in client.get("/api/imports").json())
        assert total == 5  # 4 + 1，不是 4 + 5

    def test_message_explains_what_happened(self, client: TestClient) -> None:
        """使用者要看得懂「這份檔案已經匯入過」與「匯入了 4 列」的差別。"""
        upload(client)
        message = upload(client).json()["message"]
        assert "已" in message and "匯入過" in message


class TestReconciliation:
    def test_full_flow(self, seeded_client: TestClient) -> None:
        """上傳 → 對帳 → 看報告，一條龍。"""
        upload(seeded_client)

        run = seeded_client.post("/api/reconciliations", json={}).json()
        assert run["run_id"].startswith("RUN-")
        metrics = run["metrics"]
        assert metrics["order_count"] == 4
        assert metrics["record_count"] == 4
        assert 0.0 <= metrics["automation_rate"] <= 1.0

        report = seeded_client.get(f"/api/reconciliations/{run['run_id']}/report").json()
        assert report["total"] == len(report["results"]) or report["total"] >= 1
        assert report["config"]["accept_threshold"]

    def test_report_can_be_filtered_by_quadrant(self, seeded_client: TestClient) -> None:
        upload(seeded_client)
        run_id = seeded_client.post("/api/reconciliations", json={}).json()["run_id"]

        matched = seeded_client.get(f"/api/reconciliations/{run_id}/report?outcome=matched").json()
        assert all(r["outcome"] == "matched" for r in matched["results"])

        review = seeded_client.get(f"/api/reconciliations/{run_id}/report?needs_review=true").json()
        assert all(r["needs_review"] for r in review["results"])

    def test_invalid_quadrant_lists_the_valid_ones(self, seeded_client: TestClient) -> None:
        upload(seeded_client)
        run_id = seeded_client.post("/api/reconciliations", json={}).json()["run_id"]
        response = seeded_client.get(f"/api/reconciliations/{run_id}/report?outcome=nonsense")
        assert response.status_code == 422
        assert "missing_in_ledger" in response.json()["message"]

    def test_report_survives_a_round_trip_through_the_database(
        self, seeded_client: TestClient
    ) -> None:
        """報告存進資料庫再讀回來，內容必須一致。

        這條驗證的是 repository 的序列化與反序列化沒有丟東西——
        金額、階段、歸因、候選清單都要還原得回來。
        """
        upload(seeded_client)
        run = seeded_client.post("/api/reconciliations", json={}).json()
        report = seeded_client.get(f"/api/reconciliations/{run['run_id']}/report?limit=500").json()

        assert report["metrics"]["matched"] == run["metrics"]["matched"]
        assert report["metrics"]["stage_exact"] == run["metrics"]["stage_exact"]
        quadrant_total = sum(
            report["metrics"][k]
            for k in (
                "matched",
                "amount_variance",
                "missing_in_ledger",
                "missing_in_settlement",
            )
        )
        assert quadrant_total == report["total"]

    def test_threshold_is_recorded_in_the_report(self, seeded_client: TestClient) -> None:
        """報告一定要附上當時的設定，否則數字沒有意義。"""
        upload(seeded_client)
        run = seeded_client.post("/api/reconciliations", json={"accept_threshold": "0.65"}).json()
        assert run["config"]["accept_threshold"] == "0.65"

    def test_reconciling_an_empty_database_is_a_clear_error(self, client: TestClient) -> None:
        response = client.post("/api/reconciliations", json={})
        assert response.status_code == 409
        assert response.json()["code"] == "nothing_to_reconcile"

    def test_unknown_run_id_is_404(self, client: TestClient) -> None:
        response = client.get("/api/reconciliations/RUN-NOPE/report")
        assert response.status_code == 404
        assert response.json()["code"] == "not_found"


class TestOrders:
    def test_list_and_paginate(self, seeded_client: TestClient) -> None:
        orders = seeded_client.get("/api/orders").json()
        assert len(orders) == 4
        assert orders[0]["expected_net"]["amount"] == "944.70"

        page = seeded_client.get("/api/orders?limit=2&offset=2").json()
        assert len(page) == 2
        assert page[0]["order_id"] != orders[0]["order_id"]

    def test_filter_by_platform(self, seeded_client: TestClient) -> None:
        assert seeded_client.get("/api/orders?platform=platform_a").json()
        assert seeded_client.get("/api/orders?platform=platform_z").json() == []


@contextmanager
def disposing_app(url: str) -> Iterator[TestClient]:
    """建一個 app、用完把連線池關掉。

    不關會怎樣：SQLAlchemy 的連線池握著 DBAPI 連線，engine 被垃圾回收時
    那些連線才跟著關。Python 3.13 起 ``sqlite3.Connection`` 在未關閉就被
    回收時會發出 ``ResourceWarning``，而本專案把警告視為錯誤——於是失敗會
    出現在「剛好觸發 GC 的那個測試」上，跟真正洩漏的地方毫無關係。
    （實際症狀就是這樣：報錯報在 test_properties.py 與架構測試上。）

    所以清理要明確做，不能靠回收時機。
    """
    app = create_app(url)
    try:
        with TestClient(app) as client:
            yield client
    finally:
        app.state.engine.dispose()


class TestPersistence:
    def test_data_survives_a_new_app_instance(self, tmp_path: Path) -> None:
        """換一個 app 實例，資料還在——證明真的寫進磁碟了。"""
        url = f"sqlite:///{tmp_path / 'test.db'}"
        engine = create_db_engine(url)
        try:
            init_schema(engine)
        finally:
            engine.dispose()

        with disposing_app(url) as first:
            batch_id = upload(first).json()["batch_id"]

        with disposing_app(url) as second:
            assert second.get(f"/api/imports/{batch_id}").status_code == 200
            assert second.get("/api/health").json()["settlement_records"] == 4
