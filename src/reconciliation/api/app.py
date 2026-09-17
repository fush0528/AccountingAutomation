"""FastAPI 應用程式。

這一層刻意寫得很薄：驗證輸入、呼叫服務、把領域物件轉成 DTO。
沒有任何業務規則——那些全部在 ``domain/`` 與 ``services/``。

可以這樣驗證這句話：整個檔案裡沒有一個 if 是在判斷對帳邏輯，
只有在判斷 HTTP 狀態碼。

啟動::

    uvicorn reconciliation.api.app:app --reload

然後開 http://127.0.0.1:8000/docs 看自動產生的 OpenAPI 文件。
"""

from __future__ import annotations

import logging
from collections.abc import Iterator
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Annotated

from fastapi import Depends, FastAPI, File, Query, Request, UploadFile, status
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from sqlalchemy.orm import Session, sessionmaker

from ..domain.matching import MatchingConfig
from ..domain.models import MatchOutcome
from ..parsers import ParseError, SettlementSource, available_platforms
from ..repositories.database import create_db_engine, init_schema
from ..repositories.sqlalchemy_repo import (
    SqlAlchemyImportBatchRepository,
    SqlAlchemyOrderRepository,
    SqlAlchemySettlementRepository,
)
from ..services import (
    ImportFailed,
    ImportService,
    NothingToReconcile,
    ReconciliationService,
)
from .schemas import (
    BatchOut,
    ErrorResponse,
    ImportResponse,
    OrderOut,
    PlatformOut,
    ReconciliationRequest,
    ReconciliationResponse,
    ReportResponse,
    SettlementRecordOut,
)

__all__ = ["app", "create_app"]

logger = logging.getLogger(__name__)

#: 前端 build 的產物。由 ``frontend/`` 執行 ``npm run build`` 產生——
#: vite.config.ts 直接把 outDir 指到這裡，省掉一個複製步驟。
#: 目錄不存在時（還沒 build 過）API 照常運作，只是沒有網頁介面。
STATIC_DIR = Path(__file__).parent / "static"


def get_session(request: Request) -> Iterator[Session]:
    """每個請求一個 session，也就是一個交易範圍。

    請求成功就 commit，拋例外就 rollback。這是冪等匯入的第三道防線：
    中途失敗不會留下半批資料。

    session factory 從 ``app.state`` 取，而不是從模組層級的全域變數——
    這樣同一個程序裡可以有多個指向不同資料庫的 app 實例，
    測試因此能各自使用獨立的記憶體資料庫，互不干擾。
    """
    factory: sessionmaker[Session] = request.app.state.session_factory
    session = factory()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


#: 這個別名必須定義在模組層級。放在 create_app() 裡面會變成區域變數，
#: 而 `from __future__ import annotations` 會把型別註解延後成字串，
#: FastAPI 解析時找不到那個名字。
SessionDep = Annotated[Session, Depends(get_session)]

#: Starlette 在版本之間把這個常數改過名（``HTTP_422_UNPROCESSABLE_ENTITY``
#: → ``HTTP_422_UNPROCESSABLE_CONTENT``），舊名稱現在會發出 DeprecationWarning。
#: 用字面值就不必相依於特定版本——狀態碼本身是 HTTP 標準，不會變。
HTTP_422_UNPROCESSABLE = 422

DESCRIPTION = """
把電商賣家每月的手動對帳，從逐筆核對壓縮成一次上傳。

**典型流程**

1. `POST /api/imports` 上傳結算單（平台自動偵測）
2. `POST /api/reconciliations` 執行對帳
3. `GET /api/reconciliations/{run_id}/report` 看四象限報告

**冪等匯入**：同一份檔案重複上傳不會重複入帳。回應的 `was_duplicate`
為真時代表這份檔案先前已匯入過，本次未寫入任何資料。呼叫端可以安心重試。

**金額一律以字串傳輸**，不用 JSON number。JSON 的 number 是 IEEE 754
雙精度浮點數，無法精確表示十進位小數。
"""


def create_app(database_url: str | None = None) -> FastAPI:
    """建立應用程式。

    接受 ``database_url`` 參數而不是只讀環境變數，是為了讓測試可以
    指向記憶體資料庫。這種「可注入的設定」讓 API 的整合測試不需要
    任何外部狀態，跑一輪不到一秒。
    """
    engine = create_db_engine(database_url)
    init_schema(engine)
    factory = sessionmaker(bind=engine, expire_on_commit=False)

    application = FastAPI(
        title="多平台電商對帳引擎",
        description=DESCRIPTION,
        version="0.5.0",
        openapi_tags=[
            {"name": "匯入", "description": "上傳與查詢結算單批次"},
            {"name": "對帳", "description": "執行對帳與查看報告"},
            {"name": "資料", "description": "查詢訂單與結算記錄"},
        ],
    )
    application.state.engine = engine
    application.state.session_factory = factory

    # ------------------------------------------------------------------
    # 全域錯誤處理：讓所有錯誤都長同一個樣子
    # ------------------------------------------------------------------
    @application.exception_handler(ParseError)
    async def _parse_error(_: Request, exc: ParseError) -> JSONResponse:
        return JSONResponse(
            status_code=HTTP_422_UNPROCESSABLE,
            content=ErrorResponse(code="parse_error", message=str(exc)).model_dump(),
        )

    @application.exception_handler(ImportFailed)
    async def _import_failed(_: Request, exc: ImportFailed) -> JSONResponse:
        return JSONResponse(
            status_code=HTTP_422_UNPROCESSABLE,
            content=ErrorResponse(code=exc.code, message=str(exc)).model_dump(),
        )

    @application.exception_handler(NothingToReconcile)
    async def _nothing(_: Request, exc: NothingToReconcile) -> JSONResponse:
        return JSONResponse(
            status_code=status.HTTP_409_CONFLICT,
            content=ErrorResponse(code="nothing_to_reconcile", message=str(exc)).model_dump(),
        )

    def not_found(what: str) -> JSONResponse:
        return JSONResponse(
            status_code=status.HTTP_404_NOT_FOUND,
            content=ErrorResponse(code="not_found", message=what).model_dump(),
        )

    # ------------------------------------------------------------------
    # 匯入
    # ------------------------------------------------------------------
    @application.post(
        "/api/imports",
        tags=["匯入"],
        response_model=ImportResponse,
        summary="上傳結算單",
        responses={422: {"model": ErrorResponse}},
    )
    async def create_import(
        session: SessionDep,
        file: Annotated[UploadFile, File(description="平台匯出的結算單")],
        platform: Annotated[str | None, Query(description="留空則由內容自動偵測平台")] = None,
    ) -> ImportResponse:
        """上傳並匯入一份結算單。

        **同一份檔案重複上傳是安全的。** 服務端以檔案內容的 SHA-256 判斷，
        第二次上傳會回傳原批次且 `was_duplicate` 為真，不會重複入帳。
        """
        content = await file.read()
        source = SettlementSource(filename=file.filename or "upload", content=content)
        outcome = ImportService(session).import_file(source, platform=platform)
        return ImportResponse.of(outcome)

    @application.get(
        "/api/imports",
        tags=["匯入"],
        response_model=list[BatchOut],
        summary="列出所有匯入批次",
    )
    def list_imports(session: SessionDep) -> list[BatchOut]:
        return [BatchOut.of(b) for b in SqlAlchemyImportBatchRepository(session).list_all()]

    @application.get(
        "/api/imports/{batch_id}",
        tags=["匯入"],
        summary="查看單一批次與它匯入的記錄",
        responses={404: {"model": ErrorResponse}},
    )
    def get_import(
        session: SessionDep, batch_id: str, limit: Annotated[int, Query(le=500)] = 50
    ) -> JSONResponse:
        batch = SqlAlchemyImportBatchRepository(session).get(batch_id)
        if batch is None:
            return not_found(f"查無批次 {batch_id}")
        records = SqlAlchemySettlementRepository(session).list_by_batch(batch_id)
        return JSONResponse(
            content={
                "batch": BatchOut.of(batch).model_dump(mode="json"),
                "records": [
                    SettlementRecordOut.of(r).model_dump(mode="json") for r in records[:limit]
                ],
                "record_total": len(records),
            }
        )

    # ------------------------------------------------------------------
    # 對帳
    # ------------------------------------------------------------------
    @application.post(
        "/api/reconciliations",
        tags=["對帳"],
        response_model=ReconciliationResponse,
        summary="執行一次對帳",
        responses={409: {"model": ErrorResponse}, 422: {"model": ErrorResponse}},
    )
    def create_reconciliation(
        session: SessionDep, body: ReconciliationRequest | None = None
    ) -> ReconciliationResponse:
        """把資料庫裡的訂單與結算記錄跑一次三階段匹配。

        回傳指標與 `run_id`；完整的四象限明細請用
        `GET /api/reconciliations/{run_id}/report`。
        """
        body = body or ReconciliationRequest()
        config = None
        if body.accept_threshold is not None:
            try:
                config = MatchingConfig(accept_threshold=Decimal(body.accept_threshold))
            except (InvalidOperation, ValueError) as exc:
                raise ImportFailed(f"accept_threshold 不合法：{exc}") from exc

        outcome = ReconciliationService(session).run(platform=body.platform, config=config)
        response = ReportResponse.of(
            outcome.run_id, outcome.report, [], len(outcome.report.results)
        )
        return ReconciliationResponse(
            run_id=outcome.run_id, metrics=response.metrics, config=response.config
        )

    @application.get(
        "/api/reconciliations",
        tags=["對帳"],
        response_model=list[str],
        summary="列出歷次對帳的 run_id",
    )
    def list_reconciliations(session: SessionDep) -> list[str]:
        return ReconciliationService(session).list_run_ids()

    @application.get(
        "/api/reconciliations/{run_id}/report",
        tags=["對帳"],
        response_model=ReportResponse,
        summary="查看四象限報告",
        responses={404: {"model": ErrorResponse}},
    )
    def get_report(
        session: SessionDep,
        run_id: str,
        outcome: Annotated[
            str | None,
            Query(
                description="依象限篩選：matched / amount_variance / "
                "missing_in_ledger / missing_in_settlement"
            ),
        ] = None,
        needs_review: Annotated[bool | None, Query(description="只看需要人工確認的")] = None,
        limit: Annotated[int, Query(ge=1, le=500)] = 50,
        offset: Annotated[int, Query(ge=0)] = 0,
    ) -> ReportResponse | JSONResponse:
        report = ReconciliationService(session).get(run_id)
        if report is None:
            return not_found(f"查無對帳執行 {run_id}")

        results = list(report.results)
        if outcome:
            try:
                wanted = MatchOutcome(outcome)
            except ValueError:
                return JSONResponse(
                    status_code=HTTP_422_UNPROCESSABLE,
                    content=ErrorResponse(
                        code="invalid_outcome",
                        message=f"outcome 必須是四象限之一，收到 {outcome!r}。"
                        f"可用值：{', '.join(o.value for o in MatchOutcome)}",
                    ).model_dump(),
                )
            results = [r for r in results if r.outcome is wanted]
        if needs_review is not None:
            results = [r for r in results if r.needs_review is needs_review]

        total = len(results)
        return ReportResponse.of(run_id, report, results[offset : offset + limit], total)

    # ------------------------------------------------------------------
    # 資料查詢
    # ------------------------------------------------------------------
    @application.get(
        "/api/orders",
        tags=["資料"],
        response_model=list[OrderOut],
        summary="列出我方訂單",
    )
    def list_orders(
        session: SessionDep,
        platform: str | None = None,
        limit: Annotated[int, Query(ge=1, le=500)] = 50,
        offset: Annotated[int, Query(ge=0)] = 0,
    ) -> list[OrderOut]:
        orders = SqlAlchemyOrderRepository(session).list_all(platform)
        return [OrderOut.of(o) for o in orders[offset : offset + limit]]

    @application.get(
        "/api/platforms",
        tags=["資料"],
        response_model=list[PlatformOut],
        summary="列出已支援的平台",
    )
    def list_platforms() -> list[PlatformOut]:
        """新增一個平台不需要修改這個端點——它讀的是 parser registry。"""
        return [PlatformOut(code=code, display_name=name) for code, name in available_platforms()]

    @application.get("/api/health", tags=["資料"], summary="健康檢查")
    def health(session: SessionDep) -> dict[str, object]:
        return {
            "status": "ok",
            "orders": SqlAlchemyOrderRepository(session).count(),
            "settlement_records": SqlAlchemySettlementRepository(session).count(),
            "platforms": [code for code, _ in available_platforms()],
        }

    # ------------------------------------------------------------------
    # 前端（放在最後掛載）
    # ------------------------------------------------------------------
    if STATIC_DIR.is_dir():
        assets = STATIC_DIR / "assets"
        if assets.is_dir():
            application.mount("/assets", StaticFiles(directory=assets), name="assets")

        @application.get("/", include_in_schema=False)
        def index() -> FileResponse:
            return FileResponse(STATIC_DIR / "index.html")

        # response_model=None：回傳型別是兩種 Response 的聯集，
        # FastAPI 會試著把它當成 Pydantic 模型去推導 schema 而失敗。
        # 這個端點本來就不進 OpenAPI，直接關掉推導。
        @application.get("/{path:path}", include_in_schema=False, response_model=None)
        def spa_fallback(path: str) -> JSONResponse | FileResponse:
            """把未知路徑交給前端處理。

            這條 catch-all 必須是**最後**註冊的，否則會蓋掉所有 API 路由。
            而且要明確排除 /api 開頭的路徑——不然打錯的 API 網址會回傳
            一頁 HTML，前端拿到之後解析 JSON 失敗，錯誤訊息會完全誤導人。
            寧可老實回 404。
            """
            if path.startswith("api/"):
                return JSONResponse(
                    status_code=status.HTTP_404_NOT_FOUND,
                    content=ErrorResponse(
                        code="not_found", message=f"沒有這個端點：/{path}"
                    ).model_dump(),
                )
            return FileResponse(STATIC_DIR / "index.html")

        _ = (index, spa_fallback)

    # 讓型別檢查器知道這些端點有被使用
    _ = (
        create_import,
        list_imports,
        get_import,
        create_reconciliation,
        list_reconciliations,
        get_report,
        list_orders,
        list_platforms,
        health,
        _parse_error,
        _import_failed,
        _nothing,
    )
    return application


#: 給 uvicorn 用的預設實例。測試會自己呼叫 create_app() 指向記憶體資料庫。
app = create_app()
