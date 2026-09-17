"""帳單匯入服務：把一份結算單變成資料庫裡的記錄。

冪等性的三道防線
----------------

上傳是最容易被重複觸發的操作：使用者點兩次、網路逾時後重試、
前端沒擋住連擊。同一份帳單被匯入兩次，對帳結果就會完全錯亂——
金額翻倍、假的重複交易、對不起來的總額。

所以這裡用三道互相獨立的防線：

1. **檔案內容的 SHA-256。** 整份檔案的指紋，存在
   ``import_batches.content_hash`` 並下唯一約束。同一份檔案再上傳一次，
   服務層查到既有批次就直接回傳，一列都不寫。
2. **每列的 fingerprint。** 就算檔案被改了一個字（指紋因此不同），
   裡面重複的那些列仍然會被 ``settlement_records.row_fingerprint``
   的唯一約束擋下。
3. **整批包在單一交易裡。** 中途失敗就整批回滾，不會留下半批資料
   讓下一次匯入變成「部分重複」。

為什麼不是「先查有沒有，沒有再寫」
----------------------------------

因為那在併發下會失效：兩個請求可能同時查到「沒有」，然後雙雙寫入。
唯一約束是由資料庫保證的真正不變量，應用層要做的是**處理被拒絕的情況**，
而不是試圖預先避免它。這就是為什麼 :meth:`ImportService.import_file`
是先嘗試寫入、再捕捉 ``IntegrityError`，而不是反過來。

這在分散式系統裡叫「至少一次投遞下的恰好一次語意」——
呼叫端可以放心重試，結果不會變。
"""

from __future__ import annotations

import logging
import uuid
from dataclasses import dataclass
from datetime import datetime

from sqlalchemy import select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from ..domain.models import ImportBatch
from ..parsers import ParseError, SettlementSource, detect, get_parser
from ..parsers.base import RowError
from ..repositories.mappers import record_to_row, row_fingerprint, row_to_batch
from ..repositories.tables import ImportBatchRow, SettlementRecordRow

__all__ = ["ImportFailed", "ImportOutcome", "ImportService"]

logger = logging.getLogger(__name__)


class ImportFailed(Exception):
    """整份檔案無法匯入。

    帶 ``code`` 是為了讓前端能分辨兩種完全不同的失敗：認不出平台
    （使用者可以手動指定平台重試）與格式不符（重試也沒用，要換檔案）。
    只給一段中文訊息的話，前端只能比對字串來決定行為——那很脆弱。
    """

    def __init__(self, message: str, *, code: str = "import_failed") -> None:
        super().__init__(message)
        self.code = code


@dataclass(frozen=True, slots=True)
class ImportOutcome:
    """一次匯入的結果。

    ``was_duplicate`` 是這個型別存在的理由：呼叫端必須能分辨
    「匯入成功」與「這份檔案先前已經匯入過，回傳的是原批次」。
    兩者都是成功，但意義完全不同——前者該顯示「已匯入 305 列」，
    後者該顯示「這份檔案已於 3 月 15 日匯入過」。
    """

    batch: ImportBatch
    was_duplicate: bool
    detected_platform: str
    row_errors: tuple[RowError, ...] = ()
    skipped_duplicate_rows: int = 0

    @property
    def message(self) -> str:
        if self.was_duplicate:
            return (
                f"這份檔案先前已於 {self.batch.imported_at:%Y-%m-%d %H:%M} 匯入過"
                f"（批次 {self.batch.batch_id}），未重複寫入任何資料。"
            )
        parts = [f"匯入 {self.batch.accepted_count} 列"]
        if self.skipped_duplicate_rows:
            parts.append(f"略過 {self.skipped_duplicate_rows} 列重複")
        if self.batch.rejected_count:
            parts.append(f"{self.batch.rejected_count} 列解析失敗")
        return "，".join(parts) + "。"


class ImportService:
    """把結算單匯入資料庫。

    收 :class:`~reconciliation.parsers.base.SettlementSource`（檔名加位元組），
    不收路徑——因為 API 拿到的就是上傳的位元組，沒有路徑可言。
    """

    def __init__(self, session: Session) -> None:
        self._session = session

    def import_file(
        self, source: SettlementSource, *, platform: str | None = None
    ) -> ImportOutcome:
        """匯入一份結算單。

        ``platform`` 留空時自動偵測。重複上傳同一份檔案會回傳原批次，
        且 ``was_duplicate`` 為真。
        """
        # --- 第一道防線：檔案內容指紋 --------------------------------
        existing = self._session.scalar(
            select(ImportBatchRow).where(ImportBatchRow.content_hash == source.sha256)
        )
        if existing is not None:
            logger.info(
                "偵測到重複上傳",
                extra={"filename": source.filename, "batch_id": existing.batch_id},
            )
            return ImportOutcome(
                batch=row_to_batch(existing),
                was_duplicate=True,
                detected_platform=existing.platform,
            )

        # --- 解析 -----------------------------------------------------
        try:
            parser = get_parser(platform) if platform else detect(source)
        except ParseError as exc:
            raise ImportFailed(str(exc), code="unrecognised_platform") from exc

        try:
            result = parser.parse(source)
        except ParseError as exc:
            raise ImportFailed(str(exc), code="parse_error") from exc

        # --- 寫入 -----------------------------------------------------
        batch_row = ImportBatchRow(
            batch_id=f"BATCH-{uuid.uuid4().hex[:12].upper()}",
            platform=result.platform,
            filename=source.filename,
            content_hash=source.sha256,
            imported_at=datetime.now(),
            row_count=result.row_count,
            accepted_count=0,
            rejected_count=result.rejected_count,
        )
        self._session.add(batch_row)

        try:
            self._session.flush()
        except IntegrityError as exc:
            # 併發的另一個請求剛好搶先寫入了同一份檔案。
            # 這不是錯誤，是冪等性正常運作——回滾後把它的批次撈出來回傳。
            self._session.rollback()
            duplicate = self._session.scalar(
                select(ImportBatchRow).where(ImportBatchRow.content_hash == source.sha256)
            )
            if duplicate is None:  # pragma: no cover - 理論上不會發生
                raise ImportFailed(f"寫入批次失敗：{exc}") from exc
            return ImportOutcome(
                batch=row_to_batch(duplicate),
                was_duplicate=True,
                detected_platform=duplicate.platform,
            )

        # --- 第二道防線：每列的 fingerprint ---------------------------
        #
        # 先查出已存在的指紋一次性過濾，而不是逐列 try/except。
        # 逐列捕捉 IntegrityError 在多數資料庫上會讓整個交易進入
        # aborted 狀態（PostgreSQL 尤其如此），後續寫入全部失敗。
        # 唯一約束仍然在，它負責擋住這個查詢之後才出現的併發寫入。
        fingerprints = [row_fingerprint(r) for r in result.records]
        already = set(
            self._session.scalars(
                select(SettlementRecordRow.row_fingerprint).where(
                    SettlementRecordRow.row_fingerprint.in_(fingerprints)
                )
            ).all()
        )

        accepted = 0
        skipped = 0
        seen_in_this_file: set[str] = set()
        for record, fingerprint in zip(result.records, fingerprints, strict=True):
            if fingerprint in already or fingerprint in seen_in_this_file:
                skipped += 1
                continue
            seen_in_this_file.add(fingerprint)
            self._session.add(record_to_row(record, batch_row.id))
            accepted += 1

        batch_row.accepted_count = accepted
        self._session.flush()

        logger.info(
            "匯入完成",
            extra={
                "batch_id": batch_row.batch_id,
                "platform": result.platform,
                "accepted": accepted,
                "skipped": skipped,
                "rejected": result.rejected_count,
            },
        )

        return ImportOutcome(
            batch=row_to_batch(batch_row),
            was_duplicate=False,
            detected_platform=result.platform,
            row_errors=result.errors,
            skipped_duplicate_rows=skipped,
        )
