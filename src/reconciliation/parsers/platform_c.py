"""平台 C：Big5 編碼 CSV，民國年日期，金額帶千分位。

這個 parser 存在的意義是證明抽象層真的擋得住差異。跟平台 A 比，
它的每一個維度都不一樣：

============  ==================  ==========================
維度          平台 A              平台 C
============  ==================  ==========================
編碼          UTF-8               Big5
日期          ``2026-03-15``      ``115/03/15``（民國年）
金額          ``1000``            ``"1,000"``（千分位）
負數          ``-500``            ``(500)``（會計括號）
欄位名稱      訂單編號            訂單序號
交易別        銷售／退款          銷貨／退貨／調整
============  ==================  ==========================

但 :meth:`parse` 回傳的仍然是同一種 :class:`SettlementRecord`，
下游的對帳引擎完全不知道這些差異存在。

真實檔案大概長這樣（Big5 編碼）::

    訂單序號,交易別,品名,數量,銷售金額,服務費,實撥金額,撥款日
    PC-24C9Z2N4,銷貨,USB集線器,1,"1,000","55.30","944.70",115/03/15
"""

from __future__ import annotations

import csv
import io
from typing import ClassVar

from ..domain.models import SettlementRecord, SettlementRecordType
from ..domain.money import Money, MoneyError
from .base import (
    BaseSettlementParser,
    ParseError,
    ParseResult,
    RowError,
    SettlementSource,
)
from .fields import clean_cell, parse_roc_date
from .registry import register

COL_ORDER_ID = "訂單序號"
COL_TYPE = "交易別"
COL_PRODUCT = "品名"
COL_QUANTITY = "數量"
COL_GROSS = "銷售金額"
COL_FEE = "服務費"
COL_NET = "實撥金額"
COL_SETTLED_AT = "撥款日"

SIGNATURE_COLUMNS = (COL_ORDER_ID, COL_TYPE, COL_GROSS, COL_NET, COL_SETTLED_AT)

_TYPE_MAP = {
    "銷貨": SettlementRecordType.SALE,
    "退貨": SettlementRecordType.REFUND,
    "調整": SettlementRecordType.ADJUSTMENT,
}


@register
class PlatformCParser(BaseSettlementParser):
    """平台 C 的結算單解析器。"""

    platform: ClassVar[str] = "platform_c"
    display_name: ClassVar[str] = "平台 C（Big5／民國年）"
    extensions: ClassVar[tuple[str, ...]] = (".csv", ".txt")

    #: 順序有意義：Big5 先試。很多 Big5 檔案用 UTF-8 解碼會「成功」
    #: 但得到亂碼，反過來則會直接失敗，所以把嚴格的放前面。
    ENCODINGS: ClassVar[tuple[str, ...]] = ("big5", "cp950", "utf-8-sig", "utf-8")

    @classmethod
    def sniff(cls, source: SettlementSource) -> float:
        try:
            text, _ = source.decode(*cls.ENCODINGS)
        except ParseError:
            return 0.0
        first_line = text.split("\n", 1)[0]
        header = next(csv.reader([first_line]), [])
        return cls._header_score(header, SIGNATURE_COLUMNS)

    def parse(self, source: SettlementSource) -> ParseResult:
        text, encoding = source.decode(*self.ENCODINGS)
        reader = csv.DictReader(io.StringIO(text))
        if reader.fieldnames is None:
            raise ParseError(f"{source.filename} 沒有表頭列")

        missing = [c for c in SIGNATURE_COLUMNS if c not in reader.fieldnames]
        if missing:
            raise ParseError(f"{source.filename} 缺少必要欄位：{'、'.join(missing)}")

        records: list[SettlementRecord] = []
        errors: list[RowError] = []
        row_count = 0

        for row_number, row in enumerate(reader, start=2):
            row_count += 1
            clean = {k: clean_cell(v) for k, v in row.items() if k is not None}
            try:
                records.append(self._build(clean, row_number))
            except (ParseError, MoneyError, ValueError) as exc:
                errors.append(RowError(row_number, str(exc), clean))

        return ParseResult(
            platform=self.platform,
            records=tuple(records),
            errors=tuple(errors),
            row_count=row_count,
            encoding=encoding,
        )

    # ------------------------------------------------------------------
    def _build(self, row: dict[str, str], row_number: int) -> SettlementRecord:
        record_type = _TYPE_MAP.get(row.get(COL_TYPE, ""), SettlementRecordType.SALE)

        # 千分位與會計括號都由 Money.from_str 處理，parser 不重複實作
        gross = Money.from_str(row[COL_GROSS])
        fee = Money.from_str(row[COL_FEE])
        net = Money.from_str(row[COL_NET])

        # 平台 C 的退貨列金額印成正數，靠「交易別」欄位表達方向。
        # 統一成負數，讓下游不必再判斷平台。
        if record_type is SettlementRecordType.REFUND and not gross.is_negative:
            gross, fee, net = -gross, -fee, -net

        return SettlementRecord(
            platform=self.platform,
            settled_at=parse_roc_date(row[COL_SETTLED_AT], field=COL_SETTLED_AT),
            gross_amount=gross,
            fee_amount=fee,
            net_amount=net,
            record_type=record_type,
            external_order_id=row.get(COL_ORDER_ID, "") or None,
            product_name=row.get(COL_PRODUCT) or None,
            source_row=row_number,
            raw=row,
        )
