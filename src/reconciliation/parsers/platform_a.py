"""平台 A：標準 UTF-8 CSV，單層表頭。

這是三個 parser 裡最單純的一個，當作參考實作來讀。
真實世界的平台 A 大概長這樣::

    訂單編號,訂單成立日,商品名稱,數量,訂單金額,平台手續費,撥款金額,結算日期,交易類型
    SP-24A7X9K2,2026-03-01,無線滑鼠,1,1000,55.30,944.70,2026-03-15,銷售
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
from .fields import clean_cell, parse_date
from .registry import register

COL_ORDER_ID = "訂單編號"
COL_ORDERED_AT = "訂單成立日"
COL_PRODUCT = "商品名稱"
COL_QUANTITY = "數量"
COL_GROSS = "訂單金額"
COL_FEE = "平台手續費"
COL_NET = "撥款金額"
COL_SETTLED_AT = "結算日期"
COL_TYPE = "交易類型"

#: 用來認出這個平台的特徵欄位
SIGNATURE_COLUMNS = (COL_ORDER_ID, COL_GROSS, COL_FEE, COL_NET, COL_SETTLED_AT)

_TYPE_MAP = {
    "銷售": SettlementRecordType.SALE,
    "退款": SettlementRecordType.REFUND,
    "調整": SettlementRecordType.ADJUSTMENT,
}


@register
class PlatformAParser(BaseSettlementParser):
    """平台 A 的結算單解析器。"""

    platform: ClassVar[str] = "platform_a"
    display_name: ClassVar[str] = "平台 A（標準 CSV）"
    extensions: ClassVar[tuple[str, ...]] = (".csv",)

    ENCODINGS: ClassVar[tuple[str, ...]] = ("utf-8-sig", "utf-8")

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

        for row_number, row in enumerate(reader, start=2):  # 第 1 列是表頭
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
        order_id = row.get(COL_ORDER_ID, "") or None

        # 訂單成立日這欄我們用不到（對帳看的是結算日），但存進 raw 供人工查核
        return SettlementRecord(
            platform=self.platform,
            settled_at=parse_date(row[COL_SETTLED_AT], field=COL_SETTLED_AT),
            gross_amount=Money.from_str(row[COL_GROSS]),
            fee_amount=Money.from_str(row[COL_FEE]),
            net_amount=Money.from_str(row[COL_NET]),
            record_type=record_type,
            external_order_id=order_id,
            product_name=row.get(COL_PRODUCT) or None,
            source_row=row_number,
            raw=row,
        )
