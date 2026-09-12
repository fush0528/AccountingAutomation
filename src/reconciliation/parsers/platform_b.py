"""平台 B：Excel 活頁簿，報表標題 + 兩層表頭 + 合併儲存格。

這是三個 parser 裡最麻煩的一個，也最接近真實世界——很多平台的「結算單」
其實是給人看的報表，不是給程式讀的資料檔。典型結構::

    第 1 列   平台 B 商城　賣家結算明細表          ← 報表標題
    第 2 列   結算期間：2026/03/01 - 2026/03/31    ← 期間說明
    第 3 列   (空白列)
    第 4 列   訂單資訊 |        |      | 金額明細 |      |      | 結算  ← 第一層（合併）
    第 5 列   訂單編號 | 商品   | 數量 | 訂單金額 | 手續費 | 撥款金額 | 結算日  ← 第二層
    第 6 列   起        資料

處理重點有三個：

1. **表頭不在第 1 列。** 不能寫死列號——報表標題的行數會變。
   這裡用「掃描前 N 列，找出包含特徵欄位最多的那一列」來定位表頭。
2. **合併儲存格讀出來是 None。** openpyxl 只在合併區域的左上角有值，
   其餘是 None。第一層表頭因此必須向右填補才能理解。
3. **數值不是字串。** Excel 的金額欄讀出來是 float 或 int，
   直接丟給 ``Money.from_str`` 會炸。要先經過 :func:`~.fields.clean_cell`。
"""

from __future__ import annotations

import io
from typing import Any, ClassVar

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
COL_PRODUCT = "商品"
COL_QUANTITY = "數量"
COL_GROSS = "訂單金額"
COL_FEE = "手續費"
COL_NET = "撥款金額"
COL_SETTLED_AT = "結算日"
COL_TYPE = "類別"

SIGNATURE_COLUMNS = (COL_ORDER_ID, COL_GROSS, COL_FEE, COL_NET, COL_SETTLED_AT)

#: 表頭最多可能被推到第幾列。超過就當作不是這個平台的檔案。
MAX_HEADER_SCAN = 12

_TYPE_MAP = {
    "銷售": SettlementRecordType.SALE,
    "退款": SettlementRecordType.REFUND,
    "折讓": SettlementRecordType.REFUND,
    "調整": SettlementRecordType.ADJUSTMENT,
}


@register
class PlatformBParser(BaseSettlementParser):
    """平台 B 的結算單解析器。"""

    platform: ClassVar[str] = "platform_b"
    display_name: ClassVar[str] = "平台 B（多層表頭 Excel）"
    extensions: ClassVar[tuple[str, ...]] = (".xlsx", ".xlsm")

    @classmethod
    def sniff(cls, source: SettlementSource) -> float:
        # .xlsx 是 zip，開頭一定是 PK。先擋掉非 Excel 檔，
        # 免得為了一個 CSV 去啟動 openpyxl。
        if not source.content.startswith(b"PK"):
            return 0.0
        try:
            rows = cls._load_rows(source)
        except Exception:  # noqa: BLE001 - sniff 不該拋例外
            return 0.0
        header_index = cls._find_header(rows)
        if header_index is None:
            return 0.0
        return cls._header_score(rows[header_index], SIGNATURE_COLUMNS)

    def parse(self, source: SettlementSource) -> ParseResult:
        rows = self._load_rows(source)
        header_index = self._find_header(rows)
        if header_index is None:
            raise ParseError(
                f"{source.filename} 的前 {MAX_HEADER_SCAN} 列裡找不到表頭。"
                f"需要包含這些欄位：{'、'.join(SIGNATURE_COLUMNS)}"
            )

        header = rows[header_index]
        missing = [c for c in SIGNATURE_COLUMNS if c not in header]
        if missing:
            raise ParseError(f"{source.filename} 缺少必要欄位：{'、'.join(missing)}")

        records: list[SettlementRecord] = []
        errors: list[RowError] = []
        row_count = 0

        # Excel 的列號從 1 開始，header_index 是 0-based
        for offset, values in enumerate(rows[header_index + 1 :], start=1):
            row_number = header_index + 1 + offset
            row = dict(zip(header, values, strict=False))
            if self._is_not_data(row):
                continue
            row_count += 1
            try:
                records.append(self._build(row, row_number))
            except (ParseError, MoneyError, ValueError) as exc:
                errors.append(RowError(row_number, str(exc), row))

        return ParseResult(
            platform=self.platform,
            records=tuple(records),
            errors=tuple(errors),
            row_count=row_count,
            encoding="xlsx",
        )

    # ------------------------------------------------------------------
    @staticmethod
    def _load_rows(source: SettlementSource) -> list[list[str]]:
        """把活頁簿的第一個工作表讀成純字串的二維陣列。

        ``data_only=True`` 讓公式儲存格回傳計算後的值而不是公式字串——
        平台匯出的報表常常整欄都是公式。
        """
        try:
            from openpyxl import load_workbook
        except ImportError as exc:  # pragma: no cover
            raise ParseError("解析 Excel 結算單需要 openpyxl，請執行 pip install openpyxl") from exc

        workbook = load_workbook(io.BytesIO(source.content), read_only=True, data_only=True)
        try:
            worksheet = workbook.worksheets[0]
            rows: list[list[str]] = []
            for excel_row in worksheet.iter_rows(values_only=True):
                rows.append([clean_cell(cell) for cell in excel_row])
            return rows
        finally:
            workbook.close()

    @staticmethod
    def _find_header(rows: list[list[str]]) -> int | None:
        """在前幾列裡找出最像表頭的那一列。

        判準是「命中特徵欄位最多的列」。不寫死列號，因為報表標題的
        行數會隨平台改版而變——寫死的話每次改版都要改程式。
        """
        best_index: int | None = None
        best_hits = 0
        for index, row in enumerate(rows[:MAX_HEADER_SCAN]):
            cells = {c for c in row if c}
            hits = sum(1 for name in SIGNATURE_COLUMNS if name in cells)
            if hits > best_hits:
                best_index, best_hits = index, hits
        # 至少要命中一半以上才算數，否則可能只是剛好有欄位同名的說明列
        return best_index if best_hits >= len(SIGNATURE_COLUMNS) / 2 else None

    @staticmethod
    def _is_not_data(row: dict[str, Any]) -> bool:
        """判斷這一列是不是資料列。

        報表式的結算單在資料下方常有空白列、合計列、以及「※ 本表僅供參考」
        這類附註。它們的訂單編號欄可能有字，但金額欄一定是空的——
        用這一點來辨識，比寫死列數可靠。

        這類列直接略過且不計入 row_count，因為它們不是「解析失敗的資料」，
        而是「本來就不是資料」。混為一談會讓成功率的數字失去意義。
        """
        return not any(clean_cell(row.get(c)) for c in (COL_GROSS, COL_FEE, COL_NET))

    def _build(self, row: dict[str, Any], row_number: int) -> SettlementRecord:
        record_type = _TYPE_MAP.get(clean_cell(row.get(COL_TYPE)), SettlementRecordType.SALE)
        return SettlementRecord(
            platform=self.platform,
            settled_at=parse_date(clean_cell(row[COL_SETTLED_AT]), field=COL_SETTLED_AT),
            gross_amount=Money.from_str(clean_cell(row[COL_GROSS])),
            fee_amount=Money.from_str(clean_cell(row[COL_FEE])),
            net_amount=Money.from_str(clean_cell(row[COL_NET])),
            record_type=record_type,
            external_order_id=clean_cell(row.get(COL_ORDER_ID)) or None,
            product_name=clean_cell(row.get(COL_PRODUCT)) or None,
            source_row=row_number,
            raw={k: clean_cell(v) for k, v in row.items() if k},
        )
