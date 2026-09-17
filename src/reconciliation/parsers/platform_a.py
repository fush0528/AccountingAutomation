"""平台 A：第三方金流的撥款對帳檔（UTF-8 / Big5 CSV，單層表頭）。

現實中的來源
------------
這種檔案不是電商平台給的，是**第三方金流商**給的——賣家自架官網或
自有購物車時，收單、扣手續費、撥款都由金流商處理，對帳檔也由它出。
欄位結構參考綠界科技（ECPay）公開的「撥款對帳檔」與「對帳檔 V2」
技術文件（見 ``docs/adr/0005-sample-data-provenance.md``）。
**檔案內容全部是合成的**，只有欄位規格取自公開文件。

這一份是三個 parser 裡欄位最多的，難點也最能代表金流對帳檔::

    交易日期,廠商訂單編號,金流交易編號,付款方式,手續費率,商品名稱,交易金額,
    金流手續費,平台手續費,金流處理費,退款金額,應收款項(淨額),結算日期,撥款日期
    2026-03-01,SP-24A7X9K2,EC202603010001,信用卡,2.75%,無線滑鼠,1000.00,
    27.50,20.00,7.80,0.00,944.70,2026-03-13,2026-03-15

三個要處理的地方：

1. **手續費拆成三欄。** 金流手續費（刷卡費）、平台手續費、金流處理費
   是不同的成本科目，對金流商有意義，但對「這筆訂單被扣了多少」
   沒有意義。領域模型只有一個 ``fee_amount``，所以這裡要相加。
2. **退款寫在同一列，不另起一列。** 平台 C 的退貨是獨立的一列，
   這裡卻是銷售列上的一個欄位。同一個商業事實、兩種表示法——
   parser 的工作就是把它們收斂成同一種（見 ``_build``）。
3. **檔案自帶檢核欄位。** ``應收款項(淨額)`` 是金流商自己算的結果。
   我們不直接採用它，而是用自己的解讀去**驗證**它——對不上就讓這一列
   失敗，而不是默默寫進資料庫。理由見下方 ``_cross_check``。
"""

from __future__ import annotations

import csv
import io
from typing import ClassVar

from ..domain.models import SettlementRecord, SettlementRecordType
from ..domain.money import Money, MoneyError, money_sum
from .base import (
    BaseSettlementParser,
    ParseError,
    ParseResult,
    RowError,
    SettlementSource,
)
from .fields import clean_cell, parse_date
from .registry import register

COL_TRANSACTED_AT = "交易日期"
COL_ORDER_ID = "廠商訂單編號"
COL_GATEWAY_ID = "金流交易編號"
COL_PAYMENT_METHOD = "付款方式"
COL_FEE_RATE = "手續費率"
COL_PRODUCT = "商品名稱"
COL_GROSS = "交易金額"
COL_REFUND = "退款金額"
COL_REFUND_AT = "退款日期"
COL_NET = "應收款項(淨額)"
COL_SETTLED_AT = "結算日期"
COL_PAID_AT = "撥款日期"

#: 三個扣款科目。順序固定，錯誤訊息才能穩定重現。
FEE_COLUMNS = ("金流手續費", "平台手續費", "金流處理費")

#: 用來認出這個平台的特徵欄位。
#:
#: 刻意選了 ``應收款項(淨額)`` 與兩個手續費欄位——這組合在另外兩個
#: 平台都不存在。只用「交易金額」這種通用詞會跟平台 C 撞在一起。
SIGNATURE_COLUMNS = (
    COL_ORDER_ID,
    COL_GROSS,
    COL_NET,
    COL_PAID_AT,
    FEE_COLUMNS[0],
    FEE_COLUMNS[1],
)


@register
class PlatformAParser(BaseSettlementParser):
    """平台 A（第三方金流撥款對帳檔）的解析器。"""

    platform: ClassVar[str] = "platform_a"
    display_name: ClassVar[str] = "平台 A（金流商撥款對帳檔 CSV）"
    extensions: ClassVar[tuple[str, ...]] = (".csv",)

    #: 這類檔案通常可選編碼下載，兩種都要能讀。UTF-8 先試，
    #: 因為 Big5 幾乎不會把 UTF-8 的位元組解成合法中文，反之則不然。
    ENCODINGS: ClassVar[tuple[str, ...]] = ("utf-8-sig", "utf-8", "big5")

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
                # 一列可能產生一筆或兩筆——退款欄不為零時會多一筆退款記錄。
                records.extend(self._build(clean, row_number))
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
    def _build(self, row: dict[str, str], row_number: int) -> list[SettlementRecord]:
        """把一列拆成一或兩筆領域記錄。

        回傳 list 而不是單一物件，是因為這個平台把「賣了又退一部分」
        壓縮在同一列裡。領域模型維持「一筆記錄描述一件事」，
        所以壓縮要在這裡解開——下游的歸因邏輯才不用知道
        哪個平台用哪種表示法。
        """
        settled_at = parse_date(row[COL_PAID_AT], field=COL_PAID_AT)
        gross = Money.from_str(row[COL_GROSS])
        fee = money_sum(Money.from_str(row[col]) for col in FEE_COLUMNS)
        refund = Money.from_str(row.get(COL_REFUND) or "0")
        order_id = row.get(COL_ORDER_ID, "") or None
        product = row.get(COL_PRODUCT) or None

        sale = SettlementRecord(
            platform=self.platform,
            settled_at=settled_at,
            gross_amount=gross,
            fee_amount=fee,
            # 不讀檔案上的淨額——那是**含退款後**的數字。這裡要的是
            # 這筆銷售本身的淨額，退款另外記一筆。兩者相加才等於檔案值，
            # 下面的 _cross_check 就是在驗這件事。
            net_amount=gross - fee,
            record_type=SettlementRecordType.SALE,
            external_order_id=order_id,
            product_name=product,
            source_row=row_number,
            raw=row,
        )
        built = [sale]

        if not refund.is_zero:
            # 退款有自己的發生日，但不一定填。沒填就沿用撥款日——
            # 反正這筆錢是在同一次撥款裡被扣掉的。
            refunded_at = (
                parse_date(row[COL_REFUND_AT], field=COL_REFUND_AT)
                if row.get(COL_REFUND_AT)
                else settled_at
            )
            # 檔案上的退款金額印正數（它描述的是「退了多少錢」）。
            # 領域模型要的是對帳上的符號：退款會讓撥款變少，所以是負的。
            built.append(
                SettlementRecord(
                    platform=self.platform,
                    settled_at=refunded_at,
                    gross_amount=-refund,
                    fee_amount=Money.zero(gross.currency),
                    net_amount=-refund,
                    record_type=SettlementRecordType.REFUND,
                    external_order_id=order_id,
                    product_name=product,
                    source_row=row_number,
                    raw=row,
                )
            )

        self._cross_check(built, row, row_number)
        return built

    @staticmethod
    def _cross_check(
        built: list[SettlementRecord], row: dict[str, str], row_number: int
    ) -> None:
        """用檔案自帶的淨額欄位驗證我們的解讀。

        這個檢查抓得到的真實狀況：金流商新增了第四個扣款科目而我們沒跟上、
        退款欄的正負號慣例改了、某一欄改名導致被讀成零。這些都不會讓程式
        當掉——沒有這道檢查，它們會安靜地產生少扣一筆費用的記錄，
        然後在對帳報告上變成一堆查不出原因的「金額差異」。

        寧可讓這一列明確失敗（會出現在匯入結果的 rejected 裡），
        也不要產生看起來正常但其實是錯的資料。
        """
        declared_raw = row.get(COL_NET)
        if not declared_raw:
            return  # 這欄可有可無；沒有就沒得驗，不是錯誤
        declared = Money.from_str(declared_raw)
        computed = money_sum(record.net_amount for record in built)
        if computed != declared:
            raise ParseError(
                f"第 {row_number} 列淨額對不上："
                f"檔案寫 {declared}，"
                f"依 {'＋'.join(FEE_COLUMNS)} 與退款金額算出來是 {computed}"
                f"（差 {computed - declared}）。"
                f"這通常代表對帳檔多了本 parser 未涵蓋的扣款欄位。"
            )
