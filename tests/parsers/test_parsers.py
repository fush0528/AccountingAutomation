"""三個平台 parser 的 golden file 測試。

每個平台一組測試，驗證同一件事：**輸入格式再怎麼不同，輸出都是同一種
領域模型**。三組測試的斷言結構刻意寫得很像——那正是抽象層有效的證據。
"""

from __future__ import annotations

from datetime import date

import pytest

from reconciliation.domain.models import SettlementRecordType
from reconciliation.domain.money import Money
from reconciliation.parsers.base import ParseError, SettlementSource
from reconciliation.parsers.platform_a import PlatformAParser
from reconciliation.parsers.platform_b import PlatformBParser
from reconciliation.parsers.platform_c import PlatformCParser

from .conftest import PLATFORM_A_HEADER, build_platform_b_workbook, source


class TestPlatformA:
    def test_parses_all_rows(self, platform_a_source: SettlementSource) -> None:
        result = PlatformAParser().parse(platform_a_source)
        assert result.platform == "platform_a"
        # 4 列進去，5 筆出來——最後一列同時含銷售與退款。
        # 記錄數大於列數是這個平台的正常狀態，不是 bug。
        assert result.row_count == 4
        assert result.accepted_count == 5
        assert result.rejected_count == 0

    def test_three_fee_columns_are_summed(
        self, platform_a_source: SettlementSource
    ) -> None:
        """金流手續費 ＋ 平台手續費 ＋ 金流處理費 收斂成單一 fee_amount。"""
        record = PlatformAParser().parse(platform_a_source).records[0]
        assert record.external_order_id == "SP-24A7X9K2"  # 原文保留，未正規化
        assert record.settled_at == date(2026, 3, 15)  # 用撥款日，不是結算日
        assert record.gross_amount == Money.from_str("1000")
        assert record.fee_amount == Money.from_str("55.30")  # 30.42+19.35+5.53
        assert record.net_amount == Money.from_str("944.70")
        assert record.residual.is_zero
        assert record.record_type is SettlementRecordType.SALE

    def test_blank_order_id_becomes_none(self, platform_a_source: SettlementSource) -> None:
        record = PlatformAParser().parse(platform_a_source).records[2]
        assert record.external_order_id is None
        assert not record.has_order_id

    def test_inline_refund_becomes_a_second_record(
        self, platform_a_source: SettlementSource
    ) -> None:
        """一列拆兩筆：平台把退款壓在銷售列上，parser 要把它解開。

        這是本專案最重要的 parser 斷言之一——平台 C 的退貨是獨立一列，
        這裡卻是同一列的欄位，但兩邊產生的領域記錄是同一個形狀。
        """
        records = PlatformAParser().parse(platform_a_source).records
        sale, refund = records[3], records[4]

        assert sale.record_type is SettlementRecordType.SALE
        assert sale.source_row == refund.source_row == 5  # 來自同一列
        assert sale.external_order_id == refund.external_order_id == "SP-24D1A3P5"

        assert refund.record_type is SettlementRecordType.REFUND
        # 檔案印正數 600.00，領域模型要負數
        assert refund.gross_amount == Money.from_str("-600")
        assert refund.net_amount.is_negative
        assert refund.fee_amount.is_zero  # 退款不退手續費
        assert refund.settled_at == date(2026, 3, 22)  # 用退款日，不是撥款日
        # 兩筆的淨額相加要等於檔案宣告的淨額
        assert sale.net_amount + refund.net_amount == Money.from_str("1667.28")

    def test_declared_net_that_disagrees_is_rejected(
        self, platform_a_inconsistent_source: SettlementSource
    ) -> None:
        """檔案自帶的淨額對不上時，這一列必須失敗而不是照收。

        這道檢查存在的理由：金流商多加一個扣款科目時，程式不會當掉，
        只會安靜地少扣一筆錢。沒有這個斷言，那種退化不會有人發現。
        """
        result = PlatformAParser().parse(platform_a_inconsistent_source)
        assert result.accepted_count == 0
        assert result.rejected_count == 1
        assert "淨額對不上" in result.errors[0].message
        # 錯誤訊息要講得夠清楚，人看了知道下一步查哪裡
        assert "944.70" in result.errors[0].message
        assert "900.00" in result.errors[0].message

    def test_raw_keeps_the_original_row(self, platform_a_source: SettlementSource) -> None:
        record = PlatformAParser().parse(platform_a_source).records[0]
        assert record.raw["金流交易編號"] == "EC0001"
        assert record.raw["付款方式"] == "信用卡"
        assert record.source_row == 2  # 第 1 列是表頭

    def test_bad_row_is_collected_not_fatal(self) -> None:
        """一列壞掉不能讓整份檔案作廢。"""
        good = "30.42,19.35,5.53,,0.00,944.70,2026-03-13"
        text = "\n".join(
            [
                PLATFORM_A_HEADER,
                f"2026-03-01,SP-1,EC1,信用卡,5.53%,商品,1000,{good},2026-03-15",
                f"2026-03-02,SP-2,EC2,信用卡,5.53%,商品,1000,{good},不是日期",
                f"2026-03-03,SP-3,EC3,信用卡,5.53%,商品,金額壞掉,{good},2026-03-17",
                f"2026-03-04,SP-4,EC4,信用卡,5.53%,商品,1000,{good},2026-03-18",
            ]
        )
        result = PlatformAParser().parse(source("a.csv", text))
        assert result.accepted_count == 2
        assert result.rejected_count == 2
        assert result.row_count == 4
        assert [e.row_number for e in result.errors] == [3, 4]
        assert "日期" in result.errors[0].message
        # 壞掉的原始內容要留著，人工才查得下去
        assert result.errors[0].raw["廠商訂單編號"] == "SP-2"

    def test_missing_required_column_is_fatal(self) -> None:
        """缺欄位是整份檔案的問題，不是某一列的問題。"""
        text = "廠商訂單編號,撥款日期\nSP-1,2026-03-15"
        with pytest.raises(ParseError, match="缺少必要欄位"):
            PlatformAParser().parse(source("a.csv", text))

    def test_big5_encoded_file_is_also_accepted(self) -> None:
        """同一個平台可選 UTF-8 或 Big5 下載，兩種都要能讀。"""
        text = "\n".join(
            [
                PLATFORM_A_HEADER,
                "2026-03-01,SP-1,EC1,信用卡,5.53%,無線滑鼠,1000,"
                "30.42,19.35,5.53,,0.00,944.70,2026-03-13,2026-03-15",
            ]
        )
        result = PlatformAParser().parse(source("a.csv", text, encoding="big5"))
        assert result.encoding == "big5"
        assert result.records[0].product_name == "無線滑鼠"

    def test_sniff(self, platform_a_source: SettlementSource) -> None:
        assert PlatformAParser.sniff(platform_a_source) == 1.0
        assert PlatformAParser.sniff(source("x.csv", "毫無關係,的欄位\n1,2")) == 0.0


class TestPlatformC:
    """Big5、民國年、千分位——每個維度都跟平台 A 不同。"""

    def test_parses_all_rows(self, platform_c_source: SettlementSource) -> None:
        result = PlatformCParser().parse(platform_c_source)
        assert result.encoding == "big5"
        assert result.accepted_count == 4
        assert result.rejected_count == 0

    def test_roc_date_converted_to_gregorian(self, platform_c_source: SettlementSource) -> None:
        """民國 115 年 3 月 15 日 = 西元 2026 年 3 月 15 日。"""
        record = PlatformCParser().parse(platform_c_source).records[0]
        assert record.settled_at == date(2026, 3, 15)

    def test_thousands_separator_parsed(self, platform_c_source: SettlementSource) -> None:
        record = PlatformCParser().parse(platform_c_source).records[1]
        assert record.gross_amount == Money.from_str("2580")
        assert record.net_amount == Money.from_str("2421.33")

    def test_refund_printed_positive_becomes_negative(
        self, platform_c_source: SettlementSource
    ) -> None:
        """平台 C 的退貨列印正數，靠「交易別」表達方向。

        統一成負數，下游才不必知道有這回事。
        """
        record = PlatformCParser().parse(platform_c_source).records[2]
        assert record.record_type is SettlementRecordType.REFUND
        assert record.gross_amount == Money.from_str("-590")
        assert record.net_amount.is_negative
        assert record.residual.is_zero

    def test_output_is_indistinguishable_from_platform_a(
        self, platform_a_source: SettlementSource, platform_c_source: SettlementSource
    ) -> None:
        """關鍵斷言：兩個格式天差地遠的平台，輸出的是同一種東西。"""
        a = PlatformAParser().parse(platform_a_source).records[0]
        c = PlatformCParser().parse(platform_c_source).records[0]
        assert type(a) is type(c)
        assert a.settled_at == c.settled_at
        assert a.gross_amount.currency == c.gross_amount.currency

    def test_sniff_rejects_utf8_file(self, platform_a_source: SettlementSource) -> None:
        assert PlatformCParser.sniff(platform_a_source) == 0.0


class TestPlatformB:
    def test_parses_all_rows(self, platform_b_source: SettlementSource) -> None:
        result = PlatformBParser().parse(platform_b_source)
        assert result.encoding == "xlsx"
        assert result.accepted_count == 4
        assert result.rejected_count == 0

    def test_numeric_cells_become_money(self, platform_b_source: SettlementSource) -> None:
        """Excel 的金額讀出來是 float，不是字串。"""
        record = PlatformBParser().parse(platform_b_source).records[1]
        assert record.gross_amount == Money.from_str("2500")
        assert record.fee_amount == Money.from_str("120")
        assert record.residual.is_zero

    def test_footer_row_is_not_counted_as_data(self, platform_b_source: SettlementSource) -> None:
        """「※ 本表僅供對帳參考」不是資料，也不是解析失敗。

        把它算進失敗會讓成功率的數字失去意義。
        """
        result = PlatformBParser().parse(platform_b_source)
        assert result.row_count == 4
        assert result.success_rate == 1.0

    def test_header_is_found_not_hardcoded(self) -> None:
        """報表標題佔幾列會變，parser 必須用找的。"""
        rows: list[list[object]] = [["MO-1", "商品", 1, 1000, 48.0, 952.0, "2026/03/15", "銷售"]]
        for title_rows in (0, 1, 3, 6):
            content = build_platform_b_workbook(rows, title_rows=title_rows)
            src = SettlementSource("b.xlsx", content)
            result = PlatformBParser().parse(src)
            assert result.accepted_count == 1, f"title_rows={title_rows} 時找不到表頭"

    def test_sniff_ignores_non_zip_files(self, platform_a_source: SettlementSource) -> None:
        """CSV 不是 zip，連 openpyxl 都不該啟動。"""
        assert PlatformBParser.sniff(platform_a_source) == 0.0

    def test_sniff_on_real_workbook(self, platform_b_source: SettlementSource) -> None:
        assert PlatformBParser.sniff(platform_b_source) == 1.0
