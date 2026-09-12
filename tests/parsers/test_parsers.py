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
        assert result.row_count == 5
        assert result.accepted_count == 5
        assert result.rejected_count == 0

    def test_first_record_fields(self, platform_a_source: SettlementSource) -> None:
        record = PlatformAParser().parse(platform_a_source).records[0]
        assert record.external_order_id == "SP-24A7X9K2"  # 原文保留，未正規化
        assert record.settled_at == date(2026, 3, 15)
        assert record.gross_amount == Money.from_str("1000")
        assert record.fee_amount == Money.from_str("55.30")
        assert record.net_amount == Money.from_str("944.70")
        assert record.residual.is_zero
        assert record.record_type is SettlementRecordType.SALE

    def test_blank_order_id_becomes_none(self, platform_a_source: SettlementSource) -> None:
        record = PlatformAParser().parse(platform_a_source).records[2]
        assert record.external_order_id is None
        assert not record.has_order_id

    def test_refund_row_is_typed_and_negative(self, platform_a_source: SettlementSource) -> None:
        record = PlatformAParser().parse(platform_a_source).records[4]
        assert record.record_type is SettlementRecordType.REFUND
        assert record.net_amount.is_negative

    def test_raw_keeps_the_original_row(self, platform_a_source: SettlementSource) -> None:
        record = PlatformAParser().parse(platform_a_source).records[0]
        assert record.raw["交易類型"] == "銷售"
        assert record.source_row == 2  # 第 1 列是表頭

    def test_bad_row_is_collected_not_fatal(self) -> None:
        """一列壞掉不能讓整份檔案作廢。"""
        text = "\n".join(
            [
                PLATFORM_A_HEADER,
                "SP-1,2026-03-01,商品,1,1000,55.30,944.70,2026-03-15,銷售",
                "SP-2,2026-03-02,商品,1,1000,55.30,944.70,不是日期,銷售",
                "SP-3,2026-03-03,商品,1,金額壞掉,55.30,944.70,2026-03-17,銷售",
                "SP-4,2026-03-04,商品,1,2000,110.60,1889.40,2026-03-18,銷售",
            ]
        )
        result = PlatformAParser().parse(source("a.csv", text))
        assert result.accepted_count == 2
        assert result.rejected_count == 2
        assert result.row_count == 4
        assert [e.row_number for e in result.errors] == [3, 4]
        assert "日期" in result.errors[0].message
        # 壞掉的原始內容要留著，人工才查得下去
        assert result.errors[0].raw["訂單編號"] == "SP-2"

    def test_missing_required_column_is_fatal(self) -> None:
        """缺欄位是整份檔案的問題，不是某一列的問題。"""
        text = "訂單編號,結算日期\nSP-1,2026-03-15"
        with pytest.raises(ParseError, match="缺少必要欄位"):
            PlatformAParser().parse(source("a.csv", text))

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
