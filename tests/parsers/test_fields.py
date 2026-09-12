"""欄位解析工具的測試。"""

from __future__ import annotations

from datetime import date, datetime

import pytest

from reconciliation.parsers.base import ParseError
from reconciliation.parsers.fields import (
    clean_cell,
    parse_date,
    parse_int,
    parse_roc_date,
)


class TestCleanCell:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (None, ""),
            ("  文字  ", "文字"),
            ("１２３４", "1234"),  # 全形數字
            ("ＡＢＣ", "ABC"),
            (1234, "1234"),
            (1234.5, "1234.5"),
            (date(2026, 3, 15), "2026-03-15"),
            (datetime(2026, 3, 15, 10, 30), "2026-03-15"),
        ],
    )
    def test_normalises(self, value: object, expected: str) -> None:
        assert clean_cell(value) == expected


class TestParseDate:
    @pytest.mark.parametrize(
        "text",
        ["2026-03-15", "2026/03/15", "2026.03.15", "20260315", "2026-03-15 10:30:00"],
    )
    def test_accepts_common_formats(self, text: str) -> None:
        assert parse_date(text) == date(2026, 3, 15)

    def test_empty_raises(self) -> None:
        with pytest.raises(ParseError, match="為空"):
            parse_date("")

    def test_unrecognised_format_names_the_value(self) -> None:
        with pytest.raises(ParseError, match="三月十五"):
            parse_date("三月十五", field="結算日")


class TestParseRocDate:
    @pytest.mark.parametrize(
        ("text", "expected"),
        [
            ("115/03/15", date(2026, 3, 15)),
            ("115-03-15", date(2026, 3, 15)),
            ("99/12/31", date(2010, 12, 31)),  # 兩位數的民國年
            ("100/01/01", date(2011, 1, 1)),
        ],
    )
    def test_converts_roc_year(self, text: str, expected: date) -> None:
        assert parse_roc_date(text) == expected

    def test_falls_back_to_gregorian(self) -> None:
        """同一欄混用西元是常見的事，不該直接失敗。"""
        assert parse_roc_date("2026-03-15") == date(2026, 3, 15)

    def test_invalid_day_raises(self) -> None:
        with pytest.raises(ParseError):
            parse_roc_date("115/02/30")


class TestParseInt:
    @pytest.mark.parametrize(
        ("text", "expected"),
        [("1", 1), ("1,234", 1234), ("1.0", 1), ("１２", 12)],
    )
    def test_accepts_common_formats(self, text: str, expected: int) -> None:
        assert parse_int(text) == expected

    def test_default_when_empty(self) -> None:
        assert parse_int("", default=1) == 1

    def test_no_default_raises(self) -> None:
        with pytest.raises(ParseError, match="為空"):
            parse_int("")

    def test_garbage_raises(self) -> None:
        with pytest.raises(ParseError, match="不是整數"):
            parse_int("三個")
