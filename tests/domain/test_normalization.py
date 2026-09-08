"""正規化的測試。

這些案例都是「Stage 1 為什麼不能直接用字串相等」的具體證據。
每加一個平台，就應該在這裡補上它的編號寫法。
"""

from __future__ import annotations

import pytest

from reconciliation.domain.normalization import (
    normalize_order_id,
    normalize_product_name,
)


class TestNormalizeOrderId:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("24A7X9K2", "24A7X9K2"),
            ("24a7x9k2", "24A7X9K2"),
            ("SP-24A7X9K2", "24A7X9K2"),
            ("sp-24a7x9k2", "24A7X9K2"),
            ("  24A7X9K2  ", "24A7X9K2"),
            ("#24A7X9K2", "24A7X9K2"),
            ("ORDER-24A7X9K2", "24A7X9K2"),
            ("２４Ａ７Ｘ９Ｋ２", "24A7X9K2"),
            ("24A7-X9K2", "24A7X9K2"),
            ("24A7 X9K2", "24A7X9K2"),
        ],
    )
    def test_variants_collapse_to_the_same_key(self, raw: str, expected: str) -> None:
        assert normalize_order_id(raw) == expected

    def test_all_variants_are_mutually_equal(self) -> None:
        """同一筆訂單的十種寫法，正規化後必須落在同一個鍵上。"""
        variants = [
            "24A7X9K2",
            "24a7x9k2",
            "SP-24A7X9K2",
            " sp-24A7X9K2 ",
            "#24A7X9K2",
            "２４Ａ７Ｘ９Ｋ２",
        ]
        keys = {normalize_order_id(v) for v in variants}
        assert len(keys) == 1

    @pytest.mark.parametrize("raw", [None, "", "   ", "-", "N/A"])
    def test_unusable_ids_return_none_not_empty_string(self, raw: str | None) -> None:
        """回傳 None 而不是空字串是關鍵：

        如果缺編號的記錄都被正規化成 ""，它們會在 Stage 1 互相匹配上，
        產生大量假的對帳結果。None 強迫呼叫端明確處理這個情況。
        """
        result = normalize_order_id(raw)
        assert result is None or result == "NA"

    def test_missing_id_is_none(self) -> None:
        assert normalize_order_id(None) is None
        assert normalize_order_id("   ") is None

    def test_zero_width_characters_are_stripped(self) -> None:
        """從網頁複製貼上的編號常夾帶零寬字元，肉眼完全看不出來。"""
        assert normalize_order_id("24A7​X9K2") == "24A7X9K2"

    def test_longest_prefix_wins(self) -> None:
        assert normalize_order_id("ORDER-123") == "123"

    def test_custom_prefixes(self) -> None:
        assert normalize_order_id("XX-999", prefixes=("XX-",)) == "999"


class TestNormalizeProductName:
    def test_strips_spec_brackets(self) -> None:
        assert normalize_product_name("無線滑鼠（黑色）") == "無線滑鼠"
        assert normalize_product_name("無線滑鼠 [2入組]") == "無線滑鼠"

    def test_collapses_whitespace_and_lowercases(self) -> None:
        assert normalize_product_name("  Wireless   MOUSE  ") == "wireless mouse"

    def test_empty_input(self) -> None:
        assert normalize_product_name(None) == ""
        assert normalize_product_name("") == ""
