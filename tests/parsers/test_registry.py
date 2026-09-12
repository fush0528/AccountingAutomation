"""註冊表與平台自動偵測的測試。

最重要的一個測試是 :meth:`TestExtensibility.test_new_platform_needs_no_core_changes`：
它在執行期定義一個全新的 parser，證明「新增平台不需要動到核心」不是
README 上的宣稱，而是可驗證的事實。
"""

from __future__ import annotations

from typing import ClassVar

import pytest

from reconciliation.parsers import registry
from reconciliation.parsers.base import (
    BaseSettlementParser,
    ParseError,
    ParseResult,
    SettlementSource,
)
from reconciliation.parsers.platform_a import PlatformAParser
from reconciliation.parsers.platform_b import PlatformBParser
from reconciliation.parsers.platform_c import PlatformCParser

from .conftest import source


class TestDetection:
    def test_detects_each_platform(
        self,
        platform_a_source: SettlementSource,
        platform_b_source: SettlementSource,
        platform_c_source: SettlementSource,
    ) -> None:
        assert isinstance(registry.detect(platform_a_source), PlatformAParser)
        assert isinstance(registry.detect(platform_b_source), PlatformBParser)
        assert isinstance(registry.detect(platform_c_source), PlatformCParser)

    def test_detection_does_not_rely_on_filename(self, platform_c_source: SettlementSource) -> None:
        """副檔名與檔名都不可信——內容才算數。"""
        disguised = SettlementSource(
            filename="完全沒有線索的檔名.dat", content=platform_c_source.content
        )
        assert isinstance(registry.detect(disguised), PlatformCParser)

    def test_unknown_format_raises_with_all_scores(self) -> None:
        """認不出來時要說明為什麼，不能只丟一句「不支援的格式」。"""
        with pytest.raises(ParseError) as exc_info:
            registry.detect(source("mystery.csv", "姓名,電話\n王小明,0912345678"))
        message = str(exc_info.value)
        assert "無法判斷" in message
        assert "信心分數" in message
        for _, display_name in registry.available_platforms():
            assert display_name.split("（")[0] in message

    def test_empty_file_does_not_crash_detection(self) -> None:
        with pytest.raises(ParseError):
            registry.detect(SettlementSource("empty.csv", b""))


class TestGetParser:
    def test_by_platform_code(self) -> None:
        assert isinstance(registry.get_parser("platform_a"), PlatformAParser)

    def test_unknown_code_lists_the_known_ones(self) -> None:
        with pytest.raises(ParseError) as exc_info:
            registry.get_parser("platform_z")
        assert "platform_a" in str(exc_info.value)

    def test_available_platforms_is_sorted(self) -> None:
        codes = [code for code, _ in registry.available_platforms()]
        assert codes == sorted(codes)
        assert "platform_a" in codes


class TestExtensibility:
    """新增平台不需要修改核心——這裡把它證明出來。"""

    def test_new_platform_needs_no_core_changes(self) -> None:
        signature = "神秘平台專用欄位"

        @registry.register
        class PlatformZParser(BaseSettlementParser):
            platform: ClassVar[str] = "platform_z"
            display_name: ClassVar[str] = "平台 Z（測試用）"

            @classmethod
            def sniff(cls, src: SettlementSource) -> float:
                text, _ = src.decode("utf-8")
                return 1.0 if signature in text else 0.0

            def parse(self, src: SettlementSource) -> ParseResult:
                return ParseResult(platform=self.platform, records=(), row_count=0)

        try:
            # 到這裡為止，registry.py、base.py 與其他三個 parser 都沒有被修改過
            assert ("platform_z", "平台 Z（測試用）") in registry.available_platforms()

            detected = registry.detect(source("z.csv", f"{signature}\n1"))
            assert isinstance(detected, PlatformZParser)

            # 而且不會影響既有平台的偵測
            assert isinstance(registry.get_parser("platform_a"), PlatformAParser)
        finally:
            registry.unregister("platform_z")

        assert "platform_z" not in [c for c, _ in registry.available_platforms()]

    def test_duplicate_platform_code_is_rejected(self) -> None:
        """兩個 parser 搶同一個代碼會安靜地覆蓋掉——所以直接擋下來。"""
        with pytest.raises(ValueError, match="已經被"):

            @registry.register
            class Clash(BaseSettlementParser):
                platform: ClassVar[str] = "platform_a"
                display_name: ClassVar[str] = "冒牌的平台 A"

                @classmethod
                def sniff(cls, src: SettlementSource) -> float:
                    return 0.0

                def parse(self, src: SettlementSource) -> ParseResult:
                    return ParseResult(platform=self.platform, records=())

    def test_subclass_must_declare_platform(self) -> None:
        with pytest.raises(TypeError, match="必須宣告"):

            class Incomplete(BaseSettlementParser):
                display_name: ClassVar[str] = "忘了宣告 platform"

                @classmethod
                def sniff(cls, src: SettlementSource) -> float:
                    return 0.0

                def parse(self, src: SettlementSource) -> ParseResult:
                    return ParseResult(platform="", records=())


class TestSettlementSource:
    def test_sha256_is_stable_and_content_based(self) -> None:
        """冪等匯入的第一道防線：同樣的內容必須得到同樣的指紋。"""
        a = SettlementSource("一月.csv", b"same content")
        b = SettlementSource("二月.csv", b"same content")
        c = SettlementSource("一月.csv", b"different")
        assert a.sha256 == b.sha256  # 檔名不影響指紋
        assert a.sha256 != c.sha256
        assert len(a.sha256) == 64

    def test_decode_tries_in_order(self) -> None:
        src = SettlementSource("big5.csv", "中文內容".encode("big5"))
        text, encoding = src.decode("big5", "utf-8")
        assert text == "中文內容"
        assert encoding == "big5"

    def test_decode_failure_names_the_encodings_tried(self) -> None:
        src = SettlementSource("bad.csv", b"\xff\xfe\x00\x01\xff")
        with pytest.raises(ParseError, match="utf-8"):
            src.decode("utf-8", "ascii")
