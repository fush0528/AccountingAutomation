"""解析層的抽象介面。

每個平台的結算單格式都不一樣：欄位名稱、編碼、日期寫法、數字格式、
甚至表頭有幾層都不同。如果把這些差異寫成 if-else，每新增一個平台就要
動到核心邏輯，而且無法單獨測試某個平台。

這裡的做法是把「差異」關在 parser 裡，讓核心只認識一個介面。
新增平台 = 新增一個檔案 + 一行 ``@register``，核心零改動。

三個關鍵設計：

1. **parser 收 bytes，不收路徑。** :class:`SettlementSource` 包住
   檔名與位元組內容。這樣同一個 parser 既能吃磁碟上的檔案，也能吃
   W4 從 HTTP 上傳進來的資料，不必為了上傳再寫一套。

2. **偵測用信心分數，不用副檔名。** 每個 parser 自己回報
   :meth:`~BaseSettlementParser.sniff` 分數，registry 取最高分。
   副檔名太弱——三個平台可能都匯出 .csv。

3. **解析失敗不中斷整批。** 壞掉的列收進 :class:`RowError`，
   好的列照樣回傳。實務上一份千筆的結算單不該因為第 37 列的日期
   格式怪異就整份作廢，使用者需要看到「哪幾列有問題、為什麼」。
"""

from __future__ import annotations

import hashlib
from abc import ABC, abstractmethod
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from functools import cached_property
from pathlib import Path
from typing import ClassVar

from ..domain.models import SettlementRecord

__all__ = [
    "BaseSettlementParser",
    "ParseError",
    "ParseResult",
    "RowError",
    "SettlementSource",
]


class ParseError(Exception):
    """整份檔案無法解析時拋出（單列的問題請用 :class:`RowError`）。"""


# 這裡刻意不加 slots=True：cached_property 需要實例的 __dict__ 才能存結果，
# 而 slots 會把 __dict__ 拿掉。檔案內容可能有數 MB，重算雜湊不划算。
@dataclass(frozen=True)
class SettlementSource:
    """一份待解析的結算單。

    刻意只持有檔名與位元組，不持有路徑——因為 W4 的上傳端點拿到的就是
    位元組，沒有路徑可言。
    """

    filename: str
    content: bytes

    @classmethod
    def from_path(cls, path: str | Path) -> SettlementSource:
        p = Path(path)
        return cls(filename=p.name, content=p.read_bytes())

    @property
    def suffix(self) -> str:
        return Path(self.filename).suffix.lower()

    @cached_property
    def sha256(self) -> str:
        """檔案內容的指紋。

        W4 的冪等匯入以此為第一道防線：指紋相同就是同一份檔案，
        直接回傳既有批次，不再寫入任何一列。
        """
        return hashlib.sha256(self.content).hexdigest()

    def decode(self, *encodings: str) -> tuple[str, str]:
        """依序嘗試解碼，回傳 ``(文字, 成功的編碼)``。

        台灣的結算單 Big5 與 UTF-8 都很常見，有些還帶 BOM。
        與其要求使用者先轉檔，不如在這裡試。
        """
        for enc in encodings:
            try:
                return self.content.decode(enc), enc
            except (UnicodeDecodeError, LookupError):
                continue
        raise ParseError(
            f"{self.filename} 無法以 {', '.join(encodings)} 任一編碼解讀。"
            f"請確認檔案是否為結算單原始匯出檔。"
        )


@dataclass(frozen=True, slots=True)
class RowError:
    """單一列的解析失敗。

    帶著 ``raw`` 是刻意的：使用者要能看到那一列原本長什麼樣，
    才有辦法判斷是資料真的有問題，還是我們的 parser 沒寫好。
    """

    row_number: int
    message: str
    raw: Mapping[str, str]

    def __str__(self) -> str:
        return f"第 {self.row_number} 列：{self.message}"


@dataclass(frozen=True, slots=True)
class ParseResult:
    """一次解析的完整結果，成功與失敗都在裡面。"""

    platform: str
    records: tuple[SettlementRecord, ...]
    errors: tuple[RowError, ...] = ()
    row_count: int = 0
    encoding: str = "utf-8"

    @property
    def accepted_count(self) -> int:
        return len(self.records)

    @property
    def rejected_count(self) -> int:
        return len(self.errors)

    @property
    def success_rate(self) -> float:
        return self.accepted_count / self.row_count if self.row_count else 0.0

    def summary(self) -> str:
        return (
            f"{self.platform}：{self.row_count} 列，"
            f"成功 {self.accepted_count}、失敗 {self.rejected_count}"
            f"（{self.success_rate:.1%}），編碼 {self.encoding}"
        )


class BaseSettlementParser(ABC):
    """所有平台 parser 的共同介面。

    子類別必須宣告 ``platform`` 與 ``display_name``，並實作
    :meth:`sniff` 與 :meth:`parse`。除此之外核心對它一無所知——
    這正是重點。
    """

    platform: ClassVar[str]
    display_name: ClassVar[str]
    #: 這個平台常見的副檔名，僅供提示，不作為偵測依據
    extensions: ClassVar[tuple[str, ...]] = ()

    def __init_subclass__(cls, **kwargs: object) -> None:
        super().__init_subclass__(**kwargs)
        if ABC not in cls.__bases__:
            for attr in ("platform", "display_name"):
                if not getattr(cls, attr, None):
                    raise TypeError(f"{cls.__name__} 必須宣告 {attr}")

    @classmethod
    @abstractmethod
    def sniff(cls, source: SettlementSource) -> float:
        """回報「這份檔案是不是我這個平台的」的信心分數，0 到 1。

        0 代表確定不是，1 代表確定是。registry 會選分數最高的 parser。
        實作上通常是檢查特徵欄位名稱是否出現在表頭。

        這個方法不應該拋出例外——拿到完全無關的檔案時回 0 就好。
        """

    @abstractmethod
    def parse(self, source: SettlementSource) -> ParseResult:
        """把結算單解析成領域模型。

        單列的問題收進 ``ParseResult.errors`` 並繼續處理下一列；
        只有整份檔案無法讀取時才拋出 :class:`ParseError`。
        """

    # ------------------------------------------------------------------
    # 子類別常用的小工具
    # ------------------------------------------------------------------
    @staticmethod
    def _header_score(header: Sequence[str], required: Sequence[str]) -> float:
        """特徵欄位的命中率，給 :meth:`sniff` 用。"""
        if not required:
            return 0.0
        cleaned = {h.strip() for h in header if h}
        hits = sum(1 for name in required if name in cleaned)
        return hits / len(required)

    def __repr__(self) -> str:
        return f"<{type(self).__name__} platform={self.platform!r}>"
