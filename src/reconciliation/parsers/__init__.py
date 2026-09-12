"""解析層：把各平台格式各異的結算單，變成統一的領域模型。

在這裡 import 每個 parser 模組，是為了讓 ``@register`` decorator 有機會
執行。新增平台時除了新增檔案，只需要在下面補一行 import。
"""

from .base import (
    BaseSettlementParser,
    ParseError,
    ParseResult,
    RowError,
    SettlementSource,
)
from .platform_a import PlatformAParser
from .platform_b import PlatformBParser
from .platform_c import PlatformCParser
from .registry import available_platforms, detect, get_parser, register

__all__ = [
    "BaseSettlementParser",
    "ParseError",
    "ParseResult",
    "PlatformAParser",
    "PlatformBParser",
    "PlatformCParser",
    "RowError",
    "SettlementSource",
    "available_platforms",
    "detect",
    "get_parser",
    "register",
]
