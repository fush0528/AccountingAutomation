"""Parser 註冊表與平台自動偵測。

新增一個平台的完整步驟：

1. 在 ``parsers/`` 新增一個檔案
2. 繼承 :class:`~.base.BaseSettlementParser`，加上 ``@register``
3. 在 ``parsers/__init__.py`` import 它（讓 decorator 有機會執行）

核心的任何一行程式碼都不需要修改。這就是開放封閉原則在這個專案裡的
具體樣貌——而且它是可驗證的：``tests/parsers/test_registry.py`` 裡有一個
測試會在執行期定義一個全新的 parser，證明註冊真的不需要動到既有程式碼。
"""

from __future__ import annotations

from .base import BaseSettlementParser, ParseError, SettlementSource

__all__ = [
    "available_platforms",
    "detect",
    "get_parser",
    "register",
    "unregister",
]

_REGISTRY: dict[str, type[BaseSettlementParser]] = {}

#: 低於這個信心分數就視為「認不出來」，寧可請使用者指定平台，
#: 也不要猜錯——猜錯的代價是整份帳單解析成垃圾資料。
DETECTION_THRESHOLD = 0.5


def register(cls: type[BaseSettlementParser]) -> type[BaseSettlementParser]:
    """把 parser 註冊進來的 decorator。"""
    existing = _REGISTRY.get(cls.platform)
    if existing is not None and existing is not cls:
        raise ValueError(
            f"平台代碼 {cls.platform!r} 已經被 {existing.__name__} 使用了，"
            f"{cls.__name__} 需要換一個代碼。"
        )
    _REGISTRY[cls.platform] = cls
    return cls


def unregister(platform: str) -> None:
    """移除註冊，僅供測試使用。"""
    _REGISTRY.pop(platform, None)


def available_platforms() -> list[tuple[str, str]]:
    """回傳 ``[(平台代碼, 顯示名稱), ...]``，依代碼排序。"""
    return sorted((cls.platform, cls.display_name) for cls in _REGISTRY.values())


def get_parser(platform: str) -> BaseSettlementParser:
    """依平台代碼取得 parser 實例。"""
    cls = _REGISTRY.get(platform)
    if cls is None:
        known = ", ".join(code for code, _ in available_platforms()) or "（尚未註冊任何平台）"
        raise ParseError(f"未知的平台代碼 {platform!r}。已註冊的平台：{known}")
    return cls()


def detect(source: SettlementSource) -> BaseSettlementParser:
    """自動判斷這份檔案屬於哪個平台。

    做法是問過所有 parser 的 :meth:`~.base.BaseSettlementParser.sniff`
    分數，取最高者。分數低於 :data:`DETECTION_THRESHOLD` 時直接放棄，
    並在錯誤訊息裡列出各家分數——使用者才知道為什麼認不出來，
    而不是只看到一句「不支援的格式」。
    """
    if not _REGISTRY:
        raise ParseError("尚未註冊任何 parser")

    scores: list[tuple[float, type[BaseSettlementParser]]] = []
    for cls in _REGISTRY.values():
        try:
            scores.append((cls.sniff(source), cls))
        except Exception:  # noqa: BLE001 - sniff 不該拋例外，但別讓它拖垮偵測
            scores.append((0.0, cls))

    scores.sort(key=lambda item: (-item[0], item[1].platform))
    best_score, best_cls = scores[0]

    if best_score < DETECTION_THRESHOLD:
        detail = "、".join(f"{cls.display_name} {score:.0%}" for score, cls in scores)
        raise ParseError(
            f"無法判斷 {source.filename} 是哪個平台的結算單。"
            f"各平台信心分數：{detail}。"
            f"若格式正確，請改用 get_parser() 明確指定平台。"
        )
    return best_cls()
