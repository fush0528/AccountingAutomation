"""匹配鍵的正規化。

Stage 1 精確匹配之所以不能直接用字串相等，是因為同一筆訂單在我方系統與
平台結算單上的編號寫法經常不同：

* 平台在結算單加了前綴：``"SP-24A7X9K2"`` vs ``"24A7X9K2"``
* 大小寫不一致：``"24a7x9k2"`` vs ``"24A7X9K2"``
* 匯出時混入空白或不可見字元：``"24A7X9K2 "``、``"24A7​X9K2"``
* 全形字元：``"２４Ａ７"``

正規化只用在**比對**，不會改寫 :class:`~.models.Order` 上保存的原始編號。
這個分離是刻意的：人工查核時必須看得到平台實際印出來的字串。
"""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable
from typing import Final

__all__ = ["KNOWN_PREFIXES", "normalize_order_id", "normalize_product_name"]

# 各平台在結算單上會自行加上的前綴。新增平台時在這裡補一筆即可，
# 匹配演算法本身不需要改動。
KNOWN_PREFIXES: Final[tuple[str, ...]] = (
    "SP-",
    "SPE",
    "MO-",
    "PC-",
    "ORDER-",
    "#",
)

_INVISIBLE = re.compile(r"[\s​-‏⁠﻿]+")
_NON_ALNUM = re.compile(r"[^0-9A-Z]")


def normalize_order_id(raw: str | None, *, prefixes: Iterable[str] = KNOWN_PREFIXES) -> str | None:
    """把訂單編號正規化成可比對的鍵。

    步驟依序為：Unicode NFKC 正規化（全形轉半形）→ 去除空白與零寬字元 →
    轉大寫 → 剝除已知前綴 → 移除剩餘的非英數字元。

    回傳 ``None`` 代表這個編號無法作為匹配鍵（缺失或清洗後為空），
    呼叫端應把該筆送進 Stage 3，而不是當作空字串去比對——
    空字串會讓所有缺編號的記錄互相匹配，那是災難。

    >>> normalize_order_id("SP-24a7x9k2 ")
    '24A7X9K2'
    >>> normalize_order_id("＃２４Ａ７")
    '24A7'
    >>> normalize_order_id("   ") is None
    True
    """
    if raw is None:
        return None
    text = unicodedata.normalize("NFKC", raw)
    text = _INVISIBLE.sub("", text).upper()
    if not text:
        return None
    for prefix in sorted(prefixes, key=len, reverse=True):
        upper = unicodedata.normalize("NFKC", prefix).upper()
        if upper and text.startswith(upper):
            text = text[len(upper) :]
            break
    text = _NON_ALNUM.sub("", text)
    return text or None


def normalize_product_name(raw: str | None) -> str:
    """把商品名稱正規化，供 Stage 3 的相似度比對使用。

    只做保守的處理：NFKC、去除規格括號、壓縮連續空白、轉小寫。
    刻意不做斷詞或同義詞替換——那些會引入難以解釋的行為，
    而模糊匹配的結果是要拿給人看的。
    """
    if not raw:
        return ""
    text = unicodedata.normalize("NFKC", raw).lower()
    text = re.sub(r"[（(\[【].*?[)）\]】]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()
