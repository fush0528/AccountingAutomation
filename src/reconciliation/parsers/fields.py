"""欄位層級的解析工具，給各平台 parser 共用。

放在這裡而不是各自實作，是因為「日期怎麼寫」這類問題會重複出現，
但又不屬於領域層——領域層只認識 :class:`datetime.date`，
不需要知道有人會用民國年。
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

from .base import ParseError

__all__ = ["clean_cell", "parse_date", "parse_int", "parse_roc_date"]

_DATE_PATTERNS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y%m%d",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
)

_ROC_PATTERN = re.compile(r"^(\d{2,3})[/\-.](\d{1,2})[/\-.](\d{1,2})$")


def clean_cell(value: object) -> str:
    """把儲存格內容收斂成乾淨的字串。

    NFKC 把全形轉半形，順便處理 Excel 讀出來的 None 與數值型別。
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    text = str(value)
    return unicodedata.normalize("NFKC", text).strip()


def parse_date(value: str, *, field: str = "日期") -> date:
    """解析西元日期，容忍常見的幾種寫法。"""
    text = clean_cell(value)
    if not text:
        raise ParseError(f"{field}為空")
    for pattern in _DATE_PATTERNS:
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    raise ParseError(f"{field} {text!r} 不是可辨識的日期格式")


def parse_roc_date(value: str, *, field: str = "日期") -> date:
    """解析民國年日期，例如 ``115/03/15`` 代表 2026-03-15。

    民國年加 1911 等於西元年。兩位數與三位數都要支援：民國 99 年
    與民國 115 年在同一份歷史資料裡並存是常見的事。
    """
    text = clean_cell(value)
    if not text:
        raise ParseError(f"{field}為空")
    match = _ROC_PATTERN.match(text)
    if match is None:
        # 有些平台在同一欄混用西元，先讓西元的邏輯試一次
        return parse_date(text, field=field)
    year, month, day = (int(g) for g in match.groups())
    if year > 200:  # 看起來已經是西元了
        raise ParseError(f"{field} {text!r} 的年份 {year} 不像民國年")
    try:
        return date(year + 1911, month, day)
    except ValueError as exc:
        raise ParseError(f"{field} {text!r} 不是有效的民國年日期") from exc


def parse_int(value: str, *, field: str = "數量", default: int | None = None) -> int:
    """解析整數，容忍千分位逗號與全形數字。"""
    text = clean_cell(value).replace(",", "")
    if not text:
        if default is not None:
            return default
        raise ParseError(f"{field}為空")
    try:
        return int(float(text))  # 容忍 "1.0" 這種來自 Excel 的寫法
    except ValueError as exc:
        raise ParseError(f"{field} {text!r} 不是整數") from exc
