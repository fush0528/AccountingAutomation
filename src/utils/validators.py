from datetime import datetime
from typing import Any, Union, Optional


def validate_date(date: Any) -> bool:
    """驗證日期格式"""
    return isinstance(date, datetime)


def validate_string(text: Any, allow_empty: bool = False) -> bool:
    """驗證字串格式"""
    if not isinstance(text, str):
        return False
    if not allow_empty and not text.strip():
        return False
    return True


def validate_number(value: Any, min_value: Optional[Union[int, float]] = None) -> bool:
    """驗證數字格式（整數或浮點數）"""
    if not isinstance(value, (int, float)):
        return False
    if min_value is not None and value < min_value:
        return False
    return True


def validate_boolean(value: Any) -> bool:
    """驗證布林值格式"""
    return isinstance(value, bool)


def validate_platform_fee(fee: Union[int, float], total_sales: Union[int, float]) -> bool:
    """驗證平台手續費是否合理（不能大於銷售總額）"""
    if not validate_number(fee, 0) or not validate_number(total_sales, 0):
        return False
    return fee <= total_sales