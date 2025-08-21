from datetime import datetime
from dataclasses import dataclass
from typing import Dict, Any


@dataclass
class AccountingEntry:
    """記帳項目的主要資料結構，包含所有必要的記帳資訊"""
    year: str
    month: str
    day: str
    time: str
    platform: str
    product_name: str
    order_quantity: int
    total_sales: float
    platform_fee: float
    actual_income: float = 0.0
    invoice_required: bool = False
    taxable: bool = True

    def __post_init__(self):
        """初始化後計算實收金額"""
        self.actual_income = self.calculate_actual_income()

    def calculate_actual_income(self) -> float:
        """計算實收金額（銷售總額減去平台手續費）"""
        return self.total_sales - self.platform_fee

    def to_dict(self) -> Dict[str, Any]:
        """將物件轉換為字典格式"""
        return {
            'year': self.year,
            'month': self.month,
            'day': self.day,
            'time': self.time,
            'platform': self.platform,
            'product_name': self.product_name,
            'order_quantity': self.order_quantity,
            'total_sales': self.total_sales,
            'platform_fee': self.platform_fee,
            'actual_income': self.actual_income,
            'invoice_required': self.invoice_required,
            'taxable': self.taxable
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'AccountingEntry':
        """從字典格式建立物件"""
        if 'date' in data:
            # 處理舊格式的日期
            dt = datetime.fromisoformat(data['date'])
            data['year'] = str(dt.year)
            data['month'] = str(dt.month).zfill(2)
            data['day'] = str(dt.day).zfill(2)
            data['time'] = dt.strftime('%H:%M:%S')
            del data['date']
        return cls(**data)

    def validate(self) -> bool:
        """驗證資料的正確性"""
        if not isinstance(self.year, str) or not self.year.strip():
            return False
        if not isinstance(self.month, str) or not self.month.strip():
            return False
        if not isinstance(self.day, str) or not self.day.strip():
            return False
        if not isinstance(self.time, str) or not self.time.strip():
            return False
        if not isinstance(self.platform, str) or not self.platform.strip():
            return False
        if not isinstance(self.product_name, str) or not self.product_name.strip():
            return False
        if not isinstance(self.order_quantity, int) or self.order_quantity < 0:
            return False
        if not isinstance(self.total_sales, (int, float)) or self.total_sales < 0:
            return False
        if not isinstance(self.platform_fee, (int, float)) or self.platform_fee < 0:
            return False
        if self.platform_fee > self.total_sales:
            return False
        return True