from datetime import datetime
from dataclasses import dataclass
from typing import Dict, Any


@dataclass
class AccountingEntry:
    """記帳項目的主要資料結構，包含所有必要的記帳資訊"""
    date: datetime
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
            'date': self.date.isoformat(),
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
        # 將 ISO 格式的日期字串轉換為 datetime 物件
        data['date'] = datetime.fromisoformat(data['date'])
        return cls(**data)

    def validate(self) -> bool:
        """驗證資料的正確性"""
        if not isinstance(self.date, datetime):
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