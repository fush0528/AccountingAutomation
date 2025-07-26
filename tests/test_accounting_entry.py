import unittest
from datetime import datetime
from src.models import AccountingEntry


class TestAccountingEntry(unittest.TestCase):
    """AccountingEntry 類別的單元測試"""

    def setUp(self):
        """設定測試環境"""
        self.test_date = datetime.now()
        self.valid_entry = AccountingEntry(
            date=self.test_date,
            platform="蝦皮",
            product_name="測試商品",
            order_quantity=1,
            total_sales=100.0,
            platform_fee=10.0
        )

    def test_init(self):
        """測試初始化"""
        self.assertEqual(self.valid_entry.date, self.test_date)
        self.assertEqual(self.valid_entry.platform, "蝦皮")
        self.assertEqual(self.valid_entry.product_name, "測試商品")
        self.assertEqual(self.valid_entry.order_quantity, 1)
        self.assertEqual(self.valid_entry.total_sales, 100.0)
        self.assertEqual(self.valid_entry.platform_fee, 10.0)
        self.assertEqual(self.valid_entry.actual_income, 90.0)
        self.assertFalse(self.valid_entry.invoice_required)
        self.assertTrue(self.valid_entry.taxable)

    def test_calculate_actual_income(self):
        """測試實收金額計算"""
        self.assertEqual(self.valid_entry.calculate_actual_income(), 90.0)

    def test_to_dict(self):
        """測試轉換為字典格式"""
        entry_dict = self.valid_entry.to_dict()
        self.assertEqual(entry_dict['date'], self.test_date.isoformat())
        self.assertEqual(entry_dict['platform'], "蝦皮")
        self.assertEqual(entry_dict['product_name'], "測試商品")
        self.assertEqual(entry_dict['order_quantity'], 1)
        self.assertEqual(entry_dict['total_sales'], 100.0)
        self.assertEqual(entry_dict['platform_fee'], 10.0)
        self.assertEqual(entry_dict['actual_income'], 90.0)
        self.assertFalse(entry_dict['invoice_required'])
        self.assertTrue(entry_dict['taxable'])

    def test_from_dict(self):
        """測試從字典格式建立物件"""
        entry_dict = self.valid_entry.to_dict()
        new_entry = AccountingEntry.from_dict(entry_dict)
        self.assertEqual(new_entry.date.isoformat(), self.test_date.isoformat())
        self.assertEqual(new_entry.platform, self.valid_entry.platform)
        self.assertEqual(new_entry.product_name, self.valid_entry.product_name)
        self.assertEqual(new_entry.order_quantity, self.valid_entry.order_quantity)
        self.assertEqual(new_entry.total_sales, self.valid_entry.total_sales)
        self.assertEqual(new_entry.platform_fee, self.valid_entry.platform_fee)
        self.assertEqual(new_entry.actual_income, self.valid_entry.actual_income)
        self.assertEqual(new_entry.invoice_required, self.valid_entry.invoice_required)
        self.assertEqual(new_entry.taxable, self.valid_entry.taxable)

    def test_validate(self):
        """測試資料驗證"""
        # 測試有效資料
        self.assertTrue(self.valid_entry.validate())

        # 測試無效日期
        invalid_entry = AccountingEntry(
            date="2023-01-01",  # 錯誤的日期格式
            platform="蝦皮",
            product_name="測試商品",
            order_quantity=1,
            total_sales=100.0,
            platform_fee=10.0
        )
        self.assertFalse(invalid_entry.validate())

        # 測試無效的訂單數量
        invalid_entry = AccountingEntry(
            date=self.test_date,
            platform="蝦皮",
            product_name="測試商品",
            order_quantity=-1,  # 負數訂單數量
            total_sales=100.0,
            platform_fee=10.0
        )
        self.assertFalse(invalid_entry.validate())

        # 測試手續費大於銷售額
        invalid_entry = AccountingEntry(
            date=self.test_date,
            platform="蝦皮",
            product_name="測試商品",
            order_quantity=1,
            total_sales=100.0,
            platform_fee=150.0  # 手續費大於銷售額
        )
        self.assertFalse(invalid_entry.validate())


if __name__ == '__main__':
    unittest.main()