from datetime import datetime
from src.handlers import ExcelHandler
from src.models import AccountingEntry
from openpyxl import Workbook
import os


def initialize_excel_file(file_path: str) -> bool:
    """初始化 Excel 檔案，建立標題列"""
    try:
        if not os.path.exists(file_path):
            workbook = Workbook()
            worksheet = workbook.active
            headers = [
                'date', 'platform', 'product_name', 'order_quantity',
                'total_sales', 'platform_fee', 'actual_income',
                'invoice_required', 'taxable'
            ]
            worksheet.append(headers)
            workbook.save(file_path)
            print(f"已建立新的記帳檔案：{file_path}")
        return True
    except Exception as e:
        print(f"建立 Excel 檔案時發生錯誤：{e}")
        return False


def main():
    """記帳自動化系統主程式"""
    file_path = "AccountingAutomation.xlsx"
    
    # 確保 Excel 檔案存在且格式正確
    if not initialize_excel_file(file_path):
        print("錯誤：無法初始化 Excel 檔案")
        return

    # 初始化 Excel 處理器
    handler = ExcelHandler(file_path)
    
    try:
        # 載入工作簿
        if not handler.load_workbook():
            print("錯誤：無法載入工作簿")
            return

        while True:
            print("\n=== 記帳自動化系統 ===")
            print("1. 新增記帳項目")
            print("2. 檢視所有記帳")
            print("3. 修改記帳項目")
            print("4. 刪除記帳項目")
            print("0. 離開系統")
            
            try:
                choice = input("\n請選擇操作 (0-4): ").strip()
            except EOFError:
                break
                
            if choice == "0":
                break
            elif choice == "1":
                add_entry(handler)
            elif choice == "2":
                view_entries(handler)
            elif choice == "3":
                update_entry(handler)
            elif choice == "4":
                delete_entry(handler)
            else:
                print("無效的選擇，請重試")

    except KeyboardInterrupt:
        print("\n使用者中斷程式")
    except Exception as e:
        print(f"發生錯誤：{e}")
    finally:
        # 確保變更被保存
        handler.save_workbook()
        print("\n系統已關閉")


def add_entry(handler: ExcelHandler):
    """新增記帳項目"""
    try:
        print("\n=== 新增記帳項目 ===")
        print("請依序輸入以下資料：")
        print("格式範例：")
        print("日期：2023-08-19")
        print("平台名稱：蝦皮")
        print("商品名稱：測試商品")
        print("訂單數量：2")
        print("銷售總額：1000")
        print("平台手續費：100")
        print("是否需要開發票：y")
        print("是否課稅：y")
        print("-" * 30)
        
        try:
            date_str = input("請輸入日期 (YYYY-MM-DD): ").strip()
            date = datetime.strptime(date_str, "%Y-%m-%d")
            
            platform = input("請輸入平台名稱: ").strip()
            product_name = input("請輸入商品名稱: ").strip()
            order_quantity = int(input("請輸入訂單數量: ").strip())
            total_sales = float(input("請輸入銷售總額: ").strip())
            platform_fee = float(input("請輸入平台手續費: ").strip())
            invoice_required = input("是否需要開發票 (y/n): ").strip().lower() == 'y'
            taxable = input("是否課稅 (y/n): ").strip().lower() == 'y'

            entry = AccountingEntry(
                date=date,
                platform=platform,
                product_name=product_name,
                order_quantity=order_quantity,
                total_sales=total_sales,
                platform_fee=platform_fee,
                invoice_required=invoice_required,
                taxable=taxable
            )

            if entry.validate():
                if handler.add_entry(entry):
                    print("記帳項目新增成功！")
                else:
                    print("錯誤：無法新增記帳項目")
            else:
                print("錯誤：資料驗證失敗")

        except ValueError as e:
            print(f"錯誤：輸入格式不正確 - {e}")
        except KeyboardInterrupt:
            print("\n取消新增操作")
        except Exception as e:
            print(f"錯誤：{e}")

    except Exception as e:
        print(f"錯誤：{e}")


def view_entries(handler: ExcelHandler):
    """檢視所有記帳項目"""
    print("\n=== 所有記帳項目 ===")
    entries = handler.read_entries()
    
    if not entries:
        print("目前沒有任何記帳項目")
        return

    for i, entry in enumerate(entries, 1):
        print(f"\n--- 項目 {i} ---")
        print(f"日期：{entry.date.strftime('%Y-%m-%d')}")
        print(f"平台：{entry.platform}")
        print(f"商品：{entry.product_name}")
        print(f"數量：{entry.order_quantity}")
        print(f"銷售額：{entry.total_sales}")
        print(f"手續費：{entry.platform_fee}")
        print(f"實收金額：{entry.actual_income}")
        print(f"需要發票：{'是' if entry.invoice_required else '否'}")
        print(f"課稅：{'是' if entry.taxable else '否'}")
        print("-" * 30)


def update_entry(handler: ExcelHandler):
    """修改記帳項目"""
    view_entries(handler)
    
    try:
        row_index = int(input("\n請輸入要修改的項目編號: ")) + 1  # +1 是因為 Excel 有標題列
        entry = handler.get_entry_by_index(row_index)
        
        if not entry:
            print("錯誤：找不到指定的記帳項目")
            return

        print("\n=== 修改記帳項目 ===")
        print("請輸入新的資料（直接按 Enter 保持原值）：")

        try:
            date_str = input(f"日期 ({entry.date.strftime('%Y-%m-%d')}): ").strip()
            date = datetime.strptime(date_str, "%Y-%m-%d") if date_str else entry.date
            
            platform = input(f"平台名稱 ({entry.platform}): ").strip()
            platform = platform if platform else entry.platform
            
            product_name = input(f"商品名稱 ({entry.product_name}): ").strip()
            product_name = product_name if product_name else entry.product_name
            
            order_quantity_str = input(f"訂單數量 ({entry.order_quantity}): ").strip()
            order_quantity = int(order_quantity_str) if order_quantity_str else entry.order_quantity
            
            total_sales_str = input(f"銷售總額 ({entry.total_sales}): ").strip()
            total_sales = float(total_sales_str) if total_sales_str else entry.total_sales
            
            platform_fee_str = input(f"平台手續費 ({entry.platform_fee}): ").strip()
            platform_fee = float(platform_fee_str) if platform_fee_str else entry.platform_fee
            
            invoice_required_str = input(f"是否需要開發票 ({'y' if entry.invoice_required else 'n'}): ").strip()
            invoice_required = (invoice_required_str.lower() == 'y') if invoice_required_str else entry.invoice_required
            
            taxable_str = input(f"是否課稅 ({'y' if entry.taxable else 'n'}): ").strip()
            taxable = (taxable_str.lower() == 'y') if taxable_str else entry.taxable

            updated_entry = AccountingEntry(
                date=date,
                platform=platform,
                product_name=product_name,
                order_quantity=order_quantity,
                total_sales=total_sales,
                platform_fee=platform_fee,
                invoice_required=invoice_required,
                taxable=taxable
            )

            if updated_entry.validate():
                if handler.update_entry(row_index, updated_entry):
                    print("記帳項目更新成功！")
                else:
                    print("錯誤：無法更新記帳項目")
            else:
                print("錯誤：資料驗證失敗")

        except ValueError as e:
            print(f"錯誤：輸入格式不正確 - {e}")
        except KeyboardInterrupt:
            print("\n取消修改操作")
        except Exception as e:
            print(f"錯誤：{e}")

    except ValueError as e:
        print(f"錯誤：輸入格式不正確 - {e}")
    except KeyboardInterrupt:
        print("\n取消修改操作")
    except Exception as e:
        print(f"錯誤：{e}")


def delete_entry(handler: ExcelHandler):
    """刪除記帳項目"""
    view_entries(handler)
    
    try:
        row_index = int(input("\n請輸入要刪除的項目編號: ")) + 1  # +1 是因為 Excel 有標題列
        
        confirm = input("確定要刪除這個項目嗎？(y/n): ").strip()
        if confirm.lower() == 'y':
            if handler.delete_entry(row_index):
                print("記帳項目刪除成功！")
            else:
                print("錯誤：刪除失敗")
        else:
            print("取消刪除操作")

    except ValueError as e:
        print(f"錯誤：輸入格式不正確 - {e}")
    except KeyboardInterrupt:
        print("\n取消刪除操作")
    except Exception as e:
        print(f"錯誤：{e}")


if __name__ == "__main__":
    main()