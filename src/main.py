"""
مصنع السويفي - نظام الإدارة
تطبيق رسومي لإدارة المصنع والفواتير.
"""

import sys
import os
import traceback
import flet as ft
from utils.path_utils import resource_path

# Version information
try:
    from version import __version__
except ImportError:
    __version__ = "1.0"

# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Set locale for Arabic support
import locale
try:
    # Try to set Arabic locale
    locale.setlocale(locale.LC_ALL, 'ar_SA.UTF-8')
except:
    pass

from views.dashboard_view import DashboardView
from utils.excel_utils import save_invoice, update_client_ledger


def save_callback(filepath, op_num, client, driver, date_str, phone, items):
    """
    دالة رد الاتصال لحفظ بيانات الفاتورة إلى Excel.
    
    Args:
        filepath (str): المسار لحفظ ملف Excel
        op_num (str): رقم العملية/الفاتورة
        client (str): اسم العميل
        driver (str): اسم السائق
        date_str (str): سلسلة التاريخ
        phone (str): رقم الهاتف
        items (list): قائمة عناصر الفاتورة
    """
    # Save the invoice
    save_invoice(filepath, op_num, client, driver, items, date_str=date_str, phone=phone)
    
    # Skip ledger update for "ايراد" clients
    if "ايراد" in client:
        return
    
    # Create/update client ledger
    try:
        # Calculate total amount from items
        total_amount = 0
        invoice_items_details = []
        
        # Process items to get details for the ledger
        for item in items:
            try:
                desc = item[0] or ""
                block = item[1] or ""
                thickness = item[2] or ""
                material = item[3] or ""
                count = float(item[4])
                length = float(item[5])
                height = float(item[6])
                price_val = float(item[7])
                
                # Calculate area and total for this item
                area = count * length * height
                total = area * price_val
                
                # Add to total amount
                total_amount += total
                
                # Store item details for the ledger
                invoice_items_details.append((
                    desc,
                    material,
                    thickness,
                    area,
                    total
                ))
            except (ValueError, IndexError):
                continue
        
        # Get the client folder path (parent of the invoice folder)
        client_folder = os.path.dirname(os.path.dirname(filepath))
        
        # Update or create the client's ledger
        success, error = update_client_ledger(
            client_folder, 
            client, 
            date_str, 
            op_num, 
            total_amount, 
            driver, 
            invoice_items_details
        )
        
        if not success:
            print(f"Warning: Could not update client ledger: {error}")
    except Exception as e:
        print(f"Error updating client ledger: {e}")
        traceback.print_exc()


def main(page: ft.Page):
    """نقطة الدخول الرئيسية للتطبيق."""
    try:
        # Configure window properties
        page.title = f"مصنع السويفي - الإصدار {__version__}"
        
        # Set window to maximized
        page.window.maximized = True
        
        # Set app icon
        icon_path = resource_path(os.path.join("assets", "icon.ico"))
        page.window.icon = icon_path
        
        # Create and show the Dashboard
        dashboard = DashboardView(page)
        dashboard.show(save_callback)
        
    except Exception as e:
        # Log the full traceback for debugging
        error_msg = f"حدث خطأ غير متوقع:\n{str(e)}\n\n"
        error_msg += "تفاصيل:\n" + traceback.format_exc()
        
        # Show error dialog
        def close_dlg(e):
            dlg.open = False
            page.update()
        
        dlg = ft.AlertDialog(
            title=ft.Text("خطأ في التطبيق"),
            content=ft.Text(error_msg, rtl=True),
            actions=[ft.TextButton("موافق", on_click=close_dlg)]
        )
        page.overlay.append(dlg)
        dlg.open = True
        page.update()
        sys.exit(1)


if __name__ == "__main__":
    # Run as desktop app with assets [directory]
    ft.app(target=main, view=ft.AppView.FLET_APP, assets_dir='assets')