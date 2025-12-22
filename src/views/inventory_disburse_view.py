import flet as ft
import os
import traceback
from datetime import datetime
from utils.path_utils import resource_path
from utils.inventory_utils import initialize_inventory_excel, disburse_inventory_entry, get_inventory_summary, get_available_items_with_prices, convert_existing_inventory_to_formulas


class InventoryDisburseView:
    """View for disbursing inventory items"""
    
    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - صرف مخزون"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK
        
        # Initialize data storage
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        self.inventory_path = os.path.join(self.documents_path, "مخزون الادوات")
        os.makedirs(self.inventory_path, exist_ok=True)
        
        print(f"[DEBUG] Documents path: {self.documents_path}")
        print(f"[DEBUG] Inventory path: {self.inventory_path}")
        
        # Form fields
        self.date_field = ft.TextField(
            label="تاريخ الصرف",
            value=datetime.now().strftime('%d/%m/%Y'),
            width=150,
            read_only=True
        )
        
        # Get available items for dropdown
        excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")
        print(f"[DEBUG] Excel file path: {excel_file}")
        available_items = self._get_available_items(excel_file)
        print(f"[DEBUG] Available items count: {len(available_items)}")
        print(f"[DEBUG] Available items: {available_items}")
        
        self.item_dropdown = ft.Dropdown(
            label="اسم الصنف",
            width=300,
            options=[ft.dropdown.Option(item) for item in available_items] if available_items else [ft.dropdown.Option("لا توجد أصناف متوفرة")],
            on_change=self.on_item_selected
        )
        
        self.quantity_field = ft.TextField(
            label="العدد",
            width=150,
            keyboard_type=ft.KeyboardType.NUMBER
        )
        
        self.unit_price_field = ft.TextField(
            label="ثمن الوحدة",
            width=150,
            keyboard_type=ft.KeyboardType.NUMBER,
            read_only=True  # Will be populated automatically based on selection
        )
        
        self.total_price_field = ft.TextField(
            label="الإجمالي",
            width=150,
            read_only=True
        )
        
        self.notes_field = ft.TextField(
            label="ملاحظات",
            width=300,
            multiline=True,
            min_lines=2,
            max_lines=3
        )
        
        # Bind events
        self.quantity_field.on_change = self.calculate_total

    def _get_available_items(self, excel_file):
        """Helper method to get available items with proper error handling"""
        print(f"[DEBUG] _get_available_items called with file: {excel_file}")
        available_items = []
        try:
            if os.path.exists(excel_file):
                print(f"[DEBUG] Excel file exists")
                # Convert existing file to use formulas
                try:
                    convert_existing_inventory_to_formulas(excel_file)
                    print(f"[DEBUG] Converted to formulas successfully")
                except Exception as e:
                    print(f"[ERROR] Failed to convert to formulas: {e}")
                    traceback.print_exc()
                
                # Get inventory data with calculated values
                try:
                    inventory_data = get_inventory_summary(excel_file)
                    print(f"[DEBUG] Inventory data retrieved, count: {len(inventory_data)}")
                    print(f"[DEBUG] Inventory data: {inventory_data}")
                    # Show all items, not just those with positive balance
                    # This allows users to see all items and we'll validate quantities during save
                    available_items = [item['item_name'] for item in inventory_data if item['item_name']]
                    print(f"[DEBUG] Filtered available items: {available_items}")
                except Exception as e:
                    print(f"[ERROR] Failed to get inventory summary: {e}")
                    traceback.print_exc()
            else:
                print(f"[DEBUG] Excel file does not exist")
        except Exception as e:
            print(f"[ERROR] Error in _get_available_items: {e}")
            traceback.print_exc()
        return available_items

    def on_item_selected(self, e):
        """Handle item selection from dropdown"""
        selected_item = self.item_dropdown.value
        print(f"[DEBUG] Item selected: {selected_item}")
        if selected_item and selected_item != "لا توجد أصناف متوفرة":
            # Get the unit price for the selected item
            excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")
            print(f"[DEBUG] Getting price for item: {selected_item}")
            if os.path.exists(excel_file):
                try:
                    item_prices = get_available_items_with_prices(excel_file)
                    print(f"[DEBUG] Item prices: {item_prices}")
                    if selected_item in item_prices:
                        price = round(item_prices[selected_item], 2)
                        self.unit_price_field.value = str(price)
                        print(f"[DEBUG] Set unit price: {price}")
                    else:
                        print(f"[DEBUG] Item {selected_item} not found in prices")
                        self.unit_price_field.value = "0"
                except Exception as e:
                    print(f"[ERROR] Error getting item prices: {e}")
                    traceback.print_exc()
                    self.unit_price_field.value = "0"
            else:
                print(f"[DEBUG] Excel file not found for price lookup")
        else:
            print(f"[DEBUG] No valid item selected")
            self.unit_price_field.value = ""
        self.page.update()

    def calculate_total(self, e):
        """Calculate total price based on quantity and unit price"""
        print(f"[DEBUG] Calculating total")
        try:
            quantity = float(self.quantity_field.value) if self.quantity_field.value else 0
            unit_price = float(self.unit_price_field.value) if self.unit_price_field.value else 0
            total = quantity * unit_price
            self.total_price_field.value = str(total)
            print(f"[DEBUG] Quantity: {quantity}, Unit Price: {unit_price}, Total: {total}")
        except ValueError as e:
            print(f"[ERROR] Value error in calculate_total: {e}")
            self.total_price_field.value = "0"
        except Exception as e:
            print(f"[ERROR] Unexpected error in calculate_total: {e}")
            traceback.print_exc()
            self.total_price_field.value = "0"
        self.page.update()

    def go_back(self, e):
        """Navigate back to main dashboard"""
        print(f"[DEBUG] Going back")
        if self.on_back:
            self.on_back()
        else:
            # If no back callback, go back to main dashboard
            from views.dashboard_view import DashboardView
            self.page.clean()
            self.page.appbar = None
            dashboard = DashboardView(self.page)
            dashboard.show(getattr(self.page, '_save_callback', None))

    def build_ui(self):
        """Build the inventory disbursement UI"""
        print(f"[DEBUG] Building UI")
        # Header with save button in AppBar
        self.page.appbar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK,
                on_click=self.go_back,
                tooltip="العودة"
            ),
            title=ft.Text("صرف مخزون", size=20, weight=ft.FontWeight.BOLD),
            bgcolor=ft.Colors.BLUE_GREY_900,
            actions=[
                ft.FilledButton(
                    "حفظ",
                    icon=ft.Icons.SAVE,
                    on_click=self.save_inventory,
                    style=ft.ButtonStyle(
                        bgcolor=ft.Colors.RED_700,
                        color=ft.Colors.WHITE,
                    )
                ),
            ]
        )
        
        # Refresh dropdown with current available items
        self.refresh_item_dropdown()
        
        # Form section with improved layout
        form_section = ft.Container(
            content=ft.Column(
                controls=[
                    # Date field
                    ft.Container(
                        content=ft.Row(
                            controls=[
                                ft.Icon(ft.Icons.CALENDAR_TODAY, color=ft.Colors.BLUE_300),
                                self.date_field,
                            ],
                            spacing=10,
                            alignment=ft.MainAxisAlignment.CENTER
                        ),
                        padding=10,
                        border=ft.border.all(1, ft.Colors.GREY_700),
                        border_radius=10,
                        width=300
                    ),
                    
                    # Item dropdown field
                    ft.Container(
                        content=ft.Row(
                            controls=[
                                ft.Icon(ft.Icons.INVENTORY_2, color=ft.Colors.ORANGE_300),
                                self.item_dropdown,
                            ],
                            spacing=10,
                            alignment=ft.MainAxisAlignment.CENTER
                        ),
                        padding=10,
                        border=ft.border.all(1, ft.Colors.GREY_700),
                        border_radius=10,
                        width=500
                    ),
                    
                    # Quantity and unit price fields in one row
                    ft.Row(
                        controls=[
                            ft.Container(
                                content=ft.Row(
                                    controls=[
                                        ft.Icon(ft.Icons.NUMBERS, color=ft.Colors.PURPLE_300),
                                        self.quantity_field,
                                    ],
                                    spacing=10,
                                    alignment=ft.MainAxisAlignment.CENTER
                                ),
                                padding=10,
                                border=ft.border.all(1, ft.Colors.GREY_700),
                                border_radius=10,
                            ),
                            ft.Container(
                                content=ft.Row(
                                    controls=[
                                        ft.Icon(ft.Icons.PRICE_CHANGE, color=ft.Colors.YELLOW_300),
                                        self.unit_price_field,
                                    ],
                                    spacing=10,
                                    alignment=ft.MainAxisAlignment.CENTER
                                ),
                                padding=10,
                                border=ft.border.all(1, ft.Colors.GREY_700),
                                border_radius=10,
                            ),
                        ],
                        spacing=20,
                        alignment=ft.MainAxisAlignment.CENTER
                    ),
                    
                    # Total price field
                    ft.Container(
                        content=ft.Row(
                            controls=[
                                ft.Icon(ft.Icons.ATTACH_MONEY, color=ft.Colors.GREEN_300),
                                self.total_price_field,
                            ],
                            spacing=10,
                            alignment=ft.MainAxisAlignment.CENTER
                        ),
                        padding=10,
                        border=ft.border.all(1, ft.Colors.GREY_700),
                        border_radius=10,
                        width=300
                    ),
                    
                    # Notes field
                    ft.Container(
                        content=ft.Row(
                            controls=[
                                ft.Icon(ft.Icons.NOTES, color=ft.Colors.CYAN_300),
                                self.notes_field,
                            ],
                            spacing=10,
                            alignment=ft.MainAxisAlignment.CENTER
                        ),
                        padding=10,
                        border=ft.border.all(1, ft.Colors.GREY_700),
                        border_radius=10,
                        width=500
                    ),
                ],
                spacing=20,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            ),
            padding=20,
            bgcolor=ft.Colors.GREY_900,
            border_radius=15,
            margin=ft.margin.only(bottom=10)
        )
        
        # Main layout
        main_layout = ft.Column(
            controls=[
                form_section,
            ],
            spacing=15,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )
        
        self.page.clean()
        self.page.add(main_layout)
        self.page.update()

    def refresh_item_dropdown(self):
        """Refresh the item dropdown with current available items"""
        print(f"[DEBUG] Refreshing item dropdown")
        excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")
        available_items = self._get_available_items(excel_file)
        print(f"[DEBUG] Refresh - Available items count: {len(available_items)}")
        print(f"[DEBUG] Refresh - Available items: {available_items}")
        
        # Update dropdown options
        if available_items:
            self.item_dropdown.options = [ft.dropdown.Option(item) for item in available_items]
            self.item_dropdown.value = None  # Clear selection
        else:
            self.item_dropdown.options = [ft.dropdown.Option("لا توجد أصناف في المخزون")]
            self.item_dropdown.value = "لا توجد أصناف في المخزون"
        print(f"[DEBUG] Dropdown options updated")
        self.page.update()

    def save_inventory(self, e):
        """Save inventory disbursement record to Excel file"""
        print(f"[DEBUG] Saving inventory")
        # Validate required fields
        if not self.item_dropdown.value or self.item_dropdown.value in ["لا توجد أصناف متوفرة", "لا توجد أصناف في المخزون"]:
            print(f"[DEBUG] No item selected")
            self.show_dialog("خطأ", "يرجى اختيار صنف من القائمة", ft.Colors.RED_400)
            return
            
        if not self.quantity_field.value:
            print(f"[DEBUG] No quantity entered")
            self.show_dialog("خطأ", "يرجى إدخال عدد الصنف", ft.Colors.RED_400)
            return
            
        # Validate quantity doesn't exceed available balance
        excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")
        if os.path.exists(excel_file):
            try:
                inventory_data = get_inventory_summary(excel_file)
                print(f"[DEBUG] Inventory data for validation: {inventory_data}")
                item_found = False
                for item in inventory_data:
                    if item['item_name'] == self.item_dropdown.value:
                        item_found = True
                        available_balance = float(item['current_balance'])
                        requested_quantity = float(self.quantity_field.value)
                        
                        # Check if item has any balance
                        if available_balance <= 0:
                            print(f"[DEBUG] Item has no available balance: {available_balance}")
                            self.show_dialog("خطأ", f"الصنف المحدد ({self.item_dropdown.value}) ليس له رصيد متوفر", ft.Colors.RED_400)
                            return
                        
                        # Check if requested quantity exceeds available balance
                        if requested_quantity > available_balance:
                            print(f"[DEBUG] Insufficient balance - Requested: {requested_quantity}, Available: {available_balance}")
                            self.show_dialog("خطأ", f"الكمية المطلوبة ({requested_quantity}) تتجاوز الرصيد المتاح ({available_balance})", ft.Colors.RED_400)
                            return
                        break
                
                # If item not found in inventory data
                if not item_found:
                    print(f"[DEBUG] Item not found in inventory data")
                    self.show_dialog("خطأ", f"الصنف المحدد ({self.item_dropdown.value}) غير موجود في المخزون", ft.Colors.RED_400)
                    return
            except Exception as e:
                print(f"[ERROR] Error validating inventory: {e}")
                traceback.print_exc()
                self.show_dialog("خطأ", "حدث خطأ أثناء التحقق من الرصيد", ft.Colors.RED_400)
                return
        
        # Save to Excel file using the utility function
        try:
            # Initialize Excel file if it doesn't exist
            if not os.path.exists(excel_file):
                print(f"[DEBUG] Initializing new Excel file")
                initialize_inventory_excel(excel_file)
            else:
                # Convert existing file to use formulas
                try:
                    convert_existing_inventory_to_formulas(excel_file)
                except:
                    pass  # If conversion fails, continue with existing data
            
            # Add disbursement entry using utility function
            print(f"[DEBUG] Adding disbursement entry")
            entry_number = disburse_inventory_entry(
                file_path=excel_file,
                item_name=self.item_dropdown.value,
                quantity=self.quantity_field.value,
                unit_price=self.unit_price_field.value,
                notes=self.notes_field.value,
                disburse_date=self.date_field.value
            )
            print(f"[DEBUG] Entry added with number: {entry_number}")
            
            # Clear form
            self.item_dropdown.value = None
            self.quantity_field.value = ""
            self.unit_price_field.value = ""
            self.total_price_field.value = ""
            self.notes_field.value = ""
            
            # Refresh dropdown
            self.refresh_item_dropdown()
            
            self.page.update()
            self.show_success_dialog(excel_file)
            
        except PermissionError as e:
            print(f"[ERROR] Permission error: {e}")
            self.show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
        except Exception as e:
            print(f"[ERROR] Unexpected error saving inventory: {e}")
            traceback.print_exc()
            self.show_dialog("خطأ", f"حدث خطأ أثناء حفظ البيانات: {str(e)}", ft.Colors.RED_400)

    def show_success_dialog(self, file_path):
        """Show success dialog with file path"""
        print(f"[DEBUG] Showing success dialog")
        def close_dlg(e):
            dlg.open = False
            self.page.update()
            
        def open_file(e):
            dlg.open = False
            self.page.update()
            os.startfile(file_path)
            
        def open_folder(e):
            dlg.open = False
            self.page.update()
            os.startfile(os.path.dirname(file_path))
            
        dlg = ft.AlertDialog(
            title=ft.Text("تم الحفظ بنجاح", color=ft.Colors.GREEN_400),
            content=ft.Text("تم حفظ بيانات الصرف إلى ملف Excel"),
            actions=[
                ft.TextButton(
                    "فتح الملف",
                    on_click=open_file,
                    icon=ft.Icons.FILE_OPEN,
                    style=ft.ButtonStyle(color=ft.Colors.GREEN_300)
                ),
                ft.TextButton(
                    "فتح المجلد",
                    on_click=open_folder,
                    icon=ft.Icons.FOLDER_OPEN,
                    style=ft.ButtonStyle(color=ft.Colors.BLUE_300)
                ),
                ft.TextButton("حسناً", on_click=close_dlg)
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def show_dialog(self, title, message, color):
        """Show a dialog with a message"""
        print(f"[DEBUG] Showing dialog: {title} - {message}")
        def close_dlg(e):
            dlg.open = False
            self.page.update()
            
        dlg = ft.AlertDialog(
            title=ft.Text(title, color=color),
            content=ft.Text(message),
            actions=[
                ft.TextButton("حسناً", on_click=close_dlg)
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()