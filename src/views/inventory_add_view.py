import flet as ft
import os
from datetime import datetime
from utils.path_utils import resource_path
from utils.inventory_utils import initialize_inventory_excel, add_inventory_entry, convert_existing_inventory_to_formulas


class InventoryAddView:
    """View for adding inventory items"""
    
    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - إضافة مخزون"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK
        
        # Initialize data storage
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        self.inventory_path = os.path.join(self.documents_path, "مخزون الادوات")
        os.makedirs(self.inventory_path, exist_ok=True)
        
        # Form fields
        self.date_field = ft.TextField(
            label="تاريخ الدخول",
            value=datetime.now().strftime('%d/%m/%Y'),
            width=150,
            read_only=True
        )
        
        self.item_name_field = ft.TextField(
            label="اسم الصنف",
            width=300,
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
            on_change=self.calculate_total
        )
        
        self.total_price_field = ft.TextField(
            label="الإجمالي",
            width=150,
            read_only=True,
            value="0"
        )
        
        self.notes_field = ft.TextField(
            label="ملاحظات",
            width=500,
            multiline=True,
            max_lines=3
        )
        
        # Bind events
        self.quantity_field.on_change = self.calculate_total

    def calculate_total(self, e):
        """Calculate total price based on quantity and unit price"""
        try:
            quantity = float(self.quantity_field.value) if self.quantity_field.value else 0
            unit_price = float(self.unit_price_field.value) if self.unit_price_field.value else 0
            total = quantity * unit_price
            self.total_price_field.value = str(total)
        except ValueError:
            self.total_price_field.value = "0"
        self.page.update()

    def show_dialog(self, title, message, color=ft.Colors.GREEN_400):
        """Show a dialog with the given message"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            title=ft.Text(title),
            content=ft.Text(message, rtl=True),
            actions=[
                ft.TextButton("حسناً", on_click=close_dlg)
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def save_inventory(self, e):
        """Save inventory entry to Excel file"""
        print(f"[DEBUG] Saving inventory entry")
        
        # Validate inputs
        if not self.item_name_field.value:
            self.show_dialog("تحذير", "الرجاء إدخال اسم الصنف", ft.Colors.RED_400)
            return
        
        if not self.quantity_field.value:
            self.show_dialog("تحذير", "الرجاء إدخال العدد", ft.Colors.RED_400)
            return
        
        try:
            quantity = float(self.quantity_field.value)
            if quantity <= 0:
                self.show_dialog("تحذير", "العدد يجب أن يكون أكبر من صفر", ft.Colors.RED_400)
                return
        except ValueError:
            self.show_dialog("تحذير", "الرجاء إدخال عدد صحيح", ft.Colors.RED_400)
            return
        
        try:
            unit_price = float(self.unit_price_field.value) if self.unit_price_field.value else 0
        except ValueError:
            self.show_dialog("تحذير", "الرجاء إدخال سعر صحيح", ft.Colors.RED_400)
            return
        
        # Save to Excel file using the utility function
        try:
            excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")
            
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
            
            # Add inventory entry using utility function
            print(f"[DEBUG] Adding inventory entry")
            entry_number = add_inventory_entry(
                file_path=excel_file,
                item_name=self.item_name_field.value,
                quantity=self.quantity_field.value,
                unit_price=self.unit_price_field.value,
                notes=self.notes_field.value,
                entry_date=self.date_field.value
            )
            print(f"[DEBUG] Entry added with number: {entry_number}")
            
            # Clear form
            self.item_name_field.value = ""
            self.quantity_field.value = ""
            self.unit_price_field.value = ""
            self.total_price_field.value = "0"
            self.notes_field.value = ""
            
            self.show_dialog("نجاح", f"تم حفظ إذن الإضافة برقم: {entry_number}")
            
        except Exception as e:
            print(f"[ERROR] Error saving inventory: {e}")
            self.show_dialog("خطأ", "حدث خطأ أثناء حفظ البيانات", ft.Colors.RED_400)

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
        """Build the inventory addition UI"""
        # Header with save button in AppBar
        self.page.appbar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK,
                on_click=self.go_back,
                tooltip="العودة"
            ),
            title=ft.Text("إضافة مخزون", size=20, weight=ft.FontWeight.BOLD),
            bgcolor=ft.Colors.BLUE_GREY_900,
            actions=[
                ft.FilledButton(
                    "حفظ",
                    icon=ft.Icons.SAVE,
                    on_click=self.save_inventory,
                    style=ft.ButtonStyle(
                        bgcolor=ft.Colors.GREEN_700,
                        color=ft.Colors.WHITE,
                    )
                ),
            ]
        )
        
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
                    
                    # Item name field
                    ft.Container(
                        content=ft.Row(
                            controls=[
                                ft.Icon(ft.Icons.INVENTORY_2, color=ft.Colors.ORANGE_300),
                                self.item_name_field,
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