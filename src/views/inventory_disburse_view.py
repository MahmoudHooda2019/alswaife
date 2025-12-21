import flet as ft
import os
from datetime import datetime
from utils.path_utils import resource_path
from utils.inventory_utils import initialize_inventory_excel, disburse_inventory_entry


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
        
        # Form fields
        self.date_field = ft.TextField(
            label="تاريخ الصرف",
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
            keyboard_type=ft.KeyboardType.NUMBER
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
        self.unit_price_field.on_change = self.calculate_total

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

    def go_back(self, e):
        """Navigate back to main dashboard"""
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

    def save_inventory(self, e):
        """Save inventory disbursement record to Excel file"""
        # Validate required fields
        if not self.item_name_field.value or not self.quantity_field.value or not self.unit_price_field.value:
            self.show_dialog("خطأ", "يرجى ملء جميع الحقول المطلوبة", ft.Colors.RED_400)
            return
            
        # Save to Excel file using the utility function
        excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")
        
        try:
            # Initialize Excel file if it doesn't exist
            if not os.path.exists(excel_file):
                initialize_inventory_excel(excel_file)
            
            # Add disbursement entry using utility function
            entry_number = disburse_inventory_entry(
                file_path=excel_file,
                item_name=self.item_name_field.value,
                quantity=self.quantity_field.value,
                unit_price=self.unit_price_field.value,
                notes=self.notes_field.value,
                disburse_date=self.date_field.value
            )
            
            # Clear form
            self.item_name_field.value = ""
            self.quantity_field.value = ""
            self.unit_price_field.value = ""
            self.total_price_field.value = ""
            self.notes_field.value = ""
            
            self.page.update()
            self.show_success_dialog(excel_file)
            
        except PermissionError as e:
            self.show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
        except Exception as e:
            self.show_dialog("خطأ", f"حدث خطأ أثناء حفظ البيانات: {str(e)}", ft.Colors.RED_400)

    def show_success_dialog(self, file_path):
        """Show success dialog with file path"""
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