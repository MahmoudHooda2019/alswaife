import os
import json
from datetime import datetime
import flet as ft
from utils.path_utils import resource_path
from utils.purchases_utils import export_purchases_to_excel, load_purchases_from_excel, load_item_names_from_excel
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font


class PurchasesView:
    """View for managing purchases with auto-complete functionality"""
    
    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - المشتريات"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK
        
        # Initialize data storage
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        self.purchases_path = os.path.join(self.documents_path, "المشتريات")
        os.makedirs(self.purchases_path, exist_ok=True)
        
        # Load existing items for auto-complete
        self.items_list = self.load_existing_items()
        self.suppliers_list = ["مستر مصطفي", "مستر محمد", ""]  # Predefined suppliers
        
        # Form fields
        self.date_field = ft.TextField(
            label="التاريخ",
            value=datetime.now().strftime('%d/%m/%Y'),
            width=150
        )
        
        self.code_field = ft.TextField(
            label="الكود",
            width=150,
            keyboard_type=ft.KeyboardType.NUMBER
        )
        
        self.quantity_field = ft.TextField(
            label="العدد",
            width=150,
            keyboard_type=ft.KeyboardType.NUMBER
        )
        
        # Item name with auto-complete
        self.item_name_field = ft.TextField(
            label="اسم الصنف",
            width=None,  # Remove fixed width to allow expansion
            expand=True,  # Allow field to expand
            on_change=self.on_item_name_change
        )
        
        # Suggestions list for items
        self.item_suggestions = ft.Column(
            visible=False,
            spacing=0
        )
        
        self.total_price_field = ft.TextField(
            label="إجمالي السعر",
            width=150,
            keyboard_type=ft.KeyboardType.NUMBER
        )
        
        # Supplier with auto-complete
        self.supplier_field = ft.TextField(
            label="من",
            width=200,
            on_change=self.on_supplier_change
        )
        
        # Suggestions list for suppliers
        self.supplier_suggestions = ft.Column(
            visible=False,
            spacing=0
        )
        
        self.notes_field = ft.TextField(
            label="الملاحظات",
            multiline=True,
            min_lines=3,
            max_lines=5,
            width=400
        )
        
        # Data table
        self.data_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("التاريخ")),
                ft.DataColumn(ft.Text("الكود")),
                ft.DataColumn(ft.Text("اسم الصنف")),
                ft.DataColumn(ft.Text("العدد")),
                ft.DataColumn(ft.Text("إجمالي السعر")),
                ft.DataColumn(ft.Text("الرصيد")),
                ft.DataColumn(ft.Text("من")),
                ft.DataColumn(ft.Text("الملاحظات")),
            ],
            rows=[]
        )
        
        # Load existing data
        self.load_existing_data()

    def load_existing_items(self):
        """Load existing item names for auto-complete from Excel file"""
        excel_file = os.path.join(self.purchases_path, "سجل المشتريات.xlsx")
        return load_item_names_from_excel(excel_file)

    def load_existing_data(self):
        """Load existing purchase records from Excel file"""
        excel_file = os.path.join(self.purchases_path, "سجل المشتريات.xlsx")
        records = load_purchases_from_excel(excel_file)
        
        for record in records:
            self.add_row_to_table(record)

    def on_item_name_change(self, e):
        """Handle item name change with auto-complete suggestions"""
        search_text = self.item_name_field.value.strip().lower() if self.item_name_field.value else ""
        
        if not search_text:
            self.item_suggestions.visible = False
            self.item_suggestions.controls.clear()
            self.page.update()
            return
        
        # Filter suggestions
        filtered = [item for item in self.items_list if search_text in item.lower()]
        
        if filtered:
            self.item_suggestions.controls.clear()
            for item in filtered[:5]:  # Show max 5 suggestions
                suggestion_btn = ft.TextButton(
                    text=item,
                    on_click=lambda e, i=item: self.select_item_suggestion(i),
                    style=ft.ButtonStyle(
                        padding=ft.padding.all(5),
                    )
                )
                self.item_suggestions.controls.append(suggestion_btn)
            self.item_suggestions.visible = True
        else:
            self.item_suggestions.visible = False
            self.item_suggestions.controls.clear()
        
        self.page.update()

    def select_item_suggestion(self, item_name):
        """Select an item from suggestions"""
        self.item_name_field.value = item_name
        self.item_suggestions.visible = False
        self.item_suggestions.controls.clear()
        self.page.update()

    def on_supplier_change(self, e):
        """Handle supplier change with auto-complete suggestions"""
        search_text = self.supplier_field.value.strip().lower() if self.supplier_field.value else ""
        
        if not search_text:
            self.supplier_suggestions.visible = False
            self.supplier_suggestions.controls.clear()
            self.page.update()
            return
        
        # Filter suggestions
        filtered = [supplier for supplier in self.suppliers_list if search_text in supplier.lower()]
        
        if filtered:
            self.supplier_suggestions.controls.clear()
            for supplier in filtered[:5]:  # Show max 5 suggestions
                suggestion_btn = ft.TextButton(
                    text=supplier,
                    on_click=lambda e, s=supplier: self.select_supplier_suggestion(s),
                    style=ft.ButtonStyle(
                        padding=ft.padding.all(5),
                    )
                )
                self.supplier_suggestions.controls.append(suggestion_btn)
            self.supplier_suggestions.visible = True
        else:
            self.supplier_suggestions.visible = False
            self.supplier_suggestions.controls.clear()
        
        self.page.update()

    def select_supplier_suggestion(self, supplier_name):
        """Select a supplier from suggestions"""
        self.supplier_field.value = supplier_name
        self.supplier_suggestions.visible = False
        self.supplier_suggestions.controls.clear()
        self.page.update()

    def add_row_to_table(self, record):
        """Add a record to the data table"""
        # Calculate balance for display (this is a simplified version for the table display)
        # In the actual Excel, the balance is calculated with formulas
        balance = record.get('total_price', '')  # Simplified for display purposes
        
        self.data_table.rows.append(
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(record.get('date', ''))),
                    ft.DataCell(ft.Text(record.get('code', ''))),
                    ft.DataCell(ft.Text(record.get('item_name', ''))),
                    ft.DataCell(ft.Text(record.get('quantity', ''))),
                    ft.DataCell(ft.Text(record.get('total_price', ''))),
                    ft.DataCell(ft.Text(balance)),  # Balance column
                    ft.DataCell(ft.Text(record.get('supplier', ''))),
                    ft.DataCell(ft.Text(record.get('notes', ''))),
                ]
            )
        )

    def save_purchase(self, e):
        """Save purchase record to Excel file"""
        # Validate required fields
        if not self.code_field.value or not self.item_name_field.value or not self.quantity_field.value or not self.total_price_field.value:
            self.show_dialog("خطأ", "يرجى ملء جميع الحقول المطلوبة", ft.Colors.RED_400)
            return
            
        # Create record
        record = {
            'date': self.date_field.value,
            'code': self.code_field.value,
            'item_name': self.item_name_field.value,
            'quantity': self.quantity_field.value,
            'total_price': self.total_price_field.value,
            'supplier': self.supplier_field.value,
            'notes': self.notes_field.value
        }
        
        # Save to Excel file using the utility
        excel_file = os.path.join(self.purchases_path, "سجل المشتريات.xlsx")
        
        try:
            # Export the record to Excel
            export_purchases_to_excel([record], excel_file)
                
            # Add to table
            self.add_row_to_table(record)
            
            # Add to items list for future auto-complete
            if record['item_name'] not in self.items_list:
                self.items_list.append(record['item_name'])
            
            # Clear form
            self.code_field.value = ""
            self.item_name_field.value = ""
            self.quantity_field.value = ""
            self.total_price_field.value = ""
            self.supplier_field.value = ""
            self.notes_field.value = ""
            
            self.page.update()
            self.show_success_dialog(excel_file)
            
        except PermissionError as e:
            self.show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
        except Exception as e:
            self.show_dialog("خطأ", f"حدث خطأ أثناء الحفظ: {str(e)}", ft.Colors.RED_400)

    def show_dialog(self, title, message, color):
        """Show a dialog with a message"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
            
        dlg = ft.AlertDialog(
            title=ft.Text(title, color=color),
            content=ft.Text(message, rtl=True),
            actions=[
                ft.TextButton("حسناً", on_click=close_dlg)
            ]
        )
        self.page.dialog = dlg
        dlg.open = True
        self.page.update()

    def show_success_dialog(self, filepath):
        """Show success dialog with file actions like blocks view"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()

        def open_file(e):
            close_dlg(None)
            try:
                os.startfile(filepath)
            except Exception:
                pass

        def open_folder(e):
            close_dlg(None)
            try:
                folder = os.path.dirname(filepath)
                os.startfile(folder)
            except Exception:
                pass

        dlg = ft.AlertDialog(
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=30),
                    ft.Text("تم الحفظ بنجاح", color=ft.Colors.GREEN_300, weight=ft.FontWeight.BOLD),
                ],
                spacing=10
            ),
            content=ft.Column(
                rtl=True,
                controls=[
                    ft.Text("تم إنشاء الملف بنجاح:", size=14, rtl=True),
                    ft.Container(
                        content=ft.Text(
                            os.path.basename(filepath),
                            size=13,
                            color=ft.Colors.BLUE_200,
                            weight=ft.FontWeight.W_500
                        ),
                        bgcolor=ft.Colors.BLUE_GREY_800,
                        padding=10,
                        border_radius=8,
                        margin=ft.margin.only(top=10)
                    )
                ],
                tight=True
            ),
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
                ft.TextButton("حسناً", on_click=lambda e: close_dlg(None))
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def build_ui(self):
        """Build the purchases UI"""
        # Header
        header = ft.Row(
            controls=[
                ft.Icon(ft.Icons.SHOPPING_CART, size=36, color=ft.Colors.BLUE_300),
                ft.Container(width=10),
                ft.Text(
                    "سجل المشتريات",
                    size=32,
                    weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_100
                ),
                ft.Container(expand=True),
                ft.FilledButton(
                    "حفظ البيانات",
                    icon=ft.Icons.SAVE,
                    on_click=self.save_purchase,
                    style=ft.ButtonStyle(
                        bgcolor=ft.Colors.GREEN_700,
                        color=ft.Colors.WHITE,
                        shape=ft.RoundedRectangleBorder(radius=12),
                        padding=15
                    )
                ),
            ],
            alignment=ft.MainAxisAlignment.START
        )
        
        # Form section
        form_section = ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text("إدخال بيانات المشتريات", size=20, weight=ft.FontWeight.BOLD),
                    ft.Row(
                        controls=[
                            self.date_field,
                            self.code_field,
                            self.quantity_field,
                        ],
                        spacing=20,
                        alignment=ft.MainAxisAlignment.CENTER
                    ),
                    ft.Row(
                        controls=[
                            ft.Column([
                                self.item_name_field,
                                self.item_suggestions,
                            ], spacing=0, expand=True),
                        ],
                        spacing=20,
                        alignment=ft.MainAxisAlignment.CENTER
                    ),
                    ft.Row(
                        controls=[
                            ft.Column(
                                controls=[
                                    self.total_price_field,
                                    self.supplier_field,
                                ],
                                spacing=10,
                                width=200
                            ),
                            self.notes_field,
                        ],
                        spacing=20,
                        alignment=ft.MainAxisAlignment.CENTER
                    ),
                ],
                spacing=15,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            ),
            padding=20,
            bgcolor=ft.Colors.GREY_900,
            border_radius=15,
            margin=ft.margin.only(bottom=10)
        )
        
        # Data table section
        table_section = ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text("السجلات السابقة", size=20, weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=ft.ListView(
                            controls=[self.data_table],
                            expand=True,
                            height=400
                        ),
                        border=ft.border.all(1, ft.Colors.GREY_700),
                        border_radius=10,
                        padding=10
                    )
                ],
                spacing=15,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER
            ),
            padding=20,
            bgcolor=ft.Colors.GREY_900,
            border_radius=15
        )
        
        # Main layout
        main_layout = ft.Column(
            controls=[
                header,
                form_section,
                table_section
            ],
            spacing=15,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )
        
        # Back button
        back_btn = ft.IconButton(
            icon=ft.Icons.ARROW_BACK,
            on_click=self.go_back,
            tooltip="العودة للقائمة الرئيسية"
        )
        
        self.page.clean()
        self.page.add(
            ft.Column(
                controls=[
                    ft.Row([back_btn]),
                    ft.Container(
                        content=main_layout,
                        expand=True,
                        padding=15,
                        bgcolor=ft.Colors.GREY_900
                    )
                ],
                expand=True
            )
        )
        
        self.page.update()
        return main_layout

    def go_back(self, e):
        """Navigate back to dashboard"""
        if self.on_back:
            self.on_back()