"""
Inventory Disburse View - UI for disbursing inventory items
Styled similar to blocks and slides views
"""

import flet as ft
import os
import traceback
from datetime import datetime
from utils.path_utils import resource_path
from utils.inventory_utils import (
    initialize_inventory_excel,
    disburse_inventory_entry,
    get_inventory_summary,
    get_available_items_with_prices,
    convert_existing_inventory_to_formulas,
)


class InventoryDisburseRow:
    """Row UI for inventory disbursement with styling similar to blocks view"""

    def __init__(self, page: ft.Page, delete_callback, available_items: list, item_prices: dict, inventory_balances: dict = None):
        self.page = page
        self.delete_callback = delete_callback
        self.available_items = available_items
        self.item_prices = item_prices
        self.inventory_balances = inventory_balances or {}
        self._build_controls()

    def _create_styled_textfield(self, label, width, **kwargs):
        """Create a consistently styled text field"""
        bgcolor = kwargs.pop("bgcolor", ft.Colors.BLUE_GREY_900)
        return ft.TextField(
            label=label,
            width=width,
            border_radius=10,
            filled=True,
            bgcolor=bgcolor,
            border_color=ft.Colors.GREY_700,
            focused_border_color=ft.Colors.RED_400,
            label_style=ft.TextStyle(color=ft.Colors.GREY_400),
            text_style=ft.TextStyle(size=14, weight=ft.FontWeight.W_500, color=ft.Colors.WHITE),
            cursor_color=ft.Colors.WHITE,
            **kwargs,
        )

    def _build_controls(self):
        """Build all UI controls"""
        # Date field
        self.date_field = self._create_styled_textfield(
            "التاريخ",
            140,
            value=datetime.now().strftime("%d/%m/%Y"),
            read_only=True,
            icon=ft.Icons.CALENDAR_TODAY,
        )

        # Item dropdown
        dropdown_options = [ft.dropdown.Option(item) for item in self.available_items] if self.available_items else []
        self.item_dropdown = ft.Dropdown(
            label="اسم الصنف",
            width=210,
            options=dropdown_options,
            border_radius=10,
            filled=True,
            bgcolor=ft.Colors.BLUE_GREY_900,
            border_color=ft.Colors.GREY_700,
            focused_border_color=ft.Colors.RED_400,
            label_style=ft.TextStyle(color=ft.Colors.GREY_400),
            text_style=ft.TextStyle(size=14, weight=ft.FontWeight.W_500, color=ft.Colors.WHITE),
            on_change=self._on_item_selected,
        )

        # Quantity
        self.quantity_field = self._create_styled_textfield(
            "العدد",
            105,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$"),
            on_change=self._calculate_total,
            icon=ft.Icons.NUMBERS,
        )

        # Unit price (read-only, auto-filled)
        self.unit_price_field = self._create_styled_textfield(
            "سعر الوحدة",
            120,
            read_only=True,
            suffix_text="ج",
        )

        # Total price (calculated)
        self.total_price_field = self._create_styled_textfield(
            "الإجمالي", 120, read_only=True, value="0", suffix_text="ج"
        )

        # Notes
        self.notes_field = self._create_styled_textfield(
            "ملاحظات", 180, icon=ft.Icons.NOTES
        )

        # Delete button
        self.delete_btn = ft.IconButton(
            icon=ft.Icons.DELETE_OUTLINE,
            icon_color=ft.Colors.RED_400,
            tooltip="حذف الصف",
            on_click=lambda e: self.delete_callback(self),
            bgcolor=ft.Colors.GREY_800,
            icon_size=20,
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10)),
        )

        # Build the card
        self.card = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Row(
                            controls=[ft.Container(expand=True), self.delete_btn],
                            alignment=ft.MainAxisAlignment.END,
                        ),
                        ft.Row(
                            controls=[
                                self.date_field,
                                self.item_dropdown,
                                self.quantity_field,
                                self.unit_price_field,
                                self.total_price_field,
                                self.notes_field,
                            ],
                            spacing=15,
                            wrap=True,
                            alignment=ft.MainAxisAlignment.CENTER,
                        ),
                    ],
                    spacing=10,
                ),
                padding=20,
                gradient=ft.LinearGradient(
                    begin=ft.alignment.top_left,
                    end=ft.alignment.bottom_right,
                    colors=[ft.Colors.GREY_900, ft.Colors.GREY_800],
                ),
                border_radius=15,
                border=ft.border.all(1, ft.Colors.GREY_700),
            ),
            elevation=8,
        )
        self.row = self.card

    def _on_item_selected(self, e=None):
        """Handle item selection - auto-fill unit price and update quantity hint"""
        selected_item = self.item_dropdown.value
        if selected_item and selected_item in self.item_prices:
            price = round(self.item_prices[selected_item], 2)
            self.unit_price_field.value = f"{price:.2f}"
        else:
            self.unit_price_field.value = "0"
        
        # Update quantity hint with available balance
        if selected_item and selected_item in self.inventory_balances:
            balance = self.inventory_balances[selected_item]
            # Format balance without decimals if it's a whole number
            if balance == int(balance):
                self.quantity_field.hint_text = f"{int(balance)}"
            else:
                self.quantity_field.hint_text = f"{balance:.2f}"
        else:
            self.quantity_field.hint_text = None
        
        self._calculate_total()
        self.page.update()

    def _calculate_total(self, e=None):
        """Calculate total price"""
        try:
            quantity = float(self.quantity_field.value) if self.quantity_field.value else 0
            unit_price = float(self.unit_price_field.value) if self.unit_price_field.value else 0
            total = quantity * unit_price
            self.total_price_field.value = f"{total:.2f}"
        except ValueError:
            self.total_price_field.value = "0"
        self.page.update()

    def to_dict(self):
        """Convert row data to dictionary"""
        return {
            "date": self.date_field.value,
            "item_name": self.item_dropdown.value,
            "quantity": self.quantity_field.value,
            "unit_price": self.unit_price_field.value,
            "total_price": self.total_price_field.value,
            "notes": self.notes_field.value,
        }

    def has_data(self):
        """Check if row has meaningful data"""
        return bool(self.item_dropdown.value and self.quantity_field.value)

    def clear(self):
        """Clear all fields"""
        self.item_dropdown.value = None
        self.quantity_field.value = ""
        self.unit_price_field.value = ""
        self.total_price_field.value = "0"
        self.notes_field.value = ""


class InventoryDisburseView:
    """View for disbursing inventory items with design similar to blocks section"""

    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - صرف مخزون"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK

        # Initialize paths
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        self.inventory_path = os.path.join(self.documents_path, "مخزون الادوات")
        os.makedirs(self.inventory_path, exist_ok=True)

        self.excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")
        self.rows: list[InventoryDisburseRow] = []
        self.rows_container = ft.Column(spacing=20, scroll=ft.ScrollMode.AUTO, expand=True)
        
        # Load available items and prices
        self.available_items = []
        self.item_prices = {}
        self.inventory_balances = {}
        self._load_inventory_data()

    def _load_inventory_data(self):
        """Load available items and their prices from inventory"""
        try:
            if os.path.exists(self.excel_file):
                try:
                    convert_existing_inventory_to_formulas(self.excel_file)
                except:
                    pass
                
                # Get inventory summary for balances
                inventory_data = get_inventory_summary(self.excel_file)
                self.available_items = [item['item_name'] for item in inventory_data if item['item_name']]
                self.inventory_balances = {item['item_name']: float(item['current_balance']) for item in inventory_data if item['item_name']}
                
                # Get prices
                self.item_prices = get_available_items_with_prices(self.excel_file)
        except Exception as e:
            print(f"[ERROR] Failed to load inventory data: {e}")
            traceback.print_exc()

    def build_ui(self):
        """Build the inventory disburse UI"""
        app_bar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK, on_click=self.go_back, tooltip="العودة"
            ),
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.REMOVE_SHOPPING_CART, size=24, color=ft.Colors.RED_200),
                    ft.Text(
                        "صرف من المخزون",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.RED_200,
                    ),
                ],
                spacing=10,
            ),
            actions=[
                ft.IconButton(
                    icon=ft.Icons.REFRESH, on_click=self.refresh_data, tooltip="تحديث البيانات"
                ),
                ft.IconButton(
                    icon=ft.Icons.ADD, on_click=self.add_row, tooltip="إضافة صف جديد"
                ),
                ft.Container(
                    content=ft.IconButton(
                        icon=ft.Icons.SAVE, on_click=self.save_to_excel, tooltip="حفظ البيانات"
                    ),
                    margin=ft.margin.only(left=40, right=15),
                ),
            ],
            bgcolor=ft.Colors.GREY_900,
        )

        self.page.appbar = app_bar

        # Info banner showing available items count
        info_banner = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.INFO_OUTLINE, color=ft.Colors.BLUE_300, size=18),
                    ft.Text(
                        f"عدد الأصناف المتاحة: {len(self.available_items)}",
                        size=14,
                        color=ft.Colors.BLUE_300,
                    ),
                ],
                spacing=10,
                alignment=ft.MainAxisAlignment.CENTER,
            ),
            padding=10,
            bgcolor=ft.Colors.BLUE_GREY_900,
            border_radius=10,
            margin=ft.margin.only(bottom=10),
        )

        main_column = ft.Column(
            controls=[info_banner, self.rows_container],
            spacing=15,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

        self.page.add(main_column)
        self.add_row()
        self.page.update()

    def go_back(self, e=None):
        """Navigate back"""
        if self.on_back:
            self.on_back()

    def refresh_data(self, e=None):
        """Refresh inventory data"""
        self._load_inventory_data()
        # Update existing rows with new data
        for row in self.rows:
            row.available_items = self.available_items
            row.item_prices = self.item_prices
            row.inventory_balances = self.inventory_balances
            row.item_dropdown.options = [ft.dropdown.Option(item) for item in self.available_items]
        self.page.update()
        self._show_dialog("تم التحديث", f"تم تحديث البيانات - {len(self.available_items)} صنف متاح", ft.Colors.GREEN_400)

    def add_row(self, e=None):
        """Add a new inventory disburse row"""
        row = InventoryDisburseRow(
            page=self.page,
            delete_callback=self.delete_row,
            available_items=self.available_items,
            item_prices=self.item_prices,
            inventory_balances=self.inventory_balances,
        )
        self.rows.append(row)
        self.rows_container.controls.append(row.row)
        self.page.update()

    def delete_row(self, row_obj):
        """Delete a specific row"""
        if row_obj in self.rows:
            self.rows.remove(row_obj)
            self.rows_container.controls.remove(row_obj.row)
            self.page.update()

    def save_to_excel(self, e=None):
        """Save disbursement data to Excel file"""
        if not any(row.has_data() for row in self.rows):
            self._show_dialog("تحذير", "لا توجد بيانات لحفظها", ft.Colors.ORANGE_400)
            return

        # Validate all rows first
        for row in self.rows:
            if row.has_data():
                data = row.to_dict()
                item_name = data["item_name"]
                
                # Check if item exists
                if item_name not in self.inventory_balances:
                    self._show_dialog("خطأ", f"الصنف '{item_name}' غير موجود في المخزون", ft.Colors.RED_400)
                    return
                
                # Check balance
                try:
                    requested_qty = float(data["quantity"])
                    available_balance = self.inventory_balances.get(item_name, 0)
                    
                    if available_balance <= 0:
                        self._show_dialog("خطأ", f"الصنف '{item_name}' ليس له رصيد متوفر", ft.Colors.RED_400)
                        return
                    
                    if requested_qty > available_balance:
                        self._show_dialog(
                            "خطأ",
                            f"الكمية المطلوبة ({requested_qty}) تتجاوز الرصيد المتاح ({available_balance}) للصنف '{item_name}'",
                            ft.Colors.RED_400
                        )
                        return
                except ValueError:
                    self._show_dialog("خطأ", "يرجى إدخال كمية صحيحة", ft.Colors.RED_400)
                    return

        try:
            if not os.path.exists(self.excel_file):
                initialize_inventory_excel(self.excel_file)
            else:
                try:
                    convert_existing_inventory_to_formulas(self.excel_file)
                except:
                    pass

            saved_count = 0
            for row in self.rows:
                if row.has_data():
                    data = row.to_dict()
                    disburse_inventory_entry(
                        file_path=self.excel_file,
                        item_name=data["item_name"],
                        quantity=data["quantity"],
                        unit_price=data["unit_price"],
                        notes=data["notes"],
                        disburse_date=data["date"],
                    )
                    row.clear()
                    saved_count += 1

            # Reload inventory data after saving
            self._load_inventory_data()
            self._show_success_dialog(self.excel_file, saved_count)

        except PermissionError:
            self._show_dialog(
                "خطأ",
                "الملف مفتوح في Excel. أغلقه وحاول مرة أخرى.",
                ft.Colors.RED_400,
            )
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ: {str(e)}", ft.Colors.RED_400)

    def _show_dialog(self, title: str, message: str, title_color=ft.Colors.BLUE_300):
        """Show a styled dialog"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()

        dlg = ft.AlertDialog(
            title=ft.Text(title, color=title_color, weight=ft.FontWeight.BOLD),
            content=ft.Text(message, size=16, rtl=True),
            actions=[
                ft.TextButton(
                    "إغلاق", on_click=close_dlg, style=ft.ButtonStyle(color=ft.Colors.BLUE_300)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.BLUE_GREY_900,
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def _show_success_dialog(self, filepath: str, count: int):
        """Show success dialog"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()

        def open_file(e=None):
            close_dlg()
            try:
                os.startfile(filepath)
            except:
                pass

        def open_folder(e=None):
            close_dlg()
            try:
                os.startfile(os.path.dirname(filepath))
            except:
                pass

        dlg = ft.AlertDialog(
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=30),
                    ft.Text(
                        "تم الحفظ بنجاح",
                        color=ft.Colors.GREEN_300,
                        weight=ft.FontWeight.BOLD,
                    ),
                ],
                rtl=True,
                spacing=10,
            ),
            content=ft.Text(f"تم صرف {count} صنف من المخزون", size=14, rtl=True),
            actions=[
                ft.TextButton(
                    "فتح الملف",
                    on_click=open_file,
                    icon=ft.Icons.FILE_OPEN,
                    style=ft.ButtonStyle(color=ft.Colors.GREEN_300),
                ),
                ft.TextButton(
                    "فتح المجلد",
                    on_click=open_folder,
                    icon=ft.Icons.FOLDER_OPEN,
                    style=ft.ButtonStyle(color=ft.Colors.BLUE_300),
                ),
                ft.TextButton(
                    "إغلاق", on_click=close_dlg, style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.BLUE_GREY_900,
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()
