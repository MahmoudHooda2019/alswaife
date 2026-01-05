"""
Inventory Add View - UI for adding inventory items
Styled similar to blocks and slides views
"""

import asyncio
import flet as ft
import os
from datetime import datetime
from utils.utils import resource_path, is_excel_running, get_current_date
from utils.inventory_utils import (
    initialize_inventory_excel,
    add_inventory_entry,
    convert_existing_inventory_to_formulas,
)


class InventoryRow:
    """Row UI for inventory entry with styling similar to blocks view"""

    def __init__(self, page: ft.Page, delete_callback):
        self.page = page
        self.delete_callback = delete_callback
        self._build_controls()

    def _create_styled_textfield(self, label, width, **kwargs):
        """Create a consistently styled text field"""
        # If read_only is True, use black background to distinguish it
        default_bgcolor = ft.Colors.BLACK if kwargs.get("read_only") else ft.Colors.BLUE_GREY_900
        bgcolor = kwargs.pop("bgcolor", default_bgcolor)
        return ft.TextField(
            label=label,
            width=width,
            border_radius=10,
            filled=True,
            bgcolor=bgcolor,
            border_color=ft.Colors.GREY_700,
            focused_border_color=ft.Colors.GREEN_400,
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
            value=get_current_date("%d/%m/%Y"),
            icon=ft.Icons.CALENDAR_TODAY,
        )

        # Item name
        self.item_name_field = self._create_styled_textfield(
            "اسم الصنف", 210, icon=ft.Icons.INVENTORY_2
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

        # Unit price
        self.unit_price_field = self._create_styled_textfield(
            "سعر الوحدة",
            120,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$"),
            on_change=self._calculate_total,
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
                                self.item_name_field,
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
            "item_name": self.item_name_field.value,
            "quantity": self.quantity_field.value,
            "unit_price": self.unit_price_field.value,
            "total_price": self.total_price_field.value,
            "notes": self.notes_field.value,
        }

    def has_data(self):
        """Check if row has meaningful data"""
        return bool(self.item_name_field.value and self.quantity_field.value)

    def clear(self):
        """Clear all fields"""
        self.item_name_field.value = ""
        self.quantity_field.value = ""
        self.unit_price_field.value = ""
        self.total_price_field.value = "0"
        self.notes_field.value = ""


class InventoryAddView:
    """View for adding inventory items with design similar to blocks section"""

    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - إضافة مخزون"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK

        # Initialize paths
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        self.inventory_path = os.path.join(self.documents_path, "مخزون الادوات")
        os.makedirs(self.inventory_path, exist_ok=True)

        self.rows: list[InventoryRow] = []
        self.rows_container = ft.Column(spacing=20, scroll=ft.ScrollMode.AUTO, expand=True)

    def build_ui(self):
        """Build the inventory add UI"""
        app_bar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK, on_click=self.go_back, tooltip="العودة"
            ),
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.ADD_SHOPPING_CART, size=24, color=ft.Colors.GREEN_200),
                    ft.Text(
                        "إضافة للمخزون",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.GREEN_200,
                    ),
                ],
                spacing=10,
            ),
            actions=[
                ft.IconButton(
                    icon=ft.Icons.REFRESH, on_click=self.reset_all, tooltip="مسح الكل"
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

        main_column = ft.Column(
            controls=[self.rows_container],
            spacing=15,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

        self.page.add(main_column)
        self.add_row()
        self.page.update()

    def go_back(self, e):
        """Navigate back"""
        if self.on_back:
            self.on_back()

    def reset_all(self, e=None):
        """Reset all rows - clear all data"""
        self.rows.clear()
        self.rows_container.controls.clear()
        self.add_row()
        self.page.update()

    def add_row(self, e=None):
        """Add a new inventory row"""
        row = InventoryRow(page=self.page, delete_callback=self.delete_row)
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
        """Save data to Excel file"""
        if not any(row.has_data() for row in self.rows):
            self._show_dialog("تحذير", "لا توجد بيانات لحفظها", ft.Colors.ORANGE_400)
            return

        # التحقق من أن Excel مغلق
        if is_excel_running():
            self._show_excel_warning_dialog()
            return

        self._do_save()

    def _do_save(self):
        """تنفيذ عملية الحفظ الفعلية"""
        try:
            excel_file = os.path.join(self.inventory_path, "مخزون ادوات التشغيل.xlsx")

            if not os.path.exists(excel_file):
                initialize_inventory_excel(excel_file)
            else:
                try:
                    convert_existing_inventory_to_formulas(excel_file)
                except:
                    pass

            saved_count = 0
            for row in self.rows:
                if row.has_data():
                    data = row.to_dict()
                    add_inventory_entry(
                        file_path=excel_file,
                        item_name=data["item_name"],
                        quantity=data["quantity"],
                        unit_price=data["unit_price"],
                        notes=data["notes"],
                        entry_date=data["date"],
                    )
                    saved_count += 1

            self._show_success_dialog(excel_file, saved_count)

        except PermissionError:
            self._show_dialog(
                "خطأ",
                "الملف مفتوح في Excel. أغلقه وحاول مرة أخرى.",
                ft.Colors.RED_400,
            )
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ: {str(e)}", ft.Colors.RED_400)

    async def _delayed_close(self, dlg):
        """Close dialog with delay to prevent glitch"""
        await asyncio.sleep(0.3)
        self.page.close(dlg)

    def _show_excel_warning_dialog(self):
        """Show Excel warning dialog with continue option"""
        def close_dlg(e=None):
            self.page.run_task(self._delayed_close, dlg)

        def continue_save(e=None):
            self.page.close(dlg)
            self._do_save()

        dlg = ft.AlertDialog(
            title=ft.Text("تحذير", color=ft.Colors.ORANGE_400, weight=ft.FontWeight.BOLD),
            content=ft.Text("برنامج Excel مفتوح حالياً.\nيرجى إغلاقه قبل الحفظ.", size=16, rtl=True),
            actions=[
                ft.TextButton(
                    "متابعة على أي حال",
                    on_click=continue_save,
                    style=ft.ButtonStyle(color=ft.Colors.ORANGE_400)
                ),
                ft.TextButton(
                    "إلغاء",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.BLUE_GREY_900
        )
        self.page.open(dlg)

    def _show_dialog(self, title: str, message: str, title_color=ft.Colors.BLUE_300):
        """Show a styled dialog"""
        def close_dlg(e=None):
            self.page.run_task(self._delayed_close, dlg)

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
        self.page.open(dlg)

    def _show_success_dialog(self, filepath: str, count: int):
        """Show success dialog"""
        def close_dlg(e=None):
            self.page.run_task(self._delayed_close, dlg)

        def open_file(e=None):
            self.page.close(dlg)
            try:
                os.startfile(filepath)
            except:
                pass

        def open_folder(e=None):
            self.page.close(dlg)
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
            content=ft.Text(f"تم حفظ {count} صنف في المخزون", size=14, rtl=True),
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
        self.page.open(dlg)
