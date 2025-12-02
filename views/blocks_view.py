import os
from datetime import datetime
import flet as ft

from utils.blocks_utils import export_blocks_excel


class BlockRow:
    """Row UI for block entry"""

    MATERIAL_OPTIONS = [
        ft.dropdown.Option("نيو حلايب"),
        ft.dropdown.Option("جندولا"),
    ]

    MACHINE_OPTIONS = [
        ft.dropdown.Option("1"),
        ft.dropdown.Option("2"),
        ft.dropdown.Option("3"),
    ]

    def __init__(self, page: ft.Page, delete_callback):
        self.page = page
        self.delete_callback = delete_callback
        self._build_controls()

    def _build_controls(self):
        width_small = 110
        width_medium = 140

        numeric_filter = ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$")

        self.trip_number = ft.TextField(label="رقم النقلة", width=width_small)
        self.trip_count = ft.TextField(
            label="عدد النقلات",
            width=width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.NumbersOnlyInputFilter(),
        )
        self.date_field = ft.TextField(
            label="التاريخ",
            width=width_small,
            value=datetime.now().strftime("%d/%m/%Y"),
        )
        self.quarry_field = ft.TextField(label="المحجر", width=width_medium)
        self.machine_dropdown = ft.Dropdown(
            label="رقم الماكينة",
            width=width_small,
            options=self.MACHINE_OPTIONS,
        )
        self.block_number = ft.TextField(label="رقم البلوك", width=width_small)
        self.material_dropdown = ft.Dropdown(
            label="الخامة",
            width=width_medium,
            options=self.MATERIAL_OPTIONS,
        )

        self.length_field = ft.TextField(
            label="الطول (م)",
            width=width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._calculate_values,
        )
        self.width_field = ft.TextField(
            label="العرض (م)",
            width=width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._calculate_values,
        )
        self.height_field = ft.TextField(
            label="الارتفاع (م)",
            width=width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._calculate_values,
        )

        self.volume_field = ft.TextField(label="م3", width=width_small, disabled=True)
        self.weight_field = ft.TextField(
            label="الوزن (طن/م3)",
            width=width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._calculate_values,
        )
        self.block_weight_field = ft.TextField(
            label="وزن البلوك (طن)",
            width=width_small,
            disabled=True,
        )
        self.ton_price_field = ft.TextField(
            label="سعر الطن",
            width=width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._calculate_values,
        )
        self.block_total_price_field = ft.TextField(
            label="إجمالي سعر البلوك",
            width=width_medium,
            disabled=True,
        )
        self.trip_price_field = ft.TextField(
            label="سعر النقلة",
            width=width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
        )

        self.delete_btn = ft.IconButton(
            icon=ft.Icons.DELETE,
            icon_color=ft.Colors.RED,
            tooltip="حذف الصف",
            on_click=lambda _: self.delete_callback(self),
        )

    def _calculate_values(self, e=None):
        try:
            length = float(self.length_field.value or 0)
            width = float(self.width_field.value or 0)
            height = float(self.height_field.value or 0)
            base_weight = float(self.weight_field.value or 0)
            ton_price = float(self.ton_price_field.value or 0)

            volume = length * width * height
            self.volume_field.value = f"{volume:.3f}" if volume else ""

            block_weight = volume * base_weight
            self.block_weight_field.value = f"{block_weight:.3f}" if block_weight else ""

            total_price = block_weight * ton_price
            self.block_total_price_field.value = (
                f"{total_price:.2f}" if total_price else ""
            )

        except ValueError:
            self.volume_field.value = ""
            self.block_weight_field.value = ""
            self.block_total_price_field.value = ""
        finally:
            self.page.update()

    def get_controls(self):
        controls = [
            self.trip_number,
            self.trip_count,
            self.date_field,
            self.quarry_field,
            self.machine_dropdown,
            self.block_number,
            self.material_dropdown,
            self.length_field,
            self.width_field,
            self.height_field,
            self.volume_field,
            self.weight_field,
            self.block_weight_field,
            self.ton_price_field,
            self.block_total_price_field,
            self.trip_price_field,
            self.delete_btn,
        ]
        return controls

    def to_dict(self):
        return {
            "trip_number": (self.trip_number.value or "").strip(),
            "trip_count": (self.trip_count.value or "").strip(),
            "date": (self.date_field.value or "").strip(),
            "quarry": (self.quarry_field.value or "").strip(),
            "machine_number": (self.machine_dropdown.value or "").strip(),
            "block_number": (self.block_number.value or "").strip(),
            "material": (self.material_dropdown.value or "").strip(),
            "length": self._to_float(self.length_field.value),
            "width": self._to_float(self.width_field.value),
            "height": self._to_float(self.height_field.value),
            "volume": self._to_float(self.volume_field.value),
            "weight": self._to_float(self.weight_field.value),
            "block_weight": self._to_float(self.block_weight_field.value),
            "ton_price": self._to_float(self.ton_price_field.value),
            "block_total_price": self._to_float(self.block_total_price_field.value),
            "trip_price": self._to_float(self.trip_price_field.value),
        }

    @staticmethod
    def _to_float(value):
        try:
            return float(value or 0)
        except ValueError:
            return 0.0


class BlocksView:
    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - سجل البلوكات"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK

        self.rows: list[BlockRow] = []

        self.block_size_field = ft.TextField(
            label="مقاس البلوك في الأرضية", width=250
        )
        self.notes_field = ft.TextField(
            label="ملاحظات", multiline=True, min_lines=1, max_lines=3, expand=True
        )

        self.rows_container = ft.Column(spacing=15)
        self.add_row_btn = ft.FloatingActionButton(
            icon=ft.Icons.ADD, tooltip="إضافة صف جديد", on_click=self.add_row
        )

    def build_ui(self):
        self.page.clean()
        self.page.appbar = ft.AppBar(
            title=ft.Text("سجل البلوكات"),
            bgcolor=ft.Colors.SURFACE,
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK,
                tooltip="رجوع",
                on_click=lambda _: self.on_back() if self.on_back else None,
            ),
            actions=[
                ft.IconButton(ft.Icons.SAVE, tooltip="حفظ إلى Excel", on_click=self.save_to_excel),
                ft.IconButton(ft.Icons.REFRESH, tooltip="تهيئة جديدة", on_click=self.reset_form),
            ],
        )

        header = ft.ResponsiveRow(
            controls=[
                ft.Container(self.block_size_field, col={"xs": 12, "md": 6, "lg": 4}),
                ft.Container(self.notes_field, col={"xs": 12, "md": 6, "lg": 8}),
            ],
            spacing=20,
        )

        rows_wrapper = ft.Row(
            controls=[self.rows_container],
            scroll=ft.ScrollMode.ALWAYS,
            expand=True,
        )

        layout = ft.Column(
            controls=[
                header,
                ft.Divider(),
                rows_wrapper,
                ft.Container(
                    ft.ElevatedButton(
                        "حفظ ملف Excel", icon=ft.Icons.SAVE, on_click=self.save_to_excel
                    ),
                    alignment=ft.alignment.center_right,
                    padding=ft.padding.symmetric(vertical=10, horizontal=20),
                ),
            ],
            spacing=15,
            expand=True,
            scroll=ft.ScrollMode.AUTO,
        )

        self.page.add(layout)
        self.page.floating_action_button = self.add_row_btn
        self.add_row()
        self.page.update()

    def add_row(self, e=None):
        row = BlockRow(self.page, delete_callback=self.delete_row)
        self.rows.append(row)
        row_controls = row.get_controls()
        row_widget = ft.Row(controls=row_controls, spacing=5)
        self.rows_container.controls.append(row_widget)
        if not hasattr(self, "_row_widgets"):
            self._row_widgets = {}
        self._row_widgets[row] = row_widget
        self.page.update()

    def delete_row(self, row_obj: BlockRow):
        if row_obj in self.rows:
            self.rows.remove(row_obj)
            if hasattr(self, "_row_widgets") and row_obj in self._row_widgets:
                widget = self._row_widgets[row_obj]
                if widget in self.rows_container.controls:
                    self.rows_container.controls.remove(widget)
                del self._row_widgets[row_obj]
            self.page.update()

    def reset_form(self, e=None):
        self.block_size_field.value = ""
        self.notes_field.value = ""
        self.rows.clear()
        self.rows_container.controls.clear()
        self._row_widgets = {}
        self.add_row()
        self.page.update()

    def save_to_excel(self, e=None):
        rows_data = [row.to_dict() for row in self.rows if self._row_has_data(row)]

        if not rows_data:
            self._show_dialog("تنبيه", "يرجى إدخال بيانات واحدة على الأقل قبل الحفظ.")
            return

        documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        blocks_path = os.path.join(documents_path, "البلوكات")
        os.makedirs(blocks_path, exist_ok=True)

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"blocks_{timestamp}.xlsx"
        filepath = os.path.join(blocks_path, filename)

        success, error = export_blocks_excel(
            filepath,
            block_size=self.block_size_field.value or "",
            notes=self.notes_field.value or "",
            rows=rows_data,
        )

        if success:
            self._show_success_dialog(filepath)
        else:
            if error == "file_locked":
                self._show_dialog("تحذير", "الملف مفتوح. يرجى إغلاقه ثم المحاولة مرة أخرى.")
            else:
                self._show_dialog("خطأ", f"تعذر الحفظ: {error}")

    def _row_has_data(self, row: BlockRow) -> bool:
        data = row.to_dict()
        keys_to_check = [
            "trip_number",
            "block_number",
            "material",
            "length",
            "width",
            "height",
        ]
        return any(data[key] for key in keys_to_check)

    def _show_dialog(self, title: str, message: str):
        dlg = ft.AlertDialog(title=ft.Text(title), content=ft.Text(message))
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def _show_success_dialog(self, filepath: str):
        def close_dlg(e=None):
            dialog.open = False
            self.page.update()

        def open_file(e=None):
            try:
                os.startfile(filepath)
            except Exception:
                pass

        def open_folder(e=None):
            try:
                folder = os.path.dirname(filepath)
                os.startfile(folder)
            except Exception:
                pass

        dialog = ft.AlertDialog(
            title=ft.Text("تم الحفظ"),
            content=ft.Text(f"تم إنشاء الملف بنجاح:\n{filepath}"),
            actions=[
                ft.TextButton("فتح الملف", on_click=open_file),
                ft.TextButton("فتح المجلد", on_click=open_folder),
                ft.TextButton("إغلاق", on_click=close_dlg),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.overlay.append(dialog)
        dialog.open = True
        self.page.update()


