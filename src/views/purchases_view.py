import os
import flet as ft
from utils.purchases_utils import export_purchases_to_excel, load_item_names_from_excel
from utils.utils import is_excel_running, get_current_date, is_file_locked
from utils.db_utils import get_purchases_zoom_level, set_purchases_zoom_level
from utils.bottom_sheet_utils import BottomSheetManager


class PurchaseRow:
    """Row UI for expense entry with improved styling"""
    
    def __init__(self, page: ft.Page, delete_callback, items_list=None):
        self.page = page
        self.delete_callback = delete_callback
        self.items_list = items_list or []
        self._build_controls()
    
    def get_editable_fields(self):
        """Return list of editable fields in order for navigation"""
        return [
            self.date_field,         # 0
            self.quantity_field,     # 1
            self.item_name_field,    # 2
            self.total_price_field,  # 3
        ]
    
    def focus_field(self, field_index):
        """Focus a specific field by index"""
        fields = self.get_editable_fields()
        if 0 <= field_index < len(fields):
            field = fields[field_index]
            try:
                field.focus()
                if self.page:
                    self.page.update()
                return True
            except Exception:
                return False
        return False

    def _create_styled_textfield(self, label, width, **kwargs):
        """Create a consistently styled text field"""
        bgcolor = kwargs.pop('bgcolor', ft.Colors.BLUE_GREY_900)
        
        return ft.TextField(
            label=label,
            width=width,
            border_radius=10,
            filled=True,
            bgcolor=bgcolor,
            border_color=ft.Colors.GREY_700,
            focused_border_color=ft.Colors.WHITE,
            label_style=ft.TextStyle(color=ft.Colors.GREY_400),
            text_style=ft.TextStyle(size=14, weight=ft.FontWeight.W_500, color=ft.Colors.WHITE),
            cursor_color=ft.Colors.WHITE,
            **kwargs
        )

    def _build_controls(self):
        """Build all UI controls with improved styling"""
        width_small = 130
        width_medium = 160
        width_large = 440
        numeric_filter = ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$")
        
        # Date field
        self.date_field = self._create_styled_textfield(
            "تاريخ الصرف",
            width_medium,
            value=get_current_date("%d/%m/%Y"),
            icon=ft.Icons.CALENDAR_TODAY
        )
        
        # Quantity field
        self.quantity_field = self._create_styled_textfield(
            "العدد",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            icon=ft.Icons.TAG
        )
        
        # Item name field (البيان)
        self.item_name_field = self._create_styled_textfield(
            "البيان",
            width_large,
            on_change=self._on_item_name_change,
            icon=ft.Icons.DESCRIPTION
        )
        
        # Suggestions container
        self.item_suggestions = ft.Column(
            visible=False,
            spacing=0
        )
        
        # Total price field (المبلغ)
        self.total_price_field = self._create_styled_textfield(
            "المبلغ",
            width_medium,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            icon=ft.Icons.ATTACH_MONEY
        )
        
        # Delete button
        self.delete_btn = ft.IconButton(
            icon=ft.Icons.DELETE_OUTLINE,
            icon_color=ft.Colors.RED_400,
            tooltip="حذف الصف",
            on_click=lambda e: self.delete_callback(self),
            bgcolor=ft.Colors.GREY_800,
            icon_size=20,
            style=ft.ButtonStyle(
                shape=ft.RoundedRectangleBorder(radius=10)
            )
        )
        
        # Section header style
        def create_section_header(text, color):
            return ft.Container(
                content=ft.Row(
                    controls=[
                        ft.Container(
                            width=4,
                            height=24,
                            bgcolor=color,
                            border_radius=2
                        ),
                        ft.Text(
                            text,
                            weight=ft.FontWeight.BOLD,
                            size=16,
                            color=color
                        ),
                    ],
                    spacing=10
                ),
                margin=ft.margin.only(bottom=10)
            )
        
        # Create the main card
        self.card = ft.Container(
            content=ft.Row(
                controls=[
                    self.date_field,
                    self.quantity_field,
                    ft.Column([
                        self.item_name_field,
                        self.item_suggestions,
                    ], spacing=0),
                    self.total_price_field,
                    self.delete_btn,
                ],
                spacing=12,
                vertical_alignment=ft.CrossAxisAlignment.CENTER
            ),
            padding=ft.padding.symmetric(horizontal=15, vertical=10),
            bgcolor=ft.Colors.GREY_900,
            border_radius=12,
            border=ft.border.all(1, ft.Colors.GREY_700),
        )
        
        self.row = self.card

    def _on_item_name_change(self, e):
        """Handle item name change with auto-complete suggestions"""
        search_text = self.item_name_field.value.strip().lower() if self.item_name_field.value else ""
        
        if not search_text:
            self.item_suggestions.visible = False
            self.item_suggestions.controls.clear()
            self.page.update()
            return
        
        filtered = [item for item in self.items_list if search_text in item.lower()]
        
        if filtered:
            self.item_suggestions.controls.clear()
            for item in filtered[:5]:
                suggestion_btn = ft.TextButton(
                    text=item,
                    on_click=lambda e, i=item: self._select_item_suggestion(i),
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

    def _select_item_suggestion(self, item_name):
        """Select an item from suggestions"""
        self.item_name_field.value = item_name
        self.item_suggestions.visible = False
        self.item_suggestions.controls.clear()
        self.page.update()

    def to_dict(self):
        """Convert row data to dictionary for export"""
        return {
            'date': self.date_field.value or "",
            'quantity': self.quantity_field.value or "",
            'item_name': self.item_name_field.value or "",
            'total_price': self.total_price_field.value or "",
        }

    def clear(self):
        """Clear all fields except date"""
        self.quantity_field.value = ""
        self.item_name_field.value = ""
        self.total_price_field.value = ""

    def update_scale(self, scale_factor):
        """Update the scale of all fields"""
        base_text_size = 14
        new_text_size = base_text_size * scale_factor
        
        # Default widths
        width_small = 130
        width_medium = 160
        width_large = 440
        
        # Update text fields
        fields = [
            (self.date_field, width_medium),
            (self.quantity_field, width_small),
            (self.item_name_field, width_large),
            (self.total_price_field, width_medium),
        ]
        
        for field, base_width in fields:
            field.width = base_width * scale_factor
            field.text_style = ft.TextStyle(
                size=new_text_size,
                weight=ft.FontWeight.W_500,
                color=ft.Colors.WHITE
            )
            field.label_style = ft.TextStyle(
                color=ft.Colors.GREY_400,
                size=new_text_size * 0.9
            )


class PurchasesView:
    """Main view for expenses management"""
    _instance = None

    def __init__(self, page: ft.Page, on_back=None):
        self.__class__._instance = self
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - صرف"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK

        # Initialize data storage
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        self.purchases_path = os.path.join(self.documents_path, "ايرادات ومصروفات")
        os.makedirs(self.purchases_path, exist_ok=True)
        
        # Database path for zoom level
        self.db_path = os.path.join(self.documents_path, 'invoice.db')
        
        # Load saved zoom level
        self.scale_factor = get_purchases_zoom_level(self.db_path)
        
        # Load existing items for auto-complete
        self.excel_file = os.path.join(self.purchases_path, "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
        self.items_list = load_item_names_from_excel(self.excel_file)

        self.rows: list[PurchaseRow] = []
        self.rows_container = ft.Column(
            spacing=20,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )

    @classmethod
    def get_instance(cls):
        """Get the singleton instance"""
        return cls._instance

    def _row_has_data(self, row) -> bool:
        """Check if a row has any meaningful data"""
        return bool(row.item_name_field.value and row.item_name_field.value.strip())

    def build_ui(self):
        """Build the main UI"""
        
        # Add keyboard event handler for arrow navigation
        self.page.on_keyboard_event = self.on_keyboard_event
        
        # Create AppBar
        app_bar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK,
                on_click=self.go_back,
                tooltip="العودة"
            ),
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.MONEY_OFF, size=24, color=ft.Colors.RED_300),
                    ft.Text(
                        "صرف",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.RED_200
                    ),
                ],
                spacing=10
            ),
            actions=[
                ft.Container(
                    content=ft.IconButton(
                        icon=ft.Icons.FOLDER_OPEN,
                        icon_color=ft.Colors.BLUE_300,
                        on_click=self.open_purchases_file,
                        tooltip="فتح ملف السجل",
                        icon_size=24,
                    ),
                    margin=ft.margin.only(right=5, left=5),
                    bgcolor=ft.Colors.GREY_800,
                    border_radius=20,
                ),
                ft.IconButton(
                    icon=ft.Icons.ADD,
                    on_click=self.add_row,
                    tooltip="إضافة صف جديد"
                ),
                ft.Container(
                    content=ft.Row(
                        controls=[
                            ft.IconButton(
                                ft.Icons.ZOOM_IN,
                                on_click=self.zoom_in,
                                tooltip="تكبير",
                                icon_color=ft.Colors.BLUE_300,
                            ),
                            ft.IconButton(
                                ft.Icons.ZOOM_OUT,
                                on_click=self.zoom_out,
                                tooltip="تصغير",
                                icon_color=ft.Colors.BLUE_300,
                            ),
                        ],
                        spacing=0,
                    ),
                    bgcolor=ft.Colors.GREY_800,
                    border_radius=20,
                    padding=ft.padding.symmetric(horizontal=5),
                ),
                ft.Container(
                    content=ft.IconButton(
                        icon=ft.Icons.SAVE,
                        on_click=self.save_to_excel,
                        tooltip="حفظ البيانات"
                    ),
                    margin=ft.margin.only(left=40, right=15)
                )
            ],
            bgcolor=ft.Colors.GREY_900,
        )
        
        self.page.appbar = app_bar
        
        # Main layout
        main_column = ft.Column(
            controls=[
                self.rows_container
            ],
            spacing=15,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )
        
        self.page.add(main_column)
        
        # Add initial row
        self.add_row()
        
        self.page.update()
        return main_column

    def go_back(self, e):
        """Navigate back to previous view"""
        if self.on_back:
            self.on_back()

    def add_row(self, e=None):
        """Add a new expense row"""
        row = PurchaseRow(
            page=self.page,
            delete_callback=self.delete_row,
            items_list=self.items_list
        )
        # Apply current zoom level to new row
        row.update_scale(self.scale_factor)
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
        if not any(self._row_has_data(row) for row in self.rows):
            self._show_dialog("تحذير", "لا توجد بيانات لحفظها", ft.Colors.ORANGE_400)
            return

        # التحقق من أن Excel مغلق
        if is_excel_running():
            self._show_excel_warning_dialog()
            return

        self._do_save()

    def _do_save(self):
        """تنفيذ عملية الحفظ الفعلية"""
        # التحقق من أن الملف غير مفتوح
        if is_file_locked(self.excel_file):
            self._show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
            return
        
        try:
            data = [row.to_dict() for row in self.rows if self._row_has_data(row)]
            export_purchases_to_excel(data, self.excel_file)
            
            # Add new items to auto-complete list
            for record in data:
                if record['item_name'] and record['item_name'] not in self.items_list:
                    self.items_list.append(record['item_name'])
            
            # Clear rows after save
            for row in self.rows:
                row.clear()
            
            self._show_success_dialog(self.excel_file)
            
        except PermissionError:
            self._show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ أثناء حفظ الملف:\n{str(e)}", ft.Colors.RED_400)

    def _show_excel_warning_dialog(self):
        """Show Excel warning dialog with continue option"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()

        def continue_save(e=None):
            dlg.open = False
            self.page.update()
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
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def open_purchases_file(self, e):
        """Open the purchases Excel file"""
        if os.path.exists(self.excel_file):
            try:
                os.startfile(self.excel_file)
            except Exception:
                self._show_dialog("خطأ", "لا يمكن فتح الملف حالياً", ft.Colors.RED_400)
        else:
            self._show_dialog("معلومات", "ملف السجل لم يتم إنشاؤه بعد", ft.Colors.BLUE_400)

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
                    "إغلاق",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.BLUE_300)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.BLUE_GREY_900
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def _show_success_dialog(self, filepath: str):
        """Show success bottom sheet with file actions"""
        BottomSheetManager.show_success_bottom_sheet(
            page=self.page,
            message="تم إضافة المصروف إلى السجل بنجاح",
            filepath=filepath,
            title="تم الحفظ بنجاح",
        )

        self.page.update()

    def zoom_in(self, e=None):
        """Increase the scale factor"""
        if self.scale_factor < 2.0:
            self.scale_factor += 0.05
            self._apply_zoom()
            # Save zoom level to database
            set_purchases_zoom_level(self.db_path, self.scale_factor)

    def zoom_out(self, e=None):
        """Decrease the scale factor"""
        if self.scale_factor > 0.5:
            self.scale_factor -= 0.05
            self._apply_zoom()
            # Save zoom level to database
            set_purchases_zoom_level(self.db_path, self.scale_factor)

    def _apply_zoom(self):
        """Apply the current zoom level to all rows"""
        for row in self.rows:
            row.update_scale(self.scale_factor)
        self.page.update()

    def on_keyboard_event(self, e: ft.KeyboardEvent):
        """Handle keyboard events for arrow navigation"""
        # Check if the '+' key was pressed to add a new row
        if e.key == "+" and not e.shift and not e.ctrl and not e.alt:
            self.add_row()
            return
        
        # Only handle arrow keys
        if e.key not in ["Arrow Down", "Arrow Up", "Arrow Left", "Arrow Right"]:
            return
        
        # Find currently focused element
        focused_row_idx = None
        focused_field_idx = None
        
        for row_idx, row in enumerate(self.rows):
            fields = row.get_editable_fields()
            for field_idx, field in enumerate(fields):
                try:
                    if hasattr(field, 'focus') and field == self.page.focused_control:
                        focused_row_idx = row_idx
                        focused_field_idx = field_idx
                        break
                except:
                    pass
            if focused_row_idx is not None:
                break
        
        # If no field is focused, do nothing
        if focused_row_idx is None or focused_field_idx is None:
            return
        
        # Calculate new position based on arrow key
        new_row_idx = focused_row_idx
        new_field_idx = focused_field_idx
        
        if e.key == "Arrow Down":
            new_row_idx = min(focused_row_idx + 1, len(self.rows) - 1)
        elif e.key == "Arrow Up":
            new_row_idx = max(focused_row_idx - 1, 0)
        elif e.key == "Arrow Right":
            new_field_idx = max(focused_field_idx - 1, 0)  # RTL: right goes to previous field
        elif e.key == "Arrow Left":
            fields_count = len(self.rows[focused_row_idx].get_editable_fields())
            new_field_idx = min(focused_field_idx + 1, fields_count - 1)  # RTL: left goes to next field
        
        # Focus the new field
        if new_row_idx < len(self.rows):
            self.rows[new_row_idx].focus_field(new_field_idx)
