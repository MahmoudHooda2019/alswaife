import os
from datetime import datetime
import flet as ft
from openpyxl import load_workbook
import xlsxwriter
import traceback

from utils.blocks_utils import export_simple_blocks_excel

class BlockRow:
    """Row UI for block entry with improved styling"""
    
    MATERIAL_OPTIONS = [
        ft.dropdown.Option("نيوحلايب"),
        ft.dropdown.Option("جندولا"),
        ft.dropdown.Option("احمر اسوان"),
    ]

    MACHINE_OPTIONS = [
        ft.dropdown.Option("1"),
        ft.dropdown.Option("2"),
        ft.dropdown.Option("3"),
    ]

    @staticmethod
    def _normalize_material_name(material):
        """Normalize material names to handle variations"""
        if not material:
            return material
        material = material.strip()
        if material == "نيو حلايب":
            return "نيوحلايب"
        return material

    def __init__(self, page: ft.Page, delete_callback, data=None):
        self.page = page
        self.delete_callback = delete_callback
        self._build_controls()
        self._set_default_values()

    def _create_styled_textfield(self, label, width, **kwargs):
        """Create a consistently styled text field"""
        # Extract bgcolor if provided in kwargs, otherwise use default
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

    def _create_styled_dropdown(self, label, width, options, **kwargs):
        """Create a consistently styled dropdown"""
        # Extract bgcolor if provided in kwargs, otherwise use default
        bgcolor = kwargs.pop('bgcolor', ft.Colors.BLUE_GREY_900)
        
        return ft.Dropdown(
            label=label,
            width=width,
            border_radius=10,
            filled=True,
            bgcolor=bgcolor,
            border_color=ft.Colors.GREY_700,
            focused_border_color=ft.Colors.WHITE,
            label_style=ft.TextStyle(color=ft.Colors.GREY_400),
            text_style=ft.TextStyle(size=14, weight=ft.FontWeight.W_500, color=ft.Colors.WHITE),
            options=options,
            **kwargs
        )

    def _build_controls(self):
        """Build all UI controls with improved styling"""
        width_small = 130
        width_medium = 160
        width_large = 190
        numeric_filter = ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$")
        
        # Date field
        self.date_field = self._create_styled_textfield(
            "التاريخ",
            width_medium,
            value=datetime.now().strftime("%Y-%m-%d"),
            read_only=True,
            icon=ft.Icons.CALENDAR_TODAY
        )
        
        # Block number
        self.block_number = self._create_styled_textfield(
            "رقم البلوك",
            width_small,
            on_change=self._on_field_change,
            icon=ft.Icons.TAG
        )
        
        # Material dropdown
        self.material_dropdown = self._create_styled_dropdown(
            "الخامة/النوع",
            width_medium,
            self.MATERIAL_OPTIONS,
            on_change=self._on_material_change,
            icon=ft.Icons.CATEGORY
        )
        
        # Machine number dropdown
        self.machine_dropdown = self._create_styled_dropdown(
            "رقم الماكينة",
            width_small,
            self.MACHINE_OPTIONS,
            on_change=self._on_field_change,
            icon=ft.Icons.PRECISION_MANUFACTURING
        )
        
        # Entry time
        self.entry_time = self._create_styled_textfield(
            "وقت الدخول",
            width_small,
            on_change=self._on_field_change,
            icon=ft.Icons.LOGIN
        )
        
        # Exit time
        self.exit_time = self._create_styled_textfield(
            "وقت الخروج",
            width_small,
            on_change=self._on_field_change,
            icon=ft.Icons.LOGOUT
        )
        
        # Hours count
        self.hours_count = self._create_styled_textfield(
            "عدد الساعات",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            icon=ft.Icons.ACCESS_TIME
        )
        
        # Thickness dropdown
        self.thickness_dropdown = self._create_styled_dropdown(
            "السمك",
            width_small,
            [
                ft.dropdown.Option("2سم"),
                ft.dropdown.Option("3سم"),
                ft.dropdown.Option("4سم"),
            ],
            on_change=self._on_field_change,
            value="2سم",
            icon=ft.Icons.HEIGHT
        )
        
        # Quantity
        self.quantity = self._create_styled_textfield(
            "العدد",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            icon=ft.Icons.NUMBERS
        )
        
        # Length field
        self.length_field = self._create_styled_textfield(
            "الطول",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            suffix_text="م"
        )
        
        # Height field
        self.height_field = self._create_styled_textfield(
            "الارتفاع",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            suffix_text="م"
        )
        
        # Weight field
        self.weight_field = self._create_styled_textfield(
            "الوزن",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            suffix_text="طن"
        )
        
        # Price per ton field
        self.price_per_ton_field = self._create_styled_textfield(
            "سعر الطن",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            prefix_text="ج.م "
        )
        
        # Trip price field
        self.trip_price_field = self._create_styled_textfield(
            "سعر النقلة",
            width_medium,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            prefix_text="ج.م "
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
        
        # Create the main card with gradient background
        self.card = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # Header with delete button
                        ft.Row(
                            controls=[
                                ft.Container(expand=True),
                                self.delete_btn
                            ],
                            alignment=ft.MainAxisAlignment.END
                        ),
                        
                        ft.Divider(height=20, color=ft.Colors.GREY_700),
                        
                        # Two-column layout
                        ft.Row(
                            controls=[
                                # Right column - Basic info
                                ft.Container(
                                    content=ft.Column(
                                        controls=[
                                            create_section_header("المعلومات الأساسية", ft.Colors.GREY_300),
                                            
                                            ft.Row(
                                                controls=[
                                                    self.date_field,
                                                    self.block_number,
                                                ],
                                                spacing=15,
                                                wrap=True
                                            ),
                                            
                                            ft.Row(
                                                controls=[
                                                    self.material_dropdown,
                                                    self.machine_dropdown,
                                                ],
                                                spacing=15,
                                                wrap=True
                                            ),
                                            
                                            ft.Divider(height=15, color=ft.Colors.TRANSPARENT),
                                            
                                            create_section_header("التوقيت", ft.Colors.GREY_300),
                                            
                                            ft.Row(
                                                controls=[
                                                    self.entry_time,
                                                    self.exit_time,
                                                    self.hours_count,
                                                ],
                                                spacing=15,
                                                wrap=True
                                            ),
                                        ],
                                        spacing=12
                                    ),
                                    expand=True,
                                    padding=15
                                ),
                                
                                # Left column - Measurements
                                ft.Container(
                                    content=ft.Column(
                                        controls=[
                                            create_section_header("القياسات", ft.Colors.GREY_300),
                                            
                                            ft.Row(
                                                controls=[
                                                    self.thickness_dropdown,
                                                    self.quantity,
                                                ],
                                                spacing=15,
                                                wrap=True
                                            ),
                                            
                                            ft.Row(
                                                controls=[
                                                    self.length_field,
                                                    self.height_field,
                                                ],
                                                spacing=15,
                                                wrap=True
                                            ),
                                            
                                            ft.Divider(height=15, color=ft.Colors.TRANSPARENT),
                                            
                                            create_section_header("التسعير", ft.Colors.GREY_300),
                                            
                                            ft.Row(
                                                controls=[
                                                    self.weight_field,
                                                    self.price_per_ton_field,
                                                ],
                                                spacing=15,
                                                wrap=True
                                            ),
                                            
                                            ft.Row(
                                                controls=[
                                                    self.trip_price_field,
                                                ],
                                                spacing=15,
                                                wrap=True
                                            ),
                                        ],
                                        spacing=12
                                    ),
                                    expand=True,
                                    padding=15
                                ),
                            ],
                            spacing=20,
                            vertical_alignment=ft.CrossAxisAlignment.START
                        ),
                    ],
                    spacing=15
                ),
                padding=20,
                gradient=ft.LinearGradient(
                    begin=ft.alignment.top_left,
                    end=ft.alignment.bottom_right,
                    colors=[ft.Colors.GREY_900, ft.Colors.GREY_800]
                ),
                border_radius=15,
                border=ft.border.all(1, ft.Colors.GREY_700)
            ),
            elevation=8,
        )
        
        self.row = self.card

    def _on_field_change(self, e=None):
        """Handle field changes and trigger calculations"""
        self._calculate_values()
    
    def _on_material_change(self, e=None):
        """Handle material change and update weight per m3 and price per ton"""
        material = self._normalize_material_name(self.material_dropdown.value)
        
        if material == "نيوحلايب":
            self.weight_field.value = "2.70"
            self.price_per_ton_field.value = "1150"
        elif material == "جندولا":
            self.weight_field.value = "2.85"
            self.price_per_ton_field.value = "1500"
        elif material == "احمر اسوان":
            self.weight_field.value = "0"
            self.price_per_ton_field.value = "0"
            
        self._calculate_values()
        self.page.update()
    
    def _calculate_values(self):
        """Calculate all dependent values with error handling"""
        # Calculations removed from UI per user request - only keeping for Excel export
        self.page.update()
    
    def _set_default_values(self):
        """Set default values for the row"""
        if not self.machine_dropdown.value:
            self.machine_dropdown.value = "1"
        if not self.material_dropdown.value:
            self.material_dropdown.value = "نيوحلايب"
        if not self.length_field.value:
            self.length_field.value = "1.0"
        if not self.thickness_dropdown.value:
            self.thickness_dropdown.value = "2سم"
        if not self.quantity.value:
            self.quantity.value = "1"
        if not self.height_field.value:
            self.height_field.value = "1.0"
        
        self._on_material_change()
    
    def to_dict(self):
        """Convert row data to dictionary for export"""
        # Return raw data for Excel formulas to calculate
        return {
            'trip_number': "",
            'trip_count': "",
            'date': self.date_field.value,
            'quarry': "",
            'machine_number': self.machine_dropdown.value,
            'block_number': self.block_number.value,
            'material': self.material_dropdown.value,
            'length': self.length_field.value,
            'width': "",  # Will be calculated by Excel formula
            'height': self.height_field.value,
            'weight': self.weight_field.value,
            'block_weight': "",  # Will be calculated by Excel formula
            'price_per_ton': self.price_per_ton_field.value,
            'total_price': "",  # Will be calculated by Excel formula
            'trip_price': self.trip_price_field.value,
            # Additional data for Excel export calculations
            'thickness_dropdown': self.thickness_dropdown.value,
            'quantity': self.quantity.value,
            'length_field': self.length_field.value,
            'height_field': self.height_field.value
        }


class BlocksView:
    """Main view for blocks management with improved UX"""
    _instance = None

    def __init__(self, page: ft.Page, on_back=None):
        self.__class__._instance = self
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - حساب تكلفة البلوكات مصنع محب"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK

        self.rows: list[BlockRow] = []
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
        return any([
            row.block_number.value,
            row.material_dropdown.value != "نيوحلايب",
            row.machine_dropdown.value != "1",
            row.entry_time.value,
            row.exit_time.value,
            row.hours_count.value,
            row.trip_price_field.value
        ])

    def build_ui(self):
        """Build the main UI with enhanced styling"""
        
        # Enhanced header with add button
        title_row = ft.Container(
            content=ft.Row(
                controls=[
                    ft.IconButton(
                        icon=ft.Icons.ARROW_BACK_ROUNDED,
                        on_click=self.go_back,
                        tooltip="العودة",
                        icon_size=28,
                        bgcolor=ft.Colors.GREY_800,
                        icon_color=ft.Colors.WHITE,
                        style=ft.ButtonStyle(
                            shape=ft.RoundedRectangleBorder(radius=12)
                        )
                    ),
                    ft.Container(width=15),
                    ft.Icon(ft.Icons.INVENTORY_2_ROUNDED, size=36, color=ft.Colors.BLUE_300),
                    ft.Container(width=10),
                    ft.Text(
                        "حساب تكلفة البلوكات مصنع محب",
                        size=32,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.BLUE_100
                    ),
                    ft.Container(expand=True),
                    ft.FilledButton(
                        "إضافة بلوك",
                        icon=ft.Icons.ADD_ROUNDED,
                        on_click=self.add_row,
                        style=ft.ButtonStyle(
                            bgcolor=ft.Colors.GREY_700,
                            color=ft.Colors.WHITE,
                            shape=ft.RoundedRectangleBorder(radius=12),
                            padding=15
                        )
                    ),
                ],
                alignment=ft.MainAxisAlignment.START
            ),
            padding=20,
            bgcolor=ft.Colors.GREY_900,
            border_radius=15,
            margin=ft.margin.only(bottom=10)
        )
        
        # Enhanced save button
        self.save_button = ft.FloatingActionButton(
            icon=ft.Icons.SAVE_ROUNDED,
            on_click=self.save_to_excel,
            bgcolor=ft.Colors.GREEN_700,
            tooltip="حفظ البيانات",
            elevation=8
        )
        
        # Main layout
        main_column = ft.Column(
            controls=[
                title_row,
                self.rows_container
            ],
            spacing=15,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )
        
        self.page.add(
            ft.Container(
                content=main_column,
                expand=True,
                padding=15,
                bgcolor=ft.Colors.GREY_900
            )
        )
        
        # Add initial row after the page is set up
        self.add_row()
        
        # Set floating action button (only the save button)
        self.page.floating_action_button = self.save_button
        
        self.page.update()
        return main_column

    def go_back(self, e):
        """Navigate back to previous view"""
        if self.on_back:
            self.on_back()

    def add_row(self, e=None):
        """Add a new block row and scroll to it"""
        row = BlockRow(
            page=self.page,
            delete_callback=self.delete_row
        )
        self.rows.append(row)
        self.rows_container.controls.append(row.row)
        
        # Scroll to the newly added row
        self.page.update()
        # Simple approach to scroll to bottom, but only if control is attached to page
        try:
            if hasattr(self.rows_container, 'scroll_to') and self.rows_container.uid is not None:
                self.rows_container.scroll_to(key="bottom", duration=300)
        except Exception:
            # If scrolling fails, it's not critical - just continue
            pass
        self.page.update()

    def delete_row(self, row_obj):
        """Delete a specific row"""
        if row_obj in self.rows:
            self.rows.remove(row_obj)
            self.rows_container.controls.remove(row_obj.row)
            self.page.update()

    def save_to_excel(self, e=None):
        """Save data to Excel file with validation"""
        if not any(self._row_has_data(row) for row in self.rows):
            self._show_dialog("تحذير", "لا توجد بيانات لحفظها", ft.Colors.ORANGE_400)
            return

        try:
            data = [row.to_dict() for row in self.rows if self._row_has_data(row)]
            file_path = export_simple_blocks_excel(data)
            
            # Reset rows after save
            for row in self.rows:
                row._set_default_values()
            
            self._show_success_dialog(file_path)
            
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ أثناء حفظ الملف:\n{str(e)}", ft.Colors.RED_400)
            traceback.print_exc()

    def _show_dialog(self, title: str, message: str, title_color=ft.Colors.BLUE_300):
        """Show a styled dialog"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()

        dlg = ft.AlertDialog(
            title=ft.Text(title, color=title_color, weight=ft.FontWeight.BOLD),
            content=ft.Text(message, size=16),
            actions=[
                ft.TextButton(
                    "إغلاق",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(
                        color=ft.Colors.BLUE_300
                    )
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.BLUE_GREY_900
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def _show_success_dialog(self, filepath: str):
        """Show success dialog with file actions"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()

        def open_file(e=None):
            close_dlg()
            try:
                os.startfile(filepath)
            except Exception:
                pass

        def open_folder(e=None):
            close_dlg()
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
                controls=[
                    ft.Text("تم إنشاء الملف بنجاح:", size=14),
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
                ft.TextButton(
                    "إغلاق",
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