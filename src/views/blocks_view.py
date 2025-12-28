import os
from datetime import datetime
from pickle import TRUE
import flet as ft
from openpyxl import load_workbook
import xlsxwriter

from utils.blocks_utils import export_simple_blocks_excel
from utils.log_utils import log_error, log_exception

class BlockRow:
    """Row UI for block entry with improved styling"""
    
    MATERIAL_OPTIONS = [
        ft.dropdown.Option("نيو حلايب"),
        ft.dropdown.Option("جندولا"),
        ft.dropdown.Option("احمر اسوان"),
    ]

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
        
        # Trip number
        self.trip_number = self._create_styled_textfield(
            "رقم النقله",
            width_small,
            on_change=self._on_field_change,
            icon=ft.Icons.NUMBERS
        )
        
        # Trip count
        self.trip_count = self._create_styled_textfield(
            "عدد النقله",
            width_small,
            on_change=self._on_field_change,
            icon=ft.Icons.TAG
        )
        
        # Date field
        self.date_field = self._create_styled_textfield(
            "التاريخ",
            width_medium,
            value=datetime.now().strftime("%Y-%m-%d"),
            read_only=True,
            icon=ft.Icons.CALENDAR_TODAY
        )
        
        # Quarry
        self.quarry = self._create_styled_textfield(
            "المحجر",
            width_medium,
            on_change=self._on_field_change,
            icon=ft.Icons.LOCATION_CITY
        )
        
        # Block number
        self.block_number = self._create_styled_textfield(
            "رقم البلوك",
            width_small,
            on_change=self.on_block_change,
            icon=ft.Icons.NUMBERS
        )
        
        # Material dropdown
        self.material_dropdown = self._create_styled_dropdown(
            "الخامه",
            width_medium,
            self.MATERIAL_OPTIONS,
            on_change=self._on_material_change,
            icon=ft.Icons.CATEGORY
        )
        
        # Length
        self.length_field = self._create_styled_textfield(
            "الطول",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            suffix_text="م"
        )
        
        # Width
        self.width_field = self._create_styled_textfield(
            "العرض",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            suffix_text="م"
        )
        
        # Height
        self.height_field = self._create_styled_textfield(
            "الارتفاع",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            suffix_text="م"
        )
        
        # Volume (م3) - read-only, calculated automatically
        self.volume_field = self._create_styled_textfield(
            "م3",
            width_small,
            read_only=True,
            icon=ft.Icons.CALCULATE,
            value="0.00"
        )
        
        # Weight per m3 - read-only, based on material
        self.weight_per_m3_field = self._create_styled_textfield(
            "الوزن",
            width_small,
            read_only=True,
            icon=ft.Icons.SCALE,
            value="0.00"
        )
        
        # Block weight - read-only, calculated automatically
        self.block_weight_field = self._create_styled_textfield(
            "وزن البلوك",
            width_medium,
            read_only=True,
            icon=ft.Icons.SCALE,
            value="0.00"
        )
        
        # Price per ton - read-only, based on material
        self.price_per_ton_field = self._create_styled_textfield(
            "سعر الطن",
            width_medium,
            read_only=True,
            icon=ft.Icons.ATTACH_MONEY,
            value="0.00"
        )
        
        # Total price - read-only, calculated automatically
        self.total_price_field = self._create_styled_textfield(
            "اجمالي السعر",
            width_medium,
            read_only=True,
            icon=ft.Icons.ATTACH_MONEY,
            value="0.00"
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
                        
                        # Basic Information Section
                        create_section_header("المعلومات الأساسية", ft.Colors.BLUE_400),
                        
                        ft.Row(
                            controls=[
                                self.trip_number,
                                self.trip_count,
                                self.date_field,
                                self.quarry,
                            ],
                            spacing=15,
                            wrap=True
                        ),
                        
                        ft.Row(
                            controls=[
                                self.block_number,
                                self.material_dropdown,
                            ],
                            spacing=15,
                            wrap=True
                        ),
                        
                        ft.Divider(height=15, color=ft.Colors.TRANSPARENT),
                        
                        # Measurements Section
                        create_section_header("القياسات", ft.Colors.GREEN_400),
                        
                        ft.Row(
                            controls=[
                                self.length_field,
                                self.width_field,
                                self.height_field,
                                self.volume_field,
                            ],
                            spacing=15,
                            wrap=True
                        ),
                        
                        ft.Divider(height=15, color=ft.Colors.TRANSPARENT),
                        
                        # Calculations Section
                        create_section_header("الحسابات", ft.Colors.ORANGE_400),
                        
                        ft.Row(
                            controls=[
                                self.weight_per_m3_field,
                                self.block_weight_field,
                                self.price_per_ton_field,
                                self.total_price_field,
                            ],
                            spacing=15,
                            wrap=True
                        ),
                    ],
                    spacing=12
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

    def on_block_change(self, e):
        """Handle block number changes and replace Arabic characters with English equivalents"""
        val = self.block_number.value
        if val:
            new_val = val.replace('ش', 'A').replace('لا', 'B').replace('a', 'A').replace('b', 'B').replace('أ', 'A').replace('ب', 'B').replace('ِ', 'A').replace('لآ', 'B')
            if new_val != val:
                self.block_number.value = new_val
                # Only update if value changed
                if hasattr(self, 'page') and self.page:
                    self.page.update()
        # Always trigger calculations after any change
        self._calculate_values()

    def _on_field_change(self, e=None):
        """Handle field changes and trigger calculations"""
        self._calculate_values()
    
    def _on_material_change(self, e=None):
        """Handle material change and update weight per m3 and price per ton"""
        material = self.material_dropdown.value
        
        if material == "نيو حلايب":
            self.weight_per_m3_field.value = "2.70"
            self.price_per_ton_field.value = "1150"
        elif material == "جندولا":
            self.weight_per_m3_field.value = "2.85"
            self.price_per_ton_field.value = "1500"
        elif material == "احمر اسوان":
            self.weight_per_m3_field.value = "0"
            self.price_per_ton_field.value = "0"
            
        self._calculate_values()
        self.page.update()
    
    def _calculate_values(self):
        """Calculate all dependent values with error handling"""
        try:
            # Get numeric values with error handling
            length = float(self.length_field.value) if self.length_field.value else 0
            width = float(self.width_field.value) if self.width_field.value else 0
            height = float(self.height_field.value) if self.height_field.value else 0
            weight_per_m3 = float(self.weight_per_m3_field.value) if self.weight_per_m3_field.value else 0
            price_per_ton = float(self.price_per_ton_field.value) if self.price_per_ton_field.value else 0
            
            # Calculate volume (م3) = length * width * height
            volume = length * width * height
            self.volume_field.value = f"{volume:.2f}"
            
            # Calculate block weight = volume * weight_per_m3
            block_weight = volume * weight_per_m3
            self.block_weight_field.value = f"{block_weight:.2f}"
            
            # Calculate total price = price_per_ton * block_weight
            total_price = price_per_ton * block_weight
            self.total_price_field.value = f"{total_price:.2f}"
            
        except ValueError:
            # If any field contains non-numeric values, set calculated fields to 0
            self.volume_field.value = "0.00"
            self.block_weight_field.value = "0.00"
            self.total_price_field.value = "0.00"
        
        self.page.update()
    
    def _set_default_values(self):
        """Set default values for the row"""
        if not self.material_dropdown.value:
            self.material_dropdown.value = "نيو حلايب"
        if not self.length_field.value:
            self.length_field.value = "1.0"
        if not self.width_field.value:
            self.width_field.value = "1.0"
        if not self.height_field.value:
            self.height_field.value = "1.0"
        
        self._on_material_change()
    
    def to_dict(self):
        """Convert row data to dictionary for export"""
        # Return data with calculated values for Excel export
        return {
            'trip_number': self.trip_number.value,
            'trip_count': self.trip_count.value,
            'date': self.date_field.value,
            'quarry': self.quarry.value,
            'block_number': self.block_number.value,
            'material': self.material_dropdown.value,
            'length': self.length_field.value,
            'width': self.width_field.value,
            'height': self.height_field.value,
            'volume': self.volume_field.value,  # Calculated: length * width * height
            'weight_per_m3': self.weight_per_m3_field.value,  # Based on material
            'block_weight': self.block_weight_field.value,  # Calculated: volume * weight_per_m3
            'price_per_ton': self.price_per_ton_field.value,  # Based on material
            'total_price': self.total_price_field.value,  # Calculated: price_per_ton * block_weight
        }


class BlocksView:
    """Main view for blocks management with improved UX"""
    _instance = None

    def __init__(self, page: ft.Page, on_back=None):
        self.__class__._instance = self
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - مخزون البلوكات"
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
            row.trip_number.value,
            row.trip_count.value,
            row.quarry.value,
            row.block_number.value,
            row.material_dropdown.value != "نيو حلايب",
            row.length_field.value,
            row.width_field.value,
            row.height_field.value,
        ])

    def build_ui(self):
        """Build the main UI with enhanced styling"""
        
        # Create AppBar with title and actions
        app_bar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK,
                on_click=self.go_back,
                tooltip="العودة"
            ),
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.INVENTORY_2_ROUNDED, size=24, color=ft.Colors.BLUE_200),
                    ft.Text(
                        "مخزون البلوكات",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.BLUE_200
                    ),
                ],
                spacing=10
            ),
            actions=[
                ft.IconButton(
                    icon=ft.Icons.ADD,
                    on_click=self.add_row,
                    tooltip="إضافة بلوك جديد"
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
        
        # Set the AppBar (page.clean() was already called by dashboard)
        self.page.appbar = app_bar
        
        # Main layout - Column with scroll for content below AppBar
        main_column = ft.Column(
            controls=[
                self.rows_container
            ],
            spacing=15,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )
        
        # Add the main content to the page
        self.page.add(main_column)
        
        # Add initial row after the page is set up
        self.add_row()
        
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
            
        except PermissionError as e:
            self._show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ أثناء حفظ الملف:\n{str(e)}", ft.Colors.RED_400)
            log_exception(f"Error saving blocks file: {e}")

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
                rtl=True,
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
                        margin=ft.margin.only(top=10),
                        rtl=True
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