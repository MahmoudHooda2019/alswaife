import flet as ft
import json
import os
from datetime import datetime
from utils.path_utils import resource_path
from utils.slides_utils import initialize_slides_inventory_excel, add_slides_inventory_entry, convert_existing_slides_inventory_to_formulas


class SlideRow:
    """Row UI for slide entry with styling similar to blocks view"""

    MATERIAL_OPTIONS = [
        ft.dropdown.Option("نيو حلايب"),
        ft.dropdown.Option("جندولا"),
        ft.dropdown.Option("احمر اسوان"),
    ]

    MACHINE_OPTIONS = [
        ft.dropdown.Option("1"),
        ft.dropdown.Option("2"),
        ft.dropdown.Option("3"),
    ]

    THICKNESS_OPTIONS = [
        ft.dropdown.Option("2سم"),
        ft.dropdown.Option("3سم"),
        ft.dropdown.Option("4سم"),
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
        numeric_filter = ft.InputFilter(regex_string=r"^[0-9]*$")

        # Publishing date
        self.publishing_date = self._create_styled_textfield(
            "تاريخ النشر",
            width_medium,
            value=datetime.now().strftime("%Y-%m-%d"),
            read_only=True,
            icon=ft.Icons.CALENDAR_TODAY
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
            "النوع",
            width_medium,
            self.MATERIAL_OPTIONS,
            on_change=self._on_material_change,
            icon=ft.Icons.CATEGORY
        )

        # Machine number dropdown
        self.machine_number = self._create_styled_dropdown(
            "رقم المكينه",
            width_small,
            self.MACHINE_OPTIONS,
            icon=ft.Icons.SETTINGS
        )

        # Quantity
        self.quantity_field = self._create_styled_textfield(
            "عدد",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
            on_change=self._on_field_change,
            icon=ft.Icons.NUMBERS
        )

        # Length
        self.length_field = self._create_styled_textfield(
            "الطول",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*ز*$"),
            on_change=self._on_field_change,
            suffix_text="م"
        )

        # Height
        self.height_field = self._create_styled_textfield(
            "الارتفاع",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*ز*$"),
            on_change=self._on_field_change,
            suffix_text="م"
        )

        # Thickness dropdown
        self.thickness_dropdown = self._create_styled_dropdown(
            "السمك",
            width_small,
            self.THICKNESS_OPTIONS,
            on_change=self._on_thickness_change,
            icon=ft.Icons.LINE_WEIGHT
        )

        # Area (م2) - read-only, calculated automatically
        self.area_field = self._create_styled_textfield(
            "م2",
            width_small,
            read_only=True,
            icon=ft.Icons.CALCULATE,
            value="0.00"
        )

        # Price per meter - editable, with default based on material and thickness
        self.price_per_meter_field = self._create_styled_textfield(
            "سعر المتر",
            width_medium,
            icon=ft.Icons.ATTACH_MONEY,
            value="0.00",
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$"),
            on_change=self._on_field_change
        )

        # Total price - read-only, calculated automatically
        self.total_price_field = self._create_styled_textfield(
            "اجمالي السعر",
            width_medium,
            read_only=True,
            icon=ft.Icons.ATTACH_MONEY,
            value="0.00"
        )

        # Entry time
        self.entry_time = self._create_styled_textfield(
            "وقت الدخول",
            width_medium,
            icon=ft.Icons.ACCESS_TIME
        )

        # Exit time
        self.exit_time = self._create_styled_textfield(
            "وقت الخروج",
            width_medium,
            icon=ft.Icons.ACCESS_TIME
        )

        # Hours count - read-only, calculated automatically
        self.hours_count = self._create_styled_textfield(
            "عدد الساعات",
            width_small,
            read_only=True,
            icon=ft.Icons.TIMER,
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
                                self.publishing_date,
                                self.block_number,
                                self.material_dropdown,
                                self.machine_number,
                            ],
                            spacing=15,
                            wrap=True
                        ),
                        
                        ft.Divider(height=15, color=ft.Colors.TRANSPARENT),
                        
                        # Measurements Section
                        create_section_header("القياسات", ft.Colors.GREEN_400),
                        
                        ft.Row(
                            controls=[
                                self.quantity_field,
                                self.length_field,
                                self.height_field,
                                self.thickness_dropdown,
                            ],
                            spacing=15,
                            wrap=True
                        ),
                        
                        ft.Divider(height=15, color=ft.Colors.TRANSPARENT),
                        
                        # Calculations Section
                        create_section_header("الحسابات", ft.Colors.ORANGE_400),
                        
                        ft.Row(
                            controls=[
                                self.area_field,
                                self.price_per_meter_field,
                                self.total_price_field,
                            ],
                            spacing=15,
                            wrap=True
                        ),
                        
                        ft.Divider(height=15, color=ft.Colors.TRANSPARENT),
                        
                        # Time Section
                        create_section_header("أوقات التشغيل", ft.Colors.PURPLE_400),
                        
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
        # Replace 'ز' with decimal point in length and height fields
        if e and e.control == self.length_field and self.length_field.value:
            new_value = self.length_field.value.replace('ز', '.')
            if new_value != self.length_field.value:
                self.length_field.value = new_value
        elif e and e.control == self.height_field and self.height_field.value:
            new_value = self.height_field.value.replace('ز', '.')
            if new_value != self.height_field.value:
                self.height_field.value = new_value
        
        self._calculate_values()
    
    def on_block_change(self, e):
        """Handle block number changes and replace Arabic characters with English equivalents"""
        val = self.block_number.value
        if val:
            # Replace Arabic characters with their English counterparts
            # 'ش' is 'a' on Arabic keyboard
            # 'لا' (lam-alif) is 'b' on Arabic keyboard
            new_val = val.replace('ش', 'A').replace('لا', 'B').replace('a', 'A').replace('b', 'B').replace('أ', 'A').replace('ب', 'B').replace('ِ', 'A').replace('لآ', 'B')
            if new_val != val:
                self.block_number.value = new_val
                # Only update if value changed
                if hasattr(self, 'page') and self.page:
                    self.page.update()
        # Always trigger calculations after any change
        self._calculate_values()

    def _on_material_change(self, e=None):
        """Handle material change and update price per meter"""
        material = self.material_dropdown.value
        thickness = self.thickness_dropdown.value
        
        if material and thickness:
            # Extract numeric part from thickness (e.g., "2سم" -> "2")
            thickness_value = ''.join(filter(str.isdigit, thickness))
            price = self._get_price_from_json(material, thickness_value)
            # Only update if the field is empty or has the default value
            if not self.price_per_meter_field.value or self.price_per_meter_field.value == "0.00":
                self.price_per_meter_field.value = str(price)
        
        self._calculate_values()
        self.page.update()

    def _on_thickness_change(self, e=None):
        """Handle thickness change and update price per meter"""
        self._on_material_change(e)

    def _get_price_from_json(self, material, thickness_value):
        """Get price from slides_products.json based on material and thickness"""
        try:
            # Load the JSON file - path relative to project root
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            json_path = os.path.join(project_root, "data", "slides_products.json")
            with open(json_path, 'r', encoding='utf-8') as f:
                products = json.load(f)
            
            # Find the material in the products list
            for product in products:
                if product["name"] == material:
                    prices = product["prices"]
                    if thickness_value in prices:
                        price_data = prices[thickness_value]
                        
                        # Handle different price structures (list or direct value)
                        if isinstance(price_data, list):
                            # For complex pricing with ranges, return first price
                            if len(price_data) > 0 and isinstance(price_data[0], dict) and 'price' in price_data[0]:
                                return price_data[0]['price']
                            else:
                                return 0
                        elif isinstance(price_data, (int, float)):
                            return price_data
                        else:
                            return 0
                    else:
                        return 0
            return 0
        except Exception as e:
            print(f"[ERROR] Error getting price from JSON: {e}")
            return 0

    def _calculate_values(self):
        """Calculate all dependent values with error handling"""
        try:
            # Get numeric values with error handling
            quantity = int(self.quantity_field.value) if self.quantity_field.value else 0
            length = float(self.length_field.value) if self.length_field.value else 0
            height = float(self.height_field.value) if self.height_field.value else 0
            price_per_meter = float(self.price_per_meter_field.value) if self.price_per_meter_field.value else 0
            
            # Calculate area (م2) = length * height * quantity
            area = length * height * quantity
            self.area_field.value = f"{area:.2f}"
            
            # Calculate total price = area * price_per_meter
            total_price = area * price_per_meter
            self.total_price_field.value = f"{total_price:.2f}"
            
            # Calculate hours from entry and exit time if both are provided
            if self.entry_time.value and self.exit_time.value:
                try:
                    # Parse time in HH:MM format
                    entry_parts = self.entry_time.value.split(':')
                    exit_parts = self.exit_time.value.split(':')
                    
                    if len(entry_parts) == 2 and len(exit_parts) == 2:
                        entry_hour = int(entry_parts[0])
                        entry_min = int(entry_parts[1])
                        exit_hour = int(exit_parts[0])
                        exit_min = int(exit_parts[1])
                        
                        entry_total_minutes = entry_hour * 60 + entry_min
                        exit_total_minutes = exit_hour * 60 + exit_min
                        
                        # Calculate difference in minutes and convert to hours
                        diff_minutes = exit_total_minutes - entry_total_minutes
                        if diff_minutes < 0:
                            # Handle overnight case
                            diff_minutes += 24 * 60
                        
                        hours = diff_minutes / 60.0
                        self.hours_count.value = f"{hours:.2f}"
                    else:
                        self.hours_count.value = "0.00"
                except ValueError:
                    self.hours_count.value = "0.00"
            
        except ValueError:
            # If any field contains non-numeric values, set calculated fields to 0
            self.area_field.value = "0.00"
            self.total_price_field.value = "0.00"
        
        self.page.update()

    def _set_default_values(self):
        """Set default values for the row"""
        if not self.material_dropdown.value:
            self.material_dropdown.value = "نيو حلايب"
        if not self.machine_number.value:
            self.machine_number.value = "1"
        if not self.thickness_dropdown.value:
            self.thickness_dropdown.value = "2سم"
        if not self.length_field.value:
            self.length_field.value = "1.0"
        if not self.height_field.value:
            self.height_field.value = "1.0"
        if not self.quantity_field.value:
            self.quantity_field.value = "1"
        
        self._on_material_change()

    def to_dict(self):
        """Convert row data to dictionary for export"""
        return {
            'publishing_date': self.publishing_date.value,
            'block_number': self.block_number.value,
            'material': self.material_dropdown.value,
            'machine_number': self.machine_number.value,
            'quantity': self.quantity_field.value,
            'length': self.length_field.value,
            'height': self.height_field.value,
            'thickness': self.thickness_dropdown.value,
            'area': self.area_field.value,  # Calculated: length * height * quantity
            'price_per_meter': self.price_per_meter_field.value,  # Based on material and thickness
            'total_price': self.total_price_field.value,  # Calculated: area * price_per_meter
            'entry_time': self.entry_time.value,
            'exit_time': self.exit_time.value,
            'hours_count': self.hours_count.value,  # Calculated: exit_time - entry_time
        }


class SlidesAddView:
    """View for adding slides inventory items with design similar to blocks section"""
    
    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - إضافة شرائح"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK
        
        # Initialize data storage
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        self.slides_path = os.path.join(self.documents_path, "الشرائح")
        os.makedirs(self.slides_path, exist_ok=True)
        
        self.rows: list[SlideRow] = []
        self.rows_container = ft.Column(
            spacing=20,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )

    def _row_has_data(self, row) -> bool:
        """Check if a row has any meaningful data"""
        return any([
            row.publishing_date.value,
            row.block_number.value,
            row.material_dropdown.value != "نيو حلايب",
            row.machine_number.value != "1",
            row.quantity_field.value,
            row.length_field.value,
            row.height_field.value,
        ])

    def build_ui(self):
        """Build the slides add UI with design similar to blocks section"""
        
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
                        "إضافة شرائح",
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
                    tooltip="إضافة شريحة جديدة"
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
        """Add a new slide row and scroll to it"""
        row = SlideRow(
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
            
            # Create Excel file for slides inventory
            excel_file = os.path.join(self.slides_path, "مخزون الشرائح.xlsx")
            
            # Initialize Excel file if it doesn't exist
            if not os.path.exists(excel_file):
                print(f"[DEBUG] Initializing new slides publishing Excel file")
                from utils.slides_utils import initialize_slides_publishing_excel
                initialize_slides_publishing_excel(excel_file)
            
            # Add publishing entries using utility function
            from utils.slides_utils import add_slides_publishing_entry
            add_slides_publishing_entry(excel_file, data)
            
            # Reset rows after save
            for row in self.rows:
                row._set_default_values()
            
            self._show_success_dialog(excel_file)
            
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ أثناء حفظ الملف:\n{str(e)}", ft.Colors.RED_400)
            import traceback
            traceback.print_exc()

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


def main():
    def on_back():
        print("Back button clicked")
    
    ft.app(target=lambda page: SlidesAddView(page, on_back).build_ui())


if __name__ == "__main__":
    main()