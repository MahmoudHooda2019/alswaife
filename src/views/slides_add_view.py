import flet as ft
import json
import os
from datetime import datetime
from utils.utils import resource_path, is_excel_running
from utils.slides_utils import initialize_slides_inventory_excel, add_slides_inventory_entry, convert_existing_slides_inventory_to_formulas
from utils.log_utils import log_error, log_exception


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
        # If read_only is True, use BLACK45 background to distinguish it (same as blocks view)
        is_read_only = kwargs.get('read_only', False)
        default_bgcolor = ft.Colors.BLACK45 if is_read_only else ft.Colors.BLUE_GREY_900
        bgcolor = kwargs.pop("bgcolor", default_bgcolor)
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
            **kwargs,
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
            icon=ft.Icons.CALENDAR_TODAY
        )

        # Block number
        self.block_number = self._create_styled_textfield(
            "رقم البلوك",
            width_small,
            on_change=self.on_block_change,
            on_blur=self.on_block_blur,
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

        # Entry time - with auto-formatting
        self.entry_time = self._create_styled_textfield(
            "وقت الدخول",
            220,
            icon=ft.Icons.ACCESS_TIME,
            hint_text="الساعة:الدقيقة ص/م يوم/شهر/سنة",
            on_change=self._on_entry_time_change,
        )

        # Exit time - with auto-formatting
        self.exit_time = self._create_styled_textfield(
            "وقت الخروج",
            220,
            icon=ft.Icons.ACCESS_TIME,
            hint_text="الساعة:الدقيقة ص/م يوم/شهر/سنة",
            on_change=self._on_exit_time_change,
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

        # Fold/Expand button with rotate animation
        self.is_expanded = True
        self.fold_icon = ft.Icon(
            ft.Icons.KEYBOARD_ARROW_DOWN,
            color=ft.Colors.BLUE_300,
            size=24,
            rotate=ft.Rotate(angle=3.14159),  # 180 degrees (expanded state)
            animate_rotation=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT),
        )
        
        self.fold_btn = ft.IconButton(
            content=self.fold_icon,
            tooltip="طي/فتح",
            on_click=self._toggle_fold,
            bgcolor=ft.Colors.GREY_800,
            style=ft.ButtonStyle(
                shape=ft.RoundedRectangleBorder(radius=10)
            ),
        )

        # Header summary text (shown when collapsed)
        self.summary_text = ft.Text(
            "",
            size=14,
            color=ft.Colors.GREY_400,
            weight=ft.FontWeight.W_500,
            rtl=True,
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

        # Collapsible content container
        self.content_container = ft.Container(
            content=ft.Column(
                controls=[
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
            animate=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT),
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
        )

        # Create the main card with gradient background
        self.card = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # Header with summary and action buttons
                        ft.Row(
                            controls=[
                                self.summary_text,
                                ft.Container(expand=True),
                                self.fold_btn,
                                self.delete_btn,
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        ),
                        
                        # Collapsible content
                        self.content_container,
                    ],
                    spacing=0
                ),
                padding=20,
                gradient=ft.LinearGradient(
                    begin=ft.alignment.top_left,
                    end=ft.alignment.bottom_right,
                    colors=[ft.Colors.GREY_900, ft.Colors.GREY_800]
                ),
                border_radius=15,
                border=ft.border.all(1, ft.Colors.GREY_700),
                animate=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT),
            ),
            elevation=8,
        )

        self.row = self.card

    def _toggle_fold(self, e=None):
        """Toggle the fold/expand state with animation"""
        self.is_expanded = not self.is_expanded
        
        if self.is_expanded:
            # Expand
            self.content_container.height = None
            self.content_container.opacity = 1
            self.fold_icon.rotate.angle = 3.14159  # 180 degrees (arrow up)
            self.summary_text.value = ""
        else:
            # Collapse
            self.content_container.height = 0
            self.content_container.opacity = 0
            self.fold_icon.rotate.angle = 0  # 0 degrees (arrow down)
            # Show summary
            block_num = self.block_number.value or "---"
            material = self.material_dropdown.value or "---"
            area = self.area_field.value or "0.00"
            length = self.length_field.value or "0"
            height = self.height_field.value or "0"
            ltr = "\u200E"
            self.summary_text.value = f"{ltr}{block_num} # {material} # {ltr}{length} x {ltr}{height} = {ltr}{area}"
        
        self.page.update()

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

    def _format_datetime_auto(self, value, current_period=None):
        """
        Auto-format datetime input as user types
        Format: HH:MM ص/م DD/MM/YYYY
        لا يضيف ص/م تلقائياً - ينتظر المستخدم
        """
        if not value:
            return "", None
        
        # Extract period if exists (ص or م)
        period = None
        if 'ص' in value:
            period = 'ص'
        elif 'م' in value:
            period = 'م'
        elif current_period:
            period = current_period
        
        # Remove all non-digit characters
        digits_only = ''.join(c for c in value if c.isdigit())
        
        # If no digits, return empty
        if not digits_only:
            return "", period
        
        # Build formatted string based on number of digits
        result = ""
        
        # Hour (positions 0-1)
        if len(digits_only) >= 1:
            h1 = digits_only[0]
            # Allow 0 or 1 as first digit for hours (00-12)
            if int(h1) > 1:
                # If first digit > 1, treat as single digit hour (e.g., 8 -> 08)
                result = f"0{h1}:"
                # Shift remaining digits
                digits_only = "0" + digits_only
            else:
                result = h1
        
        if len(digits_only) >= 2:
            hour = int(digits_only[0:2])
            if hour > 12:
                hour = 12
            result = f"{hour:02d}:"
        
        # Minutes (positions 2-3)
        if len(digits_only) >= 3:
            m1 = digits_only[2]
            if int(m1) > 5:
                m1 = "5"
            result += m1
        
        if len(digits_only) >= 4:
            minute = int(digits_only[2:4])
            if minute > 59:
                minute = 59
            result = result[:-1] + f"{minute:02d}"
            # Add period if exists, otherwise add space for user to type
            if period:
                result += f" {period} "
            else:
                result += " "  # Space waiting for ص or م
        
        # Only continue with date if period is set
        if period and len(digits_only) >= 4:
            # Day (positions 4-5)
            if len(digits_only) >= 5:
                d1 = digits_only[4]
                if int(d1) > 3:
                    d1 = "3"
                result += d1
            
            if len(digits_only) >= 6:
                day = int(digits_only[4:6])
                if day > 31:
                    day = 31
                if day == 0:
                    day = 1
                result = result[:-1] + f"{day:02d}/"
            
            # Month (positions 6-7)
            if len(digits_only) >= 7:
                m1 = digits_only[6]
                if int(m1) > 1:
                    m1 = "1"
                result += m1
            
            if len(digits_only) >= 8:
                month = int(digits_only[6:8])
                if month > 12:
                    month = 12
                if month == 0:
                    month = 1
                result = result[:-1] + f"{month:02d}/"
            
            # Year (positions 8-11)
            if len(digits_only) >= 9:
                year_digits = digits_only[8:12]
                result += year_digits
        
        return result, period

    def _on_entry_time_change(self, e=None):
        """Handle entry time field changes with auto-formatting"""
        if e and e.control:
            current = e.control.value or ""
            original = current  # Store original for comparison
            
            # Convert English letters to Arabic period markers
            # w → ص (morning), l → م (evening)
            current = current.replace('w', 'ص').replace('W', 'ص')
            current = current.replace('l', 'م').replace('L', 'م')
            
            # Check if user is deleting
            digits_only = ''.join(c for c in current if c.isdigit())
            has_period = 'ص' in current or 'م' in current
            
            # Allow deletion - if field is empty or only has spaces
            if not digits_only and not has_period:
                e.control.value = ""
                if hasattr(self, '_entry_period'):
                    delattr(self, '_entry_period')
                self._calculate_hours()
                self.page.update()
                return
            
            # Detect if user deleted the period marker (ص or م)
            # If previous value had period but current doesn't, clear the stored period
            if hasattr(self, '_entry_last_value'):
                last_had_period = 'ص' in self._entry_last_value or 'م' in self._entry_last_value
                if last_had_period and not has_period:
                    # User deleted the period marker
                    if hasattr(self, '_entry_period'):
                        delattr(self, '_entry_period')
            
            # Get current period from the field
            current_period = None
            if hasattr(self, '_entry_period'):
                current_period = self._entry_period
            
            formatted, period = self._format_datetime_auto(current, current_period)
            
            # Store the period
            if period:
                self._entry_period = period
            elif not period and has_period:
                # Keep the existing period if formatting didn't extract one
                pass
            else:
                # Clear period if none found
                if hasattr(self, '_entry_period'):
                    delattr(self, '_entry_period')
            
            # Only update if value actually changed
            if formatted != original:
                e.control.value = formatted
            
            # Store current value for next comparison
            self._entry_last_value = formatted
        
        self._calculate_hours()
        self.page.update()

    def _on_exit_time_change(self, e=None):
        """Handle exit time field changes with auto-formatting"""
        if e and e.control:
            current = e.control.value or ""
            original = current  # Store original for comparison
            
            # Convert English letters to Arabic period markers
            # w → ص (morning), l → م (evening)
            current = current.replace('w', 'ص').replace('W', 'ص')
            current = current.replace('l', 'م').replace('L', 'م')
            
            # Check if user is deleting
            digits_only = ''.join(c for c in current if c.isdigit())
            has_period = 'ص' in current or 'م' in current
            
            # Allow deletion - if field is empty or only has spaces
            if not digits_only and not has_period:
                e.control.value = ""
                if hasattr(self, '_exit_period'):
                    delattr(self, '_exit_period')
                self._calculate_hours()
                self.page.update()
                return
            
            # Detect if user deleted the period marker (ص or م)
            # If previous value had period but current doesn't, clear the stored period
            if hasattr(self, '_exit_last_value'):
                last_had_period = 'ص' in self._exit_last_value or 'م' in self._exit_last_value
                if last_had_period and not has_period:
                    # User deleted the period marker
                    if hasattr(self, '_exit_period'):
                        delattr(self, '_exit_period')
            
            # Get current period from the field
            current_period = None
            if hasattr(self, '_exit_period'):
                current_period = self._exit_period
            
            formatted, period = self._format_datetime_auto(current, current_period)
            
            # Store the period
            if period:
                self._exit_period = period
            elif not period and has_period:
                # Keep the existing period if formatting didn't extract one
                pass
            else:
                # Clear period if none found
                if hasattr(self, '_exit_period'):
                    delattr(self, '_exit_period')
            
            # Only update if value actually changed
            if formatted != original:
                e.control.value = formatted
            
            # Store current value for next comparison
            self._exit_last_value = formatted
        
        self._calculate_hours()
        self.page.update()

    def _parse_datetime(self, value):
        """Parse datetime string in format 'HH:MM ص/م DD/MM/YYYY' and return datetime object"""
        if not value:
            return None
        
        try:
            # Expected format: "08:30 ص 28/12/2025"
            parts = value.strip().split(' ')
            if len(parts) < 3:
                return None
            
            time_part = parts[0]  # "08:30"
            period = parts[1]     # "ص" or "م"
            date_part = parts[2]  # "28/12/2025"
            
            # Parse time
            time_parts = time_part.split(':')
            if len(time_parts) != 2:
                return None
            
            hour = int(time_parts[0])
            minute = int(time_parts[1])
            
            # Convert to 24-hour format
            if period == 'م' and hour != 12:
                hour += 12
            elif period == 'ص' and hour == 12:
                hour = 0
            
            # Parse date
            date_parts = date_part.split('/')
            if len(date_parts) != 3:
                return None
            
            day = int(date_parts[0])
            month = int(date_parts[1])
            year = int(date_parts[2])
            
            # Validate year has 4 digits
            if year < 1000:
                return None
            
            return datetime(year, month, day, hour, minute)
        except (ValueError, IndexError):
            return None

    def _calculate_hours(self):
        """Calculate hours between entry and exit datetime including date difference"""
        try:
            entry_dt = self._parse_datetime(self.entry_time.value)
            exit_dt = self._parse_datetime(self.exit_time.value)
            
            if entry_dt and exit_dt:
                # Calculate difference including date
                diff = exit_dt - entry_dt
                
                # Get total seconds and convert to hours
                total_seconds = diff.total_seconds()
                
                # Handle negative difference (exit before entry)
                if total_seconds < 0:
                    self.hours_count.value = "0.00"
                    return
                
                # Convert to hours (including days difference)
                hours = total_seconds / 3600.0
                self.hours_count.value = f"{hours:.2f}"
                return
        except (ValueError, IndexError, TypeError):
            pass
        
        self.hours_count.value = "0.00"
    
    def on_block_change(self, e):
        """Handle block number changes and replace Arabic characters with English equivalents"""
        val = self.block_number.value
        if val:
            from utils.utils import normalize_block_number
            new_val = normalize_block_number(val, reorder=False)  # Only normalize, don't reorder on change
            
            if new_val != val:
                self.block_number.value = new_val
                # Only update if value changed
                if hasattr(self, 'page') and self.page:
                    self.page.update()
        # Always trigger calculations after any change
        self._calculate_values()

    def on_block_blur(self, e):
        """Normalize and reorder block number when focus leaves"""
        val = self.block_number.value
        if val:
            from utils.utils import normalize_block_number
            new_val = normalize_block_number(val, reorder=True)  # Full normalization with reordering
            
            if new_val != val:
                self.block_number.value = new_val
                if hasattr(self, 'page') and self.page:
                    self.page.update()

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
            log_error(f"Error getting price from JSON: {e}")
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

    def get_editable_fields(self):
        """Return list of editable fields in order for navigation"""
        return [
            self.publishing_date,      # 0
            self.block_number,         # 1
            self.material_dropdown,    # 2
            self.machine_number,       # 3
            self.quantity_field,       # 4
            self.length_field,         # 5
            self.height_field,         # 6
            self.thickness_dropdown,   # 7
            self.price_per_meter_field,  # 8
            self.entry_time,           # 9
            self.exit_time,            # 10
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
        
        # Navigation tracking
        self._current_row_idx = 0
        self._current_field_idx = 0

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
        
        # Add keyboard event handler for arrow navigation
        self.page.on_keyboard_event = self.on_keyboard_event
        
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
    
    def on_keyboard_event(self, e: ft.KeyboardEvent):
        """Handle keyboard events for arrow navigation"""
        # Check if the '+' key was pressed to add a new row
        if e.key == '+' or e.key == '=':
            if not e.ctrl and not e.shift and not e.alt:
                self.add_row()
                return
        
        # Arrow key navigation
        if e.key in ["Arrow Down", "Arrow Up", "Arrow Left", "Arrow Right"]:
            self._handle_arrow_navigation(e.key)
    
    def _skip_dropdown(self, row_idx, field_idx, direction=1):
        """Skip dropdown fields in the given direction (1=forward, -1=backward)"""
        if row_idx < 0 or row_idx >= len(self.rows):
            return row_idx, field_idx
        
        fields = self.rows[row_idx].get_editable_fields()
        
        # Check current field
        while 0 <= field_idx < len(fields):
            if not isinstance(fields[field_idx], ft.Dropdown):
                return row_idx, field_idx
            field_idx += direction
        
        # If we've exhausted current row, return boundary
        if direction > 0:
            return row_idx, len(fields) - 1
        else:
            return row_idx, 0
    
    def _handle_arrow_navigation(self, key):
        """Handle arrow key navigation between fields (skipping dropdowns)"""
        if not self.rows:
            return
        
        # Ensure indices are valid
        if self._current_row_idx < 0 or self._current_row_idx >= len(self.rows):
            self._current_row_idx = 0
        
        current_row = self.rows[self._current_row_idx]
        fields = current_row.get_editable_fields()
        
        if self._current_field_idx < 0 or self._current_field_idx >= len(fields):
            self._current_field_idx = 0
            # Skip dropdown if starting field is dropdown
            self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, 1)
        
        if key == "Arrow Right":
            # Move to previous field (RTL)
            self._current_field_idx -= 1
            if self._current_field_idx < 0:
                # Move to previous row, last non-dropdown field
                if self._current_row_idx > 0:
                    self._current_row_idx -= 1
                    self._current_field_idx = len(self.rows[self._current_row_idx].get_editable_fields()) - 1
                    self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, -1)
                else:
                    self._current_field_idx = 0
                    self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, 1)
            else:
                # Skip dropdown if needed
                self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, -1)
        
        elif key == "Arrow Left":
            # Move to next field (RTL)
            self._current_field_idx += 1
            if self._current_field_idx >= len(fields):
                # Move to next row, first non-dropdown field
                if self._current_row_idx < len(self.rows) - 1:
                    self._current_row_idx += 1
                    self._current_field_idx = 0
                    self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, 1)
                else:
                    # Stay at last non-dropdown field
                    self._current_field_idx = len(fields) - 1
                    self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, -1)
            else:
                # Skip dropdown if needed
                self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, 1)
        
        elif key == "Arrow Down":
            # Move to next non-dropdown field in same row
            self._current_field_idx += 1
            if self._current_field_idx >= len(fields):
                # Wrap to first non-dropdown field in same row
                self._current_field_idx = 0
            self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, 1)
        
        elif key == "Arrow Up":
            # Move to previous non-dropdown field in same row
            self._current_field_idx -= 1
            if self._current_field_idx < 0:
                # Wrap to last non-dropdown field in same row
                self._current_field_idx = len(fields) - 1
            self._current_row_idx, self._current_field_idx = self._skip_dropdown(self._current_row_idx, self._current_field_idx, -1)
        
        # Focus the target field
        if 0 <= self._current_row_idx < len(self.rows):
            self.rows[self._current_row_idx].focus_field(self._current_field_idx)

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
        """Save data to Excel file"""
        if not any(self._row_has_data(row) for row in self.rows):
            self._show_dialog("تحذير", "لا توجد بيانات لحفظها", ft.Colors.ORANGE_400)
            return

        # التحقق من وجود رقم البلوك في كل صف
        for i, row in enumerate(self.rows):
            if self._row_has_data(row):
                block_num = row.block_number.value
                if not block_num or not block_num.strip():
                    self._show_dialog(
                        "خطأ",
                        f"الصف {i + 1}: رقم البلوك مطلوب",
                        ft.Colors.RED_400,
                    )
                    return

        # التحقق من أن Excel مغلق
        if is_excel_running():
            self._show_excel_warning_dialog()
            return

        # جمع البيانات
        data = [row.to_dict() for row in self.rows if self._row_has_data(row)]
        self._do_save(data)

    def _show_excel_warning_dialog(self):
        """Show Excel warning dialog with continue option"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()
            if dlg in self.page.overlay:
                self.page.overlay.remove(dlg)
            self.page.update()

        def continue_save(e=None):
            dlg.open = False
            self.page.update()
            data = [row.to_dict() for row in self.rows if self._row_has_data(row)]
            self._do_save(data)

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

    def _do_save(self, data):
        """تنفيذ عملية الحفظ الفعلية"""
        try:
            # Create Excel file for slides inventory
            excel_file = os.path.join(self.slides_path, "مخزون الشرائح.xlsx")
            
            # Initialize Excel file if it doesn't exist
            if not os.path.exists(excel_file):
                from utils.slides_utils import initialize_slides_inventory_excel
                initialize_slides_inventory_excel(excel_file)
            
            # Add inventory entries using utility function
            from utils.slides_utils import add_slides_inventory_from_publishing
            entries_added, blocks_warnings = add_slides_inventory_from_publishing(excel_file, data)
            
            # Reset rows after save
            for row in self.rows:
                row._set_default_values()
            
            self._show_success_dialog(excel_file, blocks_warnings)
            
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ أثناء حفظ الملف:\n{str(e)}", ft.Colors.RED_400)
            import traceback
            traceback.print_exc()

    def _show_dialog(self, title: str, message: str, title_color=ft.Colors.BLUE_300):
        """Show a styled dialog"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()
            if dlg in self.page.overlay:
                self.page.overlay.remove(dlg)
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

    def _show_success_dialog(self, filepath: str, blocks_warnings: list = None):
        """Show success bottom sheet with optional warnings and file open options"""
        from utils.bottom_sheet_utils import BottomSheetManager
        import os
        
        # Build message
        message = "تم إنشاء الملف بنجاح"
        
        # Add warnings if any
        if blocks_warnings and len(blocks_warnings) > 0:
            message += "\n\n⚠️ تحذير: لم يتم حفظ بعض الشرائح في مخزون البلوكات:\n"
            for warning in blocks_warnings:
                message += f"• {warning}\n"
        
        # Get blocks file path
        blocks_file = os.path.join(os.path.expanduser("~"), "Documents", "alswaife", "البلوكات", "مخزون البلوكات.xlsx")
        
        # Define open file callbacks
        def open_slides_file(e):
            try:
                os.startfile(filepath)
            except Exception:
                pass
        
        def open_blocks_file(e):
            try:
                if os.path.exists(blocks_file):
                    os.startfile(blocks_file)
            except Exception:
                pass
        
        def open_folder(e):
            try:
                folder = os.path.dirname(filepath)
                os.startfile(folder)
            except Exception:
                pass
        
        # Build options list
        options = [
            {
                "text": "فتح ملف الشرائح",
                "icon": ft.Icons.DESCRIPTION,
                "color": ft.Colors.GREEN_700,
                "on_click": open_slides_file,
            },
            {
                "text": "فتح ملف البلوكات",
                "icon": ft.Icons.INVENTORY_2,
                "color": ft.Colors.BLUE_700,
                "on_click": open_blocks_file,
            },
            {
                "text": "فتح المجلد",
                "icon": ft.Icons.FOLDER_OPEN,
                "color": ft.Colors.ORANGE_700,
                "on_click": open_folder,
            },
        ]
        
        # Show options bottom sheet
        BottomSheetManager.show_options_bottom_sheet(
            page=self.page,
            title="تم الحفظ بنجاح",
            description=message,
            icon=ft.Icons.CHECK_CIRCLE,
            icon_color=ft.Colors.GREEN_400,
            options=options,
        )


def main():
    def on_back():
        pass
    
    ft.app(target=lambda page: SlidesAddView(page, on_back).build_ui())


if __name__ == "__main__":
    main()