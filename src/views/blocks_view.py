import os
import flet as ft

from utils.blocks_utils import export_simple_blocks_excel
from utils.log_utils import log_exception
from utils.utils import is_excel_running, get_current_date, is_file_locked
from utils.dialog_utils import DialogManager
from utils.bottom_sheet_utils import BottomSheetManager

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
        self.is_expanded = True  # Track fold state
        self._build_controls()
        self._set_default_values()

    def _create_styled_textfield(self, label, width, **kwargs):
        """Create a consistently styled text field"""
        # Extract bgcolor if provided in kwargs, otherwise use default
        # Determine default background color based on read_only state
        is_read_only = kwargs.get('read_only', False)
        default_bgcolor = ft.Colors.BLACK45 if is_read_only else ft.Colors.BLUE_GREY_900
        
        bgcolor = kwargs.pop('bgcolor', default_bgcolor)
        
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
        # Filter that allows Arabic decimal separators (ز and ،) to be converted later
        numeric_filter_with_arabic = ft.InputFilter(regex_string=r"^[0-9ز،]*\.?[0-9ز،]*$")
        
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
            value=get_current_date("%Y-%m-%d"),
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
        
        # Fadl checkbox - styled to match input fields with fixed width
        self.fadl_checkbox = ft.Container(
            content=ft.Checkbox(
                label="فضل",
                value=False,
                on_change=self._on_fadl_change,
                fill_color={
                    ft.ControlState.DEFAULT: ft.Colors.BLUE_GREY_800,
                    ft.ControlState.SELECTED: ft.Colors.AMBER_700,
                },
                check_color=ft.Colors.WHITE,
                label_style=ft.TextStyle(
                    size=14,
                    weight=ft.FontWeight.W_500,
                    color=ft.Colors.WHITE
                ),
            ),
            bgcolor=ft.Colors.BLUE_GREY_900,
            border_radius=10,
            padding=ft.padding.symmetric(horizontal=12, vertical=10),
            border=ft.border.all(1, ft.Colors.GREY_700),
            width=100,  # Fixed width to prevent expansion
        )

        
        # Block type options for نيو حلايب
        self.BLOCK_TYPE_OPTIONS = [
            ft.dropdown.Option("A"),
            ft.dropdown.Option("B"),
            ft.dropdown.Option("F"),
            ft.dropdown.Option("K")
        ]
        
        # Block type dropdown (for نيو حلايب - user selects A or B)
        self.block_type_dropdown = self._create_styled_dropdown(
            "نوع البلوك",
            width_small,
            self.BLOCK_TYPE_OPTIONS,
            on_change=self._on_block_type_change,
            icon=ft.Icons.CATEGORY
        )
        
        # Block type text field (read-only, for جندولا and احمر اسوان)
        self.block_type_text = self._create_styled_textfield(
            "نوع البلوك",
            width_small,
            read_only=True,
            icon=ft.Icons.CATEGORY
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
            input_filter=numeric_filter_with_arabic,
            on_change=self._on_field_change,
            suffix_text="م"
        )
        
        # Width
        self.width_field = self._create_styled_textfield(
            "العرض",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter_with_arabic,
            on_change=self._on_field_change,
            suffix_text="م"
        )
        
        # Height
        self.height_field = self._create_styled_textfield(
            "الارتفاع",
            width_small,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter_with_arabic,
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
        
        # Price per ton - editable, based on material
        self.price_per_ton_field = self._create_styled_textfield(
            "سعر الطن",
            width_medium,
            icon=ft.Icons.ATTACH_MONEY,
            value="0.00",
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=numeric_filter,
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
        
        # Container for block type that switches between dropdown and text field
        self.block_type_container = ft.Container(
            content=self.block_type_dropdown,  # Default to dropdown for نيو حلايب
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
                            self.block_type_container,
                            self.material_dropdown,
                            self.fadl_checkbox,
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
            animate=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT),
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
        )
        
        # Header summary text (shown when collapsed)
        self.summary_text = ft.Text(
            "",
            size=14,
            color=ft.Colors.GREY_400,
            weight=ft.FontWeight.W_500,
            rtl=True,
            no_wrap=False,
            overflow=ft.TextOverflow.VISIBLE,
            max_lines=1,
            text_align=ft.TextAlign.RIGHT
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
            # Show summary - رقم البلوك # نوع البلوك # الخامة # الطول X العرض X الارتفاع # م3
            block_num = self.block_number.value or "---"
            block_type = self.get_block_type() or "---"
            material = self.material_dropdown.value or "---"
            length = self.length_field.value or "0"
            width = self.width_field.value or "0"
            height = self.height_field.value or "0"
            volume = self.volume_field.value or "0.00"
            # Use LTR mark (\u200E) to fix number display direction
            ltr = "\u200E"
            self.summary_text.value = f"{ltr}{block_num} # {block_type} # {material} # {ltr}{length} x {ltr}{width} x {ltr}{height} = {ltr}{volume}"
        
        self.page.update()

    def on_block_change(self, e):
        """Handle block number changes and replace Arabic characters with English equivalents"""
        val = self.block_number.value
        if val:
            from utils.utils import normalize_block_number
            new_val = normalize_block_number(val, reorder=True)  # Full normalization with reordering
            if new_val != val:
                self.block_number.value = new_val
                # Only update if value changed
                if hasattr(self, 'page') and self.page:
                    self.page.update()
        # Always trigger calculations after any change
        self._calculate_values()

    def _on_field_change(self, e=None):
        """Handle field changes and trigger calculations"""
        # Handle Arabic decimal separator for dimension fields
        if e and hasattr(e, 'control'):
            self._handle_arabic_decimal_input(e.control)
        self._calculate_values()
    
    def _handle_arabic_decimal_input(self, text_field):
        """Handle Arabic decimal separator (Zein letter) and replace with decimal point"""
        from utils.utils import normalize_numeric_input
        if not text_field or not text_field.value:
            return False
        original_value = text_field.value
        new_value = normalize_numeric_input(original_value)
        if new_value != original_value:
            text_field.value = new_value
            return True
        return False
    
    def _on_material_change(self, e=None):
        """Handle material change and update weight per m3, price per ton, and block type"""
        material = self.material_dropdown.value
        
        if material == "نيو حلايب":
            self.weight_per_m3_field.value = "2.70"
            # Show dropdown for A/B selection
            self.block_type_container.content = self.block_type_dropdown
            if not self.block_type_dropdown.value:
                self.block_type_dropdown.value = "A"
            
            # Set price based on A/B
            if self.block_type_dropdown.value == "A":
                self.price_per_ton_field.value = "1450"
            else:
                self.price_per_ton_field.value = "1125"
                
        elif material == "جندولا":
            self.weight_per_m3_field.value = "2.85"
            self.price_per_ton_field.value = "1600"
            # Show text field with "جندولا"
            self.block_type_text.value = "جندولا"
            self.block_type_container.content = self.block_type_text
        elif material == "احمر اسوان":
            self.weight_per_m3_field.value = "2.80"
            self.price_per_ton_field.value = "1500"
            # Show text field with "احمر"
            self.block_type_text.value = "احمر"
            self.block_type_container.content = self.block_type_text
            
        self._calculate_values()
        self.page.update()

    def _on_block_type_change(self, e=None):
        """Handle block type change (A/B) for نيو حلايب"""
        material = self.material_dropdown.value
        if material == "نيو حلايب":
            if self.block_type_dropdown.value == "A":
                self.price_per_ton_field.value = "1450"
            else:
                self.price_per_ton_field.value = "1125"
        
        self._calculate_values()
        self.page.update()
    
    def _on_fadl_change(self, e=None):
        """Handle fadl checkbox change"""
        # Just update the page, the actual logic is in get_block_type() and to_dict()
        self.page.update()
    
    def get_block_type(self):
        """Get the current block type value"""
        material = self.material_dropdown.value
        # Access checkbox value from container
        is_fadl = self.fadl_checkbox.content.value if hasattr(self.fadl_checkbox, 'content') else False
        
        if material == "نيو حلايب":
            block_type = self.block_type_dropdown.value or "A"
            if is_fadl:
                return f"{block_type} - فضل"
            return block_type
        elif material == "جندولا":
            if is_fadl:
                return "جندولا - فضل"
            return "جندولا"
        elif material == "احمر اسوان":
            if is_fadl:
                return "احمر - فضل"
            return "احمر"
        return ""
    
    def is_fadl_block(self):
        """Check if this block is marked as fadl"""
        if hasattr(self, 'fadl_checkbox'):
            # Access checkbox value from container
            if hasattr(self.fadl_checkbox, 'content'):
                return self.fadl_checkbox.content.value
        return False
    
    def get_editable_fields(self):
        """Return list of editable fields in order for navigation"""
        return [
            self.trip_number,      # 0
            self.trip_count,       # 1
            self.date_field,       # 2
            self.quarry,           # 3
            self.block_number,     # 4
            self.block_type_dropdown if self.material_dropdown.value == "نيو حلايب" else self.block_type_text,  # 5
            self.material_dropdown,  # 6
            self.length_field,     # 7
            self.width_field,      # 8
            self.height_field,     # 9
            self.weight_per_m3_field,  # 10
            self.price_per_ton_field,  # 11
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
            # Format: show decimals only if needed, with thousands separator
            if total_price == int(total_price):
                self.total_price_field.value = f"{int(total_price):,}"
            else:
                self.total_price_field.value = f"{total_price:,.2f}"
            
        except ValueError:
            # If any field contains non-numeric values, set calculated fields to 0
            self.volume_field.value = "0.00"
            self.block_weight_field.value = "0.00"
            self.total_price_field.value = "0"
        
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
            'block_type': self.get_block_type(),
            'material': self.material_dropdown.value,
            'length': self.length_field.value,
            'width': self.width_field.value,
            'height': self.height_field.value,
            'volume': self.volume_field.value,  # Calculated: length * width * height
            'weight_per_m3': self.weight_per_m3_field.value,  # Based on material
            'block_weight': self.block_weight_field.value,  # Calculated: volume * weight_per_m3
            'price_per_ton': self.price_per_ton_field.value,  # Based on material
            'total_price': self.total_price_field.value,  # Calculated: price_per_ton * block_weight
            'is_fadl': self.is_fadl_block(),  # Whether this is a fadl block (single row instead of merged)
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
        
        # Navigation tracking
        self._current_row_idx = 0
        self._current_field_idx = 0

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
                    icon=ft.Icons.REFRESH,
                    on_click=self.reset_all,
                    tooltip="مسح الكل"
                ),
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

    def reset_all(self, e=None):
        """Reset all rows - clear all data"""
        self.rows.clear()
        self.rows_container.controls.clear()
        self.add_row()
        self.page.update()

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

        self._do_save()

    def _do_save(self):
        """تنفيذ عملية الحفظ الفعلية"""
        # التحقق من أن الملف غير مفتوح
        blocks_file = os.path.join(
            os.path.expanduser("~"), "Documents", "alswaife", "البلوكات", "مخزون البلوكات.xlsx"
        )
        if is_file_locked(blocks_file):
            self._show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
            return
        
        try:
            data = [row.to_dict() for row in self.rows if self._row_has_data(row)]
            file_path = export_simple_blocks_excel(data)
            
            # Reset rows after save
            for row in self.rows:
                row._set_default_values()
            
            self._show_success_dialog(file_path)
            
        except ValueError as e:
            # Duplicate block numbers found
            self._show_dialog("خطأ - بلوكات مكررة", str(e), ft.Colors.RED_400)
        except PermissionError:
            self._show_dialog("خطأ", "الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى.", ft.Colors.RED_400)
        except Exception as e:
            self._show_dialog("خطأ", f"حدث خطأ أثناء حفظ الملف:\n{str(e)}", ft.Colors.RED_400)
            log_exception(f"Error saving blocks file: {e}")

    def _show_excel_warning_dialog(self):
        """Show Excel warning dialog with continue option"""
        def close_dlg(e=None):
            DialogManager.close_dialog(self.page, dlg)

        def continue_save(e=None):
            DialogManager.close_dialog(self.page, dlg)
            self._do_save()

        actions = [
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
        ]

        dlg = DialogManager.show_custom_dialog(
            self.page,
            "تحذير",
            ft.Text("برنامج Excel مفتوح حالياً.\nيرجى إغلاقه قبل الحفظ.", size=16, rtl=True),
            actions,
            icon=ft.Icons.WARNING_AMBER_ROUNDED,
            icon_color=ft.Colors.ORANGE_400,
            title_color=ft.Colors.ORANGE_400
        )

    def _show_dialog(self, title: str, message: str, title_color=ft.Colors.BLUE_300):
        """Show a styled dialog"""
        if title == "خطأ" or title_color == ft.Colors.RED_400:
            DialogManager.show_error_dialog(self.page, message, title=title)
        elif title == "تحذير" or title_color == ft.Colors.ORANGE_400:
            DialogManager.show_warning_dialog(self.page, message, title=title)
        else:
            DialogManager.show_info_dialog(self.page, message, title=title)

    def _show_success_dialog(self, filepath: str):
        """Show success bottom sheet with file actions"""
        BottomSheetManager.show_success_bottom_sheet(
            page=self.page,
            message="تم إضافة البلوكات إلى السجل بنجاح",
            filepath=filepath,
            title="تم الحفظ بنجاح",
        )
