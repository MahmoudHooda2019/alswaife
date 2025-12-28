import sys
import os
import flet as ft
import json
import re
import sqlite3
from datetime import datetime
import subprocess
import platform
import urllib.request
from pathlib import Path

# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import database utilities
try:
    from utils.db_utils import init_db as init_db_real, get_counter as get_counter_real, increment_counter as increment_counter_real, get_zoom_level, set_zoom_level, save_invoice_to_db, load_invoice_from_db, invoice_exists
    
    # Re-export with proper type annotations
    init_db = init_db_real
    get_counter = get_counter_real
    increment_counter = increment_counter_real
except ImportError:
    def init_db(db_path: str) -> None: pass
    def get_counter(db_path: str, key: str = "invoice") -> int: return 1
    def increment_counter(db_path: str, key: str = "invoice") -> int: return 1
    def get_zoom_level(db_path: str) -> float: return 1.0
    def set_zoom_level(db_path: str, zoom_level: float) -> None: pass

from utils.invoice_utils import update_client_ledger, remove_invoice_from_ledger, delete_existing_invoice_file, update_invoice_in_ledger, update_payment_in_invoice, update_payment_in_ledger, get_payment_from_invoice
from utils.path_utils import resource_path
from utils.slides_utils import disburse_slides_from_invoice
from utils.purchases_utils import add_income_record
from utils.log_utils import log_error, log_exception


def is_excel_running():
    """Check if Microsoft Excel is currently running"""
    try:
        if platform.system() == 'Windows':
            # Use tasklist to check for Excel process
            result = subprocess.run(
                ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/NH'],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return 'EXCEL.EXE' in result.stdout.upper()
        elif platform.system() == 'Darwin':  # macOS
            result = subprocess.run(
                ['pgrep', '-x', 'Microsoft Excel'],
                capture_output=True
            )
            return result.returncode == 0
        else:  # Linux
            result = subprocess.run(
                ['pgrep', '-f', 'libreoffice.*calc'],
                capture_output=True
            )
            return result.returncode == 0
    except Exception:
        return False

class InvoiceRow:
    """ كلاس صف الفاتورة (البند) """
    def __init__(self, page, row_index, product_dict, delete_callback, scale_factor=1.0):
        self.page = page
        self.products = product_dict
        self.delete_callback = delete_callback
        self.scale_factor = scale_factor
        self.row_container = None  # Reference to the UI container
        # For length calculation
        self.original_length = 0  # لحفظ الطول الأصلي
        
        # Internal material value (not displayed in UI)
        self.material_value = ""
        
        # المتغيرات
        # Default widths (minimum widths)
        self.default_widths = {
            'block': 60, # رقم البلوك 
            'thick': 105, # السمك 
            'count': 55, # العدد 
            'len_before': 60, # الطول قبل 
            'discount': 55, # الخصم
            'len': 60, # الطول 
            'height': 65, # الارتفاع
            'area': 68, # المسطح
            'price': 60, #السعر
            'total': 70, #الاجمالي
            'product': 137 # البيان
        }

        # المتغيرات
        self.block_var = ft.TextField(
            label="رقم البلوك", 
            width=self.default_widths['block'],
            on_change=self.on_block_change,
            on_blur=self.on_block_blur
        )
        self.thick_var = ft.Dropdown(
            label="السمك",
            options=[ft.dropdown.Option("2سم"), ft.dropdown.Option("3سم"), ft.dropdown.Option("4سم")],
            width=self.default_widths['thick']
        )
        self.count_var = ft.TextField(
            label="العدد", 
            width=self.default_widths['count'],
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*ز?$"),  # Allow numbers, decimal point, and 'ز' character
            on_change=self.calculate
        )
        # New field for length before discount
        self.len_before_var = ft.TextField(
            label="الطول قبل", 
            width=self.default_widths['len_before'],
            label_style=ft.TextStyle(size=10.0),
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*ز?$"),  # Allow numbers, decimal point, and 'ز' character
            on_change=self.on_len_before_change  # Use new handler
        )
        # New field for discount
        self.discount_var = ft.TextField(
            label="الخصم", 
            width=self.default_widths['discount'],
            label_style=ft.TextStyle(size=10.0),
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*ز?$"),  # Allow numbers, decimal point, and 'ز' character
            on_change=self.on_discount_change,  # Use new handler
            value="0.20"  # Set default value to 0.20
        )
        # Modified length field - now readonly
        self.len_var = ft.TextField(
            label="الطول", 
            width=self.default_widths['len'],
            disabled=True,  # Make it non-editable
            value="0"
        )
        self.height_var = ft.TextField(
            label="الارتفاع", 
            width=self.default_widths['height'],
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*ز?$"),
            on_change=self.calculate
        )
        
        self.area_var = ft.TextField(label="المسطح", width=self.default_widths['area'], disabled=True)
        self.price_var = ft.TextField(
            label="السعر", 
            width=self.default_widths['price'],
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*ز?$"),  # Allow numbers, decimal point, and 'ز' character
            on_change=self.calculate
        )
        self.total_var = ft.TextField(label="الإجمالي", width=self.default_widths['total'], disabled=True)
        
        # Product dropdown with prefixed options
        product_names = list(self.products.keys()) if self.products else []
        # Create options with "ش " prefix for display
        prefixed_options = [ft.dropdown.Option(name, "ش " + name) for name in product_names]
        self.product_dropdown = ft.Dropdown(
            label="البيان",
            options=prefixed_options,
            width=self.default_widths['product'],
            on_change=self.on_product_select
        )
        
        # Apply initial scale to ensure fields are displayed at the correct size
        self.update_scale(self.scale_factor, update_page=False)
        
        # Delete button
        self.btn_del = ft.IconButton(
            icon=ft.Icons.DELETE,
            icon_color="red",
            on_click=self.destroy
        )
        
        # Bind event for length changes
        self.len_before_var.on_change = self.on_len_before_change
        self.discount_var.on_change = self.on_discount_change
        # Bind event for thickness changes
        self.thick_var.on_change = self.on_thickness_change

    def on_block_change(self, e):
        val = self.block_var.value
        if val:
            # تحويل الحروف إلى A, B, F
            # A: ِ, ش, a
            # B: لآ, لا, b
            # F: f, [, ب
            new_val = val.replace('ِ', 'A').replace('ش', 'A').replace('a', 'A')
            new_val = new_val.replace('لآ', 'B').replace('لا', 'B').replace('b', 'B')
            new_val = new_val.replace('f', 'F').replace('[', 'F').replace('ب', 'F')
            new_val = new_val.upper()
            
            if new_val != val:
                self.block_var.value = new_val
                if hasattr(self, 'page') and self.page:
                    self.page.update()

    def on_block_blur(self, e):
        """Reorder block number when focus leaves - move letter to beginning if at end"""
        val = self.block_var.value
        if val:
            # تحويل الحروف إلى A, B, F
            new_val = val.replace('ِ', 'A').replace('ش', 'A').replace('a', 'A')
            new_val = new_val.replace('لآ', 'B').replace('لا', 'B').replace('b', 'B')
            new_val = new_val.replace('f', 'F').replace('[', 'F').replace('ب', 'F')
            new_val = new_val.upper()
            
            import re
            match = re.match(r'^(\d+)([A-Za-z]+)$', new_val)
            if match:
                numbers = match.group(1)
                letters = match.group(2).upper()
                new_val = letters + numbers
            
            if new_val != val:
                self.block_var.value = new_val
                if hasattr(self, 'page') and self.page:
                    self.page.update()

    def handle_arabic_decimal_input(self, text_field):
        """Handle Arabic decimal separator (Zein letter) and replace with decimal point"""
        if text_field.value:
            # Replace Arabic decimal separator (Zein letter 'زين') with decimal point
            # The Zein letter on Arabic keyboard typically maps to the comma key or 'z'
            new_value = text_field.value.replace('،', '.')  # Arabic comma/decimal separator
            new_value = new_value.replace('ز', '.')  # Arabic 'zein' letter often maps to 'z' key
            # Also handle potential Arabic digit inputs
            arabic_digits = {'٠': '0', '٢': '1', '٢': '2', '٣': '3', '٤': '4', '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'}
            for arabic_digit, english_digit in arabic_digits.items():
                new_value = new_value.replace(arabic_digit, english_digit)
            
            if new_value != text_field.value:
                text_field.value = new_value
                return True
        return False

    def on_len_before_change(self, e):
        """Handle input changes for length before field with Arabic decimal handling"""
        # Handle Arabic decimal separator
        changed = self.handle_arabic_decimal_input(self.len_before_var)
        # Recalculate length
        self.calculate_length()
        # Update UI if value changed
        if changed and hasattr(self, 'page') and self.page:
            self.page.update()

    def on_discount_change(self, e):
        """Handle input changes for discount field with Arabic decimal handling"""
        # Handle Arabic decimal separator
        changed = self.handle_arabic_decimal_input(self.discount_var)
        # Recalculate length
        self.calculate_length()
        # Update UI if value changed
        if changed and hasattr(self, 'page') and self.page:
            self.page.update()

    def on_product_select(self, e):
        """عند اختيار البيان، إرساله إلى خانة الخامة بدون حرف الشين"""
        selected_product = self.product_dropdown.value
        if selected_product:
            # Remove the "ش " prefix if present
            if selected_product.startswith("ش "):
                material_name = selected_product[2:]  # Remove "ش " prefix
                clean_product_name = material_name
            else:
                material_name = selected_product
                clean_product_name = selected_product
            
            # Update the material field
            self.material_value = material_name
            
            # Update the price based on product and thickness
            self.update_price(clean_product_name)
            
            # Update the page
            if hasattr(self, 'page') and self.page:
                self.page.update()

    def update_price(self, product_name):
        """Update price based on selected product and thickness"""
        try:
            thickness = self.thick_var.value
            length = float(self.len_var.value or 0)
            
            if product_name and thickness and self.products:
                # Clean the product name (remove "ش " prefix if present)
                clean_name = product_name.replace("ش ", "") if product_name.startswith("ش ") else product_name
                
                if clean_name in self.products:
                    product_prices = self.products[clean_name]
                    
                    # Extract numeric part from thickness (e.g., "2سم" -> "2")
                    thick_value = ''.join(filter(str.isdigit, thickness))
                    
                    if thick_value and thick_value in product_prices:
                        price_data = product_prices[thick_value]
                        
                        # Handle different price structures
                        if isinstance(price_data, list):
                            # Complex pricing with min/max ranges
                            selected_price = 0
                            for price_range in price_data:
                                if isinstance(price_range, dict) and 'min' in price_range and 'max' in price_range and 'price' in price_range:
                                    min_val = price_range['min']
                                    max_val = price_range['max']
                                    # Use length for range checking
                                    if min_val <= length <= max_val:
                                        selected_price = price_range['price']
                                        break
                            # If no range matches, use the first price
                            if selected_price == 0 and len(price_data) > 0:
                                first_item = price_data[0]
                                if isinstance(first_item, dict) and 'price' in first_item:
                                    selected_price = first_item['price']
                            self.price_var.value = str(selected_price)
                        elif isinstance(price_data, (int, float)):
                            # Simple pricing
                            self.price_var.value = str(price_data)
                        else:
                            # Fallback
                            self.price_var.value = "0"
                    else:
                        self.price_var.value = "0"
                else:
                    self.price_var.value = "0"
            else:
                self.price_var.value = "0"
                
            # Trigger calculation to update area and total
            self.calculate(update_page=False)
            
        except Exception as ex:
            # Handle any exceptions silently
            self.price_var.value = "0"
            # Trigger calculation to update area and total
            self.calculate(update_page=False)

    def update_scale(self, scale_factor, update_page=True):
        self.scale_factor = scale_factor
        
        # Calculate new font size (default is usually around 14-16)
        new_text_size = 14 * scale_factor
        
        # Update all text fields with scaling
        controls_map = {
            'block': self.block_var, 'thick': self.thick_var,
            'count': self.count_var, 'len_before': self.len_before_var, 'discount': self.discount_var,
            'len': self.len_var, 'height': self.height_var,
            'area': self.area_var, 'price': self.price_var,
            'total': self.total_var, 'product': self.product_dropdown
        }
        
        for key, control in controls_map.items():
            # Scale the width based on the scale factor but maintain a minimum width
            scaled_width = self.default_widths[key] * scale_factor
            control.text_size = new_text_size
            
            # Update label styles and width
            if isinstance(control, ft.TextField):
                control.label_style = ft.TextStyle(size=new_text_size * 0.9)
                control.width = max(scaled_width, self.default_widths[key] * 0.8)
            elif isinstance(control, ft.Dropdown):
                control.label_style = ft.TextStyle(size=new_text_size * 0.9)
                control.width = max(scaled_width, self.default_widths[key] * 0.8)
                
        if update_page:
            self.page.update()

    def calculate_length(self, e=None):
        """Calculate the final length based on length before discount minus discount"""
        try:
            len_before = float(self.len_before_var.value or 0)
            discount = float(self.discount_var.value or 0)
            final_length = len_before - discount
            
            # Update the readonly length field with formatted value (00.00 format)
            self.len_var.value = f"{final_length:.2f}" if final_length >= 0 else "0.00"
        except ValueError:
            self.len_var.value = "0.00"
        
        # Update calculations
        self.calculate(update_page=False)
        
        # Update the page if available
        if hasattr(self, 'page') and self.page:
            self.page.update()

    def on_thickness_change(self, e):
        """عند تغيير السمك"""
        # When thickness changes, update the price based on current product selection
        selected_product = self.product_dropdown.value
        if selected_product:
            # Remove the "ش " prefix if present
            if selected_product.startswith("ش "):
                clean_product_name = selected_product[2:]  # Remove "ش " prefix
            else:
                clean_product_name = selected_product
            
            # Update the price based on product and new thickness
            self.update_price(clean_product_name)
        else:
            # Just recalculate without updating price
            self.calculate(update_page=False)

    def calculate(self, e=None, update_page=True):
        """Calculate area and total based on current values"""
        try:
            # Handle Arabic decimal input for height
            self.handle_arabic_decimal_input(self.height_var)
            
            # Get values from input fields
            count = float(self.count_var.value or 0)
            length = float(self.len_var.value or 0)
            height = float(self.height_var.value or 0)
            price = float(self.price_var.value or 0)
            
            # Calculate area (count * length * height)
            area = count * length * height
            
            # Calculate total (area * price)
            total = area * price
            
            # Update UI fields
            self.area_var.value = f"{area:.2f}"
            self.total_var.value = f"{total:.2f}"
            
            # Update the page to reflect changes only if requested
            if update_page and hasattr(self, 'page') and self.page:
                self.page.update()
                
        except ValueError:
            # Handle case where conversion to float fails
            self.area_var.value = "0.00"
            self.total_var.value = "0.00"
            if update_page and hasattr(self, 'page') and self.page:
                self.page.update()
        except Exception:
            # Handle any other exceptions silently
            pass

    def get_controls(self):
        """Return Flet controls for this row in reversed order"""
        return [
            self.btn_del,
            self.product_dropdown,
            self.block_var,
            self.thick_var,
            self.count_var,
            self.discount_var,    # New field (moved before len_before_var)
            self.len_before_var,  # New field (moved after discount_var)
            self.len_var,         # Final calculated length
            self.height_var,
            self.area_var,
            self.price_var,
            self.total_var
        ]

    def destroy(self, e):
        # Call the delete callback to remove this row from the rows list
        self.delete_callback(self)
        # Update the page to reflect changes
        self.page.update()




class InvoiceView:
    def __init__(self, page, save_callback):
        self.page = page
        self.save_callback = save_callback
        
        # Configure the page
        self.page.title = "ادارة الفواتير"
        self.page.rtl = True  # Right-to-left support for Arabic
        self.page.theme_mode = ft.ThemeMode.DARK  # Dark theme
        
        # Add keyboard event handler
        self.page.on_keyboard_event = self.on_keyboard_event
        
        self.products_path = resource_path(os.path.join('data', 'products.json'))
        # Use Documents folder for database instead of resources (which is read-only)
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        if not os.path.exists(self.documents_path):
            try:
                os.makedirs(self.documents_path)
            except OSError as e:
                log_error(f"Could not create directory {self.documents_path}: {e}")
                # Fallback to current directory
                self.documents_path = "."
        self.db_path = os.path.join(self.documents_path, 'invoice.db')
        
        # Initialize database with error handling
        try:
            init_db(self.db_path)
        except sqlite3.Error as e:
            log_error(f"Error initializing database: {e}")
            # Try fallback location
            fallback_db_path = os.path.join(".", "invoice.db")
            init_db(fallback_db_path)
            self.db_path = fallback_db_path
        self.products = self.load_products()
        self.op_counter = get_counter(self.db_path)
        self.rows = []
        
        # Load saved zoom level from database
        self.scale_factor = get_zoom_level(self.db_path)
        
        # Form fields
        self.ent_op = ft.TextField(
            label="رقم العملية", 
            value=str(self.op_counter), 
            width=100, 
            on_blur=self.on_op_number_blur  # Check only when focus leaves the field
        )
        
        # Get date - try internet first, fallback to local time
        date_value = self.get_internet_date()
        if not date_value:
            date_value = datetime.now().strftime('%d/%m/%Y')
            
        self.date_var = ft.TextField(label="التاريخ", value=date_value, width=120)
        
        # Client selection with inline autocomplete (ghost text style like IDE)
        self.client_suggestions = self.load_clients()
        self.current_suggestion = ""  # Store current autocomplete suggestion
        
        # Suffix text for showing autocomplete suggestion
        self.client_suffix_text = ft.Text(
            "",
            color=ft.Colors.GREY_500,
            italic=True,
            size=14,
        )
        
        # Main client text field with suffix for autocomplete
        self.ent_client = ft.TextField(
            label="اسم العميل",
            width=200,
            on_change=self.on_client_text_change,
            on_submit=self.on_client_submit,
            suffix=self.client_suffix_text,
        )

        self.ent_driver = ft.TextField(label="اسم السائق", width=150)
        self.ent_phone = ft.TextField(label="رقم التليفون", width=150)
        
        # Payment field for sending payment to Excel
        self.ent_payment = ft.TextField(
            label="المدفوع",
            width=120,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$")
        )
        
        # Variable to store the current invoice file path
        self.current_invoice_path = None
        
        # Main container - using ListView for better performance with many rows
        self.rows_container = ft.ListView(
            expand=True,
            spacing=10,
            padding=10,
            auto_scroll=False
        )
        
        # Floating add button
        self.floating_add_btn = ft.FloatingActionButton(
            icon=ft.Icons.ADD,
            on_click=self.add_row
        )
        
        self.page.update()

    def get_internet_date(self):
        """Get current date from internet, fallback to local time if unavailable"""
        # Use threading to avoid blocking UI
        import threading
        import queue
        
        result_queue = queue.Queue()
        
        def fetch_date():
            try:
                # Try to get date from worldtimeapi.org
                response = urllib.request.urlopen('http://worldtimeapi.org/api/timezone/Africa/Cairo', timeout=3)
                if response.getcode() == 200:
                    import json
                    data = json.loads(response.read().decode())
                    # Parse the datetime string and format it as DD/MM/YYYY
                    dt = datetime.fromisoformat(data['datetime'].replace('Z', '+00:00'))
                    result_queue.put(dt.strftime('%d/%m/%Y'))
                    return
            except Exception as e:
                # If worldtimeapi.org fails, try another service
                try:
                    # Try to get date from httpbin.org as fallback
                    response = urllib.request.urlopen('http://httpbin.org/json', timeout=3)
                    if response.getcode() == 200:
                        # If we can connect, use local time (we just verified internet connectivity)
                        result_queue.put(datetime.now().strftime('%d/%m/%Y'))
                        return
                except Exception as e2:
                    # Internet not available, return None to use local time
                    pass
            result_queue.put(None)
        
        # Start the network request in a separate thread
        thread = threading.Thread(target=fetch_date)
        thread.daemon = True
        thread.start()
        
        # Don't wait for the result, return immediately with local time
        # The network result will be used if it arrives in time
        return datetime.now().strftime('%d/%m/%Y')
    
    def load_clients(self):
        """Load existing client names from the 'فواتير' directory"""
        # Use Documents/alswaife folder
        documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        
        self.invoices_root = os.path.join(documents_path, 'الفواتير')
        if not os.path.exists(self.invoices_root):
            try:
                os.makedirs(self.invoices_root)
            except OSError as e:
                log_error(f"Error creating invoices directory: {e}")
                pass
                
        clients = []
        if os.path.exists(self.invoices_root):
            try:
                with os.scandir(self.invoices_root) as entries:
                    clients = [entry.name for entry in entries if entry.is_dir()]
            except OSError as e:
                log_error(f"Error scanning invoices directory: {e}")
                for item in os.listdir(self.invoices_root):
                    if os.path.isdir(os.path.join(self.invoices_root, item)):
                        clients.append(item)
        return sorted(clients)

    def find_best_match(self, search_text):
        """Find the best matching client name for autocomplete"""
        if not search_text:
            return ""
        
        search_lower = search_text.lower().strip()
        
        # Find first client that starts with the search text
        for client in self.client_suggestions:
            client_lower = client.lower()
            if client_lower.startswith(search_lower):
                return client
        
        # Also try to find clients that contain the search text
        for client in self.client_suggestions:
            if search_lower in client.lower():
                return client
                
        return ""

    def convert_english_to_arabic(self, text):
        """تحويل الحروف الإنجليزية إلى العربية (بناءً على تخطيط لوحة المفاتيح)"""
        # خريطة تحويل الحروف الإنجليزية إلى العربية
        eng_to_ar = {
            'q': 'ض', 'w': 'ص', 'e': 'ث', 'r': 'ق', 't': 'ف', 'y': 'غ', 'u': 'ع', 'i': 'ه', 'o': 'خ', 'p': 'ح', '[': 'ج', ']': 'د',
            'a': 'ش', 's': 'س', 'd': 'ي', 'f': 'ب', 'g': 'ل', 'h': 'ا', 'j': 'ت', 'k': 'ن', 'l': 'م', ';': 'ك', "'": 'ط',
            'z': 'ئ', 'x': 'ء', 'c': 'ؤ', 'v': 'ر', 'b': 'لا', 'n': 'ى', 'm': 'ة', ',': 'و', '.': 'ز', '/': 'ظ',
            'Q': 'َ', 'W': 'ً', 'E': 'ُ', 'R': 'ٌ', 'T': 'لإ', 'Y': 'إ', 'U': ''', 'I': '÷', 'O': '×', 'P': '؛', '{': '<', '}': '>',
            'A': 'ِ', 'S': 'ٍ', 'D': ']', 'F': '[', 'G': 'لأ', 'H': 'أ', 'J': 'ـ', 'K': '،', 'L': '/', ':': ':', '"': '"',
            'Z': '~', 'X': 'ْ', 'C': '}', 'V': '{', 'B': 'لآ', 'N': 'آ', 'M': ''', '<': ',', '>': '.', '?': '؟',
            '`': 'ذ', '~': 'ّ',
        }
        
        result = ""
        for char in text:
            if char in eng_to_ar:
                result += eng_to_ar[char]
            else:
                result += char
        return result

    def on_client_text_change(self, e):
        """Update inline autocomplete suggestion when user types"""
        current_text = self.ent_client.value or ""
        
        # تحويل الحروف الإنجليزية إلى العربية
        converted_text = self.convert_english_to_arabic(current_text)
        
        # تحديث النص إذا تم التحويل
        if converted_text != current_text:
            self.ent_client.value = converted_text
            current_text = converted_text
        
        # Don't strip - keep the text as is for proper matching with names containing spaces
        search_text = current_text.lstrip()  # Only remove leading spaces
        
        if search_text:
            # Find best match
            best_match = self.find_best_match(search_text)
            
            if best_match and best_match.lower() != search_text.lower():
                # Show the full suggestion name as suffix
                self.client_suffix_text.value = best_match
                self.current_suggestion = best_match
            else:
                self.client_suffix_text.value = ""
                self.current_suggestion = ""
        else:
            self.client_suffix_text.value = ""
            self.current_suggestion = ""
        
        self.page.update()

    def on_client_submit(self, e):
        """Handle Tab or Enter - accept the autocomplete suggestion"""
        if self.current_suggestion:
            self.ent_client.value = self.current_suggestion
            self.client_suffix_text.value = ""
            self.current_suggestion = ""
            self.page.update()
        
        # Load client data
        client_name = self.ent_client.value.strip() if self.ent_client.value else ""
        if client_name:
            self.load_client_data(client_name)
    
    def load_client_data(self, client_name):
        """Load default data for the selected client"""
        # This method can be expanded to load default values for the client
        # In the future, this could load default driver, phone, or other client-specific settings
        pass
    
    def sanitize(self, s):
        """Sanitize string for file/folder names"""
        import re
        return re.sub(r'[\\/*?:"<>|]', "", str(s))

    def load_products(self):
        # Check if file exists first to avoid unnecessary operations
        if not os.path.exists(self.products_path):
            return {}
        
        try:
            # Use buffered reading for better performance
            with open(self.products_path, 'r', encoding='utf-8', buffering=8192) as f:
                data = json.load(f)
                products = {}
                if isinstance(data, dict):
                    # Handle dict of products
                    products = {str(k): v for k, v in data.items()}
                elif isinstance(data, list):
                    # Handle list of products
                    for item in data:
                        if 'name' in item:
                            name = str(item['name'])
                            if 'prices' in item:
                                # New structure: {'2': 310, '3': 400}
                                products[name] = item['prices']
                            else:
                                # Old structure fallback
                                products[name] = item.get('price', 0)
                return products
        except (IOError, json.JSONDecodeError) as e:
            # Log error but don't crash
            log_error(f"Could not load products file: {e}")
            return {}
        except Exception as e:
            # Handle any other unexpected errors
            log_error(f"Unexpected error loading products: {e}")
            return {}

    def add_row(self, e=None):
        row_idx = len(self.rows)
        new_row = InvoiceRow(self.page, row_idx, self.products, self.delete_row, self.scale_factor)
        self.rows.append(new_row)
        
        # Create a row container for the ListView with responsive layout
        row_controls = new_row.get_controls()
        row_container = ft.Row(
            controls=row_controls, 
            spacing=5,
            wrap=False  # Don't wrap controls to next line
        )
        
        # Wrap the row in a Container for better styling and spacing
        row_wrapper = ft.Container(
            content=row_container,
            padding=10,
            border=ft.border.all(1, ft.Colors.GREY_700),
            border_radius=5,
            bgcolor=ft.Colors.GREY_900
        )
        
        # Store reference to the row wrapper for deletion
        if not hasattr(self, 'row_wrappers'):
            self.row_wrappers = {}
        self.row_wrappers[new_row] = row_wrapper
        
        # Add to ListView instead of Column
        self.rows_container.controls.append(row_wrapper)
        
        self.page.update()
        
    def delete_row(self, row_obj):
        if row_obj in self.rows:
            # Remove the row wrapper from the UI
            if hasattr(self, 'row_wrappers') and row_obj in self.row_wrappers:
                row_wrapper = self.row_wrappers[row_obj]
                if row_wrapper in self.rows_container.controls:
                    self.rows_container.controls.remove(row_wrapper)
                # Clean up the reference
                del self.row_wrappers[row_obj]
            
            # Remove the row from the data structure
            self.rows.remove(row_obj)
            
            # Update UI
            self.page.update()

    def save_excel(self, e):
        # Check if Excel is running first
        if is_excel_running():
            self._show_excel_warning_dialog()
            return
        
        self._perform_save()
    
    def _show_excel_warning_dialog(self):
        """Show warning dialog when Excel is open"""
        def on_skip(e):
            dlg.open = False
            self.page.update()
            # Proceed with save anyway
            self._perform_save()
        
        def on_cancel(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.WARNING_AMBER_ROUNDED, color=ft.Colors.ORANGE_400, size=28),
                    ft.Text("تنبيه - برنامج Excel مفتوح", weight=ft.FontWeight.BOLD, color=ft.Colors.ORANGE_400),
                ],
                spacing=10
            ),
            content=ft.Column(
                controls=[
                    ft.Text(
                        "تم اكتشاف أن برنامج Microsoft Excel مفتوح حالياً.",
                        size=14),
                    ft.Container(height=10),
                    ft.Text(
                        "قد يؤدي ذلك إلى فشل حفظ الملفات إذا كانت مفتوحة في Excel.",
                        size=14,
                        rtl=True,
                        color=ft.Colors.GREY_400
                    ),
                    ft.Container(height=10),
                    ft.Text(
                        "يُنصح بإغلاق جميع ملفات Excel قبل المتابعة.",
                        size=14,
                        rtl=True,
                        weight=ft.FontWeight.W_500
                    ),
                ],
                tight=True),
            actions=[
                ft.TextButton(
                    "تخطي والمتابعة",
                    on_click=on_skip,
                    style=ft.ButtonStyle(color=ft.Colors.ORANGE_400)
                ),
                ft.TextButton(
                    "إلغاء",
                    on_click=on_cancel,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()
    
    def _perform_save(self):
        """Perform the actual save operation"""
        try:
            op_num = self.ent_op.value.strip() if self.ent_op.value else ""
            client = self.ent_client.value.strip() if self.ent_client.value is not None else ""
            date_str = self.date_var.value.strip() if self.date_var.value else datetime.now().strftime('%d/%m/%Y')
            driver = self.ent_driver.value.strip() if self.ent_driver.value else ""
            phone = self.ent_phone.value.strip() if self.ent_phone.value else ""

            # Check if client is empty or contains "ايراد"
            # If client is empty, set it to "ايراد"
            if not client:
                client = "ايراد"
            is_revenue = "ايراد" in client
        
            if is_revenue:
                # Show confirmation dialog
                def on_confirm_revenue(e):
                    confirm_dlg.open = False
                    self.page.update()
                    self._save_revenue_invoice(op_num, client, date_str, driver, phone)
                
                def on_cancel_revenue(e):
                    confirm_dlg.open = False
                    self.page.update()
                
                confirm_dlg = ft.AlertDialog(
                    title=ft.Text("تنبيه"),
                    content=ft.Text("العميل فارغ أو يحتوي على كلمة 'ايراد'. سيتم حفظ الفاتورة في مجلد الإيرادات دون إنشاء كشف حساب. هل توافق؟"),
                    actions=[
                        ft.TextButton("نعم", on_click=on_confirm_revenue),
                        ft.TextButton("لا", on_click=on_cancel_revenue)
                    ],
                    actions_alignment=ft.MainAxisAlignment.END
                )
                self.page.overlay.append(confirm_dlg)
                confirm_dlg.open = True
                self.page.update()
                return

            if not op_num:
                dlg = ft.AlertDialog(
                    title=ft.Text("خطأ"),
                    content=ft.Text("يرجى إدخال رقم العملية")
                )
                self.page.overlay.append(dlg)
                dlg.open = True
                self.page.update()
                return

        except Exception as ex:
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text(f"حدث خطأ أثناء الحفظ:\n{ex}\n{traceback.format_exc()}")
            )
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            return

        # Validate items before saving
        items_data = []
        validation_errors = []
        
        for row_index, row in enumerate(self.rows):
            row_num = row_index + 1
            # Check for empty required fields (except block number)
            if not row.product_dropdown.value or row.product_dropdown.value.strip() == "":
                validation_errors.append(f"الصف {row_num}: البيان فارغ")
            
            if not row.thick_var.value or row.thick_var.value.strip() == "":
                validation_errors.append(f"الصف {row_num}: السمك فارغ")
            
            count_val = row.count_var.value or "0"
            if count_val.strip() == "" or count_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: العدد فارغ أو صفر")
            
            len_before_val = row.len_before_var.value or "0"
            if len_before_val.strip() == "" or len_before_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: الطول قبل فارغ أو صفر")
            
            height_val = row.height_var.value or "0"
            if height_val.strip() == "" or height_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: الارتفاع فارغ أو صفر")
            
            price_val = row.price_var.value or "0"
            if price_val.strip() == "" or price_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: السعر فارغ أو صفر")
            
            # Collect actual data from the row controls
            item_data = (
                row.product_dropdown.value or "",  # description
                row.block_var.value or "",         # block (optional)
                row.thick_var.value or "",         # thickness
                row.material_value or "",          # material
                row.count_var.value or "0",        # count
                row.len_var.value or "0",          # length (final calculated value only)
                row.height_var.value or "0",       # height
                row.price_var.value or "0",        # price
                row.len_before_var.value or "0",   # length_before (for calculation)
                row.discount_var.value or "0"      # discount (for calculation)
            )
            items_data.append(item_data)

        # Show validation errors if any
        if validation_errors:
            error_message = "\n".join(validation_errors[:10])  # Show max 10 errors
            if len(validation_errors) > 10:
                error_message += f"\n... و {len(validation_errors) - 10} أخطاء أخرى"
            
            dlg = ft.AlertDialog(
                title=ft.Text("تنبيه - خانات فارغة"),
                content=ft.Text(error_message))
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            return

        if not items_data:
            dlg = ft.AlertDialog(
                title=ft.Text("تنبيه"),
                content=ft.Text("لا توجد بنود للحفظ")
            )
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            return

        def sanitize(s):
            return re.sub(r'[\\/*?:"<>|]', "", str(s))
        
        now = datetime.now()
        
        # Create folder structure
        # Root/الفواتير/ClientName/فواتيري/
        client_safe = sanitize(client)
        if not client_safe:
            client_safe = "General"
            
        client_dir = os.path.join(self.invoices_root, client_safe)
        my_invoices_dir = os.path.join(client_dir, "فواتيري")
        
        try:
            os.makedirs(my_invoices_dir, exist_ok=True)
        except OSError as ex:
             dlg = ft.AlertDialog(title=ft.Text("خطأ"), content=ft.Text(f"فشل إنشاء المجلد: {ex}"))
             self.page.overlay.append(dlg)
             dlg.open = True
             self.page.update()
             return

        # Update filename format
        fname = f"فاتورة رقم {sanitize(op_num)} - بتاريخ {date_str.replace('/', '-')}.xlsx"
        full_path = os.path.join(my_invoices_dir, fname)
        
        try:
            # Check if this is an update to an existing invoice
            invoice_already_exists = invoice_exists(self.db_path, op_num)
            old_file_path = None
            
            if invoice_already_exists:
                # Load the existing invoice to check if the client name has changed
                existing_invoice = load_invoice_from_db(self.db_path, op_num)
                if existing_invoice:
                    # Get the old file path to delete it
                    old_file_path = existing_invoice.get("file_path", None)
                    
                    old_client_name = existing_invoice.get("client_name", "")
                    old_client_safe = self.sanitize(old_client_name)
                    if not old_client_safe:
                        old_client_safe = "General"
                    
                    # If client name has changed, remove from old client's ledger
                    # (update_invoice_in_ledger will handle adding to the new client's ledger)
                    if old_client_name != client:
                        old_client_folder = os.path.join(self.invoices_root, old_client_safe)
                        try:
                            remove_invoice_from_ledger(old_client_folder, op_num)
                        except Exception as ledger_ex:
                            log_error(f"Error removing old ledger entry: {ledger_ex}")
                    # Note: If client name is the same, update_invoice_in_ledger will handle
                    # removing and re-adding the invoice entry internally
            
            # Delete the old invoice file if it exists
            if old_file_path and old_file_path != full_path:
                try:
                    delete_existing_invoice_file(old_file_path)
                except Exception as file_ex:
                    log_error(f"Error deleting old invoice file: {file_ex}")
            
            # Save the invoice Excel file directly
            from utils.invoice_utils import save_invoice
            # Extract only the first 8 elements for saving to Excel (excluding length_before and discount)
            items_for_excel = []
            for item in items_data:
                # Take only the first 8 elements: description, block, thickness, material, count, length, height, price
                item_excel = tuple(item[:8]) if len(item) >= 8 else item
                items_for_excel.append(item_excel)
            
            # Save the invoice
            save_invoice(full_path, op_num, client, driver, items_for_excel, date_str=date_str, phone=phone)
            
            # Save invoice data to database
            try:
                # Calculate total amount from items
                total_amount = 0
                for item in items_data:
                    try:
                        count = int(float(item[4])) if item[4] else 0  # Convert to int (count should be whole number)
                        length = float(item[5]) if item[5] else 0
                        height = float(item[6]) if item[6] else 0
                        price = float(item[7]) if item[7] else 0
                        area = count * length * height
                        total_amount += area * price
                    except (ValueError, IndexError):
                        continue
                
                # Save to database
                save_invoice_to_db(self.db_path, op_num, client, driver, phone, date_str, full_path, items_data, total_amount)
                
                # If this is an update to an existing invoice, we need to update the ledger properly
                if invoice_already_exists:
                    # Update the invoice in the ledger instead of appending
                    try:
                        # Calculate invoice items for ledger
                        invoice_items_details = []
                        for item in items_data:
                            try:
                                desc = item[0] or ""
                                material = item[3] or ""
                                thickness = item[2] or ""
                                count = int(float(item[4])) if item[4] else 0  # Convert to int (count should be whole number)
                                length = float(item[5]) if item[5] else 0
                                height = float(item[6]) if item[6] else 0
                                price_val = float(item[7]) if item[7] else 0
                                
                                # Calculate area for this item
                                area = count * length * height
                                
                                # Store item details for the ledger - format: (desc, material, thickness, area, total_price_for_item)
                                # In the original code, the last element is the total price for the item, not the unit price
                                item_total_price = area * price_val
                                invoice_items_details.append((
                                    desc,
                                    material,
                                    thickness,
                                    area,
                                    item_total_price  # This should be the total price for the item, not unit price
                                ))
                            except (ValueError, IndexError):
                                continue
                        
                        # Update the invoice in the ledger
                        success, error = update_invoice_in_ledger(
                            client_dir,
                            op_num,
                            client,
                            date_str,
                            total_amount,
                            driver,
                            invoice_items_details
                        )
                        
                        if not success:
                            log_error(f"Could not update invoice in ledger: {error}")
                    except Exception as ledger_ex:
                        log_exception(f"Error updating invoice in ledger: {ledger_ex}")
                else:
                    # For new invoices, use the original save_callback to handle ledger update
                    # Calculate items for original callback
                    items_for_callback = []
                    for item in items_data:
                        # Take only the first 8 elements: description, block, thickness, material, count, length, height, price
                        item_callback = tuple(item[:8]) if len(item) >= 8 else item
                        items_for_callback.append(item_callback)
                    
                    # Call the original save callback which will handle ledger update
                    self.save_callback(full_path, op_num, client, driver, date_str, phone, items_for_callback)
            except Exception as db_ex:
                log_error(f"Error saving invoice to database: {db_ex}")
                # Continue with the process even if database save fails
            
            # Disburse slides from inventory if invoice contains slide products
            try:
                slides_success, slides_message, disbursed_items = disburse_slides_from_invoice(
                    op_num, date_str, items_data, client
                )
            except Exception as slides_ex:
                log_error(f"Error disbursing slides: {slides_ex}")
                # Continue with the process even if slides disbursement fails
            
            # Store the current invoice path for payment updates
            self.current_invoice_path = full_path
            
            # Send payment to Excel if payment field has a value
            payment_str = self.ent_payment.value.strip() if self.ent_payment.value else ""
            payment_amount = 0
            if payment_str:
                try:
                    payment_amount = float(payment_str)
                    # Update payment in invoice file
                    payment_success = update_payment_in_invoice(full_path, payment_amount)
                    if not payment_success:
                        log_error(f"Failed to send payment to invoice")
                    
                    # Also update payment in client ledger
                    ledger_success = update_payment_in_ledger(client_dir, op_num, payment_amount)
                    if not ledger_success:
                        log_error(f"Failed to send payment to client ledger")
                except ValueError:
                    log_error(f"Invalid payment value: {payment_str}")
                except Exception as payment_ex:
                    log_error(f"Error sending payment: {payment_ex}")
            
            # Add income record to purchases/income ledger (always record invoice)
            try:
                purchases_folder = os.path.join(self.documents_path, "ايرادات ومصروفات")
                os.makedirs(purchases_folder, exist_ok=True)
                purchases_file = os.path.join(purchases_folder, "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
                income_record = {
                    'date': date_str,
                    'invoice_number': op_num,
                    'client': client if client else "بدون اسم",
                    'amount': payment_amount if payment_amount > 0 else total_amount,
                }
                result = add_income_record(purchases_file, income_record)
            except Exception as income_ex:
                log_exception(f"Error adding income record: {income_ex}")
            
            def open_file(e):
                # Use our universal function to open the file
                open_path(full_path)
            def open_folder(e):
                # Use our universal function to open the folder
                open_path(client_dir)

            def open_ledger(e):
                # Use the correct ledger file name
                ledger_path = os.path.join(client_dir, "كشف حساب.xlsx")
                # Use our universal function to open the ledger file
                open_path(ledger_path)

            def close_dlg(e):
                dlg.open = False
                self.page.update()

            dlg = ft.AlertDialog(
                title=ft.Text("نجاح"),
                content=ft.Text(f"تم حفظ الفاتورة وتحديث كشف الحساب بنجاح.\المسار: {full_path}"),
                actions=[
                    ft.TextButton("فتح الفاتورة", on_click=open_file),
                    ft.TextButton("فتح كشف الحساب", on_click=open_ledger),
                    ft.TextButton("فتح المجلد", on_click=open_folder),
                    ft.TextButton("حسنا", on_click=close_dlg)
                ],
                actions_alignment=ft.MainAxisAlignment.END
            )
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            
        except PermissionError as ex:
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text("الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى."))
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
        except Exception as ex:
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text(f"حدث خطأ أثناء الحفظ:\n{ex}\n{traceback.format_exc()}")
            )
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()

    def _save_revenue_invoice(self, op_num, client, date_str, driver, phone):
        """Save revenue invoice to a separate directory without creating a ledger"""
        if not op_num:
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text("يرجى إدخال رقم العملية")
            )
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            return

        # Validate items before saving
        items_data = []
        validation_errors = []
        
        for row_index, row in enumerate(self.rows):
            row_num = row_index + 1
            
            # Check for empty required fields (except block number)
            if not row.product_dropdown.value or row.product_dropdown.value.strip() == "":
                validation_errors.append(f"الصف {row_num}: البيان فارغ")
            
            if not row.thick_var.value or row.thick_var.value.strip() == "":
                validation_errors.append(f"الصف {row_num}: السمك فارغ")
            
            count_val = row.count_var.value or "0"
            if count_val.strip() == "" or count_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: العدد فارغ أو صفر")
            
            len_before_val = row.len_before_var.value or "0"
            if len_before_val.strip() == "" or len_before_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: الطول قبل فارغ أو صفر")
            
            height_val = row.height_var.value or "0"
            if height_val.strip() == "" or height_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: الارتفاع فارغ أو صفر")
            
            price_val = row.price_var.value or "0"
            if price_val.strip() == "" or price_val.strip() == "0":
                validation_errors.append(f"الصف {row_num}: السعر فارغ أو صفر")
            
            # Collect actual data from the row controls
            item_data = (
                row.product_dropdown.value or "",  # description
                row.block_var.value or "",         # block (optional)
                row.thick_var.value or "",         # thickness
                row.material_value or "",          # material
                row.count_var.value or "0",        # count
                row.len_var.value or "0",          # length (already net)
                row.height_var.value or "0",       # height
                row.price_var.value or "0",        # price
                row.len_before_var.value or "0",   # length_before (for calculation)
                row.discount_var.value or "0"      # discount (for calculation)
            )
            items_data.append(item_data)

        # Show validation errors if any
        if validation_errors:
            error_message = "\n".join(validation_errors[:10])  # Show max 10 errors
            if len(validation_errors) > 10:
                error_message += f"\n... و {len(validation_errors) - 10} أخطاء أخرى"
            
            dlg = ft.AlertDialog(
                title=ft.Text("تنبيه - خانات فارغة"),
                content=ft.Text(error_message))
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            return

        if not items_data:
            dlg = ft.AlertDialog(
                title=ft.Text("تنبيه"),
                content=ft.Text("لا توجد بنود للحفظ")
            )
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            return

        def sanitize(s):
            return re.sub(r'[\/*?:"<>|]', "", str(s))
        
        now = datetime.now()
        
        # Create revenue folder structure
        # Root/ايرادات/
        revenue_dir = os.path.join(self.invoices_root, "ايرادات")
        my_invoices_dir = revenue_dir  # Save directly in ايرادات folder, not in فواتيري subfolder
        
        try:
            os.makedirs(my_invoices_dir, exist_ok=True)
        except OSError as ex:
             dlg = ft.AlertDialog(title=ft.Text("خطأ"), content=ft.Text(f"فشل إنشاء المجلد: {ex}"))
             self.page.overlay.append(dlg)
             dlg.open = True
             self.page.update()
             return

        fname = f"فاتورة رقم ({sanitize(op_num)}) _ ايراد _ بتاريخ {date_str.replace('/', '-')}.xlsx"
        full_path = os.path.join(my_invoices_dir, fname)
        
        try:
            # For revenue invoices, check if it already exists
            invoice_already_exists = invoice_exists(self.db_path, op_num)
            old_file_path = None
            
            # Revenue invoices don't have ledger entries, so no need to remove them
            
            # But if it's an update, we might need to handle client changes
            if invoice_already_exists:
                # Load the existing invoice to check if the client name has changed
                existing_invoice = load_invoice_from_db(self.db_path, op_num)
                if existing_invoice:
                    # Get the old file path to delete it
                    old_file_path = existing_invoice.get("file_path", None)
                    
                    old_client_name = existing_invoice.get("client_name", "")
                    # For revenue invoices, still check if we need to clean up old client ledger
                    if old_client_name != client and "ايراد" not in old_client_name:
                        # If the old invoice was not revenue but new one is, remove from old client's ledger
                        old_client_safe = self.sanitize(old_client_name)
                        if not old_client_safe:
                            old_client_safe = "General"
                        old_client_folder = os.path.join(self.invoices_root, old_client_safe)
                        try:
                            remove_invoice_from_ledger(old_client_folder, op_num)
                        except Exception as ledger_ex:
                            log_error(f"Error removing old ledger entry: {ledger_ex}")
            
            # Delete the old invoice file if it exists
            if old_file_path and old_file_path != full_path:
                try:
                    delete_existing_invoice_file(old_file_path)
                except Exception as file_ex:
                    log_error(f"Error deleting old invoice file: {file_ex}")
            
            # Save the revenue invoice Excel file directly
            from utils.invoice_utils import save_invoice
            # Extract only the first 8 elements for saving to Excel (excluding length_before and discount)
            items_for_excel = []
            for item in items_data:
                # Take only the first 8 elements: description, block, thickness, material, count, length, height, price
                item_excel = tuple(item[:8]) if len(item) >= 8 else item
                items_for_excel.append(item_excel)
            
            # Save the invoice
            save_invoice(full_path, op_num, client, driver, items_for_excel, date_str=date_str, phone=phone)
            
            # Save revenue invoice data to database
            try:
                # Calculate total amount from items
                total_amount = 0
                for item in items_data:
                    try:
                        count = int(float(item[4])) if item[4] else 0  # Convert to int (count should be whole number)
                        length = float(item[5]) if item[5] else 0
                        height = float(item[6]) if item[6] else 0
                        price = float(item[7]) if item[7] else 0
                        area = count * length * height
                        total_amount += area * price
                    except (ValueError, IndexError):
                        continue
                
                # Save to database
                save_invoice_to_db(self.db_path, op_num, client, driver, phone, date_str, full_path, items_data, total_amount)
                
                # For revenue invoices, no ledger update is needed, but if this was a conversion
                # from a regular invoice to revenue, we already handled the ledger removal above
            except Exception as db_ex:
                log_error(f"Error saving revenue invoice to database: {db_ex}")
                # Continue with the process even if database save fails
            
            # Disburse slides from inventory if invoice contains slide products
            try:
                slides_success, slides_message, disbursed_items = disburse_slides_from_invoice(
                    op_num, date_str, items_data, client
                )
            except Exception as slides_ex:
                log_error(f"Error disbursing slides for revenue invoice: {slides_ex}")
                # Continue with the process even if slides disbursement fails
            
            # Store the current invoice path for payment updates
            self.current_invoice_path = full_path
            
            # Send payment to Excel if payment field has a value
            payment_str = self.ent_payment.value.strip() if self.ent_payment.value else ""
            payment_amount = 0
            if payment_str:
                try:
                    payment_amount = float(payment_str)
                    payment_success = update_payment_in_invoice(full_path, payment_amount)
                    if not payment_success:
                        log_error(f"Failed to send payment to revenue invoice")
                except ValueError:
                    log_error(f"Invalid payment value: {payment_str}")
                except Exception as payment_ex:
                    log_error(f"Error sending payment: {payment_ex}")
            
            # Add income record to income/expenses ledger
            try:
                purchases_folder = os.path.join(self.documents_path, "ايرادات ومصروفات")
                os.makedirs(purchases_folder, exist_ok=True)
                purchases_file = os.path.join(purchases_folder, "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
                income_record = {
                    'date': date_str,
                    'invoice_number': op_num,
                    'client': client if client else "بدون اسم",
                    'amount': payment_amount if payment_amount > 0 else total_amount,
                }
                result = add_income_record(purchases_file, income_record)
            except Exception as income_ex:
                log_exception(f"Error adding income record: {income_ex}")
            
            def open_file(e):
                try:
                    if platform.system() == 'Windows':
                        # Use subprocess with Excel-specific parameters to ensure normal window state
                        import winreg
                        try:
                            # Get Excel path from registry
                            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID") as key:
                                clsid = winreg.QueryValue(key, "")
                            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as key:
                                excel_path = winreg.QueryValue(key, "")
                            # Open with Excel in normal window state
                            subprocess.Popen([excel_path, '/e', full_path], shell=False)
                        except:
                            # Fallback to default method if registry lookup fails
                            os.startfile(full_path)
                    elif platform.system() == 'Darwin':
                        subprocess.call(('open', full_path))
                    else:
                        subprocess.call(('xdg-open', full_path))
                except Exception as ex:
                    log_error(f"Error opening file: {ex}")

            def open_folder(e):
                # Use our universal function to open the folder
                open_path(revenue_dir)

            def close_dlg(e):
                dlg.open = False
                self.page.update()

            dlg = ft.AlertDialog(
                title=ft.Text("نجاح"),
                content=ft.Text(f"تم حفظ فاتورة الإيراد بنجاح.\نالمسار: {full_path}"),
                actions=[
                    ft.TextButton("فتح الفاتورة", on_click=open_file),
                    ft.TextButton("فتح المجلد", on_click=open_folder),
                    ft.TextButton("حسنا", on_click=close_dlg)
                ],
                actions_alignment=ft.MainAxisAlignment.END
            )
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
            
        except PermissionError as ex:
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text("الملف مفتوح حالياً في برنامج Excel. يرجى إغلاق الملف والمحاولة مرة أخرى."))
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
        except Exception as ex:
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text(f"حدث خطأ أثناء الحفظ:\n{ex}\n{traceback.format_exc()}"))
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()

    def on_op_number_blur(self, e):
        """Handle when focus leaves the invoice number field"""
        op_num = self.ent_op.value.strip() if self.ent_op.value else ""
        if op_num:
            # Check if invoice exists in database
            if invoice_exists(self.db_path, op_num):
                # Ask user if they want to load the existing invoice
                def on_confirm_load(e):
                    confirm_dlg.open = False
                    self.page.update()
                    self.load_invoice_data(op_num)
                
                def on_cancel_load(e):
                    confirm_dlg.open = False
                    self.page.update()
                
                confirm_dlg = ft.AlertDialog(
                    title=ft.Text("تنبيه"),
                    content=ft.Text(f"يوجد فاتورة برقم {op_num} محفوظة مسبقاً. هل تريد تحميل بياناتها؟"),
                    actions=[
                        ft.TextButton("نعم", on_click=on_confirm_load),
                        ft.TextButton("لا", on_click=on_cancel_load)
                    ],
                    actions_alignment=ft.MainAxisAlignment.END
                )
                self.page.overlay.append(confirm_dlg)
                confirm_dlg.open = True
                self.page.update()
        
    def load_invoice_data(self, op_num):
        """Load invoice data from database and populate the form"""
        try:
            invoice_data = load_invoice_from_db(self.db_path, op_num)
            if invoice_data:
                # Populate form fields
                self.ent_op.value = invoice_data["invoice_number"]
                self.ent_client.value = invoice_data["client_name"]
                self.ent_driver.value = invoice_data["driver_name"]
                self.ent_phone.value = invoice_data["phone"]
                self.date_var.value = invoice_data["date"]
                
                # Store the invoice file path for payment updates
                self.current_invoice_path = invoice_data.get("file_path", None)
                
                # Load payment from invoice file if exists
                if self.current_invoice_path and os.path.exists(self.current_invoice_path):
                    payment_value = get_payment_from_invoice(self.current_invoice_path)
                    if payment_value > 0:
                        self.ent_payment.value = str(int(payment_value)) if payment_value == int(payment_value) else str(payment_value)
                    else:
                        self.ent_payment.value = ""
                else:
                    self.ent_payment.value = ""
                
                # Clear existing rows
                self.rows.clear()
                self.rows_container.controls.clear()
                
                # Add rows with loaded data
                for item in invoice_data["items"]:
                    self.add_row()  # Add an empty row
                    new_row = self.rows[-1]  # Get the newly added row
                    
                    # Set values for the row
                    new_row.product_dropdown.value = item[0] if len(item) > 0 else ""
                    new_row.block_var.value = item[1] if len(item) > 1 else ""
                    new_row.thick_var.value = item[2] if len(item) > 2 else ""
                    new_row.material_value = item[3] if len(item) > 3 else ""
                    
                    # Set numeric values - convert count to int to avoid "10.0" display
                    new_row.count_var.value = str(int(float(item[4]))) if len(item) > 4 and item[4] else "0"
                    
                    # Set the length fields - final length, length_before, and discount
                    saved_length = str(item[5]) if len(item) > 5 else "0"
                    saved_length_before = str(item[8]) if len(item) > 8 else "0"
                    saved_discount = str(item[9]) if len(item) > 9 else "0"
                    
                    # Set the length fields
                    new_row.len_var.value = saved_length
                    new_row.len_before_var.value = saved_length_before
                    new_row.discount_var.value = saved_discount
                    
                    new_row.height_var.value = str(item[6]) if len(item) > 6 else "0"
                    # Convert price to int to avoid "100.0" display
                    new_row.price_var.value = str(int(float(item[7]))) if len(item) > 7 and item[7] else "0"
                    
                    # Update the row calculations
                    new_row.calculate(update_page=False)
                
                # Update the UI
                self.page.update()
            
        except Exception as ex:
            log_error(f"Error loading invoice data: {ex}")
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text(f"حدث خطأ أثناء تحميل بيانات الفاتورة:\n{ex}"))
            self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()

    def increment_op(self):
        try:
            new_val = increment_counter(self.db_path)
            self.op_counter = new_val
        except:
            self.op_counter += 1
        
        self.ent_op.value = str(self.op_counter)
        self.page.update()

    def reset_form(self, e):
        """إعادة تعيين النموذج لعملية جديدة مع زيادة رقم العملية"""
        self.ent_client.value = ""
        self.ent_driver.value = ""
        self.ent_phone.value = ""
        self.ent_payment.value = ""
        
        # Reset invoice path
        self.current_invoice_path = None
        
        # تحديث التاريخ للتاريخ الحالي
        date_value = self.get_internet_date()
        if not date_value:
            date_value = datetime.now().strftime('%d/%m/%Y')
        self.date_var.value = date_value
        
        # Clear all rows
        self.rows.clear()
        self.rows_container.controls.clear()
        if hasattr(self, 'row_wrappers'):
            self.row_wrappers.clear()
        
        # زيادة رقم العملية للعملية الجديدة
        self.increment_op()
        
        # Add one empty row
        self.add_row()
        
        self.page.update()

    def update_rows_scale(self):
        # Update row fields only
        for row in self.rows:
            row.update_scale(self.scale_factor)
        self.page.update()

    def minimize_window(self, e):
        """Minimize the application window"""
        try:
            # Try different approaches to minimize the window
            # Approach 1: Direct attribute assignment
            if hasattr(self.page, 'window_minimized'):
                self.page.window_minimized = True
            # Approach 2: Try to access through window object
            elif hasattr(self.page, 'window') and hasattr(self.page.window, 'minimized'):
                self.page.window.minimized = True
            # Approach 3: Try setattr with different possible names
            else:
                success = False
                possible_attrs = ['window_minimized', 'minimized']
                for attr in possible_attrs:
                    try:
                        setattr(self.page, attr, True)
                        success = True
                        break
                    except:
                        continue
                
                # If all else fails, log for debugging
                if not success:
                    log_error("Could not minimize window - no compatible attribute found")
        except Exception as ex:
            log_error(f"Error minimizing window: {ex}")
        self.page.update()

    def close_window(self, e):
        """Close the application window"""
        try:
             self.page.window.close()
        except Exception as ex:
            log_error(f"Error closing window: {ex}")

    def on_keyboard_event(self, e: ft.KeyboardEvent):
        """Handle keyboard events"""
        # Check if the '+' or '=' key was pressed
        if e.key == '+' or e.key == '=' and not e.ctrl and not e.shift and not e.alt:
            # Add a new row when '+' is pressed
            self.add_row()
            
    def zoom_in(self, e):
        if self.scale_factor < 2.0:  # Upper limit for zoom
            self.scale_factor += 0.1  # Smaller zoom increment for smoother scaling
            self.update_rows_scale()
            set_zoom_level(self.db_path, self.scale_factor)
    def zoom_out(self, e):
        if self.scale_factor > 0.4:  # Lower minimum zoom level
            self.scale_factor -= 0.1  # Match the zoom in increment
            self.update_rows_scale()
            set_zoom_level(self.db_path, self.scale_factor)

    def go_back(self, e):
        """Go back to dashboard"""
        # Import here to avoid circular dependency
        from views.dashboard_view import DashboardView
        
        self.page.clean()
        dashboard = DashboardView(self.page)
        dashboard.show()

    def build_ui(self):
        # Create AppBar with grouped buttons (no menu)
        self.page.appbar = ft.AppBar(
            leading=ft.IconButton(ft.Icons.ARROW_BACK, on_click=self.go_back, tooltip="العودة"),
            title=ft.Text("مصنع السويفي - ادارة الفواتير"),
            bgcolor=ft.Colors.SURFACE,
            actions=[
                # Zoom buttons group
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
                    border_radius=8,
                    padding=ft.padding.symmetric(horizontal=5),
                    margin=ft.margin.only(left=10),
                ),
                # Save and New operation buttons group
                ft.Container(
                    content=ft.Row(
                        controls=[
                            ft.IconButton(
                                ft.Icons.SAVE, 
                                on_click=self.save_excel, 
                                tooltip="حفظ إلى Excel",
                                icon_color=ft.Colors.GREEN_400,
                            ),
                            ft.IconButton(
                                ft.Icons.ADD_CIRCLE_OUTLINE, 
                                on_click=self.reset_form, 
                                tooltip="عملية جديدة",
                                icon_color=ft.Colors.ORANGE_400,
                            ),
                        ],
                        spacing=0,
                    ),
                    bgcolor=ft.Colors.GREY_800,
                    border_radius=8,
                    padding=ft.padding.symmetric(horizontal=5),
                    margin=ft.margin.only(left=10, right=15),
                ),
            ]
        )
        
        # Header section - improved layout with container
        header_content = ft.Column([
            ft.Row([
                ft.Column([ft.Icon(ft.Icons.NUMBERS, color=ft.Colors.BLUE_300, size=20), self.ent_op], 
                         spacing=8, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Column([ft.Icon(ft.Icons.CALENDAR_TODAY, color=ft.Colors.BLUE_300, size=20), self.date_var], 
                         spacing=8, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Column([ft.Icon(ft.Icons.PERSON, color=ft.Colors.BLUE_300, size=20), 
                          self.ent_client], 
                         spacing=8, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Column([ft.Icon(ft.Icons.DRIVE_ETA, color=ft.Colors.BLUE_300, size=20), self.ent_driver], 
                         spacing=8, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Column([ft.Icon(ft.Icons.PHONE, color=ft.Colors.BLUE_300, size=20), self.ent_phone], 
                         spacing=8, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Column([
                    ft.Icon(ft.Icons.PAYMENTS, color=ft.Colors.GREEN_300, size=20),
                    self.ent_payment
                ], spacing=8, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            ], spacing=10, alignment=ft.MainAxisAlignment.SPACE_EVENLY),
        ], spacing=5)
        
        header = ft.Container(
            content=header_content,
            padding=20,
            bgcolor=ft.Colors.GREY_800,
            border_radius=10,
            border=ft.border.all(1, ft.Colors.GREY_600),
            margin=ft.margin.only(bottom=10)
        )
        
        # Main layout with ListView for rows
        main_layout = ft.Column([
            header,
            ft.Divider(),
            self.rows_container,
            ft.Container(height=20)  # Space at bottom
        ], expand=True, scroll=ft.ScrollMode.AUTO)
        
        # Add initial row
        self.add_row()
        
        # Add to page
        self.page.add(main_layout)
        
        # Add floating action button
        self.page.floating_action_button = self.floating_add_btn
        
        self.page.update()


# Add the utility functions
def get_excel_path():
    """
    Get Excel executable path from Windows registry.
    Returns None if not found.
    """
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID") as key:
            clsid = winreg.QueryValue(key, "")
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as key:
            excel_path = winreg.QueryValue(key, "").strip('"').split('"')[0]
            return excel_path if os.path.exists(excel_path) else None
    except Exception:
        return None

def open_path(path, select_in_folder=False):
    """
    Universal function to open files or folders.
    
    Args:
        path (str): Path to file or folder
        select_in_folder (bool): If True and path is a file, opens parent folder with file selected (Windows only)
    """
    try:
        if not os.path.exists(path):
            log_error(f"Path not found: {path}")
            return False
        
        system = platform.system()
        is_file = os.path.isfile(path)
        
        if system == 'Windows':
            if is_file and select_in_folder:
                # Open folder with file selected
                subprocess.Popen(['explorer', '/select,', path], shell=False)
            elif is_file:
                # Check if it's an Excel file
                file_ext = Path(path).suffix.lower()
                if file_ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
                    excel_path = get_excel_path()
                    if excel_path:
                        subprocess.Popen([excel_path, path], shell=False)
                        return True
                os.startfile(path)
            else:
                # Open folder
                subprocess.Popen(['explorer', path], shell=False)
                
        elif system == 'Darwin':
            if is_file and select_in_folder:
                subprocess.Popen(['open', '-R', path])
            else:
                subprocess.Popen(['open', path])
                
        else:  # Linux
            subprocess.Popen(['xdg-open', path])
        
        return True
        
    except Exception as ex:
        log_error(f"Error opening path '{path}': {ex}")
        return False
