import flet as ft
import json
import os
import sys
import re
from datetime import datetime
from tkinter import filedialog, messagebox
import traceback
import subprocess
import platform

# Import database utilities
try:
    from utils.db_utils import init_db as init_db_real, get_counter as get_counter_real, increment_counter as increment_counter_real, get_zoom_level, set_zoom_level
    
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

from utils.excel_utils import update_client_ledger
from utils.path_utils import resource_path

class InvoiceRow:
    """ كلاس صف الفاتورة (البند) """
    def __init__(self, page, row_index, product_dict, delete_callback, scale_factor=1.0):
        self.page = page
        self.row_index = row_index
        self.products = product_dict
        self.delete_callback = delete_callback
        self.scale_factor = scale_factor
        self.row_container = None  # Reference to the UI container
        self.base_length = 0.0  # Store the base length before discount
        
        # المتغيرات
        # Default widths
        self.default_widths = {
            'block': 100, 'thick': 120, 'mat': 120, 'count': 80, 
            'len': 80, 'height': 80, 'discount': 80, 'area': 100, 
            'price': 80, 'total': 100, 'product': 160
        }

        # المتغيرات
        self.block_var = ft.TextField(
            label="رقم البلوك", 
            width=self.default_widths['block'],
            on_change=self.on_block_change
        )
        self.thick_var = ft.Dropdown(
            label="السمك",
            options=[ft.dropdown.Option("2سم"), ft.dropdown.Option("3سم"), ft.dropdown.Option("4سم")],
            width=self.default_widths['thick'],
            on_change=self.update_price
        )
        self.mat_var = ft.TextField(label="الخامة", width=self.default_widths['mat'])
        self.count_var = ft.TextField(
            label="العدد", 
            width=self.default_widths['count'],
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.NumbersOnlyInputFilter()
        )
        self.len_var = ft.TextField(
            label="الطول", 
            width=self.default_widths['len'],
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9.]*$")
        )
        self.height_var = ft.TextField(
            label="الارتفاع", 
            width=self.default_widths['height'],
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9.]*$")
        )
        self.discount_var = ft.TextField(
            label="خصم", 
            width=self.default_widths['discount'], 
            value="0.0",
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9.]*$")
        )
        self.area_var = ft.TextField(label="المسطح", width=self.default_widths['area'], disabled=True)
        self.price_var = ft.TextField(
            label="السعر", 
            width=self.default_widths['price'],
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9.]*$")
        )
        self.total_var = ft.TextField(label="الإجمالي", width=self.default_widths['total'], disabled=True)
        
        # Product dropdown with prefixed options
        product_names = list(self.products.keys()) if self.products else []
        # Create options with "ش " prefix for display
        prefixed_options = [ft.dropdown.Option(name, "ش " + name) for name in product_names]
        self.product_dropdown = ft.Dropdown(
            label="البيان",
            options=prefixed_options,
            on_change=self.on_product_select,
            width=self.default_widths['product']
        )
        
        # Apply initial scale
        self.update_scale(self.scale_factor, update_page=False)
        
        # Delete button
        self.btn_del = ft.IconButton(
            icon=ft.Icons.DELETE,
            icon_color="red",
            on_click=self.destroy
        )
        
        # Bind events for calculations
        self.count_var.on_change = self.calculate
        self.height_var.on_change = self.update_price
        self.price_var.on_change = self.calculate
        
        # Special handlers for length and discount
        self.len_var.on_change = self.on_length_change
        self.discount_var.on_change = self.on_discount_change

    def on_block_change(self, e):
        val = self.block_var.value
        if val:
            # Replace Arabic characters with their English counterparts
            # 'ش' is 'a' on Arabic keyboard
            # 'لا' (lam-alif) is 'b' on Arabic keyboard
            new_val = val.replace('ش', 'A').replace('لا', 'B').replace('a', 'A').replace('b', 'B').replace('أ', 'A').replace('ب', 'B')
            if new_val != val:
                self.block_var.value = new_val
                self.page.update()

    def update_scale(self, scale_factor, update_page=True):
        self.scale_factor = scale_factor
        
        # Calculate new font size (default is usually around 14-16)
        new_text_size = 14 * scale_factor
        
        controls_map = {
            'block': self.block_var, 'thick': self.thick_var, 'mat': self.mat_var,
            'count': self.count_var, 'len': self.len_var, 'height': self.height_var,
            'discount': self.discount_var, 'area': self.area_var, 'price': self.price_var,
            'total': self.total_var, 'product': self.product_dropdown
        }
        
        for key, control in controls_map.items():
            control.width = self.default_widths[key] * scale_factor
            control.text_size = new_text_size
            if isinstance(control, ft.TextField):
                control.label_style = ft.TextStyle(size=new_text_size * 0.9)
            elif isinstance(control, ft.Dropdown):
                control.label_style = ft.TextStyle(size=new_text_size * 0.9)
                
        if update_page:
            self.page.update()

    def on_length_change(self, e):
        try:
            # If empty, just return and let the user edit
            if not self.len_var.value:
                self.calculate(e)
                return

            current_len = float(self.len_var.value) if self.len_var.value else 0
            current_disc = float(self.discount_var.value) if self.discount_var.value else 0
            
            # If this is the first time entering a length, set base_length
            # If discount was already entered, base_length should include it
            if self.base_length == 0:
                self.base_length = current_len + current_disc
            
            # If discount was entered first, update the displayed length now
            if current_disc > 0 and self.base_length > 0:
                new_net_len = self.base_length - current_disc
                if new_net_len < 0:
                    new_net_len = 0
                # Only update if the value is significantly different to avoid fighting the user
                if abs(float(self.len_var.value) - new_net_len) > 0.01:
                     # Format to remove trailing zeros if integer
                    self.len_var.value = f"{new_net_len:g}"
                    self.page.update()
        except ValueError:
            pass
        self.calculate(e)

    def on_discount_change(self, e):
        try:
            current_disc = float(self.discount_var.value) if self.discount_var.value else 0
            
            # Store the discount value regardless of whether length is entered
            # If length is already entered, apply the discount immediately
            if self.len_var.value and self.len_var.value.strip() != "":
                # Length is entered, apply discount now
                if self.base_length == 0:
                    # Initialize base_length if not already set
                    current_len = float(self.len_var.value) if self.len_var.value else 0
                    self.base_length = current_len + current_disc
                
                # Update displayed length: Net = Base - Discount
                if self.base_length > 0:
                    new_net_len = self.base_length - current_disc
                    if new_net_len < 0:
                        new_net_len = 0
                    # Format to remove trailing zeros if integer
                    self.len_var.value = f"{new_net_len:g}"
                    self.page.update()
            # If length is not entered yet, just store the discount value for later use
        except ValueError:
            pass
        self.calculate(e)

    def on_product_select(self, e):
        # Just call update_price which handles everything
        self.update_price(e)

    def update_price(self, e):
        choice = self.product_dropdown.value
        thick_val = self.thick_var.value
        
        if choice and choice in self.products:
            # Get thickness key (2, 3, or 4)
            thick_key = "2" # Default
            if thick_val:
                # Extract number from "2سم" -> "2"
                import re
                match = re.search(r'\d+', thick_val)
                if match:
                    thick_key = match.group(0)
            
            product_data = self.products[choice]
            price = 0
            
            # Handle new structure (dict of prices) vs old structure (single price)
            if isinstance(product_data, dict):
                price_data = product_data.get(thick_key, 0)
                
                # Check if price_data is a list (Ranges)
                if isinstance(price_data, list):
                    # Get current height
                    try:
                        current_height = float(self.height_var.value) if self.height_var.value else 0
                    except ValueError:
                        current_height = 0
                        
                    # Find matching range
                    found_price = 0
                    for range_item in price_data:
                        min_h = range_item.get('min', 0)
                        max_h = range_item.get('max', float('inf'))
                        if min_h <= current_height <= max_h:
                            found_price = range_item.get('price', 0)
                            break
                    
                    # If no range found, maybe use the last one or 0? 
                    # For now, if found_price is 0, let's try to use the first range as default if height is 0
                    if found_price == 0 and current_height == 0 and price_data:
                         found_price = price_data[0].get('price', 0)
                         
                    price = found_price
                else:
                    # Simple number
                    price = price_data
                
                # Fallback if specific thickness not found, try to find any price
                if price == 0 and product_data and not isinstance(price_data, list):
                     # Only fallback if we didn't find a specific thickness entry
                     # If we found an entry but it was 0, that's different.
                     # But here price_data being 0 means key wasn't found (get default)
                     if price_data == 0:
                        first_val = next(iter(product_data.values()))
                        if not isinstance(first_val, list):
                            price = first_val
            else:
                # Old structure or simple value
                price = product_data
                
            self.price_var.value = str(price)
            # Set the material field with original product name (without prefix)
            self.mat_var.value = choice
            self.calculate(None)
            self.page.update()

    def calculate(self, e):
        try:
            cnt_str = self.count_var.value or "0"
            cnt = float(cnt_str) if cnt_str else 0
            l = float(self.len_var.value) if self.len_var.value else 0
            h = float(self.height_var.value) if self.height_var.value else 0
            # Discount is already applied to 'l' (Length field)
            p = float(self.price_var.value) if self.price_var.value else 0

            area = cnt * l * h
            total = area * p

            # Format area with 2 decimal places (always show 2 decimals)
            self.area_var.value = f"{area:.2f}"
            # Format total with 2 decimal places (always show 2 decimals)
            self.total_var.value = f"{int(round(total))}"
        except ValueError:
            # If user input is invalid, just ignore calculation
            self.area_var.value = "0.00"
            self.total_var.value = "0"
        self.page.update()

    def get_controls(self):
        """Return Flet controls for this row in reversed order"""
        return [
            self.btn_del,
            self.product_dropdown,
            self.block_var,
            self.thick_var,
            self.mat_var,
            self.count_var,
            self.len_var,
            self.height_var,
            self.discount_var,
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
        self.page.title = "مصنع السويفي - ادارة الفواتير"
        self.page.rtl = True  # Right-to-left support for Arabic
        self.page.theme_mode = ft.ThemeMode.DARK  # Dark theme
        
        self.products_path = resource_path(os.path.join('res', 'products.json'))
        self.db_path = resource_path(os.path.join('res', 'invoice.db'))
        
        init_db(self.db_path)
        self.products = self.load_products()
        self.op_counter = get_counter(self.db_path)
        self.products = self.load_products()
        self.op_counter = get_counter(self.db_path)
        self.rows = []
        
        # Load saved zoom level from database
        self.scale_factor = get_zoom_level(self.db_path)
        
        # Form fields
        self.ent_op = ft.TextField(label="رقم العملية", value=str(self.op_counter))
        self.date_var = ft.TextField(label="التاريخ", value=datetime.now().strftime('%d/%m/%Y'))
        
        # Client selection with autocomplete suggestions
        self.client_suggestions = self.load_clients()
        self.ent_client = ft.TextField(
            label="اسم العميل",
            on_change=self.on_client_text_change
        )
        
        # Suggestions list container (hidden by default)
        self.suggestions_list = ft.Column(
            visible=False,
            spacing=0
        )

        self.ent_driver = ft.TextField(label="اسم السائق")
        self.ent_phone = ft.TextField(label="رقم التليفون")
        
        # Main container - no scroll to prevent conflicts
        self.rows_container = ft.Column()
        
        # Floating add button
        self.floating_add_btn = ft.FloatingActionButton(
            icon=ft.Icons.ADD,
            on_click=self.add_row
        )
        
        self.page.update()

    def load_clients(self):
        """Load existing client names from the 'الفواتير' directory"""
        # Use Documents/alswaife folder
        documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        
        self.invoices_root = os.path.join(documents_path, 'الفواتير')
        if not os.path.exists(self.invoices_root):
            try:
                os.makedirs(self.invoices_root)
            except OSError:
                pass
                
        clients = []
        if os.path.exists(self.invoices_root):
            for item in os.listdir(self.invoices_root):
                if os.path.isdir(os.path.join(self.invoices_root, item)):
                    clients.append(item)
        return sorted(clients)

    def on_client_text_change(self, e):
        """Show filtered suggestions when user types"""
        search_text = self.ent_client.value.strip().lower() if self.ent_client.value is not None else ""
        
        if not search_text:
            self.suggestions_list.visible = False
            self.suggestions_list.controls.clear()
            self.page.update()
            return
        
        # Filter suggestions
        filtered = [c for c in self.client_suggestions if search_text in c.lower()]
        
        if filtered:
            self.suggestions_list.controls.clear()
            for client in filtered[:5]:  # Show max 5 suggestions
                suggestion_btn = ft.TextButton(
                    text=client,
                    on_click=lambda e, c=client: self.select_suggestion(c),
                    style=ft.ButtonStyle(
                        padding=ft.padding.all(5),
                    )
                )
                self.suggestions_list.controls.append(suggestion_btn)
            self.suggestions_list.visible = True
        else:
            self.suggestions_list.visible = False
            self.suggestions_list.controls.clear()
        
        self.page.update()

    def select_suggestion(self, client_name):
        """Set the selected client name and hide suggestions"""
        self.ent_client.value = client_name
        self.suggestions_list.visible = False
        self.suggestions_list.controls.clear()
        self.page.update()




    def load_products(self):
        if not os.path.exists(self.products_path):
            return {}
        try:
            with open(self.products_path, 'r', encoding='utf-8') as f:
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
        except Exception as e:
            return {}

    def add_row(self, e=None):
        row_idx = len(self.rows)
        new_row = InvoiceRow(self.page, row_idx, self.products, self.delete_row, self.scale_factor)
        self.rows.append(new_row)
        
        # Create a row container WITHOUT individual scrolling
        row_controls = new_row.get_controls()
        row_container = ft.Row(
            controls=row_controls, 
            spacing=5
            # Removed scroll to prevent individual row scrolling
        )
        
        # Store reference to the row container for deletion in the InvoiceUI class
        if not hasattr(self, 'row_containers'):
            self.row_containers = {}
        self.row_containers[new_row] = row_container
        
        self.rows_container.controls.append(row_container)
        
        # Add spacing after each row to prevent scroll overlap
        spacer = ft.Container(height=20)
        self.rows_container.controls.append(spacer)
        # Store reference to the spacer as well
        if not hasattr(self, 'row_spacers'):
            self.row_spacers = {}
        self.row_spacers[new_row] = spacer
        
        self.page.update()
        
    def delete_row(self, row_obj):
        if row_obj in self.rows:
            # Remove the row container from the UI
            if hasattr(self, 'row_containers') and row_obj in self.row_containers:
                row_container = self.row_containers[row_obj]
                if row_container in self.rows_container.controls:
                    self.rows_container.controls.remove(row_container)
                # Clean up the reference
                del self.row_containers[row_obj]
            
            # Remove the spacer from the UI
            if hasattr(self, 'row_spacers') and row_obj in self.row_spacers:
                spacer = self.row_spacers[row_obj]
                if spacer in self.rows_container.controls:
                    self.rows_container.controls.remove(spacer)
                # Clean up the reference
                del self.row_spacers[row_obj]
            
            # Remove the row from the data structure
            self.rows.remove(row_obj)
            
            # Update UI
            self.page.update()

    def save_excel(self, e):
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

        items_data = []
        for row in self.rows:
            # Collect actual data from the row controls and format it as expected by excel_utils
            item_data = (
                row.product_dropdown.value or "",  # description
                row.block_var.value or "",         # block
                row.thick_var.value or "",         # thickness
                row.mat_var.value or "",           # material
                row.count_var.value or "0",        # count
                row.len_var.value or "0",          # length (already net)
                row.height_var.value or "0",       # height
                row.price_var.value or "0"         # price
            )
            items_data.append(item_data)

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

        fname = f"{sanitize(op_num)}_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        full_path = os.path.join(my_invoices_dir, fname)
        
        try:
            # Save the invoice
            self.save_callback(full_path, op_num, client, driver, date_str, phone, items_data)
            
            # Update ledger with aggregated data
            # Aggregate items by (description, material, thickness)
            aggregated_items = {}
            total_amount = 0
            
            for item in items_data:
                try:
                    # item structure: (desc, block, thick, mat, count, len, height, price)
                    desc = item[0] or ""
                    material = item[3] or ""
                    thickness = item[2] or ""
                    count = float(item[4])
                    length = float(item[5])
                    height = float(item[6])
                    price = float(item[7])
                    
                    area = count * length * height
                    item_total = area * price
                    total_amount += item_total
                    
                    # Create key for aggregation (material, thickness)
                    key = (desc, material, thickness)
                    
                    if key in aggregated_items:
                        # Add to existing entry
                        aggregated_items[key]['area'] += area
                        aggregated_items[key]['total'] += item_total
                    else:
                        # Create new entry
                        aggregated_items[key] = {
                            'desc': desc,
                            'material': material,
                            'thickness': thickness,
                            'area': area,
                            'price': price,  # Use first price encountered
                            'total': item_total
                        }
                except:
                    pass
            
            # Convert aggregated items to list format for ledger
            # Format: (desc, material, thickness, area, price_per_unit)
            ledger_items = []
            for key, data in aggregated_items.items():
                # Calculate average price per square meter
                price_per_sqm = data['total'] / data['area'] if data['area'] > 0 else 0
                ledger_items.append((
                    data['desc'],
                    data['material'],
                    data['thickness'],
                    data['area'],
                    price_per_sqm
                ))
            
            update_result = update_client_ledger(client_dir, client, date_str, op_num, total_amount, driver, ledger_items)
            
            # Check if ledger update succeeded
            if not update_result[0]:
                error_type = update_result[1]
                if error_type == "file_locked":
                    # Show warning dialog about file being open
                    def close_warning(e):
                        warning_dlg.open = False
                        self.page.update()
                    
                    warning_dlg = ft.AlertDialog(
                        title=ft.Text("⚠️ تنبيه", color="orange"),
                        content=ft.Text(
                            f"تم حفظ الفاتورة بنجاح ولكن لم يتم تحديث كشف الحساب.\n\n"
                            f"السبب: ملف كشف الحساب مفتوح حالياً.\n\n"
                            f"الحل: الرجاء إغلاق ملف:\n{client}.xlsx\n\n"
                            f"ثم حاول حفظ الفاتورة مرة أخرى.",
                            text_align=ft.TextAlign.RIGHT
                        ),
                        actions=[
                            ft.TextButton("حسناً", on_click=close_warning)
                        ],
                        actions_alignment=ft.MainAxisAlignment.END
                    )
                    self.page.overlay.append(warning_dlg)
                    warning_dlg.open = True
                    self.page.update()
                    return  # Don't show success dialog
            
            # Refresh client list if new client
            if client_safe not in self.client_suggestions:
                self.client_suggestions = self.load_clients()
            
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
                    print(f"Error opening file: {ex}")

            def open_folder(e):
                try:
                    folder_path = os.path.dirname(full_path)
                    if platform.system() == 'Windows':
                        # Use explorer to open folder in normal window state
                        subprocess.Popen(['explorer', folder_path], shell=False)
                    elif platform.system() == 'Darwin':
                        subprocess.call(('open', folder_path))
                    else:
                        subprocess.call(('xdg-open', folder_path))
                except Exception as ex:
                    print(f"Error opening folder: {ex}")

            def open_ledger(e):
                try:
                    ledger_path = os.path.join(client_dir, f"{client}.xlsx")
                    if os.path.exists(ledger_path):
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
                                subprocess.Popen([excel_path, '/e', ledger_path], shell=False)
                            except:
                                # Fallback to default method if registry lookup fails
                                os.startfile(ledger_path)
                        elif platform.system() == 'Darwin':
                            subprocess.call(('open', ledger_path))
                        else:
                            subprocess.call(('xdg-open', ledger_path))
                    else:
                        print(f"Ledger file not found: {ledger_path}")
                except Exception as ex:
                    print(f"Error opening ledger: {ex}")

            def close_dlg(e):
                dlg.open = False
                self.page.update()

            dlg = ft.AlertDialog(
                title=ft.Text("نجاح"),
                content=ft.Text(f"تم حفظ الفاتورة وتحديث كشف الحساب بنجاح.\nالمسار: {full_path}"),
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
            self.increment_op()
            
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

        items_data = []
        for row in self.rows:
            # Collect actual data from the row controls and format it as expected by excel_utils
            item_data = (
                row.product_dropdown.value or "",  # description
                row.block_var.value or "",         # block
                row.thick_var.value or "",         # thickness
                row.mat_var.value or "",           # material
                row.count_var.value or "0",        # count
                row.len_var.value or "0",          # length (already net)
                row.height_var.value or "0",       # height
                row.price_var.value or "0"         # price
            )
            items_data.append(item_data)

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

        fname = f"{sanitize(op_num)}_{now.strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        full_path = os.path.join(my_invoices_dir, fname)
        
        try:
            # Save the invoice
            self.save_callback(full_path, op_num, client, driver, date_str, phone, items_data)
            
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
                    print(f"Error opening file: {ex}")

            def open_folder(e):
                try:
                    folder_path = os.path.dirname(full_path)
                    if platform.system() == 'Windows':
                        # Use explorer to open folder in normal window state
                        subprocess.Popen(['explorer', folder_path], shell=False)
                    elif platform.system() == 'Darwin':
                        subprocess.call(('open', folder_path))
                    else:
                        subprocess.call(('xdg-open', folder_path))
                except Exception as ex:
                    print(f"Error opening folder: {ex}")

            def close_dlg(e):
                dlg.open = False
                self.page.update()

            dlg = ft.AlertDialog(
                title=ft.Text("نجاح"),
                content=ft.Text(f"تم حفظ فاتورة الإيراد بنجاح.\nالمسار: {full_path}"),
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
            self.increment_op()
            
        except Exception as ex:
            dlg = ft.AlertDialog(
                title=ft.Text("خطأ"),
                content=ft.Text(f"حدث خطأ أثناء الحفظ:\n{ex}\n{traceback.format_exc()}")
            )
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
        self.ent_client.value = ""
        self.ent_driver.value = ""
        self.ent_phone.value = ""
        
        # Clear all rows
        self.rows.clear()
        self.rows_container.controls.clear()
        
        # Add one empty row
        self.add_row()
        
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
                
                # If all else fails, print available attributes for debugging
                if not success:
                    print("Available page attributes:", [attr for attr in dir(self.page) if not attr.startswith('_')][:20])
        except Exception as ex:
            print(f"Error minimizing window: {ex}")
        self.page.update()

    def close_window(self, e):
        """Close the application window"""
        try:
             self.page.window.close()
        except Exception as ex:
            print(f"Error closing window: {ex}")

    def zoom_in(self, e):
        self.scale_factor += 0.1
        self.update_rows_scale()
        set_zoom_level(self.db_path, self.scale_factor)

    def zoom_out(self, e):
        if self.scale_factor > 0.5:
            self.scale_factor -= 0.1
            self.update_rows_scale()
            set_zoom_level(self.db_path, self.scale_factor)

    def update_rows_scale(self):
        for row in self.rows:
            row.update_scale(self.scale_factor)
        self.page.update()
        
    def build_ui(self):
        # Create AppBar with menu
        self.page.appbar = ft.AppBar(
            title=ft.Text("مصنع السويفي - ادارة الفواتير"),
            bgcolor=ft.Colors.SURFACE,
            actions=[
                ft.IconButton(ft.Icons.ZOOM_IN, on_click=self.zoom_in, tooltip="تكبير"),
                ft.IconButton(ft.Icons.ZOOM_OUT, on_click=self.zoom_out, tooltip="تصغير"),
                ft.PopupMenuButton(
                    items=[
                        ft.PopupMenuItem(text="حفظ إلى Excel", on_click=self.save_excel),
                        ft.PopupMenuItem(text="عملية جديدة", on_click=self.reset_form),
                        ft.PopupMenuItem(text="تصغير", on_click=self.minimize_window),
                        ft.PopupMenuItem(text="إغلاق", on_click=self.close_window)
                    ]
                )
            ]
        )
        
        # Header section - changed to horizontal layout
        header = ft.Row([
            ft.Column([
                ft.Text("بيانات الفاتورة", size=18, weight=ft.FontWeight.BOLD),
                self.ent_op,
            ]),
            ft.Column([
                self.date_var,
                ft.Column([
                    self.ent_client,
                    self.suggestions_list,
                ], spacing=0),
            ]),

            ft.Column([
                self.ent_driver,
                self.ent_phone,
            ]),
        ], spacing=20)
        
        # Wrap rows container in a horizontally scrollable container
        rows_wrapper = ft.Row(
            controls=[self.rows_container],
            scroll=ft.ScrollMode.ALWAYS,  # Always show horizontal scrollbar
            expand=True
        )
        
        # Main layout with reduced bottom spacing since we're adding spacing in rows container
        main_layout = ft.Column([
            header,
            ft.Divider(),
            rows_wrapper,
            ft.Container(height=20)  # Reduced space at bottom
        ], expand=True, scroll=ft.ScrollMode.AUTO)
        
        # Add initial row
        self.add_row()
        
        # Add to page
        self.page.add(main_layout)
        
        # Add floating action button
        self.page.floating_action_button = self.floating_add_btn
        
        self.page.update()

