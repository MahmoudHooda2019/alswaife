import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime, timedelta
import os
import json
import re

# استيراد دوال قاعدة البيانات
try:
    from db_utils import init_db as init_db_real, get_counter as get_counter_real, increment_counter as increment_counter_real
    
    # Re-export with proper type annotations
    init_db = init_db_real
    get_counter = get_counter_real
    increment_counter = increment_counter_real
except ImportError:
    def init_db(db_path: str) -> None: pass
    def get_counter(db_path: str, key: str = "invoice") -> int: return 1
    def increment_counter(db_path: str, key: str = "invoice") -> int: return 1

class InvoiceRow:
    """ كلاس صف الفاتورة (البند) """
    def __init__(self, parent_frame, row_index, product_dict, delete_callback):
        self.parent = parent_frame
        self.row_index = row_index
        self.products = product_dict
        self.delete_callback = delete_callback
        
        self.entry_font = ("Arial", 13)
        
        # المتغيرات
        self.block_var = tk.StringVar()
        self.thick_var = tk.StringVar()
        self.mat_var = tk.StringVar()
        self.count_var = tk.StringVar()
        self.len_var = tk.StringVar()
        self.height_var = tk.StringVar()
        self.area_var = tk.StringVar()
        self.price_var = tk.StringVar()
        self.total_var = tk.StringVar()

        # إعدادات الشبكة (Grid)
        pad_opts = {'padx': 1, 'pady': 2, 'sticky': "nsew"}
        
        # 0:Total, 1:Price, 2:Area, 3:Height, 4:Len, 5:Count, 6:Mat, 7:Thick, 8:Block, 9:Desc, 10:Del

        self.ent_total = ctk.CTkEntry(parent_frame, textvariable=self.total_var, state='disabled', font=self.entry_font, justify="center")
        self.ent_total.grid(row=row_index, column=0, **pad_opts)

        self.ent_price = ctk.CTkEntry(parent_frame, textvariable=self.price_var, font=self.entry_font, justify="center")
        self.ent_price.grid(row=row_index, column=1, **pad_opts)
        
        self.ent_area = ctk.CTkEntry(parent_frame, textvariable=self.area_var, state='disabled', font=self.entry_font, justify="center")
        self.ent_area.grid(row=row_index, column=2, **pad_opts)

        self.ent_height = ctk.CTkEntry(parent_frame, textvariable=self.height_var, font=self.entry_font, justify="center")
        self.ent_height.grid(row=row_index, column=3, **pad_opts)

        self.ent_len = ctk.CTkEntry(parent_frame, textvariable=self.len_var, font=self.entry_font, justify="center")
        self.ent_len.grid(row=row_index, column=4, **pad_opts)

        self.ent_count = ctk.CTkEntry(parent_frame, textvariable=self.count_var, font=self.entry_font, justify="center")
        self.ent_count.grid(row=row_index, column=5, **pad_opts)

        self.ent_mat = ctk.CTkEntry(parent_frame, textvariable=self.mat_var, font=self.entry_font, justify="center")
        self.ent_mat.grid(row=row_index, column=6, **pad_opts)

        # القائمة المنسدلة للسمك
        self.ent_thick = ctk.CTkComboBox(
            parent_frame, 
            values=["2سم", "3سم", "4سم"], 
            variable=self.thick_var,
            font=self.entry_font,
            justify="center"
        )
        self.ent_thick.grid(row=row_index, column=7, **pad_opts)

        self.ent_block = ctk.CTkEntry(parent_frame, textvariable=self.block_var, font=self.entry_font, justify="center")
        self.ent_block.grid(row=row_index, column=8, **pad_opts)

        product_names = list(self.products.keys())
        self.combo_desc = ctk.CTkComboBox(
            parent_frame, 
            values=product_names, 
            command=self.on_product_select,
            justify="right",
            font=self.entry_font
        )
        self.combo_desc.set("") 
        self.combo_desc.grid(row=row_index, column=9, **pad_opts)

        self.btn_del = ctk.CTkButton(
            parent_frame, text="×", width=25, fg_color="#ff4444", hover_color="#cc0000",
            command=self.destroy
        )
        self.btn_del.grid(row=row_index, column=10, **pad_opts)

        for entry in [self.ent_count, self.ent_len, self.ent_height, self.ent_price]:
            entry.bind("<KeyRelease>", self.calculate)

    def on_product_select(self, choice):
        if choice in self.products:
            self.price_var.set(str(self.products[choice]))
            self.mat_var.set(choice)
            self.calculate()

    def calculate(self, event=None):
        try:
            cnt = float(self.count_var.get()) if self.count_var.get() else 0
            l = float(self.len_var.get()) if self.len_var.get() else 0
            h = float(self.height_var.get()) if self.height_var.get() else 0
            p = float(self.price_var.get()) if self.price_var.get() else 0
            
            area = cnt * l * h
            total = area * p

            self.area_var.set(f"{area:.3f}")
            self.total_var.set(f"{total:.2f}")
        except ValueError:
            pass

    def get_data(self):
        desc = self.combo_desc.get().strip()
        block = self.block_var.get().strip()
        if not desc and not block:
            return None
        return (desc, block, self.thick_var.get().strip(), self.mat_var.get().strip(),
                self.count_var.get().strip(), self.len_var.get().strip(), 
                self.height_var.get().strip(), self.price_var.get().strip())

    def destroy(self):
        self.ent_total.destroy()
        self.ent_price.destroy()
        self.ent_area.destroy()
        self.ent_height.destroy()
        self.ent_len.destroy()
        self.ent_count.destroy()
        self.ent_mat.destroy()
        self.ent_thick.destroy()
        self.ent_block.destroy()
        self.combo_desc.destroy()
        self.btn_del.destroy()
        self.delete_callback(self)


class InvoiceUI(ctk.CTk):
    def __init__(self, save_callback):
        super().__init__()
        self.save_callback = save_callback
        
        self.title("Invoice Creator")
        self.geometry("1150x700")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.base_dir = os.path.dirname(__file__)
        self.products_path = os.path.join(self.base_dir, 'res', 'products.json')
        self.db_path = os.path.join(self.base_dir, 'res', 'invoice.db')
        
        init_db(self.db_path)
        self.products = self.load_products()
        self.op_counter = get_counter(self.db_path)
        self.rows = [] 

        self.grid_rowconfigure(1, weight=1) 
        self.grid_columnconfigure(0, weight=1)

        self.create_header()
        self.create_items_area()
        self.create_footer()

    def create_header(self):
        """
        إنشاء قسم البيانات باستخدام Grid موحدة.
        """
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=15, pady=10, sticky="ew")
        
        # تخصيص أوزان الأعمدة (Weights) لإعطاء مساحات منطقية
        # 0:Phone, 1:Driver, 2:Client(كبير), 3:Date, 4:Op(صغير)
        header_frame.columnconfigure(0, weight=2) # Phone
        header_frame.columnconfigure(1, weight=2) # Driver
        header_frame.columnconfigure(2, weight=3) # Client (أكبر)
        header_frame.columnconfigure(3, weight=2) # Date
        header_frame.columnconfigure(4, weight=1) # Operation Num

        lbl_font = ("Arial", 14, "bold")
        ent_font = ("Arial", 14)

        # --- الصف 0: العناوين ---
        # sticky="ew" + anchor="center" يجبر النص على التوسط في كامل عرض العمود
        
        ctk.CTkLabel(header_frame, text="رقم العملية", font=lbl_font, anchor="center").grid(row=0, column=4, sticky="ew", pady=(0, 2))
        ctk.CTkLabel(header_frame, text="التاريخ", font=lbl_font, anchor="center").grid(row=0, column=3, sticky="ew", pady=(0, 2))
        ctk.CTkLabel(header_frame, text="اسم العميل", font=lbl_font, anchor="center").grid(row=0, column=2, sticky="ew", pady=(0, 2))
        ctk.CTkLabel(header_frame, text="اسم السائق", font=lbl_font, anchor="center").grid(row=0, column=1, sticky="ew", pady=(0, 2))
        ctk.CTkLabel(header_frame, text="رقم التليفون", font=lbl_font, anchor="center").grid(row=0, column=0, sticky="ew", pady=(0, 2))

        # --- الصف 1: الخانات ---
        
        # 1. رقم العملية
        self.ent_op = ctk.CTkEntry(header_frame, justify="center", font=ent_font)
        self.ent_op.grid(row=1, column=4, padx=5, sticky="ew")
        self.ent_op.insert(0, str(self.op_counter))

        # 2. التاريخ
        self.date_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        
        # Date entry field with calendar tooltip
        self.ent_date = ctk.CTkEntry(
            header_frame,
            textvariable=self.date_var,
            font=ent_font,
            justify="center"
        )
        self.ent_date.grid(row=1, column=3, padx=5, sticky="ew")
        
        # Create tooltip window (initially hidden)
        self.cal_tooltip = None
        
        # Bind click event to show calendar tooltip
        self.ent_date.bind('<Button-1>', self.toggle_calendar_tooltip)
        
        # Also bind to focus in case user tabs to the field
        self.ent_date.bind('<FocusIn>', self.show_calendar_tooltip)

        # 3. اسم العميل
        self.ent_client = ctk.CTkEntry(header_frame, justify="right", font=ent_font)
        self.ent_client.grid(row=1, column=2, padx=5, sticky="ew")

        # 4. اسم السائق
        self.ent_driver = ctk.CTkEntry(header_frame, justify="right", font=ent_font)
        self.ent_driver.grid(row=1, column=1, padx=5, sticky="ew")

        # 5. رقم التليفون
        self.ent_phone = ctk.CTkEntry(header_frame, justify="right", font=ent_font)
        self.ent_phone.grid(row=1, column=0, padx=5, sticky="ew")

    def create_items_area(self):
        self.list_container = ctk.CTkFrame(self)
        self.list_container.grid(row=1, column=0, padx=15, pady=5, sticky="nsew")
        
        # 0:Total, 1:Price, 2:Area, 3:Height, 4:Len, 5:Count, 6:Mat, 7:Thick, 8:Block, 9:Desc, 10:Del
        self.cols_config = {
            0: 1, 1: 1, 2: 1, 3: 1, 4: 1, 
            5: 1, 6: 2, 7: 1, 8: 1, 9: 4, 10: 0
        }
        
        for c, w in self.cols_config.items():
            self.list_container.columnconfigure(c, weight=w)
            if c == 10:
                self.list_container.columnconfigure(c, minsize=35)

        # عناوين الجدول
        headers = [
            (10, ""), (9, "البيان"), (8, "رقم البلوك"), (7, "السمك"), 
            (6, "الخامة"), (5, "العدد"), (4, "الطول"), (3, "الارتفاع"), 
            (2, "بالمتر"), (1, "السعر"), (0, "الإجمالي")
        ]
        
        header_font = ("Arial", 13, "bold")
        for col_idx, text in headers:
            lbl = ctk.CTkLabel(self.list_container, text=text, font=header_font, anchor="center")
            lbl.grid(row=0, column=col_idx, padx=1, pady=5, sticky="ew")

        self.scroll_frame = ctk.CTkScrollableFrame(self.list_container)
        self.scroll_frame.grid(row=1, column=0, columnspan=11, sticky="nsew")
        self.list_container.rowconfigure(1, weight=1)
        
        for c, w in self.cols_config.items():
            self.scroll_frame.columnconfigure(c, weight=w)
            if c == 10:
                self.scroll_frame.columnconfigure(c, minsize=35)

        self.add_row()

    def create_footer(self):
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")

        self.btn_add = ctk.CTkButton(
            footer_frame, text="+ إضافة بند", font=("Arial", 16), 
            command=self.add_row
        )
        self.btn_add.pack(side="right", padx=10)

        self.btn_save = ctk.CTkButton(
            footer_frame, text="حفظ إلى Excel", font=("Arial", 16),
            fg_color="#2ecc71", hover_color="#27ae60",
            command=self.save_excel
        )
        self.btn_save.pack(side="left", padx=10)
        
        self.btn_new = ctk.CTkButton(
            footer_frame, text="عملية جديدة", font=("Arial", 16),
            fg_color="#e67e22", hover_color="#d35400",
            command=self.reset_form
        )
        self.btn_new.pack(side="left", padx=10)

    def load_products(self):
        if not os.path.exists(self.products_path):
            return {}
        try:
            with open(self.products_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                products = {}
                if isinstance(data, dict):
                    products = {str(k): v for k, v in data.items()}
                elif isinstance(data, list):
                    for item in data:
                        if 'name' in item:
                            products[str(item['name'])] = item.get('price', 0)
                return products
        except Exception as e:
            return {}

    def add_row(self):
        row_idx = len(self.rows)
        new_row = InvoiceRow(self.scroll_frame, row_idx, self.products, self.delete_row)
        self.rows.append(new_row)

    def delete_row(self, row_obj):
        if row_obj in self.rows:
            self.rows.remove(row_obj)

    def toggle_calendar_tooltip(self, event=None):
        """Toggle the calendar tooltip visibility"""
        if self.cal_tooltip and self.cal_tooltip.winfo_exists():
            self.hide_calendar_tooltip()
        else:
            self.show_calendar_tooltip()
    
    def show_calendar_tooltip(self, event=None):
        """Show the calendar tooltip"""
        try:
            from tkcalendar import Calendar
            
            # Prevent multiple tooltips
            if hasattr(self, '_tooltip_showing') and self._tooltip_showing:
                return
                
            self._tooltip_showing = True
            
            # Hide any existing tooltip
            if self.cal_tooltip and self.cal_tooltip.winfo_exists():
                self.hide_calendar_tooltip()
                return
            
            # Create tooltip window
            self.cal_tooltip = ctk.CTkToplevel(self)
            self.cal_tooltip.overrideredirect(True)
            self.cal_tooltip.attributes('-topmost', True)
            self.cal_tooltip.wm_attributes('-topmost', True)
            
            # Add calendar
            cal_frame = ctk.CTkFrame(self.cal_tooltip)
            cal_frame.pack(padx=1, pady=1)
            
            cal = Calendar(
                cal_frame,
                locale='ar_SA',
                date_pattern='dd/MM/yyyy',
                font=("Arial", 10),
                selectmode='day',
                showweeknumbers=False,
                firstweekday='sunday',
                showothermonthdays=False
            )
            
            # Set current date
            try:
                current_date = datetime.strptime(self.date_var.get(), '%d/%m/%Y')
                cal.selection_set(current_date)
            except:
                pass
                
            cal.pack(padx=1, pady=1)
            
            # Position the tooltip below the date field
            x = self.ent_date.winfo_rootx()
            y = self.ent_date.winfo_rooty() + self.ent_date.winfo_height()
            self.cal_tooltip.geometry(f'+{x}+{y}')
            
            # Handle window close on outside click
            def on_click(event):
                x, y = self.winfo_pointerxy()
                widget = self.winfo_containing(x, y)
                if widget not in [cal, cal_frame, self.ent_date]:
                    self.hide_calendar_tooltip()
            
            # Bind events
            self.cal_tooltip.bind('<ButtonPress>', on_click, add='+')
            self.cal_tooltip.bind('<FocusOut>', lambda e: None)  # Block focus out
            
            # Bind date selection
            def on_date_selected(e):
                selected_date = cal.get_date()
                # Safely convert to desired date format
                try:
                    # Try to use strftime (works for datetime/date objects)
                    if not isinstance(selected_date, str):
                        formatted_date = selected_date.strftime('%d/%m/%Y')
                    else:
                        raise AttributeError("strftime not available for str")
                except AttributeError:
                    # If strftime fails, it's likely a string
                    try:
                        # Try to parse as a date string and reformat
                        date_obj = datetime.strptime(str(selected_date), '%Y-%m-%d')
                        formatted_date = date_obj.strftime('%d/%m/%Y')
                    except ValueError:
                        # If parsing fails, use as-is
                        formatted_date = str(selected_date)
                self.date_var.set(formatted_date)
                self.after(100, self.hide_calendar_tooltip)
                
            cal.bind('<<CalendarSelected>>', on_date_selected, add='+')
            
            # Focus the calendar
            self.cal_tooltip.focus_force()
            
        except ImportError:
            self.show_fallback_date_entry()
        except Exception as e:
            print(f"Error showing calendar: {e}")
            self._tooltip_showing = False
    
    def hide_calendar_tooltip(self, event=None):
        """Hide the calendar tooltip"""
        if hasattr(self, 'cal_tooltip') and self.cal_tooltip and self.cal_tooltip.winfo_exists():
            self.cal_tooltip.destroy()
            self.cal_tooltip = None
        if hasattr(self, '_tooltip_showing'):
            self._tooltip_showing = False
    
    def show_fallback_date_entry(self):
        """Fallback to simple date entry if tkcalendar is not available"""
        from tkinter import simpledialog
        current_date = self.ent_date.get()
        new_date = simpledialog.askstring(
            "إدخال التاريخ", 
            "أدخل التاريخ (DD/MM/YYYY):", 
            initialvalue=current_date
        )
        if new_date:
            self.date_var.set(new_date)
    
    def get_date_str(self):
        """Get the selected date string"""
        try:
            return self.date_var.get() or datetime.now().strftime('%d/%m/%Y')
        except Exception as e:
            print(f"Error getting date: {e}")
            return datetime.now().strftime('%d/%m/%Y')

    def save_excel(self):
        op_num = self.ent_op.get().strip()
        client = self.ent_client.get().strip()
        date_str = self.get_date_str()
        driver = self.ent_driver.get().strip()
        phone = self.ent_phone.get().strip()

        if not op_num:
            messagebox.showerror("خطأ", "يرجى إدخال رقم العملية")
            return

        items_data = []
        for row in self.rows:
            data = row.get_data()
            if data:
                items_data.append(data)

        if not items_data:
            messagebox.showwarning("تنبيه", "لا توجد بنود للحفظ")
            return

        def sanitize(s):
            return re.sub(r'[\\/*?:"<>|]', "", str(s)).replace(" ", "_")
        
        now = datetime.now()
        fname = f"{sanitize(op_num)}_{sanitize(client)}_{now.day}_{now.month}_{now.year}.xlsx"
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=fname,
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            try:
                self.save_callback(save_path, op_num, client, driver, date_str, phone, items_data)
                messagebox.showinfo("نجاح", f"تم حفظ الملف بنجاح:\n{save_path}")
                self.increment_op()
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء الحفظ:\n{e}")

    def increment_op(self):
        try:
            new_val = increment_counter(self.db_path)
            self.op_counter = new_val
        except:
            self.op_counter += 1
        
        self.ent_op.delete(0, tk.END)
        self.ent_op.insert(0, str(self.op_counter))

    def reset_form(self):
        self.ent_client.delete(0, tk.END)
        self.ent_driver.delete(0, tk.END)
        self.ent_phone.delete(0, tk.END)
        for row in list(self.rows):
            row.destroy()
        self.add_row()