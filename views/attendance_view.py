"""
Attendance View - UI for employee attendance tracking
"""

import flet as ft
import os
from datetime import datetime
import platform
import subprocess
from utils.attendance_utils import create_or_update_attendance, load_attendance_data
from tkinter import filedialog
import tkinter as tk
from utils.path_utils import resource_path
import json
from typing import Optional


class AttendanceView:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "الحضور والانصراف"
        self.page.rtl = True
        
        # Data storage
        self.employees_list = self.load_employees()
        self.attendance_data = {}  # Dictionary to store attendance status
        self.current_file = None
        
        # UI components
        self.date_field: Optional[ft.TextField] = None
        self.day_field: Optional[ft.TextField] = None
        self.shift_dropdown: Optional[ft.Dropdown] = None
        self.employees_container: Optional[ft.Column] = None
        
    def load_employees(self):
        """Load employees from JSON file"""
        try:
            employees_path = resource_path(os.path.join('res', 'employees.json'))
            if os.path.exists(employees_path):
                with open(employees_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data if isinstance(data, list) else []
        except:
            pass
        return []
        
    def build_ui(self):
        """Build the attendance tracking UI"""
        
        # Title and back button
        title_row = ft.Row(
            controls=[
                ft.IconButton(
                    icon=ft.Icons.ARROW_BACK,
                    on_click=self.go_back,
                    tooltip="العودة"
                ),
                ft.Text(
                    "الحضور والانصراف",
                    size=30,
                    weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_200
                ),
            ],
            alignment=ft.MainAxisAlignment.START
        )
        
        # Date selection
        self.date_field = ft.TextField(
            label="التاريخ",
            value=datetime.now().strftime('%d/%m/%Y'),
            width=150,
            text_align=ft.TextAlign.CENTER,
            on_change=self.on_date_change
        )
        
        # Day field (automatically populated based on date)
        self.day_field = ft.TextField(
            label="اليوم",
            width=150,
            text_align=ft.TextAlign.CENTER,
            disabled=True
        )
        
        # Populate day field based on current date
        self.update_day_field(self.date_field.value)
        
        self.shift_dropdown = ft.Dropdown(
            label="الوردية",
            options=[
                ft.dropdown.Option("الاولي", "الاولي"),
                ft.dropdown.Option("الثانية", "الثانية")
            ],
            width=150,
            on_change=self.on_shift_change
        )
        
        date_info_row = ft.Row(
            controls=[
                self.date_field,
                self.day_field,
                self.shift_dropdown
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=20
        )
        
        # Employees container
        self.employees_container = ft.Column(
            controls=[],
            spacing=10
        )
        
        # Load existing data for current date if available
        self.load_existing_data()
        
        # Floating buttons will be created in the Stack layout
        
        # Main layout
        main_column = ft.Column(
            controls=[
                title_row,
                ft.Divider(),
                date_info_row,
                ft.Container(height=20),
                self.employees_container
            ],
            spacing=10,
            scroll=ft.ScrollMode.AUTO,
            expand=True
        )
        
        # Create main layout with floating buttons
        self.page.add(
            ft.Stack(
                controls=[
                    main_column,
                    # Floating buttons container
                    ft.Container(
                        content=ft.Column(
                            controls=[
                                # Add employee button
                                ft.Container(
                                    content=ft.FloatingActionButton(
                                        icon=ft.Icons.PERSON_ADD,
                                        on_click=self.add_new_employee,
                                        bgcolor=ft.Colors.BLUE_600,
                                        foreground_color=ft.Colors.WHITE,
                                        tooltip="إضافة موظف جديد",
                                        shape=ft.RoundedRectangleBorder(radius=20),
                                        elevation=6
                                    ),
                                    margin=ft.margin.only(bottom=15),
                                    shadow=ft.BoxShadow(
                                        spread_radius=1,
                                        blur_radius=8,
                                        color=ft.Colors.with_opacity(0.3, ft.Colors.BLUE_600),
                                        offset=ft.Offset(0, 2)
                                    )
                                ),
                                # Save button
                                ft.Container(
                                    content=ft.FloatingActionButton(
                                        icon=ft.Icons.SAVE,
                                        on_click=self.save_to_excel,
                                        bgcolor=ft.Colors.GREEN_600,
                                        foreground_color=ft.Colors.WHITE,
                                        tooltip="حفظ البيانات",
                                        shape=ft.RoundedRectangleBorder(radius=20),
                                        elevation=6
                                    ),
                                    shadow=ft.BoxShadow(
                                        spread_radius=1,
                                        blur_radius=10,
                                        color=ft.Colors.with_opacity(0.3, ft.Colors.GREEN_600),
                                        offset=ft.Offset(0, 3)
                                    )
                                ),
                            ],
                            spacing=0,
                            alignment=ft.MainAxisAlignment.END
                        ),
                        right=24,
                        bottom=24,
                        alignment=ft.alignment.bottom_right
                    )
                ],
                expand=True
            )
        )
        
        self.page.update()
    
    def on_date_change(self, e):
        """Handle date change"""
        if self.date_field is not None:
            date_str = self.date_field.value
            if date_str:
                self.update_day_field(date_str)
                # Load existing data for the selected date
                self.load_existing_data()
                self.page.update()
    
    def on_shift_change(self, e):
        """Handle shift change"""
        # Reload data when shift changes
        self.load_existing_data()
        self.page.update()
    
    def update_day_field(self, date_str):
        """Update day field based on date"""
        if self.day_field is None:
            return
            
        try:
            # Parse the date string
            date_obj = datetime.strptime(date_str, '%d/%m/%Y')
            # Get day name in Arabic
            arabic_days = ['الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت', 'الأحد']
            day_name = arabic_days[date_obj.weekday()]
            self.day_field.value = day_name
        except:
            self.day_field.value = ""
    
    def load_existing_data(self):
        """Load existing attendance data for current date and shift"""
        # Clear existing controls
        if self.employees_container is not None:
            self.employees_container.controls.clear()
        
        # Check if both date and shift are selected
        if not self.date_field or not self.date_field.value or not self.shift_dropdown or not self.shift_dropdown.value:
            # Show message to select date and shift first
            if self.employees_container is not None:
                self.employees_container.controls.append(
                    ft.Container(
                        content=ft.Text(
                            "الرجاء اختيار التاريخ والوردية أولاً لعرض قائمة الموظفين",
                            size=18,
                            text_align=ft.TextAlign.CENTER,
                            color=ft.Colors.GREY_600
                        ),
                        alignment=ft.alignment.center,
                        padding=50
                    )
                )
            self.page.update()
            return
        
        # Ensure the directory exists
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        alswaife_path = os.path.join(documents_path, "alswaife")
        attendance_path = os.path.join(alswaife_path, "حضور وانصراف")
        
        # Use single attendance file
        try:
            filename = "attendance.xlsx"
            filepath = os.path.join(attendance_path, filename)
            
            if os.path.exists(filepath):
                # Load from Excel
                success, data, error = load_attendance_data(filepath)
                
                if success and data:
                    self.current_file = filepath
                    
                    # Filter data for the current date
                    filtered_data = self.filter_data_for_date_and_shift(data)
                    
                    # Create a dictionary for quick lookup
                    employee_data = {emp['name']: emp for emp in filtered_data}
                    
                    # Create employee rows with existing data
                    for emp in self.employees_list:
                        emp_name = emp['name']
                        price = emp.get('price', 0)
                        
                        # Check if employee has attendance data for current date/shift
                        is_present = False
                        emp_price = price  # Default to JSON price
                        if emp_name in employee_data:
                            emp_record = employee_data[emp_name]
                            shift_key = self.get_shift_key()
                            if shift_key and emp_record.get(shift_key, 0) > 0:
                                is_present = True
                            # Use price from existing record if available
                            if 'price' in emp_record and emp_record['price'] != 0:
                                emp_price = emp_record['price']
                        
                        self.add_employee_row(emp_name, emp_price, is_present)
                    
                    # Load any additional employees from Excel that are not in JSON
                    for emp_record in filtered_data:
                        emp_name = emp_record['name']
                        # Check if this employee is not in the JSON list
                        if not any(emp['name'] == emp_name for emp in self.employees_list):
                            price = emp_record.get('price', 0)
                            shift_key = self.get_shift_key()
                            is_present = shift_key and emp_record.get(shift_key, 0) > 0
                            self.add_employee_row(emp_name, price, is_present)
                else:
                    # Create employee rows without existing data
                    for emp in self.employees_list:
                        emp_name = emp['name']
                        price = emp.get('price', 0)
                        self.add_employee_row(emp_name, price, False)
            else:
                # Create employee rows without existing data
                for emp in self.employees_list:
                    emp_name = emp['name']
                    price = emp.get('price', 0)
                    self.add_employee_row(emp_name, price, False)
        except Exception as e:
            print(f"Error loading existing data: {e}")
            # Create employee rows without existing data
            for emp in self.employees_list:
                emp_name = emp['name']
                price = emp.get('price', 0)
                self.add_employee_row(emp_name, price, False)
        
        self.page.update()
    
    def filter_data_for_date_and_shift(self, all_data):
        """Filter data to show only records for the current date"""
        if not self.date_field or not self.date_field.value:
            return []
        
        current_date = self.date_field.value
        filtered_data = []
        
        for emp_record in all_data:
            # Check if the record matches the current date
            # Handle potential date format differences
            record_date = emp_record.get('date', '')
            if record_date and str(record_date) == str(current_date):
                filtered_data.append(emp_record)
        
        return filtered_data
    
    def get_shift_key(self):
        """Get the shift key based on current day and shift selection"""
        if self.day_field is None or self.shift_dropdown is None:
            return None
            
        if not self.day_field.value or not self.shift_dropdown.value:
            return None
            
        # Map Arabic day names to English keys
        day_mapping = {
            'الاثنين': 'monday',
            'الثلاثاء': 'tuesday',
            'الأربعاء': 'wednesday',
            'الخميس': 'thursday',
            'الجمعة': 'friday',
            'السبت': 'saturday',
            'الأحد': 'sunday'
        }
        
        shift_mapping = {
            'الاولي': 'shift1',
            'الثانية': 'shift2'
        }
        
        day_key = day_mapping.get(self.day_field.value, '')
        shift_key = shift_mapping.get(self.shift_dropdown.value, '')
        
        if day_key and shift_key:
            return f"{day_key}_{shift_key}"
        return None
    
    def add_employee_row(self, name, price, is_present=False):
        """Add an employee row to the UI with improved layout"""
        # Determine if controls should be enabled (only when a shift is selected)
        shift_selected = bool(self.shift_dropdown and self.shift_dropdown.value)
        
        # Create price field first
        price_field = ft.TextField(
            value=str(price),
            width=120,
            text_align=ft.TextAlign.CENTER,
            keyboard_type=ft.KeyboardType.NUMBER,
            label="السعر",
            dense=True,
            disabled=not shift_selected,
        )
        
        # Check if this employee is from JSON or added manually
        is_json_employee = any(emp['name'] == name for emp in self.employees_list)
        
        # Create delete button for manually added employees
        delete_button = None
        if not is_json_employee:
            delete_button = ft.IconButton(
                icon=ft.Icons.DELETE,
                icon_color=ft.Colors.RED_400,
                tooltip="حذف الموظف",
                on_click=lambda e, n=name: self.delete_employee(n)
            )
        
        # Create row controls
        row_controls = [
            # Checkbox for attendance
            ft.Checkbox(
                value=is_present if shift_selected else False,
                disabled=not shift_selected,
                on_change=lambda e, n=name: self.on_attendance_change(n, e.control.value)
            ),
            # Employee name
            ft.Text(name, size=18, weight=ft.FontWeight.BOLD),
            # Spacer
            ft.Container(expand=True),
            # Price field
            price_field
        ]
        
        # Add delete button if not from JSON
        if delete_button:
            row_controls.append(delete_button)
        
        # Create card container for better visual appearance
        card = ft.Card(
            content=ft.Container(
                content=ft.Row(
                    controls=row_controls,
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                padding=15,
                border_radius=12,
            ),
            elevation=3,
            shadow_color=ft.Colors.with_opacity(0.2, ft.Colors.GREY_400),
            shape=ft.RoundedRectangleBorder(radius=12),
            margin=ft.margin.symmetric(vertical=4, horizontal=8),
            color=ft.Colors.GREY_800
        )
        
        # Store attendance status and field references
        self.attendance_data[name] = {
            'card': card,
            'price_field': price_field,
            'present': is_present if shift_selected else False,
            'is_json_employee': is_json_employee
        }
        
        if self.employees_container is not None:
            self.employees_container.controls.append(card)
    
    def on_attendance_change(self, employee_name, is_present):
        """Handle attendance checkbox change"""
        if employee_name in self.attendance_data:
            self.attendance_data[employee_name]['present'] = is_present
    
    def save_to_excel(self, e):
        """Save attendance data to Excel file in Documents/alswaife/حضور وانصراف/"""
        # Ensure the directory exists
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        alswaife_path = os.path.join(documents_path, "alswaife")
        attendance_path = os.path.join(alswaife_path, "حضور وانصراف")
        
        try:
            os.makedirs(attendance_path, exist_ok=True)
        except OSError as ex:
            self.show_message(f"فشل إنشاء المجلد: {ex}", error=True)
            return
        
        # Check if shift is selected
        if self.shift_dropdown is None or not self.shift_dropdown.value:
            self.show_message("الرجاء اختيار الوردية", error=True)
            return
        
        # Generate filename for single attendance file
        try:
            filename = "attendance.xlsx"  # Single file for all attendance data
            filepath = os.path.join(attendance_path, filename)
        except Exception as ex:
            self.show_message(f"خطأ في إنشاء اسم الملف: {ex}", error=True)
            return
        
        # Load existing data if file exists
        existing_data = []
        if os.path.exists(filepath):
            success, data, error = load_attendance_data(filepath)
            if success and data:
                existing_data = data
        
        # Prepare employees data
        employees_data = []
        
        # Process all employees (from JSON and manually added)
        all_employee_names = set()
        
        # First, process employees from JSON
        for emp in self.employees_list:
            emp_name = emp['name']
            all_employee_names.add(emp_name)
            
            # Check if employee already has a record for this date
            existing_record = None
            for record in existing_data:
                if record['name'] == emp_name and record.get('date', '') == (self.date_field.value if self.date_field else ""):
                    existing_record = record.copy()
                    break
            
            # Create or update employee record
            if existing_record:
                emp_record = existing_record
            else:
                emp_record = {
                    'name': emp_name,
                    'friday_shift1': 0, 'friday_shift2': 0,
                    'saturday_shift1': 0, 'saturday_shift2': 0,
                    'sunday_shift1': 0, 'sunday_shift2': 0,
                    'monday_shift1': 0, 'monday_shift2': 0,
                    'tuesday_shift1': 0, 'tuesday_shift2': 0,
                    'wednesday_shift1': 0, 'wednesday_shift2': 0,
                    'thursday_shift1': 0, 'thursday_shift2': 0,
                    'date': self.date_field.value if self.date_field is not None else "",
                    'advance': 0,
                    'price': emp.get('price', 0)
                }
            
            # Update attendance and price based on UI
            if emp_name in self.attendance_data:
                attendance_info = self.attendance_data[emp_name]
                is_present = attendance_info['present']
                price_field = attendance_info['price_field']
                
                # Get price from field
                try:
                    price = float(price_field.value) if price_field.value else emp.get('price', 0)
                except:
                    price = emp.get('price', 0)
                
                # Update price in record
                emp_record['price'] = price
                
                # Update attendance based on checkbox
                shift_key = self.get_shift_key()
                if shift_key:
                    # Set to price value if present, 0 if not
                    emp_record[shift_key] = price if is_present else 0
            
            employees_data.append(emp_record)
        
        # Then, process manually added employees
        for emp_name, attendance_info in self.attendance_data.items():
            if emp_name not in all_employee_names:  # This is a manually added employee
                # Check if employee already has a record for this date
                existing_record = None
                for record in existing_data:
                    if record['name'] == emp_name and record.get('date', '') == (self.date_field.value if self.date_field else ""):
                        existing_record = record.copy()
                        break
                
                # Create or update employee record
                if existing_record:
                    emp_record = existing_record
                else:
                    emp_record = {
                        'name': emp_name,
                        'friday_shift1': 0, 'friday_shift2': 0,
                        'saturday_shift1': 0, 'saturday_shift2': 0,
                        'sunday_shift1': 0, 'sunday_shift2': 0,
                        'monday_shift1': 0, 'monday_shift2': 0,
                        'tuesday_shift1': 0, 'tuesday_shift2': 0,
                        'wednesday_shift1': 0, 'wednesday_shift2': 0,
                        'thursday_shift1': 0, 'thursday_shift2': 0,
                        'date': self.date_field.value if self.date_field is not None else "",
                        'advance': 0,
                        'price': 0
                    }
                
                # Update attendance and price based on UI
                is_present = attendance_info['present']
                price_field = attendance_info['price_field']
                
                # Get price from field
                try:
                    price = float(price_field.value) if price_field.value else 0
                except:
                    price = 0
                
                # Update price in record
                emp_record['price'] = price
                
                # Update attendance based on checkbox
                shift_key = self.get_shift_key()
                if shift_key:
                    # Set to price value if present, 0 if not
                    emp_record[shift_key] = price if is_present else 0
                
                employees_data.append(emp_record)
        
        # Combine existing data with new/updated data
        # Remove existing records for today's date and employees
        final_data = []
        current_date = self.date_field.value if self.date_field else ""
        employee_names = [emp['name'] for emp in employees_data]
        
        for record in existing_data:
            # Keep records that are not for today's date or not for current employees
            if record.get('date', '') != current_date or record['name'] not in employee_names:
                final_data.append(record)
        
        # Add updated/new records
        final_data.extend(employees_data)
        
        # Save to Excel
        success, error = create_or_update_attendance(filepath, final_data)
        
        if success:
            self.current_file = filepath
            self.show_message(f"تم الحفظ بنجاح: {os.path.basename(filepath)}", filepath=filepath)
        else:
            if error == "file_locked":
                self.show_message("الملف مفتوح في برنامج آخر، الرجاء إغلاقه", error=True)
            else:
                self.show_message(f"خطأ في الحفظ: {error}", error=True)
    
    def show_message(self, message, error=False, filepath=None):
        """Show status message with dialog notification"""
        # Create dialog
        if not error and filepath:
            # Success message with action buttons
            self.dialog = ft.AlertDialog(
                title=ft.Text("الحضور والانصراف"),
                content=ft.Text(message),
                actions=[
                    ft.TextButton("فتح الملف", on_click=lambda e: self.open_file(filepath)),
                    ft.TextButton("فتح المسار", on_click=lambda e: self.open_folder(filepath)),
                    ft.TextButton("إغلاق", on_click=lambda e: self.close_dialog()),
                ],
            )
        else:
            # Error message or no filepath
            self.dialog = ft.AlertDialog(
                title=ft.Text("الحضور والانصراف"),
                content=ft.Text(message),
                actions=[
                    ft.TextButton("إغلاق", on_click=lambda e: self.close_dialog()),
                ],
            )
        
        # Show dialog
        self.page.overlay.append(self.dialog)
        self.dialog.open = True
        self.page.update()
    
    def close_dialog(self):
        """Close the dialog"""
        if hasattr(self, 'dialog') and self.dialog:
            self.dialog.open = False
            self.page.update()
    
    def open_file(self, filepath):
        """Open the saved Excel file"""
        try:
            if platform.system() == 'Windows':
                os.startfile(filepath)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', filepath])
            else:  # Linux
                subprocess.call(['xdg-open', filepath])
        except Exception as ex:
            self.show_message(f"فشل في فتح الملف: {ex}", error=True)
        finally:
            self.close_dialog()
    
    def open_folder(self, filepath):
        """Open the folder containing the saved Excel file"""
        try:
            folder_path = os.path.dirname(filepath)
            if platform.system() == 'Windows':
                os.startfile(folder_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', folder_path])
            else:  # Linux
                subprocess.call(['xdg-open', folder_path])
        except Exception as ex:
            self.show_message(f"فشل في فتح المسار: {ex}", error=True)
        finally:
            self.close_dialog()
    
    def add_new_employee(self, e):
        """Add a new employee not in the JSON file"""
        # Check if date and shift are selected
        if not self.date_field or not self.date_field.value or not self.shift_dropdown or not self.shift_dropdown.value:
            self.show_message("الرجاء اختيار التاريخ والوردية أولاً", error=True)
            return
        
        # Create dialog for adding new employee
        self.name_field = ft.TextField(
            label="اسم الموظف",
            width=300,
            autofocus=True,
            border_radius=8,
            prefix_icon=ft.Icons.PERSON,
            border_color=ft.Colors.BLUE_300,
            focused_border_color=ft.Colors.BLUE_600
        )
        
        self.price_field = ft.TextField(
            label="السعر",
            width=300,
            keyboard_type=ft.KeyboardType.NUMBER,
            value="400",  # Default price
            border_radius=8,
            prefix_icon=ft.Icons.ATTACH_MONEY,
            border_color=ft.Colors.GREEN_300,
            focused_border_color=ft.Colors.GREEN_600,
            suffix_text="جنيه"
        )
        
        self.add_employee_dialog = ft.AlertDialog(
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.PERSON_ADD, color=ft.Colors.BLUE_600, size=28),
                    ft.Text("إضافة موظف جديد", size=20, weight=ft.FontWeight.BOLD)
                ],
                spacing=10
            ),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Container(
                            content=self.name_field,
                            padding=ft.padding.symmetric(vertical=5)
                        ),
                        ft.Container(
                            content=self.price_field,
                            padding=ft.padding.symmetric(vertical=5)
                        )
                    ],
                    spacing=15,
                    tight=True
                ),
                padding=20,
                width=350
            ),
            actions=[
                ft.Container(
                    content=ft.Row(
                        controls=[
                            ft.TextButton(
                                "إلغاء", 
                                on_click=self.close_add_employee_dialog,
                                style=ft.ButtonStyle(
                                    color=ft.Colors.GREY_600,
                                    overlay_color=ft.Colors.GREY_100
                                )
                            ),
                            ft.ElevatedButton(
                                "إضافة", 
                                on_click=self.confirm_add_employee,
                                bgcolor=ft.Colors.BLUE_600,
                                color=ft.Colors.WHITE,
                                style=ft.ButtonStyle(
                                    shape=ft.RoundedRectangleBorder(radius=8)
                                )
                            ),
                        ],
                        alignment=ft.MainAxisAlignment.END,
                        spacing=10
                    ),
                    padding=ft.padding.only(top=10)
                )
            ],
            shape=ft.RoundedRectangleBorder(radius=16),
            bgcolor=ft.Colors.WHITE,
            shadow_color=ft.Colors.with_opacity(0.3, ft.Colors.BLACK),
        )
        
        self.page.overlay.append(self.add_employee_dialog)
        self.add_employee_dialog.open = True
        self.page.update()
    
    def close_add_employee_dialog(self, e=None):
        """Close the add employee dialog"""
        if hasattr(self, 'add_employee_dialog') and self.add_employee_dialog:
            self.add_employee_dialog.open = False
            self.page.update()
    
    def confirm_add_employee(self, e):
        """Confirm adding the new employee"""
        if not self.name_field.value or not self.name_field.value.strip():
            self.show_message("الرجاء إدخال اسم الموظف", error=True)
            return
        
        employee_name = self.name_field.value.strip()
        
        # Check if employee already exists
        if employee_name in self.attendance_data:
            self.show_message("الموظف موجود بالفعل", error=True)
            return
        
        try:
            price = float(self.price_field.value) if self.price_field.value else 400
        except:
            price = 400
        
        # Add the employee row
        self.add_employee_row(employee_name, price, False)
        
        # Close dialog
        self.close_add_employee_dialog()
        
        # Update the page
        self.page.update()
        
        self.show_message(f"تم إضافة الموظف: {employee_name}")
    
    def delete_employee(self, employee_name):
        """Delete an employee (only for manually added employees)"""
        if employee_name in self.attendance_data:
            employee_info = self.attendance_data[employee_name]
            
            # Only allow deletion of manually added employees
            if not employee_info.get('is_json_employee', True):
                # Remove from UI
                card = employee_info['card']
                if self.employees_container and card in self.employees_container.controls:
                    self.employees_container.controls.remove(card)
                
                # Remove from data
                del self.attendance_data[employee_name]
                
                # Update page
                self.page.update()
                
                self.show_message(f"تم حذف الموظف: {employee_name}")
            else:
                self.show_message("لا يمكن حذف الموظفين الأساسيين", error=True)
    
    def go_back(self, e):
        """Go back to dashboard"""
        # Import here to avoid circular dependency
        from views.dashboard_view import DashboardView
        
        self.page.clean()
        dashboard = DashboardView(self.page)
        
        # Get save_callback from main
        save_callback = getattr(self.page, '_save_callback', None)
        if save_callback is not None:
            dashboard.show(save_callback)
        else:
            # Fallback to import
            try:
                from main import save_callback
                dashboard.show(save_callback)
            except:
                # Last resort fallback
                dashboard.show(None)