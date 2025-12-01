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
        
        # Floating save button
        self.floating_save_btn = ft.FloatingActionButton(
            icon=ft.Icons.SAVE,
            on_click=self.save_to_excel,
            bgcolor=ft.Colors.GREEN_700
        )
        
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
        
        self.page.add(main_column)
        self.page.floating_action_button = self.floating_save_btn
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
        """Load existing attendance data for current date"""
        # Clear existing controls
        if self.employees_container is not None:
            self.employees_container.controls.clear()
        
        # Ensure the directory exists
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        alswaife_path = os.path.join(documents_path, "alswaife")
        attendance_path = os.path.join(alswaife_path, "حضور وانصراف")
        
        # Generate filename with date only
        try:
            if self.date_field is not None and self.date_field.value:
                date_obj = datetime.strptime(self.date_field.value, '%d/%m/%Y')
                
                # Format date as DD-MM
                date_str = date_obj.strftime('%d-%m')
                
                # Create filename with date only
                filename = f"{date_str}.xlsx"
                filepath = os.path.join(attendance_path, filename)
                
                if os.path.exists(filepath):
                    # Load from Excel
                    success, data, error = load_attendance_data(filepath)
                    
                    if success and data:
                        self.current_file = filepath
                        # Create a dictionary for quick lookup
                        employee_data = {emp['name']: emp for emp in data}
                        
                        # Create employee rows with existing data
                        for emp in self.employees_list:
                            emp_name = emp['name']
                            price = emp.get('price', 0)
                            
                            # Check if employee has attendance data
                            is_present = False
                            emp_price = price  # Default to JSON price
                            if emp_name in employee_data:
                                emp_record = employee_data[emp_name]
                                shift_key = self.get_shift_key()
                                if shift_key and emp_record.get(shift_key, 0) > 0:
                                    is_present = True
                                # Use price from existing record if available
                                if 'price' in emp_record:
                                    emp_price = emp_record['price']
                            
                            self.add_employee_row(emp_name, emp_price, is_present)
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
        except:
            # Create employee rows without existing data
            for emp in self.employees_list:
                emp_name = emp['name']
                price = emp.get('price', 0)
                self.add_employee_row(emp_name, price, False)
        
        self.page.update()
    
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
        # Create price field first
        price_field = ft.TextField(
            value=str(price),
            width=120,
            text_align=ft.TextAlign.CENTER,
            keyboard_type=ft.KeyboardType.NUMBER,
            label="السعر",
            dense=True
        )
        
        # Create card container for better visual appearance
        card = ft.Card(
            content=ft.Container(
                content=ft.Row(
                    controls=[
                        # Checkbox for attendance
                        ft.Checkbox(
                            value=is_present,
                            on_change=lambda e, n=name: self.on_attendance_change(n, e.control.value)
                        ),
                        # Employee name
                        ft.Text(name, size=18, weight=ft.FontWeight.BOLD),
                        # Spacer
                        ft.Container(expand=True),
                        # Price field
                        price_field
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                padding=10,
            ),
            elevation=2,
        )
        
        # Store attendance status and price field reference
        self.attendance_data[name] = {
            'card': card,
            'price_field': price_field,
            'present': is_present
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
        
        # Generate filename with only the date
        try:
            if self.date_field is not None and self.date_field.value:
                date_obj = datetime.strptime(self.date_field.value, '%d/%m/%Y')
                
                # Format date as DD-MM
                date_str = date_obj.strftime('%d-%m')
                
                # Create filename with date only
                filename = f"{date_str}.xlsx"
                filepath = os.path.join(attendance_path, filename)
            else:
                self.show_message("تنسيق التاريخ غير صحيح", error=True)
                return
        except Exception as ex:
            self.show_message(f"خطأ في إنشاء اسم الملف: {ex}", error=True)
            return
        
        # Prepare employees data
        employees_data = []
        
        # Load existing data if file exists
        existing_data = {}
        if os.path.exists(filepath):
            success, data, error = load_attendance_data(filepath)
            if success and data:
                existing_data = {emp['name']: emp for emp in data}
        
        # Process each employee
        for emp in self.employees_list:
            emp_name = emp['name']
            
            # Start with existing data or create new record
            if emp_name in existing_data:
                emp_record = existing_data[emp_name].copy()
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
        
        # Save to Excel (removed shift_name and day_name parameters as they're no longer needed)
        success, error = create_or_update_attendance(filepath, employees_data)
        
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