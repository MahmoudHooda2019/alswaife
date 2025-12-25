"""
Excel Utilities for Purchases Management
This module provides functions to generate and manage Excel files for income and expenses data.
Contains two side-by-side tables: Income (الإيرادات) and Expenses (المصروفات)
"""

import xlsxwriter
import openpyxl
import os
from typing import List, Dict
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font

# Column positions for income table (columns A-D, 0-3)
INCOME_START_COL = 0
INCOME_END_COL = 3

# Column positions for expenses table (columns E-H, 4-7) - بدون عمود فاصل
EXPENSES_START_COL = 4
EXPENSES_END_COL = 7


def export_purchases_to_excel(records: List[Dict], filepath: str) -> str:
    """
    Export purchases data to the expenses section of the Excel file.
    """
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    
    if os.path.exists(filepath):
        append_to_expenses(filepath, records)
    else:
        create_purchases_excel_file(filepath, [], records)
    
    return filepath


def add_income_record(filepath: str, record: Dict) -> str:
    """
    Add an income record to the income section of the Excel file.
    Called from invoice saving to record income.
    If a record with the same invoice_number exists, it will be updated.
    """
    print(f"[DEBUG] add_income_record called with filepath: {filepath}")
    print(f"[DEBUG] Record: {record}")
    
    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        print(f"[DEBUG] Directory created/exists: {os.path.dirname(filepath)}")
        
        if os.path.exists(filepath):
            print(f"[DEBUG] File exists, checking for existing record...")
            # Check if record with same invoice_number exists and update it
            if update_income_record(filepath, record):
                print(f"[DEBUG] Updated existing record")
            else:
                print(f"[DEBUG] No existing record found, appending new...")
                append_to_income(filepath, [record])
        else:
            print(f"[DEBUG] File does not exist, creating new file...")
            create_purchases_excel_file(filepath, [record], [])
        
        print(f"[DEBUG] Operation completed successfully")
        return filepath
    except Exception as e:
        print(f"[ERROR] add_income_record failed: {e}")
        import traceback
        traceback.print_exc()
        raise


def update_income_record(filepath: str, record: Dict) -> bool:
    """
    Update an existing income record by invoice_number.
    Returns True if record was found and updated, False otherwise.
    """
    invoice_number = record.get('invoice_number', '')
    if not invoice_number:
        return False
    
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active
        
        # Search for the invoice number in column A (column 1)
        for row in range(6, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=row, column=1).value
            if str(cell_value) == str(invoice_number):
                # Found the record, update it
                worksheet.cell(row=row, column=2, value=record.get('client', ''))
                worksheet.cell(row=row, column=3, value=record.get('amount', ''))
                worksheet.cell(row=row, column=4, value=record.get('date', ''))
                
                workbook.save(filepath)
                workbook.close()
                print(f"[DEBUG] Updated income record for invoice {invoice_number}")
                return True
        
        workbook.close()
        return False
    except Exception as e:
        print(f"[ERROR] update_income_record failed: {e}")
        return False
        raise


def create_purchases_excel_file(filepath: str, income_records: List[Dict] = None, expense_records: List[Dict] = None):
    """
    Create a new Excel file with the new layout:
    - Main title at top
    - Summary row with totals
    - Two side-by-side tables: Income and Expenses
    """
    if income_records is None:
        income_records = []
    if expense_records is None:
        expense_records = []
    
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet("سجل الحسابات")
    worksheet.right_to_left()
    
    # ==================== FORMATS ====================
    # Main title format
    main_title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#1F4E78',
        'font_color': 'white',
        'font_size': 18,
        'border': 2
    })
    
    # Summary label format
    summary_label_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#D9E1F2',
        'font_color': '#1F4E78',
        'font_size': 12,
        'border': 1
    })
    
    # Summary value format - Income (green)
    summary_income_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#C6EFCE',
        'font_color': '#006100',
        'font_size': 14,
        'border': 1,
        'num_format': '#,##0'
    })
    
    # Summary value format - Expenses (red)
    summary_expenses_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#FFC7CE',
        'font_color': '#9C0006',
        'font_size': 14,
        'border': 1,
        'num_format': '#,##0'
    })
    
    # Summary value format - Balance (blue)
    summary_balance_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#BDD7EE',
        'font_color': '#1F4E78',
        'font_size': 14,
        'border': 1,
        'num_format': '#,##0'
    })
    
    # Income title format (green)
    income_title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#00B050',
        'font_color': 'white',
        'font_size': 14,
        'border': 2
    })
    
    # Expenses title format (red)
    expenses_title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#C00000',
        'font_color': 'white',
        'font_size': 14,
        'border': 2
    })
    
    # Income header format
    income_header_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#92D050',
        'font_color': 'white',
        'font_size': 11
    })
    
    # Expenses header format
    expenses_header_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#FF6B6B',
        'font_color': 'white',
        'font_size': 11
    })
    
    # Cell formats
    cell_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11
    })
    
    cell_format_alt_income = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#E2EFDA'
    })
    
    cell_format_alt_expenses = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#FCE4D6'
    })
    
    number_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'num_format': '#,##0'
    })
    
    number_format_alt_income = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#E2EFDA',
        'num_format': '#,##0'
    })
    
    number_format_alt_expenses = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#FCE4D6',
        'num_format': '#,##0'
    })
    
    # ==================== ROW 0: MAIN TITLE ====================
    worksheet.merge_range(0, 0, 0, 7, 'بيان مصروفات وإيرادات مصنع جرانيت السويفي', main_title_format)
    worksheet.set_row(0, 30)
    
    # ==================== ROW 1: SUMMARY - 3 boxes متساوية ====================
    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)
    
    # Box 1: إجمالي الإيرادات (أعمدة 0-1)
    worksheet.merge_range(1, 0, 1, 1, 'إجمالي الإيرادات', summary_label_format)
    worksheet.merge_range(2, 0, 2, 1, '=SUM(C6:C1000)', summary_income_format)
    
    # Box 2: الرصيد المتبقي (أعمدة 2-5)
    worksheet.merge_range(1, 2, 1, 5, 'الرصيد المتبقي', summary_label_format)
    worksheet.merge_range(2, 2, 2, 5, '=A3-G3', summary_balance_format)
    
    # Box 3: إجمالي المصروفات (أعمدة 6-7)
    worksheet.merge_range(1, 6, 1, 7, 'إجمالي المصروفات', summary_label_format)
    worksheet.merge_range(2, 6, 2, 7, '=SUM(G6:G1000)', summary_expenses_format)
    
    # ==================== ROW 3: TABLE TITLES ====================
    worksheet.merge_range(3, INCOME_START_COL, 3, INCOME_END_COL, 'الإيرادات', income_title_format)
    worksheet.merge_range(3, EXPENSES_START_COL, 3, EXPENSES_END_COL, 'المصروفات', expenses_title_format)
    worksheet.set_row(3, 25)
    
    # ==================== ROW 4: TABLE HEADERS ====================
    # Income headers
    income_headers = ["رقم الفاتورة", "اسم العميل", "المبلغ", "تاريخ الإيراد"]
    for col, header in enumerate(income_headers):
        worksheet.write(4, INCOME_START_COL + col, header, income_header_format)
    
    # Expenses headers
    expenses_headers = ["العدد", "البيان", "المبلغ", "تاريخ الصرف"]
    for col, header in enumerate(expenses_headers):
        worksheet.write(4, EXPENSES_START_COL + col, header, expenses_header_format)
    
    worksheet.set_row(4, 22)
    
    # ==================== DATA ROWS (starting from row 5) ====================
    # Write income records
    for row_idx, record in enumerate(income_records, start=5):
        is_alt_row = (row_idx - 5) % 2 == 1
        current_cell_format = cell_format_alt_income if is_alt_row else cell_format
        current_number_format = number_format_alt_income if is_alt_row else number_format
        
        worksheet.write(row_idx, INCOME_START_COL + 0, record.get('invoice_number', ''), current_cell_format)
        worksheet.write(row_idx, INCOME_START_COL + 1, record.get('client', ''), current_cell_format)
        worksheet.write(row_idx, INCOME_START_COL + 2, record.get('amount', ''), current_number_format)
        worksheet.write(row_idx, INCOME_START_COL + 3, record.get('date', ''), current_cell_format)
    
    # Write expense records
    for row_idx, record in enumerate(expense_records, start=5):
        is_alt_row = (row_idx - 5) % 2 == 1
        current_cell_format = cell_format_alt_expenses if is_alt_row else cell_format
        current_number_format = number_format_alt_expenses if is_alt_row else number_format
        
        worksheet.write(row_idx, EXPENSES_START_COL + 0, record.get('quantity', ''), current_cell_format)
        worksheet.write(row_idx, EXPENSES_START_COL + 1, record.get('item_name', ''), current_cell_format)
        worksheet.write(row_idx, EXPENSES_START_COL + 2, record.get('total_price', ''), current_number_format)
        worksheet.write(row_idx, EXPENSES_START_COL + 3, record.get('date', ''), current_cell_format)
    
    # ==================== COLUMN WIDTHS (مصغرة) ====================
    # Income columns
    worksheet.set_column(0, 0, 12)   # رقم الفاتورة
    worksheet.set_column(1, 1, 18)   # اسم العميل
    worksheet.set_column(2, 2, 12)   # المبلغ
    worksheet.set_column(3, 3, 12)   # تاريخ الإيراد
    
    # Expenses columns (بدون عمود فاصل)
    worksheet.set_column(4, 4, 8)    # العدد
    worksheet.set_column(5, 5, 28)   # البيان
    worksheet.set_column(6, 6, 12)   # المبلغ
    worksheet.set_column(7, 7, 12)   # تاريخ الصرف
    
    try:
        workbook.close()
    except PermissionError as e:
        raise PermissionError("File is currently open in Excel. Please close the file and try again.") from e


def append_to_income(filepath: str, new_records: List[Dict]):
    """
    Append new income records to the income table.
    """
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active
        
        # Find the last row with data in income table (column A)
        start_row = 6  # Default start after headers (row 5 in 0-indexed = row 6 in 1-indexed)
        for row in range(6, worksheet.max_row + 2):
            if worksheet.cell(row=row, column=1).value is None:
                start_row = row
                break
        
        # Define styles
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        white_fill = PatternFill(fill_type=None)
        alt_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
        
        # Add new records
        for row_idx, record in enumerate(new_records, start=start_row):
            is_alt_row = (row_idx - 6) % 2 == 1
            current_fill = alt_fill if is_alt_row else white_fill
            
            # Invoice number
            cell = worksheet.cell(row=row_idx, column=1, value=record.get('invoice_number', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            # Client
            cell = worksheet.cell(row=row_idx, column=2, value=record.get('client', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            # Amount
            cell = worksheet.cell(row=row_idx, column=3, value=record.get('amount', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            cell.number_format = '#,##0'
            
            # Date
            cell = worksheet.cell(row=row_idx, column=4, value=record.get('date', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
        
        workbook.save(filepath)
        workbook.close()
    except PermissionError as e:
        raise PermissionError("File is currently open in Excel. Please close the file and try again.") from e


def append_to_expenses(filepath: str, new_records: List[Dict]):
    """
    Append new expense records to the expenses table.
    """
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active
        
        # Find the last row with data in expenses table (column E = column 5)
        start_row = 6  # Default start after headers
        for row in range(6, worksheet.max_row + 2):
            if worksheet.cell(row=row, column=5).value is None:
                start_row = row
                break
        
        # Define styles
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        white_fill = PatternFill(fill_type=None)
        alt_fill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
        
        # Add new records
        for row_idx, record in enumerate(new_records, start=start_row):
            is_alt_row = (row_idx - 6) % 2 == 1
            current_fill = alt_fill if is_alt_row else white_fill
            
            # Quantity (column E = 5)
            cell = worksheet.cell(row=row_idx, column=5, value=record.get('quantity', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            # Item name (column F = 6)
            cell = worksheet.cell(row=row_idx, column=6, value=record.get('item_name', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            # Total price (column G = 7)
            cell = worksheet.cell(row=row_idx, column=7, value=record.get('total_price', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            cell.number_format = '#,##0'
            
            # Date (column H = 8)
            cell = worksheet.cell(row=row_idx, column=8, value=record.get('date', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
        
        workbook.save(filepath)
        workbook.close()
    except PermissionError as e:
        raise PermissionError("File is currently open in Excel. Please close the file and try again.") from e


def load_item_names_from_excel(filepath: str) -> List[str]:
    """
    Load existing item names from the expenses table for auto-complete.
    """
    items = set()
    
    if not os.path.exists(filepath):
        return list(items)
    
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active
        
        # Read item names from column F (column 6) in expenses table
        for row in range(6, worksheet.max_row + 1):
            item_name = worksheet.cell(row=row, column=6).value
            if item_name:
                items.add(str(item_name))
                
        workbook.close()
    except PermissionError:
        return []
    except Exception as e:
        print(f"Error loading items from Excel: {e}")
    
    return list(items)
