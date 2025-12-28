"""
Excel Utilities for Purchases Management
This module provides functions to generate and manage Excel files for income and expenses data.
Contains three sheets: Income (الإيرادات), Expenses (المصروفات), Summary (الإجمالي)
"""

import xlsxwriter
import openpyxl
import os
from typing import List, Dict
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font

from utils.log_utils import log_error, log_exception


def export_purchases_to_excel(records: List[Dict], filepath: str) -> str:
    """
    Export purchases data to the expenses sheet of the Excel file.
    """
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    
    if os.path.exists(filepath):
        append_to_expenses(filepath, records)
    else:
        create_purchases_excel_file(filepath, [], records)
    
    return filepath


def add_income_record(filepath: str, record: Dict) -> str:
    """
    Add an income record to the income sheet of the Excel file.
    Called from invoice saving to record income.
    If a record with the same invoice_number exists, it will be updated.
    """
    try:
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        
        if os.path.exists(filepath):
            if update_income_record(filepath, record):
                pass  # Updated existing record
            else:
                append_to_income(filepath, [record])
        else:
            create_purchases_excel_file(filepath, [record], [])
        
        return filepath
    except Exception as e:
        log_exception(f"add_income_record failed: {e}")
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
        
        # Get income sheet
        if 'الإيرادات' not in workbook.sheetnames:
            workbook.close()
            return False
        
        worksheet = workbook['الإيرادات']
        
        # Search for the invoice number in column A (column 1)
        for row in range(3, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=row, column=1).value
            if str(cell_value) == str(invoice_number):
                # Found the record, update it
                worksheet.cell(row=row, column=2, value=record.get('client', ''))
                worksheet.cell(row=row, column=3, value=record.get('amount', ''))
                worksheet.cell(row=row, column=4, value=record.get('date', ''))
                
                workbook.save(filepath)
                workbook.close()
                return True
        
        workbook.close()
        return False
    except Exception as e:
        log_error(f"update_income_record failed: {e}")
        return False


def create_purchases_excel_file(filepath: str, income_records: List[Dict] = None, expense_records: List[Dict] = None):
    """
    Create a new Excel file with 3 sheets:
    - الإيرادات (Income)
    - المصروفات (Expenses)
    - الإجمالي (Summary)
    """
    if income_records is None:
        income_records = []
    if expense_records is None:
        expense_records = []
    
    workbook = xlsxwriter.Workbook(filepath)
    
    # ==================== FORMATS ====================
    title_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#1F4E78', 'font_color': 'white', 'font_size': 16, 'border': 2
    })
    
    income_header_format = workbook.add_format({
        'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#00B050', 'font_color': 'white', 'font_size': 12
    })
    
    expenses_header_format = workbook.add_format({
        'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#C00000', 'font_color': 'white', 'font_size': 12
    })
    
    summary_header_format = workbook.add_format({
        'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#4472C4', 'font_color': 'white', 'font_size': 12
    })
    
    cell_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11
    })
    
    cell_format_alt = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'bg_color': '#F2F2F2'
    })
    
    number_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'num_format': '#,##0'
    })
    
    number_format_alt = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11,
        'bg_color': '#F2F2F2', 'num_format': '#,##0'
    })
    
    summary_label_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#D9E1F2', 'font_color': '#1F4E78', 'font_size': 14, 'border': 1
    })
    
    summary_income_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#C6EFCE', 'font_color': '#006100', 'font_size': 16, 'border': 2, 'num_format': '#,##0'
    })
    
    summary_expenses_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'font_size': 16, 'border': 2, 'num_format': '#,##0'
    })
    
    summary_balance_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#BDD7EE', 'font_color': '#1F4E78', 'font_size': 16, 'border': 2, 'num_format': '#,##0'
    })
    
    # ==================== SHEET 1: الإيرادات ====================
    ws_income = workbook.add_worksheet('الإيرادات')
    ws_income.right_to_left()
    
    # Title
    ws_income.merge_range(0, 0, 0, 3, 'سجل الإيرادات', title_format)
    ws_income.set_row(0, 30)
    
    # Headers
    income_headers = ["رقم الفاتورة", "اسم العميل", "المبلغ", "التاريخ"]
    for col, header in enumerate(income_headers):
        ws_income.write(1, col, header, income_header_format)
    ws_income.set_row(1, 25)
    
    # Data
    for row_idx, record in enumerate(income_records, start=2):
        is_alt = (row_idx - 2) % 2 == 1
        cf = cell_format_alt if is_alt else cell_format
        nf = number_format_alt if is_alt else number_format
        
        ws_income.write(row_idx, 0, record.get('invoice_number', ''), cf)
        ws_income.write(row_idx, 1, record.get('client', ''), cf)
        ws_income.write(row_idx, 2, record.get('amount', ''), nf)
        ws_income.write(row_idx, 3, record.get('date', ''), cf)
    
    # Column widths
    ws_income.set_column(0, 0, 12)
    ws_income.set_column(1, 1, 20)
    ws_income.set_column(2, 2, 12)
    ws_income.set_column(3, 3, 12)
    
    # ==================== SHEET 2: المصروفات ====================
    ws_expenses = workbook.add_worksheet('المصروفات')
    ws_expenses.right_to_left()
    
    # Title
    ws_expenses.merge_range(0, 0, 0, 3, 'سجل المصروفات', title_format)
    ws_expenses.set_row(0, 30)
    
    # Headers
    expenses_headers = ["العدد", "البيان", "المبلغ", "التاريخ"]
    for col, header in enumerate(expenses_headers):
        ws_expenses.write(1, col, header, expenses_header_format)
    ws_expenses.set_row(1, 25)
    
    # Data
    for row_idx, record in enumerate(expense_records, start=2):
        is_alt = (row_idx - 2) % 2 == 1
        cf = cell_format_alt if is_alt else cell_format
        nf = number_format_alt if is_alt else number_format
        
        ws_expenses.write(row_idx, 0, record.get('quantity', ''), cf)
        ws_expenses.write(row_idx, 1, record.get('item_name', ''), cf)
        ws_expenses.write(row_idx, 2, record.get('total_price', ''), nf)
        ws_expenses.write(row_idx, 3, record.get('date', ''), cf)
    
    # Column widths
    ws_expenses.set_column(0, 0, 8)
    ws_expenses.set_column(1, 1, 30)
    ws_expenses.set_column(2, 2, 12)
    ws_expenses.set_column(3, 3, 12)
    
    # ==================== SHEET 3: الإجمالي ====================
    ws_summary = workbook.add_worksheet('الإجمالي')
    ws_summary.right_to_left()
    
    # Title
    ws_summary.merge_range(0, 0, 0, 1, 'ملخص الحسابات', title_format)
    ws_summary.set_row(0, 35)
    
    # Summary rows
    ws_summary.write(2, 0, 'إجمالي الإيرادات', summary_label_format)
    ws_summary.write(2, 1, "=SUM(الإيرادات!C:C)", summary_income_format)
    
    ws_summary.write(4, 0, 'إجمالي المصروفات', summary_label_format)
    ws_summary.write(4, 1, "=SUM(المصروفات!C:C)", summary_expenses_format)
    
    ws_summary.write(6, 0, 'الرصيد المتبقي', summary_label_format)
    ws_summary.write(6, 1, '=B3-B5', summary_balance_format)
    
    # Row heights
    ws_summary.set_row(2, 35)
    ws_summary.set_row(4, 35)
    ws_summary.set_row(6, 35)
    
    # Column widths
    ws_summary.set_column(0, 0, 20)
    ws_summary.set_column(1, 1, 20)
    
    try:
        workbook.close()
    except PermissionError as e:
        raise PermissionError("الملف مفتوح في Excel. أغلقه وحاول مرة أخرى.") from e


def append_to_income(filepath: str, new_records: List[Dict]):
    """
    Append new income records to the income sheet.
    """
    try:
        workbook = openpyxl.load_workbook(filepath)
        
        if 'الإيرادات' not in workbook.sheetnames:
            workbook.close()
            raise ValueError("Income sheet not found")
        
        worksheet = workbook['الإيرادات']
        
        # Find the last row with data
        start_row = 3
        for row in range(3, worksheet.max_row + 2):
            if worksheet.cell(row=row, column=1).value is None:
                start_row = row
                break
        
        # Styles
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')
        white_fill = PatternFill(fill_type=None)
        alt_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        
        for row_idx, record in enumerate(new_records, start=start_row):
            is_alt = (row_idx - 3) % 2 == 1
            current_fill = alt_fill if is_alt else white_fill
            
            cell = worksheet.cell(row=row_idx, column=1, value=record.get('invoice_number', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            cell = worksheet.cell(row=row_idx, column=2, value=record.get('client', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            cell = worksheet.cell(row=row_idx, column=3, value=record.get('amount', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            cell.number_format = '#,##0'
            
            cell = worksheet.cell(row=row_idx, column=4, value=record.get('date', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
        
        workbook.save(filepath)
        workbook.close()
    except PermissionError as e:
        raise PermissionError("الملف مفتوح في Excel. أغلقه وحاول مرة أخرى.") from e


def append_to_expenses(filepath: str, new_records: List[Dict]):
    """
    Append new expense records to the expenses sheet.
    """
    try:
        workbook = openpyxl.load_workbook(filepath)
        
        if 'المصروفات' not in workbook.sheetnames:
            workbook.close()
            raise ValueError("Expenses sheet not found")
        
        worksheet = workbook['المصروفات']
        
        # Find the last row with data
        start_row = 3
        for row in range(3, worksheet.max_row + 2):
            if worksheet.cell(row=row, column=1).value is None:
                start_row = row
                break
        
        # Styles
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')
        white_fill = PatternFill(fill_type=None)
        alt_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        
        for row_idx, record in enumerate(new_records, start=start_row):
            is_alt = (row_idx - 3) % 2 == 1
            current_fill = alt_fill if is_alt else white_fill
            
            cell = worksheet.cell(row=row_idx, column=1, value=record.get('quantity', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            cell = worksheet.cell(row=row_idx, column=2, value=record.get('item_name', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            
            cell = worksheet.cell(row=row_idx, column=3, value=record.get('total_price', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
            cell.number_format = '#,##0'
            
            cell = worksheet.cell(row=row_idx, column=4, value=record.get('date', ''))
            cell.border = thin_border
            cell.fill = current_fill
            cell.alignment = center_alignment
        
        workbook.save(filepath)
        workbook.close()
    except PermissionError as e:
        raise PermissionError("الملف مفتوح في Excel. أغلقه وحاول مرة أخرى.") from e


def load_item_names_from_excel(filepath: str) -> List[str]:
    """
    Load existing item names from the expenses sheet for auto-complete.
    """
    items = set()
    
    if not os.path.exists(filepath):
        return list(items)
    
    try:
        workbook = openpyxl.load_workbook(filepath)
        
        if 'المصروفات' not in workbook.sheetnames:
            workbook.close()
            return list(items)
        
        worksheet = workbook['المصروفات']
        
        # Read item names from column B (column 2)
        for row in range(3, worksheet.max_row + 1):
            item_name = worksheet.cell(row=row, column=2).value
            if item_name:
                items.add(str(item_name))
                
        workbook.close()
    except PermissionError:
        return []
    except Exception as e:
        log_error(f"Error loading items from Excel: {e}")
    
    return list(items)
