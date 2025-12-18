"""
Excel Utilities for Purchases Management
This module provides functions to generate and manage Excel files for purchases data.
"""

import xlsxwriter
import openpyxl
import os
from typing import List, Dict
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font


def export_purchases_to_excel(records: List[Dict], filepath: str) -> str:
    """
    Export purchases data to an Excel file.
    
    Args:
        records (List[Dict]): List of purchase records
        filepath (str): Path to save the Excel file
    
    Returns:
        str: Path to the created Excel file
    """
    # Create directory if it doesn't exist
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    
    if os.path.exists(filepath):
        append_purchases_to_excel(filepath, records)
    else:
        create_purchases_excel_file(filepath, records)
    
    return filepath


def create_purchases_excel_file(filepath: str, records: List[Dict]):
    """
    Create a new Excel file with purchases data and headers.
    
    Args:
        filepath (str): Path to save the Excel file
        records (List[Dict]): List of purchase records
    """
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet()
    worksheet.right_to_left()
    
    # Formats
    title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#1F4E78',
        'font_color': 'white',
        'font_size': 16,
        'border': 2
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#4472C4',
        'font_color': 'white',
        'font_size': 12
    })
    
    cell_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11
    })
    
    cell_format_alt = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#F2F2F2'
    })
    
    number_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'num_format': '#,##0.00'
    })
    
    number_format_alt = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#F2F2F2',
        'num_format': '#,##0.00'
    })
    
    balance_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bold': True,
        'bg_color': '#E7E6E6',
        'num_format': '#,##0.00'
    })
    
    # Title row (merged cells)
    worksheet.merge_range('A1:H1', 'بيان مشتريات مصنع الجرانيت', title_format)
    worksheet.set_row(0, 30)  # Set title row height
    
    # Headers in row 2
    headers = ["التاريخ", "الكود", "اسم الصنف", "العدد", "إجمالي السعر", "الرصيد", "من", "الملاحظات"]
    for col, header in enumerate(headers):
        worksheet.write(1, col, header, header_format)
    
    worksheet.set_row(1, 25)  # Set header row height
    
    # Write records starting from row 3
    for row_idx, record in enumerate(records, start=2):
        # Alternate row colors
        is_alt_row = (row_idx - 2) % 2 == 1
        current_cell_format = cell_format_alt if is_alt_row else cell_format
        current_number_format = number_format_alt if is_alt_row else number_format
        
        worksheet.write(row_idx, 0, record.get('date', ''), current_cell_format)
        worksheet.write(row_idx, 1, record.get('code', ''), current_cell_format)
        worksheet.write(row_idx, 2, record.get('item_name', ''), current_cell_format)
        worksheet.write(row_idx, 3, record.get('quantity', ''), current_number_format)
        worksheet.write(row_idx, 4, record.get('total_price', ''), current_number_format)
        
        # Balance formula
        if row_idx == 2:
            # First row: balance = total_price
            worksheet.write_formula(row_idx, 5, f'=E{row_idx+1}', balance_format)
        else:
            # Subsequent rows: balance = previous_balance + current_total_price
            worksheet.write_formula(row_idx, 5, f'=F{row_idx}+E{row_idx+1}', balance_format)
        
        worksheet.write(row_idx, 6, record.get('supplier', ''), current_cell_format)
        worksheet.write(row_idx, 7, record.get('notes', ''), current_cell_format)
    
    # Auto-adjust column widths
    worksheet.set_column(0, 0, 14)  # التاريخ
    worksheet.set_column(1, 1, 10)  # الكود
    worksheet.set_column(2, 2, 65)  # اسم الصنف
    worksheet.set_column(3, 3, 10)  # العدد
    worksheet.set_column(4, 4, 15)  # إجمالي السعر
    worksheet.set_column(5, 5, 20)  # الرصيد
    worksheet.set_column(6, 6, 15)  # من
    worksheet.set_column(7, 7, 30)  # الملاحظات
    
    # Freeze panes (freeze title and header rows)
    worksheet.freeze_panes(2, 0)
    
    workbook.close()


def append_purchases_to_excel(filepath: str, new_records: List[Dict]):
    """
    Append new purchase records to an existing Excel file.
    
    Args:
        filepath (str): Path to the existing Excel file
        new_records (List[Dict]): List of new purchase records to append
    """
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active
        
        # Find the starting row for new data
        start_row = worksheet.max_row + 1
        
        # Define styles
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')
        
        # Define fills for alternating rows
        white_fill = PatternFill(fill_type=None)
        gray_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        balance_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
        
        # Bold font for balance
        bold_font = Font(bold=True)
        
        # Add new records
        for row_idx, record in enumerate(new_records, start=start_row):
            # Determine if this is an alternate row (considering title row at row 1, header at row 2)
            is_alt_row = (row_idx - 2) % 2 == 1
            current_fill = gray_fill if is_alt_row else white_fill
            
            worksheet.cell(row=row_idx, column=1, value=record.get('date', '')).border = thin_border
            worksheet.cell(row=row_idx, column=1).fill = current_fill
            
            worksheet.cell(row=row_idx, column=2, value=record.get('code', '')).border = thin_border
            worksheet.cell(row=row_idx, column=2).fill = current_fill
            
            worksheet.cell(row=row_idx, column=3, value=record.get('item_name', '')).border = thin_border
            worksheet.cell(row=row_idx, column=3).fill = current_fill
            
            worksheet.cell(row=row_idx, column=4, value=record.get('quantity', '')).border = thin_border
            worksheet.cell(row=row_idx, column=4).fill = current_fill
            worksheet.cell(row=row_idx, column=4).number_format = '#,##0.00'
            
            worksheet.cell(row=row_idx, column=5, value=record.get('total_price', '')).border = thin_border
            worksheet.cell(row=row_idx, column=5).fill = current_fill
            worksheet.cell(row=row_idx, column=5).number_format = '#,##0.00'
            
            if row_idx == 3 and start_row == 3:
                # First data row in file (row 3, after title and header), balance = total_price
                balance_formula = f'=E{row_idx}'
            else:
                # Subsequent rows, balance = previous_balance + current_total_price
                balance_formula = f'=F{row_idx-1}+E{row_idx}'
            
            balance_cell = worksheet.cell(row=row_idx, column=6, value=balance_formula)
            balance_cell.border = thin_border
            balance_cell.fill = balance_fill
            balance_cell.font = bold_font
            balance_cell.number_format = '#,##0.00'
            
            worksheet.cell(row=row_idx, column=7, value=record.get('supplier', '')).border = thin_border
            worksheet.cell(row=row_idx, column=7).fill = current_fill
            
            worksheet.cell(row=row_idx, column=8, value=record.get('notes', '')).border = thin_border
            worksheet.cell(row=row_idx, column=8).fill = current_fill
            
            # Apply alignment to all cells
            for col in range(1, 9):
                worksheet.cell(row=row_idx, column=col).alignment = center_alignment
        
        workbook.save(filepath)
    except Exception as e:
        raise Exception(f"Error appending to Excel file: {str(e)}")


def load_purchases_from_excel(filepath: str) -> List[Dict]:
    """
    Load purchase records from an Excel file.
    
    Args:
        filepath (str): Path to the Excel file
        
    Returns:
        List[Dict]: List of purchase records
    """
    records = []
    
    if not os.path.exists(filepath):
        return records
    
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active
        
        # Skip title row (row 1) and header row (row 2), start reading from row 3
        for row in range(3, worksheet.max_row + 1):
            record = {
                'date': worksheet.cell(row=row, column=1).value or "",
                'code': worksheet.cell(row=row, column=2).value or "",
                'item_name': worksheet.cell(row=row, column=3).value or "",
                'quantity': worksheet.cell(row=row, column=4).value or "",
                'total_price': worksheet.cell(row=row, column=5).value or "",
                'supplier': worksheet.cell(row=row, column=7).value or "",
                'notes': worksheet.cell(row=row, column=8).value or ""
            }
            records.append(record)
            
        workbook.close()
    except Exception as e:
        print(f"Error loading purchases from Excel: {e}")
    
    return records


def load_item_names_from_excel(filepath: str) -> List[str]:
    """
    Load existing item names from an Excel file for auto-complete.
    
    Args:
        filepath (str): Path to the Excel file
        
    Returns:
        List[str]: List of unique item names
    """
    items = set()
    
    if not os.path.exists(filepath):
        return list(items)
    
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active
        
        # Skip title row (row 1) and header row (row 2), read item names from column 3
        for row in range(3, worksheet.max_row + 1):
            item_name = worksheet.cell(row=row, column=3).value
            if item_name:
                items.add(str(item_name))
                
        workbook.close()
    except Exception as e:
        print(f"Error loading items from Excel: {e}")
    
    return list(items)