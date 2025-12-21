import os
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime


def initialize_inventory_excel(file_path):
    """
    Initialize the inventory Excel file with proper formatting
    
    Args:
        file_path (str): Path to the Excel file
    """
    # Create workbook
    wb = Workbook()
    
    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    # Create sheets with Arabic right-to-left orientation
    add_sheet = wb.create_sheet("اذن الاضافه", 0)
    disburse_sheet = wb.create_sheet("اذن الصرف", 1)
    inventory_sheet = wb.create_sheet("المخزون", 2)
    
    # Set right-to-left for all sheets
    for sheet in wb.sheetnames:
        wb[sheet].sheet_view.rightToLeft = True
    
    # Define styles
    header_font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Add headers to add sheet
    add_headers = ["رقم اذن الاضافه", "تاريخ الدخول", "اسم الصنف", "العدد", "ثمن الوحدة", "الإجمالي", "ملاحظات"]
    for col_num, header in enumerate(add_headers, 1):
        cell = add_sheet.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
    
    # Set column widths for add sheet
    column_widths = [18, 15, 25, 12, 15, 15, 30]
    for col_num, width in enumerate(column_widths, 1):
        add_sheet.column_dimensions[get_column_letter(col_num)].width = width
    
    # Add headers to disburse sheet
    disburse_headers = ["رقم اذن الصرف", "تاريخ الصرف", "اسم الصنف", "العدد", "ثمن الوحدة", "الإجمالي", "ملاحظات"]
    for col_num, header in enumerate(disburse_headers, 1):
        cell = disburse_sheet.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
    
    # Set column widths for disburse sheet
    for col_num, width in enumerate(column_widths, 1):
        disburse_sheet.column_dimensions[get_column_letter(col_num)].width = width
    
    # Add headers to inventory sheet
    inventory_headers = ["اسم الصنف", "إجمالي الإضافات", "إجمالي الصرف", "الرصيد الحالي"]
    for col_num, header in enumerate(inventory_headers, 1):
        cell = inventory_sheet.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
    
    # Set column widths for inventory sheet
    inventory_column_widths = [30, 20, 20, 20]
    for col_num, width in enumerate(inventory_column_widths, 1):
        inventory_sheet.column_dimensions[get_column_letter(col_num)].width = width
    
    # Save the workbook
    wb.save(file_path)
    return wb


def add_inventory_entry(file_path, item_name, quantity, unit_price, notes="", entry_date=None):
    """
    Add an inventory entry to the additions sheet
    
    Args:
        file_path (str): Path to the Excel file
        item_name (str): Name of the item
        quantity (float): Quantity of the item
        unit_price (float): Price per unit
        notes (str): Additional notes
        entry_date (str): Date of entry (defaults to today)
        
    Returns:
        int: Entry number
    """
    # Load workbook
    wb = openpyxl.load_workbook(file_path)
    add_sheet = wb["اذن الاضافه"]
    
    # Determine the next entry number
    next_entry_number = add_sheet.max_row
    
    # Get today's date if not provided
    if entry_date is None:
        entry_date = datetime.now().strftime('%d/%m/%Y')
    
    # Calculate total price
    total_price = float(quantity) * float(unit_price)
    
    # Add data row
    row_data = [
        next_entry_number,  # Auto entry number
        entry_date,
        item_name,
        float(quantity),
        float(unit_price),
        total_price,
        notes
    ]
    
    # Add row to sheet
    add_sheet.append(row_data)
    
    # Apply styles to the new row
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    alignment = Alignment(horizontal='center', vertical='center')
    
    row_num = add_sheet.max_row
    for col_num, value in enumerate(row_data, 1):
        cell = add_sheet.cell(row=row_num, column=col_num)
        cell.border = border
        cell.alignment = alignment
        # Apply number formatting for numeric columns
        if col_num in [4, 5, 6]:  # Quantity, Unit Price, Total Price
            cell.number_format = '#,##0.00'
    
    # Save the workbook
    wb.save(file_path)
    
    # Update inventory sheet
    update_inventory_summary(file_path, item_name, float(quantity), 0)
    
    return next_entry_number


def disburse_inventory_entry(file_path, item_name, quantity, unit_price, notes="", disburse_date=None):
    """
    Add an inventory disbursement entry to the disbursements sheet
    
    Args:
        file_path (str): Path to the Excel file
        item_name (str): Name of the item
        quantity (float): Quantity of the item
        unit_price (float): Price per unit
        notes (str): Additional notes
        disburse_date (str): Date of disbursement (defaults to today)
        
    Returns:
        int: Disbursement entry number
    """
    # Load workbook
    wb = openpyxl.load_workbook(file_path)
    disburse_sheet = wb["اذن الصرف"]
    
    # Determine the next entry number
    next_entry_number = disburse_sheet.max_row
    
    # Get today's date if not provided
    if disburse_date is None:
        disburse_date = datetime.now().strftime('%d/%m/%Y')
    
    # Calculate total price
    total_price = float(quantity) * float(unit_price)
    
    # Add data row
    row_data = [
        next_entry_number,  # Auto entry number
        disburse_date,
        item_name,
        float(quantity),
        float(unit_price),
        total_price,
        notes
    ]
    
    # Add row to sheet
    disburse_sheet.append(row_data)
    
    # Apply styles to the new row
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    alignment = Alignment(horizontal='center', vertical='center')
    
    row_num = disburse_sheet.max_row
    for col_num, value in enumerate(row_data, 1):
        cell = disburse_sheet.cell(row=row_num, column=col_num)
        cell.border = border
        cell.alignment = alignment
        # Apply number formatting for numeric columns
        if col_num in [4, 5, 6]:  # Quantity, Unit Price, Total Price
            cell.number_format = '#,##0.00'
    
    # Save the workbook
    wb.save(file_path)
    
    # Update inventory sheet
    update_inventory_summary(file_path, item_name, 0, float(quantity))
    
    return next_entry_number


def update_inventory_summary(file_path, item_name, added_quantity, disbursed_quantity):
    """
    Update the inventory summary sheet
    
    Args:
        file_path (str): Path to the Excel file
        item_name (str): Name of the item
        added_quantity (float): Quantity added
        disbursed_quantity (float): Quantity disbursed
    """
    # Load workbook
    wb = openpyxl.load_workbook(file_path)
    inventory_sheet = wb["المخزون"]
    
    # Find if item already exists in inventory sheet
    item_row = None
    for row_num in range(2, inventory_sheet.max_row + 1):
        if inventory_sheet.cell(row=row_num, column=1).value == item_name:
            item_row = row_num
            break
    
    # If item doesn't exist, add it
    if item_row is None:
        item_row = inventory_sheet.max_row + 1
        inventory_sheet.cell(row=item_row, column=1, value=item_name)
    
    # Get current values
    current_additions = inventory_sheet.cell(row=item_row, column=2).value or 0
    current_disbursements = inventory_sheet.cell(row=item_row, column=3).value or 0
    
    # Update values
    new_additions = float(current_additions) + added_quantity
    new_disbursements = float(current_disbursements) + disbursed_quantity
    balance = new_additions - new_disbursements
    
    # Apply styles
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    alignment = Alignment(horizontal='center', vertical='center')
    
    # Update cells
    inventory_sheet.cell(row=item_row, column=2, value=new_additions).border = border
    inventory_sheet.cell(row=item_row, column=3, value=new_disbursements).border = border
    inventory_sheet.cell(row=item_row, column=4, value=balance).border = border
    
    # Apply number formatting
    for col in range(2, 5):
        inventory_sheet.cell(row=item_row, column=col).number_format = '#,##0.00'
        inventory_sheet.cell(row=item_row, column=col).alignment = alignment
    
    # Color the balance cell based on value
    balance_cell = inventory_sheet.cell(row=item_row, column=4)
    if balance > 0:
        balance_cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')  # Green
    elif balance < 0:
        balance_cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')  # Red
    else:
        balance_cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')  # Yellow
    
    # Save the workbook
    wb.save(file_path)


def get_inventory_summary(file_path):
    """
    Get inventory summary data
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        list: List of dictionaries containing inventory data
    """
    if not os.path.exists(file_path):
        return []
    
    wb = openpyxl.load_workbook(file_path)
    inventory_sheet = wb["المخزون"]
    
    inventory_data = []
    for row_num in range(2, inventory_sheet.max_row + 1):
        item_data = {
            'item_name': inventory_sheet.cell(row=row_num, column=1).value,
            'total_additions': inventory_sheet.cell(row=row_num, column=2).value or 0,
            'total_disbursements': inventory_sheet.cell(row=row_num, column=3).value or 0,
            'current_balance': inventory_sheet.cell(row=row_num, column=4).value or 0
        }
        inventory_data.append(item_data)
    
    return inventory_data