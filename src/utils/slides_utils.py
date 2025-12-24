import os
import openpyxl
import traceback
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime


def initialize_slides_inventory_excel(file_path):
    """
    Initialize the slides inventory Excel file with proper formatting and formulas
    
    Args:
        file_path (str): Path to the Excel file
    """
    print(f"[DEBUG] initialize_slides_inventory_excel called with file: {file_path}")
    try:
        # Create workbook
        wb = Workbook()
        
        # Remove default sheet
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        # Create sheets with Arabic right-to-left orientation
        add_sheet = wb.create_sheet("اذن اضافة الشرائح", 0)
        disburse_sheet = wb.create_sheet("اذن صرف الشرائح", 1)
        inventory_sheet = wb.create_sheet("مخزون الشرائح", 2)
        
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
        print(f"[DEBUG] Slides Excel file initialized successfully")
        return wb
    except Exception as e:
        print(f"[ERROR] Failed to initialize slides Excel file: {e}")
        traceback.print_exc()
        raise


def convert_existing_slides_inventory_to_formulas(file_path):
    """
    Convert an existing slides inventory file to use formulas instead of manual calculations
    
    Args:
        file_path (str): Path to the Excel file
    """
    print(f"[DEBUG] convert_existing_slides_inventory_to_formulas called with file: {file_path}")
    try:
        if not os.path.exists(file_path):
            print(f"[DEBUG] File does not exist, creating new one")
            # If file doesn't exist, create a new one
            initialize_slides_inventory_excel(file_path)
            return
        
        # Load workbook
        wb = openpyxl.load_workbook(file_path)
        add_sheet = wb["اذن اضافة الشرائح"]
        disburse_sheet = wb["اذن صرف الشرائح"]
        inventory_sheet = wb["مخزون الشرائح"]
        
        # Get all unique item names from both sheets
        item_names = set()
        
        # Get items from additions sheet (skip header row)
        for row_num in range(2, add_sheet.max_row + 1):
            item_name = add_sheet.cell(row=row_num, column=3).value  # Item name column
            if item_name:
                item_names.add(item_name)
        
        # Get items from disbursements sheet (skip header row)
        for row_num in range(2, disburse_sheet.max_row + 1):
            item_name = disburse_sheet.cell(row=row_num, column=3).value  # Item name column
            if item_name:
                item_names.add(item_name)
        
        print(f"[DEBUG] Found {len(item_names)} unique items")
        
        # Clear existing data in inventory sheet (keep header)
        for row_num in range(2, inventory_sheet.max_row + 1):
            for col_num in range(1, 5):
                inventory_sheet.cell(row=row_num, column=col_num).value = None
        
        # Apply styles
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        alignment = Alignment(horizontal='center', vertical='center')
        
        # Add items to inventory sheet with corrected formulas
        for row_num, item_name in enumerate(sorted(item_names), 2):
            # Item name
            inventory_sheet.cell(row=row_num, column=1, value=item_name).border = border
            inventory_sheet.cell(row=row_num, column=1).alignment = alignment
            
            # Formula for total additions (SUMIF for additions)
            additions_formula = f"=SUMIF('اذن اضافة الشرائح'!C:C,A{row_num},'اذن اضافة الشرائح'!D:D)"
            inventory_sheet.cell(row=row_num, column=2).value = additions_formula
            inventory_sheet.cell(row=row_num, column=2).border = border
            inventory_sheet.cell(row=row_num, column=2).alignment = alignment
            inventory_sheet.cell(row=row_num, column=2).number_format = '#,##0.00'
            
            # Formula for total disbursements (SUMIF for disbursements)
            disbursements_formula = f"=SUMIF('اذن صرف الشرائح'!C:C,A{row_num},'اذن صرف الشرائح'!D:D)"
            inventory_sheet.cell(row=row_num, column=3).value = disbursements_formula
            inventory_sheet.cell(row=row_num, column=3).border = border
            inventory_sheet.cell(row=row_num, column=3).alignment = alignment
            inventory_sheet.cell(row=row_num, column=3).number_format = '#,##0.00'
            
            # Formula for current balance (additions - disbursements)
            balance_formula = f"=B{row_num}-C{row_num}"
            inventory_sheet.cell(row=row_num, column=4).value = balance_formula
            inventory_sheet.cell(row=row_num, column=4).border = border
            inventory_sheet.cell(row=row_num, column=4).alignment = alignment
            inventory_sheet.cell(row=row_num, column=4).number_format = '#,##0.00'
        
        # Save the workbook
        wb.save(file_path)
        print(f"[DEBUG] Slides formulas converted successfully")
    except Exception as e:
        print(f"[ERROR] Failed to convert slides to formulas: {e}")
        traceback.print_exc()
        raise


def add_slides_inventory_entry(file_path, item_name, quantity, unit_price, notes="", entry_date=None):
    """
    Add a slides inventory entry to the additions sheet
    
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
    print(f"[DEBUG] add_slides_inventory_entry called")
    try:
        # Load workbook
        wb = openpyxl.load_workbook(file_path)
        add_sheet = wb["اذن اضافة الشرائح"]
        
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
        
        # Update inventory sheet with formulas
        convert_existing_slides_inventory_to_formulas(file_path)
        
        print(f"[DEBUG] Slides entry added successfully with number: {next_entry_number}")
        return next_entry_number
    except Exception as e:
        print(f"[ERROR] Failed to add slides inventory entry: {e}")
        traceback.print_exc()
        raise


def disburse_slides_inventory_entry(file_path, item_name, quantity, unit_price, notes="", disburse_date=None):
    """
    Add a slides inventory disbursement entry to the disbursements sheet
    
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
    print(f"[DEBUG] disburse_slides_inventory_entry called")
    try:
        # Load workbook
        wb = openpyxl.load_workbook(file_path)
        disburse_sheet = wb["اذن صرف الشرائح"]
        
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
        
        # Update inventory sheet with formulas
        convert_existing_slides_inventory_to_formulas(file_path)
        
        print(f"[DEBUG] Slides disbursement entry added successfully with number: {next_entry_number}")
        return next_entry_number
    except Exception as e:
        print(f"[ERROR] Failed to add slides disbursement entry: {e}")
        traceback.print_exc()
        raise


def get_slides_inventory_summary(file_path):
    """
    Get slides inventory summary data by evaluating formulas
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        list: List of dictionaries containing inventory data
    """
    print(f"[DEBUG] get_slides_inventory_summary called with file: {file_path}")
    if not os.path.exists(file_path):
        print(f"[DEBUG] File does not exist")
        return []
    
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)  # data_only=True to get calculated values
        inventory_sheet = wb["مخزون الشرائح"]
        
        inventory_data = []
        for row_num in range(2, inventory_sheet.max_row + 1):
            item_name = inventory_sheet.cell(row=row_num, column=1).value
            if item_name:  # Only process rows with item names
                total_additions = inventory_sheet.cell(row=row_num, column=2).value or 0
                total_disbursements = inventory_sheet.cell(row=row_num, column=3).value or 0
                current_balance = inventory_sheet.cell(row=row_num, column=4).value or 0
                
                # Convert to float and handle None values
                try:
                    total_additions = float(total_additions)
                except (ValueError, TypeError):
                    total_additions = 0
                    
                try:
                    total_disbursements = float(total_disbursements)
                except (ValueError, TypeError):
                    total_disbursements = 0
                    
                try:
                    current_balance = float(current_balance)
                except (ValueError, TypeError):
                    current_balance = 0
                
                item_data = {
                    'item_name': item_name,
                    'total_additions': total_additions,
                    'total_disbursements': total_disbursements,
                    'current_balance': current_balance
                }
                inventory_data.append(item_data)
        
        print(f"[DEBUG] Retrieved {len(inventory_data)} slides inventory items")
        return inventory_data
    except Exception as e:
        print(f"[ERROR] Error reading slides inventory summary: {e}")
        traceback.print_exc()
        return []


def get_available_slides_items_with_prices(file_path):
    """
    Get available slides items with their average unit prices from additions
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        dict: Dictionary with item names as keys and average unit prices as values
    """
    print(f"[DEBUG] get_available_slides_items_with_prices called with file: {file_path}")
    if not os.path.exists(file_path):
        print(f"[DEBUG] File does not exist")
        return {}
    
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)  # data_only=True to get calculated values
        add_sheet = wb["اذن اضافة الشرائح"]
        inventory_sheet = wb["مخزون الشرائح"]
        
        # Get current inventory balances
        inventory_balances = {}
        for row_num in range(2, inventory_sheet.max_row + 1):
            item_name = inventory_sheet.cell(row=row_num, column=1).value
            if item_name:  # Only process rows with item names
                balance = inventory_sheet.cell(row=row_num, column=4).value or 0
                try:
                    inventory_balances[item_name] = float(balance)
                except (ValueError, TypeError):
                    inventory_balances[item_name] = 0
        
        print(f"[DEBUG] Slides inventory balances: {inventory_balances}")
        
        # Get item prices from additions
        item_prices = {}
        item_quantities = {}
        
        # Skip header row (row 1)
        for row_num in range(2, add_sheet.max_row + 1):
            item_name = add_sheet.cell(row=row_num, column=3).value  # Item name column
            quantity = add_sheet.cell(row=row_num, column=4).value or 0  # Quantity column
            unit_price = add_sheet.cell(row=row_num, column=5).value or 0  # Unit price column
            
            try:
                quantity = float(quantity)
                unit_price = float(unit_price)
            except (ValueError, TypeError):
                continue  # Skip invalid rows
            
            if item_name and quantity > 0:
                # Only include items that have positive balance in inventory
                if inventory_balances.get(item_name, 0) > 0:
                    if item_name not in item_prices:
                        item_prices[item_name] = 0
                        item_quantities[item_name] = 0
                    
                    # Accumulate weighted prices
                    item_prices[item_name] += unit_price * quantity
                    item_quantities[item_name] += quantity
        
        # Calculate average prices
        avg_prices = {}
        for item_name in item_prices:
            if item_quantities[item_name] > 0:
                avg_prices[item_name] = item_prices[item_name] / item_quantities[item_name]
            else:
                avg_prices[item_name] = 0
        
        print(f"[DEBUG] Available slides items with prices: {avg_prices}")
        return avg_prices
    except Exception as e:
        print(f"[ERROR] Error getting available slides items with prices: {e}")
        traceback.print_exc()
        return []


def initialize_slides_publishing_excel(file_path):
    """
    Initialize the slides publishing Excel file with proper formatting
    
    Args:
        file_path (str): Path to the Excel file
    """
    print(f"[DEBUG] initialize_slides_publishing_excel called with file: {file_path}")
    try:
        # Create workbook
        wb = Workbook()
        
        # Remove default sheet
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        # Create publishing sheet
        publish_sheet = wb.create_sheet("اذن اضافه", 0)
        
        # Set right-to-left
        publish_sheet.sheet_view.rightToLeft = True
        
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
        
        # Add headers to publishing sheet
        publish_headers = [
            "تاريخ النشر", "رقم البلوك", "النوع", "رقم المكينه", "عدد", 
            "الطول", "الارتفاع", "السمك", "م2", "سعر المتر", "اجمالي السعر", 
            "وقت الدخول", "وقت الخروج", "عدد الساعات"
        ]
        for col_num, header in enumerate(publish_headers, 1):
            cell = publish_sheet.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border
        
        # Set column widths for publishing sheet
        column_widths = [15, 12, 15, 12, 10, 12, 12, 12, 12, 15, 15, 15, 15, 15]
        for col_num, width in enumerate(column_widths, 1):
            publish_sheet.column_dimensions[get_column_letter(col_num)].width = width
        
        # Save the workbook
        wb.save(file_path)
        print(f"[DEBUG] Slides publishing Excel file initialized successfully")
        return wb
    except Exception as e:
        print(f"[ERROR] Failed to initialize slides publishing Excel file: {e}")
        traceback.print_exc()
        raise


def add_slides_publishing_entry(file_path, publishing_data):
    """
    Add a slides inventory entry to the publishing sheet
    
    Args:
        file_path (str): Path to the Excel file
        publishing_data (list): List of dictionaries containing publishing data
        
    Returns:
        int: Number of entries added
    """
    print(f"[DEBUG] add_slides_publishing_entry called")
    try:
        # Load workbook
        wb = openpyxl.load_workbook(file_path)
        publish_sheet = wb["اذن اضافه"]
        
        # Add data rows
        for entry in publishing_data:
            row_data = [
                entry.get('publishing_date', ''),
                entry.get('block_number', ''),
                entry.get('material', ''),
                entry.get('machine_number', ''),
                int(entry.get('quantity', 0)),
                float(entry.get('length', 0)),
                float(entry.get('height', 0)),
                entry.get('thickness', ''),
                '',  # Area will be calculated as formula
                float(entry.get('price_per_meter', 0)),
                '',  # Total price will be calculated as formula
                entry.get('entry_time', ''),
                entry.get('exit_time', ''),
                float(entry.get('hours_count', 0))
            ]
            
            # Add row to sheet
            publish_sheet.append(row_data)
            
            # Apply styles to the new row
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            alignment = Alignment(horizontal='center', vertical='center')
            
            row_num = publish_sheet.max_row
            for col_num, value in enumerate(row_data, 1):
                cell = publish_sheet.cell(row=row_num, column=col_num)
                cell.border = border
                cell.alignment = alignment
                # Convert quantity (col 5) to int, keep others as float
                if col_num == 5:  # Quantity column should be integer
                    cell.value = int(float(value)) if value else 0
                # Apply number formatting for numeric columns
                if col_num in [5, 6, 7, 9, 10, 11, 14]:  # Quantity, Length, Height, Area, Price, Total, Hours
                    if col_num == 5:  # Quantity column should be integer
                        cell.number_format = '#,##0'
                    else:
                        cell.number_format = '#,##0.00'
            
            # Add formulas for Area (م2) and Total Price
            # Area = Quantity * Length * Height (Column 5 * Column 6 * Column 7)
            area_formula = f'={get_column_letter(5)}{row_num}*{get_column_letter(6)}{row_num}*{get_column_letter(7)}{row_num}'
            publish_sheet.cell(row=row_num, column=9, value=area_formula)  # Column I (9) is Area
            
            # Total Price = Area * Price per meter (Column 9 * Column 10)
            total_price_formula = f'={get_column_letter(9)}{row_num}*{get_column_letter(10)}{row_num}'
            publish_sheet.cell(row=row_num, column=11, value=total_price_formula)  # Column K (11) is Total Price
        
        # Save the workbook
        wb.save(file_path)
        
        print(f"[DEBUG] {len(publishing_data)} slides inventory entries added successfully")
        return len(publishing_data)
    except Exception as e:
        print(f"[ERROR] Failed to add slides publishing entries: {e}")
        traceback.print_exc()
        raise