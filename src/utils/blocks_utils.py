import xlsxwriter
from typing import List, Dict
import os
import openpyxl
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# Table column definitions
TABLE1_COLUMNS = [
    "رقم النقله", "عدد النقله", "التاريخ", "المحجر", 
    "رقم البلوك", "الخامه", "الطول", 
    "العرض", "الارتفاع", "م3", "الوزن", 
    "وزن البلوك", "سعر الطن", "اجمالي السعر"
]



# Column width definitions for each table
TABLE1_WIDTH = [12, 10, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12]

def export_simple_blocks_excel(rows: List[Dict]) -> str:
    """إنشاء أو تحديث ملف Excel لحساب مخزون البلوكات"""
    documents_folder = os.path.join(
        os.path.expanduser("~"), "Documents", "alswaife", "البلوكات"
    )
    os.makedirs(documents_folder, exist_ok=True)
    
    filepath = os.path.join(documents_folder, "مخزون البلوكات.xlsx")
    
    if os.path.exists(filepath):
        result = append_to_existing_file(filepath, rows)
        if result is False:  # File is locked
            raise PermissionError("File is currently open in Excel. Please close the file and try again.")
    else:
        result = create_new_excel_file(filepath, rows)
        if result is False:  # File is locked
            raise PermissionError("File is currently open in Excel. Please close the file and try again.")
    
    return filepath


def append_to_existing_file(filepath: str, new_rows: List[Dict]):
    """
    Append new rows to an existing Excel file.
    
    This function adds new block data to an existing Excel file while maintaining
    the existing structure and formatting. It handles one table with specific
    column arrangements and formulas.
    
    Args:
        filepath (str): Path to the existing Excel file
        new_rows (List[Dict]): List of dictionaries containing block data
    
    Returns:
        bool: True if successful, False if file is locked
    """
    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook["البلوكات"]
        
        # Starting row for new data (after existing data)
        start_row = worksheet.max_row + 1
        
        # Define styles
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')
        gap_fill = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")  # Gray for gaps
        # Custom style for headers
        header_fill_table1 = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")  # Dark blue for Table 1 headers
        header_font = Font(color="FFFFFF", bold=True)  # White bold font for headers
        
        # Define different colors for columns with improved contrast
        column_colors = [
            "FFB3BA",  # Light Red-Pink
            "BAFFC9",  # Light Mint Green
            "BAE1FF",  # Light Blue
            "FFFFBA",  # Light Yellow
            "FFDFBA",  # Light Orange
            "E4BAFF",  # Light Purple
            "FFC9DE",  # Light Pink
            "C9FFE4",  # Light Aqua
            "C9D6FF",  # Light Indigo
            "FFE8C9",  # Light Peach
            "D6FFC9",  # Light Lime
            "FFC9F3",  # Light Magenta
            "C9FFF3",  # Light Turquoise
            "F3C9FF",  # Light Lavender
        ]
        

        # Add new data
        for i, block_data in enumerate(new_rows):
            excel_row = start_row + i
            
            # --- الجدول الأول ---
            # Write values and set formulas for calculated fields
            # Column 1: رقم النقله
            cell = worksheet.cell(row=excel_row, column=1, value=block_data.get("trip_number", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[0], end_color=column_colors[0], fill_type="solid")
            
            # Column 2: عدد النقله
            cell = worksheet.cell(row=excel_row, column=2, value=block_data.get("trip_count", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[1], end_color=column_colors[1], fill_type="solid")
            
            # Column 3: التاريخ
            cell = worksheet.cell(row=excel_row, column=3, value=block_data.get("date", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[2], end_color=column_colors[2], fill_type="solid")
            
            # Column 4: المحجر
            cell = worksheet.cell(row=excel_row, column=4, value=block_data.get("quarry", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[3], end_color=column_colors[3], fill_type="solid")
            
            # Column 5: رقم البلوك
            cell = worksheet.cell(row=excel_row, column=5, value=block_data.get("block_number", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[4], end_color=column_colors[4], fill_type="solid")
            
            # Column 6: الخامه
            cell = worksheet.cell(row=excel_row, column=6, value=block_data.get("material", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[5], end_color=column_colors[5], fill_type="solid")
            
            # Column 7: الطول
            cell = worksheet.cell(row=excel_row, column=7, value=block_data.get("length", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[6], end_color=column_colors[6], fill_type="solid")
            
            # Column 8: العرض
            cell = worksheet.cell(row=excel_row, column=8, value=block_data.get("width", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[7], end_color=column_colors[7], fill_type="solid")
            
            # Column 9: الارتفاع
            cell = worksheet.cell(row=excel_row, column=9, value=block_data.get("height", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[8], end_color=column_colors[8], fill_type="solid")
            
            # Column 10: م3 (volume) - Calculate using formula: length * width * height
            length_cell = get_column_letter(7)
            width_cell = get_column_letter(8)
            height_cell = get_column_letter(9)
            formula_str = f'={length_cell}{excel_row}*{width_cell}{excel_row}*{height_cell}{excel_row}'
            cell = worksheet.cell(row=excel_row, column=10)
            cell.value = formula_str
            cell.data_type = 'f'  # Set as formula
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[9], end_color=column_colors[9], fill_type="solid")
            
            # Column 11: الوزن
            cell = worksheet.cell(row=excel_row, column=11, value=block_data.get("weight_per_m3", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[10], end_color=column_colors[10], fill_type="solid")
            
            # Column 12: وزن البلوك - Calculate using formula: volume * weight_per_m3
            volume_cell = get_column_letter(10)
            weight_per_m3_cell = get_column_letter(11)
            formula_str = f'={volume_cell}{excel_row}*{weight_per_m3_cell}{excel_row}'
            cell = worksheet.cell(row=excel_row, column=12)
            cell.value = formula_str
            cell.data_type = 'f'  # Set as formula
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[11], end_color=column_colors[11], fill_type="solid")
            
            # Column 13: سعر الطن
            cell = worksheet.cell(row=excel_row, column=13, value=block_data.get("price_per_ton", ""))
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[12], end_color=column_colors[12], fill_type="solid")
            
            # Column 14: اجمالي السعر - Calculate using formula: block_weight * price_per_ton
            block_weight_cell = get_column_letter(12)
            price_per_ton_cell = get_column_letter(13)
            formula_str = f'={block_weight_cell}{excel_row}*{price_per_ton_cell}{excel_row}'
            cell = worksheet.cell(row=excel_row, column=14)
            cell.value = formula_str
            cell.data_type = 'f'  # Set as formula
            cell.border = thin_border
            cell.alignment = center_alignment
            cell.fill = PatternFill(start_color=column_colors[13], end_color=column_colors[13], fill_type="solid")
            
            
        workbook.save(filepath)
        
    except PermissionError as e:
        print(f"File is locked: {e}")
        return False  # File is locked
    except Exception as e:
        print(f"Error adding data: {e}")
        # In case of severe error, can try to recreate, but cautiously
        # create_new_excel_file(filepath, new_rows)
    return True 

def create_new_excel_file(filepath: str, rows: List[Dict]):
    """
    Create a new Excel file with block data.
    
    This function creates a new Excel file with one table containing block data.
    It sets up the proper structure, formatting, and formulas for the table.
    
    Args:
        filepath (str): Path where the new Excel file should be created
        rows (List[Dict]): List of dictionaries containing block data
    
    Returns:
        bool: True if successful, False if file is locked
    """
    workbook = xlsxwriter.Workbook(filepath, {'constant_memory': False})
    worksheet = workbook.add_worksheet("البلوكات")
    worksheet.right_to_left()

    # تعريف ألوان مختلفة للأعمدة مع تحسين التباين
    column_colors = [
        "#FFB3BA",  # Light Red-Pink
        "#BAFFC9",  # Light Mint Green
        "#BAE1FF",  # Light Blue
        "#FFFFBA",  # Light Yellow
        "#FFDFBA",  # Light Orange
        "#E4BAFF",  # Light Purple
        "#FFC9DE",  # Light Pink
        "#C9FFE4",  # Light Aqua
        "#C9D6FF",  # Light Indigo
        "#FFE8C9",  # Light Peach
        "#D6FFC9",  # Light Lime
        "#FFC9F3",  # Light Magenta
        "#C9FFF3",  # Light Turquoise
        "#F3C9FF",  # Light Lavender
    ]

    # أنماط محسّنة للعناوين والجداول
    title_fmt = workbook.add_format({
        "bold": True, 
        "font_size": 18, 
        "align": "center", 
        "valign": "vcenter", 
        "bg_color": "#1F4E79",  # أزرق داكن أكثر احترافية
        "font_color": "white", 
        "border": 1
    })
    
    table1_title_fmt = workbook.add_format({
        "bold": True, 
        "font_size": 14, 
        "align": "center", 
        "valign": "vcenter", 
        "bg_color": "#4472C4",  # أزرق متوسط لجدول 1
        "font_color": "white", 
        "border": 1
    })
    
    
    gap_fmt = workbook.add_format({
        "bg_color": "#A6A6A6",  # رمادي للفراغات
        "border": 1
    })
    
    # تعريف أنماط ملونة للبيانات مع تحسين التباين
    data_formats = []
    for i, color in enumerate(column_colors):
        fmt = workbook.add_format({
            "border": 1, 
            "align": "center", 
            "valign": "vcenter", 
            "font_size": 11, 
            "bg_color": color
        })
        data_formats.append(fmt)

    # Write title
    total_cols = len(TABLE1_COLUMNS)
    worksheet.merge_range(0, 0, 0, len(TABLE1_COLUMNS) - 1, "مقاس البلوك علي الارضية", table1_title_fmt)

    # Write column headers with distinct colors
    header_fmt_table1 = workbook.add_format({
        "bold": True, 
        "border": 1, 
        "align": "center", 
        "valign": "vcenter", 
        "bg_color": "#2F5597",  # Dark blue for Table 1 headers
        "font_color": "white", 
        "font_size": 12
    })
    
    
    for idx, col in enumerate(TABLE1_COLUMNS):
        worksheet.write(1, idx, col, header_fmt_table1)

    worksheet.set_row(0, 30)
    worksheet.set_row(1, 25)

    start_row = 2
    
    for i, block_data in enumerate(rows):
        excel_row = start_row + i
        
        # --- الجدول الأول ---
        # Write values and formulas for calculations
        worksheet.write(excel_row, 0, block_data.get("trip_number", ""), data_formats[0])  # رقم النقله
        worksheet.write(excel_row, 1, block_data.get("trip_count", ""), data_formats[1])  # عدد النقله
        worksheet.write(excel_row, 2, block_data.get("date", ""), data_formats[2])  # التاريخ
        worksheet.write(excel_row, 3, block_data.get("quarry", ""), data_formats[3])  # المحجر
        worksheet.write(excel_row, 4, block_data.get("block_number", ""), data_formats[4])  # رقم البلوك
        worksheet.write(excel_row, 5, block_data.get("material", ""), data_formats[5])  # الخامه
        worksheet.write(excel_row, 6, block_data.get("length", ""), data_formats[6])  # الطول
        worksheet.write(excel_row, 7, block_data.get("width", ""), data_formats[7])  # العرض
        worksheet.write(excel_row, 8, block_data.get("height", ""), data_formats[8])  # الارتفاع
        
        # Calculate volume (م3) = length * width * height using Excel formula
        length_col = get_column_letter(6 + 1)  # Column F (0-indexed)
        width_col = get_column_letter(7 + 1)   # Column G (0-indexed)
        height_col = get_column_letter(8 + 1)  # Column H (0-indexed)
        volume_formula = f'={length_col}{excel_row + 1}*{width_col}{excel_row + 1}*{height_col}{excel_row + 1}'
        worksheet.write_formula(excel_row, 9, volume_formula, data_formats[9])  # م3
        
        worksheet.write(excel_row, 10, block_data.get("weight_per_m3", ""), data_formats[10])  # الوزن
        
        # Calculate block weight (وزن البلوك) = volume * weight_per_m3 using Excel formula
        volume_col = get_column_letter(9 + 1)  # Column J (0-indexed)
        weight_per_m3_col = get_column_letter(10 + 1)  # Column K (0-indexed)
        weight_formula = f'={volume_col}{excel_row + 1}*{weight_per_m3_col}{excel_row + 1}'
        worksheet.write_formula(excel_row, 11, weight_formula, data_formats[11])  # وزن البلوك
        
        worksheet.write(excel_row, 12, block_data.get("price_per_ton", ""), data_formats[12])  # سعر الطن
        
        # Calculate total price (اجمالي السعر) = block_weight * price_per_ton using Excel formula
        block_weight_col = get_column_letter(11 + 1)  # Column L (0-indexed)
        price_per_ton_col = get_column_letter(12 + 1)  # Column M (0-indexed)
        total_price_formula = f'={block_weight_col}{excel_row + 1}*{price_per_ton_col}{excel_row + 1}'
        worksheet.write_formula(excel_row, 13, total_price_formula, data_formats[13])  # اجمالي السعر
    
        
    worksheet.freeze_panes(2, 0)
    
    # تنسيق عرض الأعمدة
    for i in range(min(len(TABLE1_COLUMNS), len(TABLE1_WIDTH))): 
        worksheet.set_column(i, i, TABLE1_WIDTH[i])
    
    try:
        workbook.close()
        return True
    except PermissionError as e:
        print(f"File is locked: {e}")
        return False  # File is locked