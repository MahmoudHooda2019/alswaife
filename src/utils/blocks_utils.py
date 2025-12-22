import xlsxwriter
from typing import List, Dict
import os
import openpyxl
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# Table column definitions
TABLE1_COLUMNS = [
    "رقم النقله", "عدد النقله", "التاريخ", "المحجر", 
    "رقم الماكينة", "رقم البلوك", "الخام", "الطول", 
    "العرض", "الارتفاع", "م3", "الوزن", 
    "وزن البلوك", "السعر للطن", "الإجمالي", "سعر الرحلة"
]

TABLE2_COLUMNS = [
    "تاريخ النشر", "رقم البلوك", "النوع", "رقم الماكينة", "وقت الدخول",
    "وقت الخروج", "عدد الساعات", "الاكراميه", "السمك", "العدد",
    "الطول", "الطول بعد", "الخصم", "الارتفاع", "الكميه م2",
    "سعر النشر", "إجمالي سعر النشر", "إجمالي تكلفه البلوك"
]



# Column width definitions for each table
TABLE1_WIDTH = [12, 8, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12]
TABLE2_WIDTH = [12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 16, 18]



def export_simple_blocks_excel(rows: List[Dict]) -> str:
    """إنشاء أو تحديث ملف Excel لحساب تكلفه البلوكات مصنع محب"""
    documents_folder = os.path.join(
        os.path.expanduser("~"), "Documents", "alswaife", "البلوكات"
    )
    os.makedirs(documents_folder, exist_ok=True)
    
    filepath = os.path.join(documents_folder, "حساب تكلفه البلوكات مصنع محب.xlsx")
    
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
    the existing structure and formatting. It handles three tables with specific
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
        header_fill_table2 = PatternFill(start_color="548235", end_color="548235", fill_type="solid")  # Dark green for Table 2 headers
        # Table 3 has been removed as per user request
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
            "FFDEC9",  # Light Salmon
            "C9E7FF",  # Light Sky Blue
        ]
        
        # Additional colors for repetition with better contrast
        additional_colors = [
            "FF9AA2",  # Medium Red-Pink
            "B5FFAD",  # Medium Green
            "ADD8FF",  # Medium Blue
            "FFFFAD",  # Medium Yellow
            "FFD8AD",  # Medium Orange
            "DBADFF",  # Medium Purple
            "FFADD8",  # Medium Pink
            "ADF8D8",  # Medium Aqua
            "ADBFFF",  # Medium Indigo
            "FFD6AD",  # Medium Peach
            "D1FFAD",  # Medium Lime
            "FFADE8",  # Medium Magenta
            "ADF8E8",  # Medium Turquoise
            "E8ADFF",  # Medium Lavender
            "FFD1AD",  # Medium Salmon
            "ADE0FF",  # Medium Sky Blue
        ]

        # Add new data
        for i, block_data in enumerate(new_rows):
            excel_row = start_row + i
            
            thickness_text = block_data.get("thickness_dropdown", "2سم") or "2سم"
            
            # --- الجدول الأول (بدون تغيير) ---
            table1_data = [
                block_data.get("trip_number", ""),
                block_data.get("trip_count", ""),
                block_data.get("date", ""),
                block_data.get("quarry", ""),
                block_data.get("machine_number", ""),
                block_data.get("block_number", ""),
                block_data.get("material", ""),
                "",  # الطول (معادلة) - سيتم حسابه من UI + 0.20
                "",  # العرض (معادلة)
                "",  # الارتفاع (معادلة) - سيتم حسابه من مرحلة النشر
                "",  # م3 (معادلة)
                block_data.get("weight", ""),
                "",  # وزن البلوك (معادلة)
                block_data.get("price_per_ton", ""),
                "",  # إجمالي السعر (معادلة)
                block_data.get("trip_price", "")
            ]
            
            for col_idx, value in enumerate(table1_data, start=1):
                cell = worksheet.cell(row=excel_row, column=col_idx, value=value)
                cell.border = thin_border
                cell.alignment = center_alignment
                # Apply different color to each column
                color_index = (col_idx - 1) % len(column_colors)
                cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            
            # Length formula: length from UI + 0.20
            # Store the raw length value in a temporary cell and create a formula referencing it
            length_value = block_data.get("length", "")
            if length_value != "":
                # Put the raw length value in a temporary cell within our table structure
                # Use the last column of Table 1 (column 16, 0-indexed = 15) to store the raw length
                temp_cell = worksheet.cell(row=excel_row, column=16, value=length_value)
                temp_cell.border = thin_border
                temp_cell.fill = PatternFill(start_color=column_colors[15], end_color=column_colors[15], fill_type="solid")
                # Create formula that references this cell and adds 0.20
                length_formula = f'={get_column_letter(16)}{excel_row}+0.20'
                length_cell = worksheet.cell(row=excel_row, column=8, value=length_formula)
                length_cell.border = thin_border
                color_index = 7 % len(column_colors)
                length_cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            else:
                empty_cell = worksheet.cell(row=excel_row, column=8, value="")
                empty_cell.border = thin_border
                color_index = 7 % len(column_colors)
                empty_cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            
            # معادلات الجدول الأول (نفس كودك السابق)
            thickness_col = get_column_letter(25) # Y (السمك الآن في العمود 25)
            count_col = get_column_letter(26)     # Z (العدد الآن في العمود 26)
            count_col_fixed = get_column_letter(26) # Z هو موقع العدد الجديد
            
            width_formula = f'=((VALUE(SUBSTITUTE({thickness_col}{excel_row},"سم",""))+1)*{count_col_fixed}{excel_row})'
            width_cell = worksheet.cell(row=excel_row, column=9, value=width_formula)
            width_cell.border = thin_border
            width_cell.alignment = center_alignment
            color_index = 8 % len(column_colors)
            width_cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            
            # Height formula: height in publishing stage + 0.5
            # Publishing stage height is in column 29 (AC)
            publish_height_col = get_column_letter(29)
            height_formula = f'={publish_height_col}{excel_row}+0.5'
            height_cell = worksheet.cell(row=excel_row, column=10, value=height_formula)
            height_cell.border = thin_border
            color_index = 9 % len(column_colors)
            height_cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            
            length_col = get_column_letter(8)
            width_col = get_column_letter(9)
            height_col = get_column_letter(10)
            volume_formula = f"={length_col}{excel_row}*{width_col}{excel_row}*{height_col}{excel_row}"
            volume_cell = worksheet.cell(row=excel_row, column=11, value=volume_formula)
            volume_cell.border = thin_border
            color_index = 10 % len(column_colors)
            volume_cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            
            m3_col = get_column_letter(11)
            weight_col = get_column_letter(12)
            weight_formula = f"={m3_col}{excel_row}*{weight_col}{excel_row}"
            weight_cell = worksheet.cell(row=excel_row, column=13, value=weight_formula)
            weight_cell.border = thin_border
            color_index = 12 % len(column_colors)
            weight_cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            
            price_col = get_column_letter(14)
            blk_weight_col = get_column_letter(13)
            total_price_formula = f"={price_col}{excel_row}*{blk_weight_col}{excel_row}"
            total_price_cell = worksheet.cell(row=excel_row, column=15, value=total_price_formula)
            total_price_cell.border = thin_border
            color_index = 14 % len(column_colors)
            total_price_cell.fill = PatternFill(start_color=column_colors[color_index], end_color=column_colors[color_index], fill_type="solid")
            
            # إزالة كتابة الخانة الفارغة بين الجدول الأول والثاني
            # gap_cell = worksheet.cell(row=excel_row, column=17, value="")
            # gap_cell.border = thin_border
            # gap_cell.fill = gap_fill
            
            # --- الجدول الثاني (تم التصحيح هنا) ---
            # البداية من العمود 17 (بدلاً من 18 بعد إزالة الخانة الفارغة)
            
            # 17-25: البيانات الأساسية (مع تحديث الأعمود)
            date_cell = worksheet.cell(row=excel_row, column=17, value=block_data.get("date", ""))
            date_cell.border = thin_border
            color_index = 16 % len(additional_colors)
            date_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            block_cell = worksheet.cell(row=excel_row, column=18, value=block_data.get("block_number", ""))
            block_cell.border = thin_border
            color_index = 17 % len(additional_colors)
            block_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            material_cell = worksheet.cell(row=excel_row, column=19, value=block_data.get("material", ""))
            material_cell.border = thin_border
            color_index = 18 % len(additional_colors)
            material_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            machine_cell = worksheet.cell(row=excel_row, column=20, value=block_data.get("machine_number", ""))
            machine_cell.border = thin_border
            color_index = 19 % len(additional_colors)
            machine_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            entry_cell = worksheet.cell(row=excel_row, column=21, value="")  # دخول
            entry_cell.border = thin_border
            color_index = 20 % len(additional_colors)
            entry_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            exit_cell = worksheet.cell(row=excel_row, column=22, value="")  # خروج
            exit_cell.border = thin_border
            color_index = 21 % len(additional_colors)
            exit_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            hours_cell = worksheet.cell(row=excel_row, column=23, value="")  # ساعات
            hours_cell.border = thin_border
            color_index = 22 % len(additional_colors)
            hours_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            tips_cell = worksheet.cell(row=excel_row, column=24, value=150)  # إكرامية
            tips_cell.border = thin_border
            color_index = 23 % len(additional_colors)
            tips_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            thickness_cell = worksheet.cell(row=excel_row, column=25, value=thickness_text)  # السمك
            thickness_cell.border = thin_border
            color_index = 24 % len(additional_colors)
            thickness_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 26: Quantity (was previously placed in a gap column incorrectly)
            quantity_cell = worksheet.cell(row=excel_row, column=26, value=block_data.get("quantity", 1))
            quantity_cell.border = thin_border
            color_index = 25 % len(additional_colors)
            quantity_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 27: Length (copied from Table 1)
            # Get length value from Table 1 (column 8)
            table1_length_col = get_column_letter(8)
            length_value = block_data.get("length", "")
            length_cell = worksheet.cell(row=excel_row, column=27, value=length_value)
            length_cell.border = thin_border
            color_index = 26 % len(additional_colors)
            length_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 28: Discount
            discount_cell = worksheet.cell(row=excel_row, column=28, value=0.20)
            discount_cell.border = thin_border
            color_index = 27 % len(additional_colors)
            discount_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 29: Length After = Length from Table 2 - Discount
            # Using length from Table 2 (column 27) instead of Table 1
            table2_length_col = get_column_letter(27)  # Column containing length in Table 2
            disc_col = get_column_letter(28)  # AB (Discount in column 28)
            length_after_formula = f"={table2_length_col}{excel_row}-{disc_col}{excel_row}"
            length_after_cell = worksheet.cell(row=excel_row, column=29, value=length_after_formula)
            length_after_cell.border = thin_border
            length_after_cell.alignment = center_alignment
            color_index = 28 % len(additional_colors)
            length_after_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 30: Height (Publishing stage height)
            publish_height = block_data.get("publish_height", float(block_data.get("height", 0) or 0))
            height_cell = worksheet.cell(row=excel_row, column=30, value=publish_height)
            height_cell.border = thin_border
            color_index = 29 % len(additional_colors)
            height_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 31: Quantity m2 (Formula) = Length After × Height × Quantity
            # Note: Usually calculated based on "Length After" (net). 
            # If you want to calculate based on raw length, change column 29 to 28
            len_after_col = get_column_letter(29)  # AC (Length After is now in column 29)
            height_pub_col = get_column_letter(30)  # AD
            qty_col = get_column_letter(26)  # Z (Quantity)
            qty_m2_formula = f"={len_after_col}{excel_row}*{height_pub_col}{excel_row}*{qty_col}{excel_row}"
            qty_m2_cell = worksheet.cell(row=excel_row, column=31, value=qty_m2_formula)
            qty_m2_cell.border = thin_border
            qty_m2_cell.alignment = center_alignment
            color_index = 30 % len(additional_colors)
            qty_m2_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 32: Publish Price (based on thickness)
            thickness_value = thickness_text.replace("سم", "")
            if thickness_value == "2":
                publish_price = 125
            elif thickness_value == "3":
                publish_price = 145
            elif thickness_value == "4":
                publish_price = 155
            else:
                publish_price = 125
            price_cell = worksheet.cell(row=excel_row, column=32, value=publish_price)
            price_cell.border = thin_border
            color_index = 31 % len(additional_colors)
            price_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 33: Total Publish Price (Formula) = Publish Price × Quantity m2
            pub_price_col = get_column_letter(32)  # AF
            qty_m2_col = get_column_letter(31)  # AE
            total_publish_formula = f"={pub_price_col}{excel_row}*{qty_m2_col}{excel_row}"
            total_publish_cell = worksheet.cell(row=excel_row, column=33, value=total_publish_formula)
            total_publish_cell.border = thin_border
            total_publish_cell.alignment = center_alignment
            color_index = 32 % len(additional_colors)
            total_publish_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # 34: Total Block Cost (Formula) = Total Publish Price + Tips
            tot_pub_col = get_column_letter(33)  # AG
            tips_col = get_column_letter(24)  # X
            total_cost_formula = f"={tot_pub_col}{excel_row}+{tips_col}{excel_row}"
            total_cost_cell = worksheet.cell(row=excel_row, column=34, value=total_cost_formula)
            total_cost_cell.border = thin_border
            total_cost_cell.alignment = center_alignment
            color_index = 33 % len(additional_colors)
            total_cost_cell.fill = PatternFill(start_color=additional_colors[color_index], end_color=additional_colors[color_index], fill_type="solid")
            
            # Table 3 has been removed as per user request
            
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
    
    This function creates a new Excel file with three tables containing block data.
    It sets up the proper structure, formatting, and formulas for all tables.
    
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
        "#FFDEC9",  # Light Salmon
        "#C9E7FF",  # Light Sky Blue
    ]
    
    # ألوان إضافية للتكرار مع تباين أفضل
    additional_colors = [
        "#FF9AA2",  # Medium Red-Pink
        "#B5FFAD",  # Medium Green
        "#ADD8FF",  # Medium Blue
        "#FFFFAD",  # Medium Yellow
        "#FFD8AD",  # Medium Orange
        "#DBADFF",  # Medium Purple
        "#FFADD8",  # Medium Pink
        "#ADF8D8",  # Medium Aqua
        "#ADBFFF",  # Medium Indigo
        "#FFD6AD",  # Medium Peach
        "#D1FFAD",  # Medium Lime
        "#FFADE8",  # Medium Magenta
        "#ADF8E8",  # Medium Turquoise
        "#E8ADFF",  # Medium Lavender
        "#FFD1AD",  # Medium Salmon
        "#ADE0FF",  # Medium Sky Blue
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
    
    table2_title_fmt = workbook.add_format({
        "bold": True, 
        "font_size": 14, 
        "align": "center", 
        "valign": "vcenter", 
        "bg_color": "#70AD47",  # أخضر لجدول 2
        "font_color": "white", 
        "border": 1
    })
    
    # Table 3 has been removed as per user request
    
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
    
    # أنماط إضافية للأعمدة المتكررة مع تباين أفضل
    additional_formats = []
    for i, color in enumerate(additional_colors):
        fmt = workbook.add_format({
            "border": 1, 
            "align": "center", 
            "valign": "vcenter", 
            "font_size": 11, 
            "bg_color": color
        })
        additional_formats.append(fmt)

    # Write title
    total_cols = len(TABLE1_COLUMNS) + len(TABLE2_COLUMNS)
    worksheet.merge_range(0, 0, 0, total_cols - 1, "حساب تكلفه البلوكات مصنع محب", title_fmt)
    worksheet.merge_range(1, 0, 1, len(TABLE1_COLUMNS) - 1, "مقاس البلوك في الأرضية", table1_title_fmt)
    worksheet.merge_range(1, len(TABLE1_COLUMNS), 1, len(TABLE1_COLUMNS) + len(TABLE2_COLUMNS) - 1, "مرحلة النشر", table2_title_fmt)
    # Table 3 has been removed as per user request
    
    # Table 3 has been removed as per user request

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
    
    header_fmt_table2 = workbook.add_format({
        "bold": True, 
        "border": 1, 
        "align": "center", 
        "valign": "vcenter", 
        "bg_color": "#548235",  # Dark green for Table 2 headers
        "font_color": "white", 
        "font_size": 12
    })
    
    # Table 3 has been removed as per user request
    
    for idx, col in enumerate(TABLE1_COLUMNS):
        worksheet.write(2, idx, col, header_fmt_table1)
    for idx, col in enumerate(TABLE2_COLUMNS):
        worksheet.write(2, len(TABLE1_COLUMNS) + idx, col, header_fmt_table2)
    # Table 3 has been removed as per user request
    worksheet.set_row(1, 30)
    worksheet.set_row(2, 25)

    start_row = 3
    # تحديث مواضع الأعمدة بعد إزالة الخانة الفارغة
    col_offset = len(TABLE1_COLUMNS)  # بدون إضافة 1 لأننا أزلنا الخانة الفارغة
    # Table 3 has been removed as per user request
    
    for i, block_data in enumerate(rows):
        excel_row = start_row + i
        thickness_text = block_data.get("thickness_dropdown", "2سم") or "2سم"
        
        # --- الجدول الأول ---
        table1_values = [
            block_data.get("trip_number", ""), block_data.get("trip_count", ""), block_data.get("date", ""), 
            block_data.get("quarry", ""), block_data.get("machine_number", ""), block_data.get("block_number", ""),
            block_data.get("material", ""), "", "", "",
            "", block_data.get("weight", ""), "", block_data.get("price_per_ton", ""), "", block_data.get("trip_price", "")
        ]
        
        for col_idx, value in enumerate(table1_values):
            # تطبيق لون مختلف لكل عمود
            color_index = col_idx % len(data_formats)
            worksheet.write(excel_row, col_idx, value, data_formats[color_index])
        
        # Length formula: length from UI + 0.20
        length_value = block_data.get("length", "")
        if length_value != "":
            # Put the raw length value in a temporary cell within our table structure
            # Use the last column of Table 1 (column 16, 0-indexed = 15) to store the raw length
            color_index = 15 % len(data_formats)
            worksheet.write(excel_row, 15, length_value, data_formats[color_index])  # Column 16 (0-indexed = 15)
            # Create formula that references this cell and adds 0.20
            length_formula = f'={get_column_letter(16)}{excel_row + 1}+0.20'
            color_index = 7 % len(data_formats)
            worksheet.write_formula(excel_row, 7, length_formula, data_formats[color_index])  # Column 8 (0-indexed = 7)
        
        # معادلات الجدول الأول وتصحيح مرجع العدد
        # العدد موجود الآن في col_offset + 9 (index 9 في TABLE2_COLUMNS)
        thickness_col = get_column_letter(26) # Z
        count_col_fixed = get_column_letter(col_offset + 9 + 1) # (offset + index + 1 for A1 notation)
        
        width_formula = f'=((VALUE(SUBSTITUTE({thickness_col}{excel_row + 1},"سم",""))+1)*{count_col_fixed}{excel_row + 1})'
        color_index = 8 % len(data_formats)
        worksheet.write_formula(excel_row, 8, width_formula, data_formats[color_index])
        
        # Height formula: height in publishing stage + 0.5
        # Publishing stage height is in col_offset + 13
        publish_height_col = get_column_letter(col_offset + 13 + 1)
        height_formula = f'={publish_height_col}{excel_row + 1}+0.5'
        color_index = 9 % len(data_formats)
        worksheet.write_formula(excel_row, 9, height_formula, data_formats[color_index])
        
        # باقي معادلات الجدول الأول
        l_col = get_column_letter(8)
        w_col = get_column_letter(9)
        h_col = get_column_letter(10)
        volume_formula = f"={l_col}{excel_row + 1}*{w_col}{excel_row + 1}*{h_col}{excel_row + 1}"
        color_index = 10 % len(data_formats)
        worksheet.write_formula(excel_row, 10, volume_formula, data_formats[color_index])
        
        m3_col = get_column_letter(11)
        wt_col = get_column_letter(12)
        weight_formula = f"={m3_col}{excel_row + 1}*{wt_col}{excel_row + 1}"
        color_index = 12 % len(data_formats)
        worksheet.write_formula(excel_row, 12, weight_formula, data_formats[color_index])
        
        pr_col = get_column_letter(14)
        bw_col = get_column_letter(13)
        total_price_formula = f"={pr_col}{excel_row + 1}*{bw_col}{excel_row + 1}"
        color_index = 14 % len(data_formats)
        worksheet.write_formula(excel_row, 14, total_price_formula, data_formats[color_index])
            
        # إزالة كتابة الخانة الفارغة بين الجدول الأول والثاني
        # worksheet.write(excel_row, len(TABLE1_COLUMNS), "", gap_fmt)
        
        # --- الجدول الثاني (تم التصحيح) ---
        # 0: Date, 1: Block, 2: Material, 3: Machine
        color_index = 0 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 0, block_data.get("date", ""), additional_formats[color_index])
        color_index = 1 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 1, block_data.get("block_number", ""), additional_formats[color_index])
        color_index = 2 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 2, block_data.get("material", ""), additional_formats[color_index])
        color_index = 3 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 3, block_data.get("machine_number", ""), additional_formats[color_index])
        
        # 4,5,6: Time In, Out, Hours
        color_index = 4 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 4, "", additional_formats[color_index])
        color_index = 5 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 5, "", additional_formats[color_index])
        color_index = 6 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 6, "", additional_formats[color_index])
        
        # 7: Tips
        color_index = 7 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 7, 150, additional_formats[color_index])
        
        # Column 8: Block thickness
        color_index = 8 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 8, thickness_text, additional_formats[color_index])
        
        # Column 9: Quantity (removed gap column and placed quantity here)
        color_index = 9 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 9, block_data.get("quantity", 1), additional_formats[color_index])
        
        # Column 10: Length (copied from Table 1)
        length_value = block_data.get("length", "")
        color_index = 10 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 10, length_value, additional_formats[color_index])
        
        # Column 11: Length After (Formula) = Length from Table 2 - Discount
        # Using length from Table 2 (column 10) instead of Table 1
        table2_length_col = get_column_letter(col_offset + 10 + 1)  # Column containing length in Table 2
        disc_cell = get_column_letter(col_offset + 12 + 1)
        length_after_formula = f'={table2_length_col}{excel_row + 1}-{disc_cell}{excel_row + 1}'
        color_index = 11 % len(additional_formats)
        worksheet.write_formula(excel_row, col_offset + 11, length_after_formula, additional_formats[color_index])
        
        # Column 12: Discount (fixed value)
        color_index = 12 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 12, 0.20, additional_formats[color_index])
        
        # Column 13: Height (Publishing stage height)
        publish_height = block_data.get("publish_height", float(block_data.get("height", 0) or 0))
        color_index = 13 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 13, publish_height, additional_formats[color_index])
        
        # Column 14: Quantity m2 (Formula) = Length After × Height × Quantity
        len_aft_cell = get_column_letter(col_offset + 11 + 1)
        h_cell = get_column_letter(col_offset + 13 + 1)
        cnt_cell = get_column_letter(col_offset + 9 + 1)
        qty_m2_formula = f'={len_aft_cell}{excel_row + 1}*{h_cell}{excel_row + 1}*{cnt_cell}{excel_row + 1}'
        color_index = 14 % len(additional_formats)
        worksheet.write_formula(excel_row, col_offset + 14, qty_m2_formula, additional_formats[color_index])
        
        # Column 15: Publish Price (based on thickness)
        thickness_value = thickness_text.replace("سم", "")
        if thickness_value == "2":
            publish_price = 125
        elif thickness_value == "3":
            publish_price = 145
        elif thickness_value == "4":
            publish_price = 155
        else:
            publish_price = 125
        color_index = 15 % len(additional_formats)
        worksheet.write(excel_row, col_offset + 15, publish_price, additional_formats[color_index])
        
        # Column 16: Total Publish Price (Formula) = Publish Price × Quantity m2
        pr_cell = get_column_letter(col_offset + 15 + 1)
        qm2_cell = get_column_letter(col_offset + 14 + 1)
        total_publish_formula = f'={pr_cell}{excel_row + 1}*{qm2_cell}{excel_row + 1}'
        color_index = 16 % len(additional_formats)
        worksheet.write_formula(excel_row, col_offset + 16, total_publish_formula, additional_formats[color_index])
        
        # Column 17: Total Cost (Formula) = Total Publish Price + Tips
        tot_pub_cell = get_column_letter(col_offset + 16 + 1)
        tips_cell = get_column_letter(col_offset + 7 + 1)
        total_cost_formula = f'={tot_pub_cell}{excel_row + 1}+{tips_cell}{excel_row + 1}'
        color_index = 17 % len(additional_formats)
        worksheet.write_formula(excel_row, col_offset + 17, total_cost_formula, additional_formats[color_index])
        
    worksheet.freeze_panes(3, 0)
    
    # تنسيق عرض الأعمدة
    for i in range(min(len(TABLE1_COLUMNS), len(TABLE1_WIDTH))): 
        worksheet.set_column(i, i, TABLE1_WIDTH[i])
    
    for i in range(min(len(TABLE2_COLUMNS), len(TABLE2_WIDTH))): 
        worksheet.set_column(len(TABLE1_COLUMNS) + i, len(TABLE1_COLUMNS) + i, TABLE2_WIDTH[i])
    
    # Table 3 has been removed as per user request
    
    try:
        workbook.close()
        return True
    except PermissionError as e:
        print(f"File is locked: {e}")
        return False  # File is locked
