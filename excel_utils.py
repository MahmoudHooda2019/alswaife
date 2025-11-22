"""
Excel Utilities for Invoice Creation
This module provides functions to generate Excel invoices from invoice data.
"""

import xlsxwriter
from typing import List, Tuple


def save_invoice(filepath: str, op_num: str, client: str, driver: str,
                 items: List[Tuple], date_str: str = "", phone: str = ""):
    """
    Save invoice data to an Excel file with proper formatting.
    
    Args:
        filepath (str): Path to save the Excel file
        op_num (str): Operation/invoice number
        client (str): Client name
        driver (str): Driver name
        items (List[Tuple]): List of invoice items, each tuple contains:
            (description, block, thickness, material, count, length, height, price)
        date_str (str, optional): Date string. Defaults to "".
        phone (str, optional): Phone number. Defaults to "".
    """
    
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet("فاتورة")

    # RTL + Page Break Preview
    worksheet.right_to_left()
    worksheet.set_pagebreak_view()
    worksheet.hide_gridlines(2)  # Hide screen and printed gridlines

    # ================
    #  PAGE SETTINGS
    # ================
    worksheet.set_paper(9)         # A4
    worksheet.set_landscape()
    worksheet.set_margins(0.3, 0.3, 0.5, 0.5)
    worksheet.fit_to_pages(1, 1)   # صفحة واحدة عرض + صفحة واحدة طول

    # ==========================
    #   FORMATS
    # ==========================

    header_fmt = workbook.add_format({
        'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 14, 'bg_color': '#F0F0F0'
    })

    label_fmt = workbook.add_format({
        'bold': True, 'border': 1,
        'align': 'center', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 12
    })

    border_fmt = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 10
    })

    # Format for integer numbers (no decimal places)
    integer_fmt = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 10,
        'num_format': '0'  # أعداد صحيحة فقط
    })

    # Format for phone numbers (text format to prevent scientific notation)
    phone_fmt = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 10,
        'num_format': '@'  # Text format
    })

    # Format for area (always 2 decimal places, يحافظ على الصفر الأخير)
    area_fmt = workbook.add_format({
        'border': 1, 'align': 'center',
        'font_name': 'Arial', 'font_size': 10,
        'num_format': '0.00'  # يظهر دائمًا منزلتين عشريتين
    })

    # Format for money (always 2 decimal places, يحافظ على الصفر الأخير)
    money_fmt = workbook.add_format({
        'border': 1, 'align': 'center',
        'font_name': 'Arial', 'font_size': 10,
        'num_format': '0'
    })



    # ==========================
    #   SPACE
    # ==========================
    worksheet.set_row(0, 12)
    worksheet.set_row(1, 12)

    start_row = 2

    # ==========================
    #   عنوان — Merge
    # ==========================
    worksheet.merge_range(start_row, 3, start_row, 8,
                          f"فاتورة رقم ( {op_num} )", header_fmt)

    # ==========================
    #   جدول العميل / التاريخ
    # ==========================
    r = start_row + 2

    # العميل
    worksheet.write(r, 1, "العميل", label_fmt)
    worksheet.merge_range(r, 2, r, 3, client or "", border_fmt)

    # التاريخ
    worksheet.write(r, 4, "التاريخ", label_fmt)
    worksheet.merge_range(r, 5, r, 6, date_str or "", border_fmt)

    # عدد السيارات - expand label to span two columns, remove merge from value cell
    worksheet.merge_range(r, 7, r, 8, "عدد السيارات", label_fmt)
    worksheet.write(r, 9, "1", integer_fmt)

    r += 1

    # السائق
    worksheet.write(r, 1, "اسم السائق", label_fmt)
    worksheet.merge_range(r, 2, r, 3, driver or "", border_fmt)

    # تليفون
    worksheet.write(r, 4, "ت", label_fmt)
    worksheet.merge_range(r, 5, r, 6, phone or "", phone_fmt)

    # نوع السيارة - expand label to span two columns, remove merge from value cell
    worksheet.merge_range(r, 7, r, 8, "سيارة", label_fmt)
    worksheet.write(r, 9, "", border_fmt)

    r += 2

    # ==========================
    #   جدول الصنف مع المقاسات الفرعية
    # ==========================
    worksheet.merge_range(r, 1, r+1, 1, "البيان", header_fmt)
    worksheet.merge_range(r, 2, r+1, 2, "رقم البلوك", header_fmt)
    worksheet.merge_range(r, 3, r+1, 3, "السمك", header_fmt)
    worksheet.merge_range(r, 4, r+1, 4, "الخامة", header_fmt)
    
    # دمج 3 أعمدة للعنوان العام للمقاس
    worksheet.merge_range(r, 5, r, 7, "المقاس", header_fmt)
    worksheet.write(r+1, 5, "العدد", header_fmt)
    worksheet.write(r+1, 6, "الطول", header_fmt)
    worksheet.write(r+1, 7, "الارتفاع", header_fmt)
    
    worksheet.merge_range(r, 8, r+1, 8, "المسطح م2", header_fmt)
    worksheet.merge_range(r, 9, r+1, 9, "السعر", header_fmt)
    worksheet.merge_range(r, 10, r+1, 10, "إجمالي السعر", header_fmt)

    first_item_row = None
    r += 2

    # ==========================
    #   العناصر
    # ==========================
    for item in items:
        try:
            desc = item[0]
            block = item[1]
            thickness = item[2]
            material = item[3]
            count = float(item[4])
            length = float(item[5])
            height = float(item[6])
            price_val =  int(item[7])
        except (ValueError, IndexError):
            continue

        if first_item_row is None:
            first_item_row = r

        worksheet.write(r, 1, desc, border_fmt)
        worksheet.write(r, 2, block, border_fmt)
        worksheet.write(r, 3, thickness, border_fmt)
        worksheet.write(r, 4, material, border_fmt)
        # القيم توضع مباشرة في الأعمدة الثلاثة
        worksheet.write_number(r, 5, count, integer_fmt)      # العدد
        worksheet.write_number(r, 6, length, border_fmt)     # الطول
        worksheet.write_number(r, 7, height, border_fmt)     # الارتفاع
        
        # الصيغ يحسبها Excel تلقائياً
        excel_row = r + 1
        worksheet.write_formula(r, 8, f"=F{excel_row}*G{excel_row}*H{excel_row}", area_fmt)
        worksheet.write_number(r, 9, price_val, money_fmt)
        # تقريب الناتج إلى أقرب عدد صحيح
        worksheet.write_formula(r, 10, f"=ROUND(I{excel_row}*J{excel_row},0)", money_fmt)

        r += 1

    # ==========================
    #   المجموع
    # ==========================
    if first_item_row is not None:
        total_row = r
        worksheet.merge_range(total_row, 1, total_row, 4, "المجموع", header_fmt)
        excel_first = first_item_row + 1
        excel_last = r - 1  # Exclude the sum row itself to avoid circular reference
        
        # Sum for count (column 5)
        worksheet.write(total_row, 5, f"=SUM(F{excel_first}:F{excel_last})", integer_fmt)
        
        # Sum for area (column 8)
        worksheet.write(total_row, 8, f"=SUM(I{excel_first}:I{excel_last})", area_fmt)
        
        # Sum for total price (column 10)
        worksheet.write(total_row, 10, f"=SUM(K{excel_first}:K{excel_last})", money_fmt)

    # ==========================
    #   عرض الأعمدة
    # ==========================
    worksheet.set_column(0, 0, 3)
    worksheet.set_column(1, 1, 16)  # البيان
    worksheet.set_column(2, 2, 10)  # رقم البلوك
    worksheet.set_column(3, 3, 10)  # السمك
    worksheet.set_column(4, 4, 12)  # الخامة
    worksheet.set_column(5, 5, 8)   # العدد
    worksheet.set_column(6, 6, 8)   # الطول
    worksheet.set_column(7, 7, 8)   # # الارتفاع
    worksheet.set_column(8, 8, 10)  # الإجمالي م2
    worksheet.set_column(9, 9, 10)  # السعر
    worksheet.set_column(10, 10, 14)  # إجمالي السعر

    workbook.close()
