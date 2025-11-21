"""
Excel Utilities for Invoice Creation
This module provides functions to generate Excel invoices from invoice data.
"""

import xlsxwriter
from typing import List, Tuple, Union


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
    worksheet.set_portrait()
    worksheet.set_margins(0.3, 0.3, 0.5, 0.5)

    # أهم شيء هنا:
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
        'align': 'right', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 12
    })

    # أهم شيء: خط أصغر في الصفوف حتى لا تمتد لصفحتين
    border_fmt = workbook.add_format({
        'border': 1, 'align': 'right', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 10
    })

    number_fmt = workbook.add_format({
        'border': 1, 'align': 'right', 'valign': 'vcenter',
        'font_name': 'Arial', 'font_size': 10
    })

    empty_fmt = workbook.add_format({})

    area_fmt = workbook.add_format({
        'border': 1, 'align': 'right',
        'font_name': 'Arial', 'font_size': 10,
        'num_format': '0.000'
    })

    money_fmt = workbook.add_format({
        'border': 1, 'align': 'right',
        'font_name': 'Arial', 'font_size': 10,
        'num_format': '0.00'
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
    worksheet.merge_range(start_row, 1, start_row, 4,
                          f"فاتورة رقم ( {op_num} )", header_fmt)

    # ==========================
    #   جدول العميل / التاريخ
    # ==========================
    r = start_row + 2

    worksheet.write(r, 1, "العميل", label_fmt)
    worksheet.merge_range(r, 2, r, 3, client or "", border_fmt)

    worksheet.write(r, 4, "التاريخ", label_fmt)
    worksheet.merge_range(r, 5, r, 6, date_str or "", border_fmt)

    r += 1

    worksheet.write(r, 1, "اسم السائق", label_fmt)
    worksheet.merge_range(r, 2, r, 3, driver or "", border_fmt)

    worksheet.write(r, 4, "ت", label_fmt)
    worksheet.merge_range(r, 5, r, 6, phone or "", border_fmt)

    r += 2

    # ==========================
    #   جدول الصنف
    # ==========================
    worksheet.write(r, 1, "البيان", header_fmt)
    worksheet.write(r, 2, "رقم البلوك", header_fmt)
    worksheet.write(r, 3, "السمك", header_fmt)
    worksheet.write(r, 4, "الخامة", header_fmt)
    worksheet.write(r, 5, "بالمتر", header_fmt)
    worksheet.write(r, 6, "الإجمالي م2", header_fmt)
    worksheet.write(r, 7, "السعر", header_fmt)
    worksheet.write(r, 8, "إجمالي السعر", header_fmt)

    first_item_row = None

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
            price_val = float(item[7])
        except (ValueError, IndexError) as e:
            # Skip invalid items
            continue

        r += 1
        if first_item_row is None:
            first_item_row = r

        worksheet.write(r, 1, desc, border_fmt)
        worksheet.write(r, 2, block, border_fmt)
        worksheet.write(r, 3, thickness, border_fmt)
        worksheet.write(r, 4, material, border_fmt)
        worksheet.write(r, 5, f"{count} * {length} * {height}", border_fmt)
        
        # Calculate area
        area_val = count * length * height
        worksheet.write_number(r, 6, area_val, area_fmt)
        worksheet.write_number(r, 7, price_val, money_fmt)

        excel_row = r + 1
        worksheet.write_formula(r, 8, f"=H{excel_row}*G{excel_row}", money_fmt)

    # ==========================
    #   المجموع
    # ==========================
    if first_item_row:
        total_row = r + 1
        worksheet.merge_range(total_row, 1, total_row, 7, "المجموع", header_fmt)

        excel_first = first_item_row + 1
        excel_last = r + 1
        worksheet.write_formula(total_row, 8,
                                f"=SUM(I{excel_first}:I{excel_last})",
                                money_fmt)

    # ==========================
    #   عرض الأعمدة
    # ==========================
    worksheet.set_column(0, 0, 3)
    worksheet.set_column(1, 1, 16)
    worksheet.set_column(2, 2, 10)
    worksheet.set_column(3, 3, 10)
    worksheet.set_column(4, 4, 12)
    worksheet.set_column(5, 5, 16)
    worksheet.set_column(6, 6, 10)
    worksheet.set_column(7, 7, 10)
    worksheet.set_column(8, 8, 14)

    workbook.close()