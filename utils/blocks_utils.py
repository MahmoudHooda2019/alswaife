import xlsxwriter
from typing import List, Dict, Tuple, Optional


COLUMNS = [
    "رقم النقلة",
    "عدد النقلات",
    "التاريخ",
    "المحجر",
    "رقم الماكينة",
    "رقم البلوك",
    "الخامة",
    "الطول",
    "العرض",
    "الارتفاع",
    "م3",
    "الوزن",
    "وزن البلوك",
    "سعر الطن",
    "إجمالي سعر البلوك",
    "سعر النقلة",
]


def export_blocks_excel(
    filepath: str, block_size: str, notes: str, rows: List[Dict]
) -> Tuple[bool, Optional[str]]:
    """Create an Excel sheet for blocks entries."""
    try:
        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet("البلوكات")
        worksheet.right_to_left()

        title_fmt = workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#F4B084",
            }
        )
        info_fmt = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "align": "right",
                "bg_color": "#FFF2CC",
            }
        )
        header_fmt = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#4F81BD",
                "font_color": "white",
            }
        )
        data_fmt = workbook.add_format(
            {"border": 1, "align": "center", "valign": "vcenter"}
        )
        number_fmt = workbook.add_format(
            {"border": 1, "align": "center", "valign": "vcenter", "num_format": "0.000"}
        )

        worksheet.merge_range(0, 0, 0, len(COLUMNS) - 1, "سجل البلوكات", title_fmt)
        worksheet.write(1, 0, "مقاس البلوك:", info_fmt)
        worksheet.merge_range(1, 1, 1, 4, block_size or "-", data_fmt)
        worksheet.write(2, 0, "ملاحظات:", info_fmt)
        worksheet.merge_range(2, 1, 2, 4, notes or "-", data_fmt)

        for idx, col in enumerate(COLUMNS):
            worksheet.write(4, idx, col, header_fmt)

        start_row = 5
        for i, row in enumerate(rows):
            excel_row = start_row + i
            worksheet.write(excel_row, 0, row["trip_number"], data_fmt)
            worksheet.write(excel_row, 1, row["trip_count"], data_fmt)
            worksheet.write(excel_row, 2, row["date"], data_fmt)
            worksheet.write(excel_row, 3, row["quarry"], data_fmt)
            worksheet.write(excel_row, 4, row["machine_number"], data_fmt)
            worksheet.write(excel_row, 5, row["block_number"], data_fmt)
            worksheet.write(excel_row, 6, row["material"], data_fmt)

            worksheet.write_number(excel_row, 7, row["length"], number_fmt)
            worksheet.write_number(excel_row, 8, row["width"], number_fmt)
            worksheet.write_number(excel_row, 9, row["height"], number_fmt)

            # Formulas for M3 and related values
            length_col = "H"
            width_col = "I"
            height_col = "J"
            volume_col = "K"
            weight_col = "L"
            block_weight_col = "M"
            ton_price_col = "N"

            row_number = excel_row + 1  # Excel is 1-indexed
            worksheet.write_formula(
                excel_row,
                10,
                f"={length_col}{row_number}*{width_col}{row_number}*{height_col}{row_number}",
                number_fmt,
            )

            worksheet.write_number(excel_row, 11, row["weight"], number_fmt)
            worksheet.write_formula(
                excel_row,
                12,
                f"={volume_col}{row_number}*{weight_col}{row_number}",
                number_fmt,
            )
            worksheet.write_number(excel_row, 13, row["ton_price"], number_fmt)
            worksheet.write_formula(
                excel_row,
                14,
                f"={block_weight_col}{row_number}*{ton_price_col}{row_number}",
                number_fmt,
            )
            worksheet.write_number(excel_row, 15, row["trip_price"], number_fmt)

        worksheet.set_column(0, 6, 18)
        worksheet.set_column(7, 15, 14)

        workbook.close()
        return True, None
    except PermissionError:
        return False, "file_locked"
    except Exception as exc:
        return False, str(exc)


