"""
Reports Utilities - Offline report generation
Uses Pandas to process Excel data
"""

import os
import pandas as pd
from datetime import datetime
from typing import Dict, Optional
import xlsxwriter


def execute_report(query: Dict, documents_path: str) -> Optional[str]:
    """Execute the report based on parsed query and return the output file path."""
    report_type = query.get("report_type", "")
    date_from = query.get("date_from")
    date_to = query.get("date_to")
    
    # 1. البلوكات المنشورة كاملة (A و B)
    if report_type == "blocks_published":
        return generate_blocks_published_report(documents_path, date_from, date_to)
    
    # 2. العملاء المدينين
    elif report_type == "clients_debts":
        return generate_clients_debts_report(documents_path)
    
    # 3. إنتاج الماكينات
    elif report_type == "machine_production":
        machine_number = query.get("machine_number")
        return generate_machine_production_report(documents_path, machine_number, date_from, date_to)
    
    # 4. الإيرادات فقط
    elif report_type == "income":
        return generate_income_report(documents_path, date_from, date_to)
    
    # 5. المصروفات فقط
    elif report_type == "expenses":
        return generate_expenses_report(documents_path, date_from, date_to)
    
    # 6. الإيرادات والمصروفات معاً
    elif report_type == "income_expenses_both":
        return generate_income_expenses_both_report(documents_path, date_from, date_to)
    
    # 7. استهلاك الأدوات
    elif report_type == "inventory_consumption":
        return generate_inventory_consumption_report(documents_path, date_from, date_to)
    
    return None


def generate_blocks_published_report(documents_path: str, date_from: str, date_to: str) -> Optional[str]:
    """تقرير البلوكات المنشورة كاملة (التي تحتوي على A و B)"""
    filepath = os.path.join(documents_path, "البلوكات", "مخزون البلوكات.xlsx")
    
    if not os.path.exists(filepath):
        return None
    
    try:
        df = pd.read_excel(filepath, sheet_name='البلوكات', skiprows=1)
        if df.empty:
            return None
        
        # تحديد أسماء الأعمدة
        expected_cols = ["رقم النقله", "عدد النقله", "التاريخ", "المحجر", 
                        "رقم البلوك", "الخامه", "الطول", "العرض", "الارتفاع", 
                        "م3", "الوزن", "وزن البلوك", "سعر الطن", "اجمالي السعر"]
        if len(df.columns) >= len(expected_cols):
            df.columns = expected_cols + list(df.columns[len(expected_cols):])
        
        df = df.dropna(subset=["رقم البلوك"])
        
        # استخراج رقم البلوك الأساسي (بدون الحرف)
        df["رقم_اساسي"] = df["رقم البلوك"].astype(str).str.extract(r'(\d+)')[0]
        
        # البحث عن البلوكات التي لها A و B
        published_blocks = []
        for block_num in df["رقم_اساسي"].unique():
            if pd.isna(block_num):
                continue
            block_rows = df[df["رقم_اساسي"] == block_num]
            block_ids = block_rows["رقم البلوك"].astype(str).str.upper().tolist()
            
            has_a = any('A' in bid for bid in block_ids)
            has_b = any('B' in bid for bid in block_ids)
            
            if has_a and has_b:
                published_blocks.append(block_num)
        
        if not published_blocks:
            return None
        
        # فلترة البلوكات المنشورة
        result_df = df[df["رقم_اساسي"].isin(published_blocks)].copy()
        result_df = result_df.drop(columns=["رقم_اساسي"])
        
        # تطبيق فلتر التاريخ
        if date_from or date_to:
            result_df = apply_date_filter(result_df, date_from, date_to)
        
        if result_df.empty:
            return None
        
        return save_report_to_excel(result_df, documents_path, "البلوكات المنشورة كاملة")
        
    except Exception as e:
        print(f"[ERROR] generate_blocks_published_report: {e}")
        return None


def generate_clients_debts_report(documents_path: str) -> Optional[str]:
    """تقرير العملاء المدينين فقط"""
    invoices_path = os.path.join(documents_path, "الفواتير")
    
    if not os.path.exists(invoices_path):
        return None
    
    clients_data = []
    
    try:
        for client_folder in os.listdir(invoices_path):
            client_path = os.path.join(invoices_path, client_folder)
            if not os.path.isdir(client_path):
                continue
            
            ledger_file = os.path.join(client_path, "كشف حساب.xlsx")
            if not os.path.exists(ledger_file):
                continue
            
            try:
                # قراءة كشف الحساب
                df = pd.read_excel(ledger_file, skiprows=2)
                
                if df.empty:
                    continue
                
                # البحث عن عمود الرصيد
                balance_col = None
                for col in df.columns:
                    col_str = str(col).strip()
                    if 'رصيد' in col_str or 'الرصيد' in col_str:
                        balance_col = col
                        break
                
                if balance_col is None and len(df.columns) >= 6:
                    balance_col = df.columns[5]  # عادة العمود السادس
                
                if balance_col is None:
                    continue
                
                # الحصول على آخر رصيد
                last_balance = df[balance_col].dropna().iloc[-1] if not df[balance_col].dropna().empty else 0
                
                try:
                    last_balance = float(last_balance)
                except:
                    last_balance = 0
                
                # إضافة العميل إذا كان عليه دين (رصيد موجب)
                if last_balance > 0:
                    clients_data.append({
                        "اسم العميل": client_folder,
                        "إجمالي الدين": last_balance
                    })
                    
            except Exception as e:
                print(f"[ERROR] Reading ledger for {client_folder}: {e}")
                continue
        
        if not clients_data:
            return None
        
        result_df = pd.DataFrame(clients_data)
        result_df = result_df.sort_values("إجمالي الدين", ascending=False)
        
        # إضافة صف الإجمالي
        total_debt = result_df["إجمالي الدين"].sum()
        total_row = pd.DataFrame([{"اسم العميل": "الإجمالي", "إجمالي الدين": total_debt}])
        result_df = pd.concat([result_df, total_row], ignore_index=True)
        
        return save_report_to_excel(result_df, documents_path, "العملاء المدينين")
        
    except Exception as e:
        print(f"[ERROR] generate_clients_debts_report: {e}")
        return None


def generate_machine_production_report(documents_path: str, machine_number: str, date_from: str, date_to: str) -> Optional[str]:
    """تقرير إنتاج ماكينة معينة"""
    # محاولة قراءة من ملف الشرائح أولاً
    slides_filepath = os.path.join(documents_path, "الشرائح", "مخزون الشرائح.xlsx")
    blocks_filepath = os.path.join(documents_path, "البلوكات", "مخزون البلوكات.xlsx")
    
    df = pd.DataFrame()
    
    if os.path.exists(slides_filepath):
        try:
            df = pd.read_excel(slides_filepath, sheet_name='اذن اضافة الشرائح', skiprows=0)
        except:
            pass
    
    if df.empty and os.path.exists(blocks_filepath):
        try:
            df = pd.read_excel(blocks_filepath, sheet_name='الشرائح', skiprows=1)
        except:
            pass
    
    if df.empty:
        return None
    
    try:
        # البحث عن عمود رقم الماكينة
        machine_col = None
        for col in df.columns:
            col_str = str(col).strip()
            if 'مكينه' in col_str or 'مكينة' in col_str or 'ماكينه' in col_str or 'ماكينة' in col_str or 'المكينه' in col_str:
                machine_col = col
                break
        
        if machine_col is None:
            return None
        
        # فلترة حسب رقم الماكينة
        df[machine_col] = df[machine_col].astype(str).str.strip()
        machine_str = str(machine_number).strip()
        
        result_df = df[df[machine_col] == machine_str].copy()
        
        if result_df.empty:
            result_df = df[df[machine_col].str.contains(machine_str, na=False)].copy()
        
        if result_df.empty:
            return None
        
        # تطبيق فلتر التاريخ
        if date_from or date_to:
            result_df = apply_date_filter(result_df, date_from, date_to)
        
        if result_df.empty:
            return None
        
        return save_report_to_excel(result_df, documents_path, f"إنتاج ماكينة {machine_number}")
        
    except Exception as e:
        print(f"[ERROR] generate_machine_production_report: {e}")
        return None


def generate_income_report(documents_path: str, date_from: str, date_to: str) -> Optional[str]:
    """تقرير الإيرادات فقط"""
    filepath = os.path.join(documents_path, "ايرادات ومصروفات", "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
    
    if not os.path.exists(filepath):
        return None
    
    try:
        df = pd.read_excel(filepath, sheet_name='الإيرادات', skiprows=1, header=None)
        df.columns = ["رقم الفاتورة", "العميل", "المبلغ", "التاريخ"]
        df = df.dropna(subset=["رقم الفاتورة"])
        
        if df.empty:
            return None
        
        # تطبيق فلتر التاريخ
        if date_from or date_to:
            df = apply_date_filter(df, date_from, date_to)
        
        if df.empty:
            return None
        
        # إضافة صف الإجمالي
        total = df["المبلغ"].sum()
        total_row = pd.DataFrame([{"رقم الفاتورة": "", "العميل": "الإجمالي", "المبلغ": total, "التاريخ": ""}])
        df = pd.concat([df, total_row], ignore_index=True)
        
        return save_report_to_excel(df, documents_path, "الإيرادات")
        
    except Exception as e:
        print(f"[ERROR] generate_income_report: {e}")
        return None


def generate_expenses_report(documents_path: str, date_from: str, date_to: str) -> Optional[str]:
    """تقرير المصروفات فقط"""
    filepath = os.path.join(documents_path, "ايرادات ومصروفات", "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
    
    if not os.path.exists(filepath):
        return None
    
    try:
        df = pd.read_excel(filepath, sheet_name='المصروفات', skiprows=1, header=None)
        df.columns = ["العدد", "البيان", "المبلغ", "التاريخ"]
        df = df.dropna(subset=["البيان"])
        
        if df.empty:
            return None
        
        # تطبيق فلتر التاريخ
        if date_from or date_to:
            df = apply_date_filter(df, date_from, date_to)
        
        if df.empty:
            return None
        
        # إضافة صف الإجمالي
        total = df["المبلغ"].sum()
        total_row = pd.DataFrame([{"العدد": "", "البيان": "الإجمالي", "المبلغ": total, "التاريخ": ""}])
        df = pd.concat([df, total_row], ignore_index=True)
        
        return save_report_to_excel(df, documents_path, "المصروفات")
        
    except Exception as e:
        print(f"[ERROR] generate_expenses_report: {e}")
        return None


def generate_income_expenses_both_report(documents_path: str, date_from: str, date_to: str) -> Optional[str]:
    """تقرير الإيرادات والمصروفات معاً"""
    filepath = os.path.join(documents_path, "ايرادات ومصروفات", "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
    
    if not os.path.exists(filepath):
        return None
    
    try:
        # قراءة الإيرادات
        df_income = pd.read_excel(filepath, sheet_name='الإيرادات', skiprows=1, header=None)
        df_income.columns = ["رقم الفاتورة", "العميل", "المبلغ", "التاريخ"]
        df_income = df_income.dropna(subset=["رقم الفاتورة"])
        df_income["النوع"] = "إيراد"
        
        # قراءة المصروفات
        df_expenses = pd.read_excel(filepath, sheet_name='المصروفات', skiprows=1, header=None)
        df_expenses.columns = ["العدد", "البيان", "المبلغ", "التاريخ"]
        df_expenses = df_expenses.dropna(subset=["البيان"])
        df_expenses["النوع"] = "مصروف"
        
        # تطبيق فلتر التاريخ
        if date_from or date_to:
            df_income = apply_date_filter(df_income, date_from, date_to)
            df_expenses = apply_date_filter(df_expenses, date_from, date_to)
        
        # حساب الإجماليات
        total_income = df_income["المبلغ"].sum() if not df_income.empty else 0
        total_expenses = df_expenses["المبلغ"].sum() if not df_expenses.empty else 0
        net_profit = total_income - total_expenses
        
        # إنشاء ملخص
        summary_df = pd.DataFrame({
            "البيان": ["إجمالي الإيرادات", "إجمالي المصروفات", "صافي الربح"],
            "المبلغ": [total_income, total_expenses, net_profit]
        })
        
        return save_report_to_excel(summary_df, documents_path, "ملخص الإيرادات والمصروفات")
        
    except Exception as e:
        print(f"[ERROR] generate_income_expenses_both_report: {e}")
        return None


def generate_inventory_consumption_report(documents_path: str, date_from: str, date_to: str) -> Optional[str]:
    """تقرير استهلاك الأدوات (المصروفات من أذون الصرف)"""
    filepath = os.path.join(documents_path, "مخزون الادوات", "مخزون ادوات التشغيل.xlsx")
    
    if not os.path.exists(filepath):
        return None
    
    try:
        df = pd.read_excel(filepath, sheet_name='اذن الصرف', skiprows=0)
        
        if len(df.columns) >= 7:
            df.columns = ["رقم الإذن", "التاريخ", "اسم الصنف", "العدد", "سعر الوحدة", "الإجمالي", "ملاحظات"]
        
        df = df.dropna(subset=["اسم الصنف"])
        
        if df.empty:
            return None
        
        # تطبيق فلتر التاريخ
        if date_from or date_to:
            df = apply_date_filter(df, date_from, date_to)
        
        if df.empty:
            return None
        
        # تجميع حسب الصنف
        consumption = df.groupby("اسم الصنف").agg({
            "العدد": "sum",
            "الإجمالي": "sum"
        }).reset_index()
        
        consumption.columns = ["اسم الصنف", "إجمالي الكمية المصروفة", "إجمالي التكلفة"]
        consumption = consumption.sort_values("إجمالي التكلفة", ascending=False)
        
        # إضافة صف الإجمالي
        total_qty = consumption["إجمالي الكمية المصروفة"].sum()
        total_cost = consumption["إجمالي التكلفة"].sum()
        total_row = pd.DataFrame([{
            "اسم الصنف": "الإجمالي",
            "إجمالي الكمية المصروفة": total_qty,
            "إجمالي التكلفة": total_cost
        }])
        consumption = pd.concat([consumption, total_row], ignore_index=True)
        
        return save_report_to_excel(consumption, documents_path, "استهلاك الأدوات")
        
    except Exception as e:
        print(f"[ERROR] generate_inventory_consumption_report: {e}")
        return None


def apply_date_filter(df: pd.DataFrame, date_from: str, date_to: str) -> pd.DataFrame:
    """تطبيق فلتر التاريخ على DataFrame"""
    date_col = None
    for col in ["التاريخ", "تاريخ النشر", "تاريخ الدخول", "date"]:
        if col in df.columns:
            date_col = col
            break
    
    if date_col is None:
        return df
    
    try:
        df = df.copy()
        df["date_parsed"] = pd.to_datetime(df[date_col], format="%d/%m/%Y", errors="coerce")
        
        if df["date_parsed"].isna().all():
            df["date_parsed"] = pd.to_datetime(df[date_col], errors="coerce")
        
        if date_from:
            from_dt = datetime.strptime(date_from, "%d/%m/%Y")
            df = df[df["date_parsed"] >= from_dt]
        
        if date_to:
            to_dt = datetime.strptime(date_to, "%d/%m/%Y")
            df = df[df["date_parsed"] <= to_dt]
        
        df = df.drop(columns=["date_parsed"])
    except Exception as e:
        print(f"[ERROR] Date filter failed: {e}")
    
    return df


def save_report_to_excel(df: pd.DataFrame, documents_path: str, report_name: str) -> str:
    """حفظ التقرير في ملف Excel"""
    reports_path = os.path.join(documents_path, "التقارير")
    os.makedirs(reports_path, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{report_name}_{timestamp}.xlsx"
    filepath = os.path.join(reports_path, filename)
    
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet("التقرير")
    worksheet.right_to_left()
    
    # تنسيقات
    header_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#4472C4', 'font_color': 'white', 'font_size': 12, 'border': 1
    })
    
    cell_fmt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1
    })
    
    number_fmt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1,
        'num_format': '#,##0.00'
    })
    
    total_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#FFC107', 'font_size': 12, 'border': 2
    })
    
    # كتابة العناوين
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_fmt)
        worksheet.set_column(col_num, col_num, 18)
    
    # كتابة البيانات
    for row_num, row_data in enumerate(df.values, 1):
        is_total_row = row_num == len(df)
        for col_num, value in enumerate(row_data):
            if is_total_row:
                fmt = total_fmt
            elif isinstance(value, (int, float)) and not pd.isna(value):
                fmt = number_fmt
            else:
                fmt = cell_fmt
            
            if pd.isna(value):
                worksheet.write(row_num, col_num, "", fmt)
            else:
                worksheet.write(row_num, col_num, value, fmt)
    
    workbook.close()
    return filepath
