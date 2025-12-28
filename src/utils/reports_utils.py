"""
Reports Utilities - AI-powered report generation
Uses Groq/Gemini to parse user requests and Pandas to process Excel data
Supports all sections: Income, Expenses, Inventory, Attendance, Blocks, Slides
"""

import os
import json
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional
import xlsxwriter

# Load environment variables from .env file
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# API configuration - Groq as primary, Gemini as fallback
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

REPORT_PROMPT = """أنت مساعد لتحليل طلبات التقارير. قم بتحويل طلب المستخدم إلى JSON بالتنسيق التالي:
{
    "report_type": نوع التقرير (انظر القائمة أدناه),
    "date_from": "DD/MM/YYYY" أو null,
    "date_to": "DD/MM/YYYY" أو null,
    "client_name": "اسم العميل" أو null,
    "item_name": "اسم الصنف" أو null,
    "block_number": "رقم البلوك" أو null,
    "machine_number": "رقم الماكينة" أو null,
    "filters": {},
    "group_by": "day" أو "month" أو "client" أو "item" أو null,
    "sort_by": "date" أو "amount" أو null,
    "sort_order": "asc" أو "desc"
}

أنواع التقارير المتاحة:
- income: تقرير الإيرادات
- expenses: تقرير المصروفات
- inventory: تقرير المخزون (الرصيد الحالي)
- inventory_add: تقرير إضافات المخزون
- inventory_disburse: تقرير صرف المخزون
- attendance: تقرير الحضور والانصراف
- blocks: تقرير البلوكات
- slides: تقرير الشرائح
- machine_production: تقرير إنتاج ماكينة معينة (استخدمه عند ذكر مكنه أو ماكينه أو انتاج مكينه)
- clients: تقرير العملاء

ملاحظات مهمة:
- إذا ذكر المستخدم "مكنه" أو "ماكينه" أو "انتاج" استخدم report_type = "machine_production" واستخرج رقم الماكينة في machine_number
- السنة الحالية هي 2025. إذا لم يذكر المستخدم السنة، استخدم 2025.
- أرجع JSON فقط بدون أي نص إضافي."""


def parse_with_groq(user_request: str) -> Dict:
    """Parse user request using Groq API (fast and free)"""
    try:
        from groq import Groq
        client = Groq(api_key=GROQ_API_KEY)
        
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": REPORT_PROMPT},
                {"role": "user", "content": user_request}
            ],
            temperature=0.1,
            max_tokens=500
        )
        
        result = response.choices[0].message.content.strip()
        print(f"[DEBUG] Groq response: {result}")
        
        # Extract JSON from response
        if "```json" in result:
            result = result.split("```json")[1].split("```")[0]
        elif "```" in result:
            result = result.split("```")[1].split("```")[0]
        
        # Find JSON object in response
        start = result.find("{")
        end = result.rfind("}") + 1
        if start != -1 and end > start:
            result = result[start:end]
        
        return json.loads(result)
    
    except Exception as e:
        print(f"[ERROR] Groq parsing failed: {e}")
        return None


def parse_with_gemini(user_request: str) -> Dict:
    """Parse user request using Google Gemini API"""
    try:
        import google.generativeai as genai
        genai.configure(api_key=GEMINI_API_KEY)
        
        model = genai.GenerativeModel('gemini-1.5-flash')
        response = model.generate_content(f"{REPORT_PROMPT}\n\nطلب المستخدم: {user_request}")
        
        result = response.text.strip()
        print(f"[DEBUG] Gemini response: {result}")
        
        # Extract JSON from response
        if "```json" in result:
            result = result.split("```json")[1].split("```")[0]
        elif "```" in result:
            result = result.split("```")[1].split("```")[0]
        
        # Find JSON object in response
        start = result.find("{")
        end = result.rfind("}") + 1
        if start != -1 and end > start:
            result = result[start:end]
        
        return json.loads(result)
    
    except Exception as e:
        print(f"[ERROR] Gemini parsing failed: {e}")
        return None


def parse_user_request_with_ai(user_request: str) -> Dict:
    """
    Use AI to parse user's natural language request into structured JSON.
    Tries Groq first (fastest), then Gemini, then falls back to simple parser.
    """
    # Try Groq first (very fast and free)
    if GROQ_API_KEY:
        result = parse_with_groq(user_request)
        if result:
            return result
    
    # Try Gemini as fallback
    if GEMINI_API_KEY:
        result = parse_with_gemini(user_request)
        if result:
            return result
    
    # Fall back to simple parser
    return parse_request_simple(user_request)


def parse_request_simple(user_request: str) -> Dict:
    """Simple fallback parser without AI"""
    import re
    
    result = {
        "report_type": "income",
        "date_from": None,
        "date_to": None,
        "client_name": None,
        "item_name": None,
        "block_number": None,
        "machine_number": None,
        "filters": {},
        "group_by": None,
        "sort_by": "date",
        "sort_order": "desc"
    }
    
    # Detect report type
    if "مكن" in user_request or "مكينه" in user_request or "ماكينه" in user_request or "انتاج" in user_request:
        result["report_type"] = "machine_production"
        # Extract machine number
        machine_pattern = r'(?:مكن|مكينه|ماكينه|machine)\s*(\d+)'
        machine_match = re.search(machine_pattern, user_request, re.IGNORECASE)
        if machine_match:
            result["machine_number"] = machine_match.group(1)
        else:
            # Try to find standalone number after machine keywords
            num_pattern = r'(\d+)'
            nums = re.findall(num_pattern, user_request)
            if nums:
                result["machine_number"] = nums[0]
    elif "مصروف" in user_request or "صرف" in user_request:
        if "مخزون" in user_request:
            result["report_type"] = "inventory_disburse"
        else:
            result["report_type"] = "expenses"
    elif "إيراد" in user_request or "ايراد" in user_request or "دخل" in user_request:
        result["report_type"] = "income"
    elif "عميل" in user_request or "عملاء" in user_request:
        result["report_type"] = "clients"
    elif "مخزون" in user_request or "مخزن" in user_request:
        if "اضاف" in user_request:
            result["report_type"] = "inventory_add"
        else:
            result["report_type"] = "inventory"
    elif "حضور" in user_request or "انصراف" in user_request:
        result["report_type"] = "attendance"
    elif "بلوك" in user_request or "بلوكات" in user_request:
        result["report_type"] = "blocks"
    elif "شرائح" in user_request or "شريحة" in user_request:
        result["report_type"] = "slides"
    
    # Extract dates
    date_pattern = r'(\d{1,2})[/\\-](\d{1,2})(?:[/\\-](\d{2,4}))?'
    dates = re.findall(date_pattern, user_request)
    
    current_year = datetime.now().year
    
    if len(dates) >= 1:
        d, m, y = dates[0][0], dates[0][1], dates[0][2] or str(current_year)
        if len(y) == 2:
            y = "20" + y
        result["date_from"] = f"{d.zfill(2)}/{m.zfill(2)}/{y}"
    
    if len(dates) >= 2:
        d, m, y = dates[1][0], dates[1][1], dates[1][2] or str(current_year)
        if len(y) == 2:
            y = "20" + y
        result["date_to"] = f"{d.zfill(2)}/{m.zfill(2)}/{y}"
    
    return result


# ============ DATA LOADING FUNCTIONS ============

def load_income_expenses_data(documents_path: str):
    """Load income and expenses data from Excel file (3 sheets version)"""
    filepath = os.path.join(documents_path, "ايرادات ومصروفات", "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
    
    if not os.path.exists(filepath):
        return pd.DataFrame(), pd.DataFrame()
    
    try:
        df_income = pd.read_excel(filepath, sheet_name='الإيرادات', skiprows=1, header=None)
        df_income.columns = ["رقم الفاتورة", "العميل", "المبلغ", "التاريخ"]
        df_income["النوع"] = "إيراد"
        df_income = df_income.dropna(subset=["رقم الفاتورة"])
        
        df_expenses = pd.read_excel(filepath, sheet_name='المصروفات', skiprows=1, header=None)
        df_expenses.columns = ["العدد", "البيان", "المبلغ", "التاريخ"]
        df_expenses["النوع"] = "مصروف"
        df_expenses = df_expenses.dropna(subset=["البيان"])
        
        return df_income, df_expenses
    except Exception as e:
        print(f"[ERROR] Failed to load income/expenses data: {e}")
        return pd.DataFrame(), pd.DataFrame()


def load_inventory_data(documents_path: str):
    """Load inventory data from Excel file"""
    filepath = os.path.join(documents_path, "مخزون الادوات", "مخزون ادوات التشغيل.xlsx")
    
    if not os.path.exists(filepath):
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    try:
        # Load inventory summary
        df_inventory = pd.read_excel(filepath, sheet_name='المخزون', skiprows=0)
        df_inventory.columns = ["اسم الصنف", "إجمالي الإضافات", "إجمالي الصرف", "الرصيد الحالي"]
        df_inventory = df_inventory.dropna(subset=["اسم الصنف"])
        
        # Load additions
        df_add = pd.read_excel(filepath, sheet_name='اذن الاضافه', skiprows=0)
        df_add.columns = ["رقم الإذن", "التاريخ", "اسم الصنف", "العدد", "سعر الوحدة", "الإجمالي", "ملاحظات"]
        df_add = df_add.dropna(subset=["اسم الصنف"])
        
        # Load disbursements
        df_disburse = pd.read_excel(filepath, sheet_name='اذن الصرف', skiprows=0)
        df_disburse.columns = ["رقم الإذن", "التاريخ", "اسم الصنف", "العدد", "سعر الوحدة", "الإجمالي", "ملاحظات"]
        df_disburse = df_disburse.dropna(subset=["اسم الصنف"])
        
        return df_inventory, df_add, df_disburse
    except Exception as e:
        print(f"[ERROR] Failed to load inventory data: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


def load_attendance_data(documents_path: str) -> pd.DataFrame:
    """Load attendance data from Excel file"""
    filepath = os.path.join(documents_path, "حضور وانصراف", "سجل الحضور والانصراف.xlsx")
    
    if not os.path.exists(filepath):
        return pd.DataFrame()
    
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        
        data = []
        for row_idx in range(3, sheet.max_row + 1):
            name = sheet.cell(row=row_idx, column=1).value
            if not name or str(name).strip() in ['', 'اسم الموظف', 'الإجمالي الأسبوعي']:
                continue
            
            # Sum all shifts
            total_shifts = 0
            for col in range(2, 16):
                val = sheet.cell(row=row_idx, column=col).value
                if val and isinstance(val, (int, float)):
                    total_shifts += val
            
            total = sheet.cell(row=row_idx, column=16).value or 0
            date = sheet.cell(row=row_idx, column=17).value or ''
            advance = sheet.cell(row=row_idx, column=18).value or 0
            
            data.append({
                "الاسم": str(name).strip(),
                "إجمالي الورديات": total_shifts,
                "الإجمالي": total,
                "التاريخ": str(date),
                "السلفة": advance
            })
        
        return pd.DataFrame(data)
    except Exception as e:
        print(f"[ERROR] Failed to load attendance data: {e}")
        return pd.DataFrame()


def load_blocks_data(documents_path: str) -> pd.DataFrame:
    """Load blocks data from Excel file"""
    filepath = os.path.join(documents_path, "البلوكات", "مخزون البلوكات.xlsx")
    
    if not os.path.exists(filepath):
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(filepath, sheet_name='البلوكات', skiprows=1)
        # Rename columns based on TABLE1_COLUMNS
        expected_cols = ["رقم النقله", "عدد النقله", "التاريخ", "المحجر", 
                        "رقم البلوك", "الخامه", "الطول", "العرض", "الارتفاع", 
                        "م3", "الوزن", "وزن البلوك", "سعر الطن", "اجمالي السعر"]
        if len(df.columns) >= len(expected_cols):
            df.columns = expected_cols + list(df.columns[len(expected_cols):])
        df = df.dropna(subset=["رقم البلوك"])
        return df
    except Exception as e:
        print(f"[ERROR] Failed to load blocks data: {e}")
        return pd.DataFrame()


def load_slides_data(documents_path: str) -> pd.DataFrame:
    """Load slides data from Excel file"""
    filepath = os.path.join(documents_path, "البلوكات", "مخزون البلوكات.xlsx")
    
    if not os.path.exists(filepath):
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(filepath, sheet_name='البلوكات', skiprows=1)
        # Slides start from column 15 (index 14)
        slides_cols = ["تاريخ النشر", "رقم البلوك", "النوع", "رقم المكينه",
                      "وقت الدخول", "وقت الخروج", "عدد الساعات",
                      "السمك", "العدد", "الطول", "الخصم", "الطول بعد",
                      "الارتفاع", "الكمية م2", "سعر المتر", "اجمالي السعر"]
        
        if len(df.columns) >= 15 + len(slides_cols):
            slides_df = df.iloc[:, 15:15+len(slides_cols)].copy()
            slides_df.columns = slides_cols
            slides_df = slides_df.dropna(subset=["رقم البلوك"])
            return slides_df
        return pd.DataFrame()
    except Exception as e:
        print(f"[ERROR] Failed to load slides data: {e}")
        return pd.DataFrame()


def load_machine_production_data(documents_path: str, machine_number: str = None) -> pd.DataFrame:
    """Load machine production data from slides, optionally filtered by machine number"""
    filepath = os.path.join(documents_path, "البلوكات", "مخزون البلوكات.xlsx")
    
    print(f"[DEBUG] load_machine_production_data - filepath: {filepath}")
    print(f"[DEBUG] machine_number: {machine_number}")
    
    if not os.path.exists(filepath):
        print(f"[DEBUG] File does not exist")
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(filepath, sheet_name='البلوكات', skiprows=1)
        print(f"[DEBUG] DataFrame shape: {df.shape}")
        print(f"[DEBUG] DataFrame columns: {list(df.columns)}")
        
        # Slides columns start after blocks columns (14 columns for blocks + 1 gap = 15)
        # But let's find the actual position by looking for slides header
        slides_cols = ["تاريخ النشر", "رقم البلوك", "النوع", "رقم المكينه",
                      "وقت الدخول", "وقت الخروج", "عدد الساعات",
                      "السمك", "العدد", "الطول", "الخصم", "الطول بعد",
                      "الارتفاع", "الكمية م2", "سعر المتر", "اجمالي السعر"]
        
        # Try to find slides data starting from column 14 (0-indexed)
        slides_start = 14
        
        if len(df.columns) >= slides_start + len(slides_cols):
            slides_df = df.iloc[:, slides_start:slides_start+len(slides_cols)].copy()
            slides_df.columns = slides_cols
            print(f"[DEBUG] Slides DataFrame shape: {slides_df.shape}")
            print(f"[DEBUG] Slides first rows:\n{slides_df.head()}")
            
            # Drop rows where block number is empty
            slides_df = slides_df.dropna(subset=["رقم البلوك"])
            print(f"[DEBUG] After dropna shape: {slides_df.shape}")
            
            # Filter by machine number if provided
            if machine_number and "رقم المكينه" in slides_df.columns:
                print(f"[DEBUG] Filtering by machine: {machine_number}")
                print(f"[DEBUG] Machine column values: {slides_df['رقم المكينه'].unique()}")
                slides_df = slides_df[slides_df["رقم المكينه"].astype(str).str.contains(str(machine_number), na=False)]
                print(f"[DEBUG] After filter shape: {slides_df.shape}")
            
            # Select relevant columns for machine production report
            production_cols = ["تاريخ النشر", "رقم البلوك", "رقم المكينه", "عدد الساعات", 
                             "السمك", "العدد", "الكمية م2", "اجمالي السعر"]
            available_cols = [col for col in production_cols if col in slides_df.columns]
            
            return slides_df[available_cols]
        else:
            print(f"[DEBUG] Not enough columns. Have {len(df.columns)}, need {slides_start + len(slides_cols)}")
        return pd.DataFrame()
    except Exception as e:
        print(f"[ERROR] Failed to load machine production data: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


def load_invoices_data(documents_path: str) -> pd.DataFrame:
    """Load all invoices data from client folders"""
    invoices_path = os.path.join(documents_path, "الفواتير")
    
    if not os.path.exists(invoices_path):
        return pd.DataFrame()
    
    all_data = []
    
    try:
        for client_folder in os.listdir(invoices_path):
            client_path = os.path.join(invoices_path, client_folder)
            if not os.path.isdir(client_path):
                continue
            
            ledger_file = os.path.join(client_path, "كشف حساب.xlsx")
            if os.path.exists(ledger_file):
                try:
                    df = pd.read_excel(ledger_file, skiprows=2)
                    df["العميل"] = client_folder
                    all_data.append(df)
                except:
                    pass
        
        if all_data:
            return pd.concat(all_data, ignore_index=True)
        return pd.DataFrame()
    except Exception as e:
        print(f"[ERROR] Failed to load invoices data: {e}")
        return pd.DataFrame()


# ============ REPORT EXECUTION ============

def execute_report(query: Dict, documents_path: str) -> Optional[str]:
    """Execute the report based on parsed query and return the output file path."""
    report_type = query.get("report_type", "income")
    date_from = query.get("date_from")
    date_to = query.get("date_to")
    
    df = pd.DataFrame()
    
    # Load data based on report type
    if report_type == "income":
        df_income, _ = load_income_expenses_data(documents_path)
        df = df_income
    
    elif report_type == "expenses":
        _, df_expenses = load_income_expenses_data(documents_path)
        df = df_expenses
    
    elif report_type == "inventory":
        df_inv, _, _ = load_inventory_data(documents_path)
        df = df_inv
    
    elif report_type == "inventory_add":
        _, df_add, _ = load_inventory_data(documents_path)
        df = df_add
    
    elif report_type == "inventory_disburse":
        _, _, df_disburse = load_inventory_data(documents_path)
        df = df_disburse
    
    elif report_type == "attendance":
        df = load_attendance_data(documents_path)
    
    elif report_type == "blocks":
        df = load_blocks_data(documents_path)
    
    elif report_type == "slides":
        df = load_slides_data(documents_path)
    
    elif report_type == "machine_production":
        machine_number = query.get("machine_number")
        df = load_machine_production_data(documents_path, machine_number)
    
    elif report_type == "clients":
        df = load_invoices_data(documents_path)
    
    if df.empty:
        return None
    
    # Apply date filters
    if date_from or date_to:
        df = apply_date_filter(df, date_from, date_to)
    
    # Apply item filter for inventory
    item_name = query.get("item_name")
    if item_name and "اسم الصنف" in df.columns:
        df = df[df["اسم الصنف"].str.contains(item_name, na=False)]
    
    # Apply block filter
    block_number = query.get("block_number")
    if block_number and "رقم البلوك" in df.columns:
        df = df[df["رقم البلوك"].str.contains(str(block_number), na=False)]
    
    # Apply client filter
    client_name = query.get("client_name")
    if client_name:
        if "العميل" in df.columns:
            df = df[df["العميل"].str.contains(client_name, na=False)]
    
    return generate_report_excel(df, query, documents_path)


def apply_date_filter(df: pd.DataFrame, date_from: str, date_to: str) -> pd.DataFrame:
    """Apply date range filter to dataframe"""
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


def generate_report_excel(df: pd.DataFrame, query: Dict, documents_path: str) -> str:
    """Generate Excel report from dataframe"""
    reports_path = os.path.join(documents_path, "التقارير")
    os.makedirs(reports_path, exist_ok=True)
    
    report_type = query.get("report_type", "report")
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    type_names = {
        "income": "الإيرادات",
        "expenses": "المصروفات",
        "inventory": "المخزون",
        "inventory_add": "إضافات المخزون",
        "inventory_disburse": "صرف المخزون",
        "attendance": "الحضور والانصراف",
        "blocks": "البلوكات",
        "slides": "الشرائح",
        "machine_production": "إنتاج الماكينات",
        "clients": "العملاء"
    }
    
    report_name = type_names.get(report_type, "تقرير")
    
    # Add machine number to report name if available
    machine_number = query.get("machine_number")
    if report_type == "machine_production" and machine_number:
        report_name = f"إنتاج ماكينة {machine_number}"
    
    filename = f"تقرير {report_name}_{timestamp}.xlsx"
    filepath = os.path.join(reports_path, filename)
    
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet("التقرير")
    worksheet.right_to_left()
    
    # Formats
    title_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#1F4E78', 'font_color': 'white', 'font_size': 16, 'border': 2
    })
    
    header_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#4472C4', 'font_color': 'white', 'font_size': 12, 'border': 1
    })
    
    cell_fmt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1
    })
    
    cell_fmt_alt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1,
        'bg_color': '#F2F2F2'
    })
    
    number_fmt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1,
        'num_format': '#,##0.00'
    })
    
    number_fmt_alt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1,
        'num_format': '#,##0.00', 'bg_color': '#F2F2F2'
    })
    
    summary_label_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#1F4E78', 'font_color': 'white', 'font_size': 12, 'border': 2
    })
    
    summary_value_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#C6EFCE', 'font_color': '#006100', 'font_size': 14, 'border': 2,
        'num_format': '#,##0.00'
    })
    
    # Title with date range
    date_range = ""
    if query.get("date_from") and query.get("date_to"):
        date_range = f" من {query['date_from']} إلى {query['date_to']}"
    elif query.get("date_from"):
        date_range = f" من {query['date_from']}"
    elif query.get("date_to"):
        date_range = f" حتى {query['date_to']}"
    
    title = f"تقرير {report_name}{date_range}"
    num_cols = len(df.columns) if not df.empty else 4
    
    worksheet.merge_range(0, 0, 0, num_cols - 1, title, title_fmt)
    worksheet.set_row(0, 35)
    
    if not df.empty:
        # Headers
        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(1, col_idx, col_name, header_fmt)
        worksheet.set_row(1, 25)
        
        # Data rows
        data_start_row = 2
        for i, (_, row) in enumerate(df.iterrows()):
            current_row = data_start_row + i
            is_alt = i % 2 == 1
            
            for col_idx, value in enumerate(row):
                if isinstance(value, (int, float)) and not pd.isna(value):
                    fmt = number_fmt_alt if is_alt else number_fmt
                    worksheet.write(current_row, col_idx, value, fmt)
                else:
                    fmt = cell_fmt_alt if is_alt else cell_fmt
                    worksheet.write(current_row, col_idx, str(value) if not pd.isna(value) else "", fmt)
        
        # Column widths
        for col_idx, col_name in enumerate(df.columns):
            max_len = max(len(str(col_name)), df.iloc[:, col_idx].astype(str).str.len().max() if len(df) > 0 else 10)
            worksheet.set_column(col_idx, col_idx, min(max_len + 2, 25))
        
        # Summary row - find numeric columns to sum
        summary_row = data_start_row + len(df) + 1
        worksheet.merge_range(summary_row, 0, summary_row, num_cols - 2, "الإجمالي", summary_label_fmt)
        
        # Find amount/total column
        amount_cols = ["المبلغ", "الإجمالي", "اجمالي السعر", "الرصيد الحالي", "الكمية م2"]
        total_value = 0
        for col_name in amount_cols:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name)
                if df.iloc[:, col_idx].dtype in ['int64', 'float64']:
                    total_value = df.iloc[:, col_idx].sum()
                    break
        
        worksheet.write(summary_row, num_cols - 1, total_value, summary_value_fmt)
        worksheet.set_row(summary_row, 30)
    else:
        worksheet.merge_range(1, 0, 1, 3, "لا توجد بيانات متاحة", cell_fmt)
    
    workbook.close()
    return filepath
