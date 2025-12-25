"""
Reports Utilities - AI-powered report generation
Uses OpenAI to parse user requests and Pandas to process Excel data
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

# OpenAI API configuration
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")


def parse_user_request_with_ai(user_request: str) -> Dict:
    """
    Use AI to parse user's natural language request into structured JSON.
    
    Example input: "أعطني تقرير الإيرادات من 1/10 حتى 10/10"
    Example output: {
        "report_type": "income",
        "date_from": "01/10/2025",
        "date_to": "10/10/2025",
        "filters": {},
        "columns": ["date", "invoice_number", "client", "amount"]
    }
    """
    if not OPENAI_API_KEY:
        # Fallback: simple parsing without AI
        return parse_request_simple(user_request)
    
    try:
        from openai import OpenAI
        client = OpenAI(api_key=OPENAI_API_KEY)
        
        system_prompt = """أنت مساعد لتحليل طلبات التقارير. قم بتحويل طلب المستخدم إلى JSON بالتنسيق التالي:
{
    "report_type": "income" أو "expenses" أو "clients" أو "inventory" أو "attendance" أو "sales",
    "date_from": "DD/MM/YYYY" أو null,
    "date_to": "DD/MM/YYYY" أو null,
    "client_name": "اسم العميل" أو null,
    "filters": {},
    "group_by": "day" أو "month" أو "client" أو null,
    "sort_by": "date" أو "amount" أو null,
    "sort_order": "asc" أو "desc"
}

السنة الحالية هي 2025. إذا لم يذكر المستخدم السنة، استخدم 2025.
أرجع JSON فقط بدون أي نص إضافي."""

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_request}
            ],
            temperature=0.1,
            max_tokens=500
        )
        
        result = response.choices[0].message.content.strip()
        # Clean up the response
        if result.startswith("```"):
            result = result.split("```")[1]
            if result.startswith("json"):
                result = result[4:]
        
        return json.loads(result)
    
    except Exception as e:
        print(f"[ERROR] AI parsing failed: {e}")
        return parse_request_simple(user_request)


def parse_request_simple(user_request: str) -> Dict:
    """Simple fallback parser without AI"""
    import re
    
    result = {
        "report_type": "income",
        "date_from": None,
        "date_to": None,
        "client_name": None,
        "filters": {},
        "group_by": None,
        "sort_by": "date",
        "sort_order": "desc"
    }
    
    # Detect report type
    if "مصروف" in user_request or "صرف" in user_request:
        result["report_type"] = "expenses"
    elif "إيراد" in user_request or "ايراد" in user_request or "دخل" in user_request:
        result["report_type"] = "income"
    elif "عميل" in user_request or "عملاء" in user_request:
        result["report_type"] = "clients"
    elif "مخزون" in user_request or "مخزن" in user_request:
        result["report_type"] = "inventory"
    elif "حضور" in user_request or "انصراف" in user_request:
        result["report_type"] = "attendance"
    elif "مبيعات" in user_request or "بيع" in user_request:
        result["report_type"] = "sales"
    
    # Extract dates (format: DD/MM or DD/MM/YYYY)
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


def load_income_expenses_data(documents_path: str) -> pd.DataFrame:
    """Load income and expenses data from Excel file"""
    filepath = os.path.join(documents_path, "ايرادات ومصروفات", "بيان مصروفات وايرادات مصنع جرانيت السويفى.xlsx")
    
    if not os.path.exists(filepath):
        return pd.DataFrame(), pd.DataFrame()
    
    try:
        # Read income data (columns A-D, starting from row 6)
        df_income = pd.read_excel(filepath, sheet_name=0, usecols="A:D", skiprows=5, header=None)
        df_income.columns = ["رقم الفاتورة", "العميل", "المبلغ", "التاريخ"]
        df_income["النوع"] = "إيراد"
        df_income = df_income.dropna(subset=["رقم الفاتورة"])
        
        # Read expenses data (columns E-H, starting from row 6)
        df_expenses = pd.read_excel(filepath, sheet_name=0, usecols="E:H", skiprows=5, header=None)
        df_expenses.columns = ["العدد", "البيان", "المبلغ", "التاريخ"]
        df_expenses["النوع"] = "مصروف"
        df_expenses = df_expenses.dropna(subset=["البيان"])
        
        return df_income, df_expenses
    except Exception as e:
        print(f"[ERROR] Failed to load income/expenses data: {e}")
        return pd.DataFrame(), pd.DataFrame()


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
            
            # Check for ledger file
            ledger_file = os.path.join(client_path, "كشف حساب.xlsx")
            if os.path.exists(ledger_file):
                try:
                    df = pd.read_excel(ledger_file, skiprows=2)
                    df["client"] = client_folder
                    all_data.append(df)
                except:
                    pass
        
        if all_data:
            return pd.concat(all_data, ignore_index=True)
        return pd.DataFrame()
    except Exception as e:
        print(f"[ERROR] Failed to load invoices data: {e}")
        return pd.DataFrame()


def execute_report(query: Dict, documents_path: str) -> Optional[str]:
    """
    Execute the report based on parsed query and return the output file path.
    """
    report_type = query.get("report_type", "income")
    date_from = query.get("date_from")
    date_to = query.get("date_to")
    client_name = query.get("client_name")
    group_by = query.get("group_by")
    
    # Load data based on report type
    if report_type in ["income", "expenses", "income_expenses"]:
        df_income, df_expenses = load_income_expenses_data(documents_path)
        
        if report_type == "income":
            df = df_income
        elif report_type == "expenses":
            df = df_expenses
        else:
            df = pd.concat([df_income, df_expenses], ignore_index=True)
    
    elif report_type in ["clients", "sales"]:
        df = load_invoices_data(documents_path)
    
    else:
        df = pd.DataFrame()
    
    if df.empty:
        return None
    
    # Apply date filters
    if date_from or date_to:
        df = apply_date_filter(df, date_from, date_to)
    
    # Apply client filter
    if client_name:
        # Check for client column (Arabic or English)
        client_col = None
        if "العميل" in df.columns:
            client_col = "العميل"
        elif "client" in df.columns:
            client_col = "client"
        
        if client_col:
            df = df[df[client_col].str.contains(client_name, na=False)]
    
    # Generate report file
    return generate_report_excel(df, query, documents_path)


def apply_date_filter(df: pd.DataFrame, date_from: str, date_to: str) -> pd.DataFrame:
    """Apply date range filter to dataframe"""
    # Check for date column (Arabic or English)
    date_col = None
    if "التاريخ" in df.columns:
        date_col = "التاريخ"
    elif "date" in df.columns:
        date_col = "date"
    
    if date_col is None:
        return df
    
    try:
        # Convert date column to datetime
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
        "income_expenses": "الإيرادات والمصروفات",
        "clients": "العملاء",
        "sales": "المبيعات",
        "inventory": "المخزون",
        "attendance": "الحضور"
    }
    
    report_name = type_names.get(report_type, "تقرير")
    filename = f"تقرير {report_name}_{timestamp}.xlsx"
    filepath = os.path.join(reports_path, filename)
    
    # Create Excel file
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet("التقرير")
    worksheet.right_to_left()
    
    # Formats
    title_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#1F4E78', 'font_color': 'white', 'font_size': 16, 'border': 1
    })
    
    header_fmt = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#4472C4', 'font_color': 'white', 'font_size': 12, 'border': 1
    })
    
    cell_fmt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1
    })
    
    number_fmt = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'font_size': 11, 'border': 1,
        'num_format': '#,##0'
    })
    
    # Title
    date_range = ""
    if query.get("date_from") and query.get("date_to"):
        date_range = f" من {query['date_from']} إلى {query['date_to']}"
    elif query.get("date_from"):
        date_range = f" من {query['date_from']}"
    elif query.get("date_to"):
        date_range = f" حتى {query['date_to']}"
    
    title = f"تقرير {report_name}{date_range}"
    num_cols = len(df.columns) if not df.empty else 5
    worksheet.merge_range(0, 0, 0, num_cols - 1, title, title_fmt)
    worksheet.set_row(0, 35)
    
    # Write data
    if not df.empty:
        # Headers
        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(2, col_idx, col_name, header_fmt)
        
        # Data rows
        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                if isinstance(value, (int, float)) and not pd.isna(value):
                    worksheet.write(row_idx + 3, col_idx, value, number_fmt)
                else:
                    worksheet.write(row_idx + 3, col_idx, str(value) if not pd.isna(value) else "", cell_fmt)
        
        # Set column widths
        for col_idx in range(len(df.columns)):
            worksheet.set_column(col_idx, col_idx, 15)
        
        # Summary row
        summary_row = len(df) + 4
        worksheet.write(summary_row, 0, "الإجمالي", header_fmt)
        
        # Sum numeric columns (check for Arabic column names)
        for col_idx, col_name in enumerate(df.columns):
            col_lower = str(col_name).lower()
            if col_name == "المبلغ" or col_name == "amount" or "مبلغ" in col_lower:
                total = df[col_name].sum() if df[col_name].dtype in ['int64', 'float64'] else 0
                worksheet.write(summary_row, col_idx, total, number_fmt)
    else:
        worksheet.write(2, 0, "لا توجد بيانات متاحة", cell_fmt)
    
    workbook.close()
    return filepath
