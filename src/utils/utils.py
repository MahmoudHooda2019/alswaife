"""
General Utilities - دوال مساعدة عامة
"""

import sys
import os
import platform
import subprocess
from datetime import datetime


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def get_current_date(format_str="%Y-%m-%d"):
    """
    الحصول على التاريخ الحالي - أونلاين إذا كان الإنترنت متاحاً، أوفلاين إذا لم يكن متاحاً
    
    Args:
        format_str: صيغة التاريخ المطلوبة (افتراضي: YYYY-MM-DD)
    
    Returns:
        str: التاريخ الحالي بالصيغة المطلوبة
    """
    try:
        import urllib.request
        import json
        
        # محاولة الحصول على التاريخ من الإنترنت (worldtimeapi.org)
        url = "http://worldtimeapi.org/api/timezone/Africa/Cairo"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        
        with urllib.request.urlopen(req, timeout=3) as response:
            data = json.loads(response.read().decode())
            # التاريخ يأتي بصيغة ISO: "2025-01-15T14:30:00.123456+02:00"
            datetime_str = data.get('datetime', '')
            if datetime_str:
                # استخراج التاريخ والوقت
                dt = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
                return dt.strftime(format_str)
    except Exception:
        pass
    
    # إذا فشل الاتصال بالإنترنت، استخدم التاريخ المحلي
    return datetime.now().strftime(format_str)


def get_documents_path():
    """الحصول على مسار مجلد المستندات الخاص بالتطبيق"""
    return os.path.join(os.path.expanduser("~"), "Documents", "alswaife")


def ensure_folder_exists(folder_path):
    """التأكد من وجود المجلد وإنشائه إذا لم يكن موجوداً"""
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def format_number(value, decimals=2):
    """تنسيق الأرقام مع فواصل الآلاف"""
    try:
        num = float(value)
        if decimals == 0:
            return f"{int(num):,}"
        return f"{num:,.{decimals}f}"
    except (ValueError, TypeError):
        return str(value)


def safe_float(value, default=0.0):
    """تحويل آمن إلى float"""
    try:
        return float(value) if value else default
    except (ValueError, TypeError):
        return default


def safe_int(value, default=0):
    """تحويل آمن إلى int"""
    try:
        return int(value) if value else default
    except (ValueError, TypeError):
        return default


def is_excel_running():
    """التحقق مما إذا كان برنامج Excel مفتوحاً"""
    try:
        if platform.system() == "Windows":
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            return "EXCEL.EXE" in result.stdout.upper()
        elif platform.system() == "Darwin":  # macOS
            result = subprocess.run(
                ["pgrep", "-x", "Microsoft Excel"], capture_output=True
            )
            return result.returncode == 0
        else:  # Linux
            result = subprocess.run(
                ["pgrep", "-f", "libreoffice.*calc"], capture_output=True
            )
            return result.returncode == 0
    except Exception:
        return False


def is_file_locked(filepath):
    """التحقق مما إذا كان الملف مقفلاً (مفتوح في برنامج آخر)"""
    if not os.path.exists(filepath):
        return False
    try:
        with open(filepath, "a"):
            pass
        return False
    except (IOError, PermissionError):
        return True
