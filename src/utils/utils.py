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
    # Check if running as PyInstaller bundle
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        # First check _MEIPASS (for onefile mode)
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            # For onedir mode, use the exe directory
            base_path = os.path.dirname(sys.executable)
    else:
        # Running in development
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


def convert_english_to_arabic(text):
    """
    تحويل الحروف الإنجليزية إلى العربية (بناءً على تخطيط لوحة المفاتيح)
    
    Args:
        text: النص المراد تحويله
    
    Returns:
        str: النص بعد تحويل الحروف الإنجليزية إلى العربية
    """
    if not text:
        return text
    
    # خريطة تحويل الحروف الإنجليزية إلى العربية
    eng_to_ar = {
        'q': 'ض', 'w': 'ص', 'e': 'ث', 'r': 'ق', 't': 'ف', 'y': 'غ', 'u': 'ع', 'i': 'ه', 'o': 'خ', 'p': 'ح', '[': 'ج', ']': 'د',
        'a': 'ش', 's': 'س', 'd': 'ي', 'f': 'ب', 'g': 'ل', 'h': 'ا', 'j': 'ت', 'k': 'ن', 'l': 'م', ';': 'ك', "'": 'ط',
        'z': 'ئ', 'x': 'ء', 'c': 'ؤ', 'v': 'ر', 'b': 'لا', 'n': 'ى', 'm': 'ة', ',': 'و', '.': 'ز', '/': 'ظ',
        'Q': 'َ', 'W': 'ً', 'E': 'ُ', 'R': 'ٌ', 'T': 'لإ', 'Y': 'إ', 'U': ''', 'I': '÷', 'O': '×', 'P': '؛', '{': '<', '}': '>',
        'A': 'ِ', 'S': 'ٍ', 'D': ']', 'F': '[', 'G': 'لأ', 'H': 'أ', 'J': 'ـ', 'K': '،', 'L': '/', ':': ':', '"': '"',
        'Z': '~', 'X': 'ْ', 'C': '}', 'V': '{', 'B': 'لآ', 'N': 'آ', 'M': ''', '<': ',', '>': '.', '?': '؟',
        '`': 'ذ', '~': 'ّ',
    }
    
    result = ""
    for char in text:
        if char in eng_to_ar:
            result += eng_to_ar[char]
        else:
            result += char
    return result

def normalize_block_number(block_value, reorder=True):
    """
    تحويل رقم البلوك إلى الصيغة الصحيحة (A, B, F, K) مع إعادة الترتيب
    
    يقوم بتحويل الحروف المختلفة إلى الحروف الصحيحة:
    - A: ِ, ش, a, أ
    - B: لآ, لا, b, ب  
    - F: f, [, ب
    - K: k, ن, ،
    
    ثم يعيد ترتيب رقم البلوك ليبدأ بالحروف ثم الأرقام (مثل: "12A" -> "A12")
    
    Args:
        block_value (str): رقم البلوك المدخل
        reorder (bool): إعادة ترتيب الحروف والأرقام (افتراضي: True)
        
    Returns:
        str: رقم البلوك بعد التحويل والتنسيق وإعادة الترتيب
    """
    if not block_value:
        return ""
    
    # تحويل إلى نص للتأكد
    val = str(block_value).strip()
    if not val:
        return ""
    
    # تحويل الحروف إلى A, B, F, K
    # A: ِ, ش, a, أ
    new_val = val.replace('ِ', 'A').replace('ش', 'A').replace('a', 'A').replace('أ', 'A')
    
    # B: لآ, لا, b, ب
    new_val = new_val.replace('لآ', 'B').replace('لا', 'B').replace('b', 'B').replace('ب', 'B')
    
    # F: f, [, ب (ملاحظة: ب مكررة في B و F، سنعطي الأولوية لـ B)
    new_val = new_val.replace('f', 'F').replace('[', 'F')
    
    # K: k, ن, ،
    new_val = new_val.replace('k', 'K').replace('ن', 'K').replace('،', 'K')
    
    # تحويل إلى أحرف كبيرة
    new_val = new_val.upper()
    
    # إعادة ترتيب: إذا كان يبدأ بأرقام ويحتوي على حروف، ضع الحروف في البداية
    if reorder:
        import re
        match = re.match(r'^(\d+)([A-Za-z]+)$', new_val)
        if match:
            numbers = match.group(1)
            letters = match.group(2).upper()
            new_val = letters + numbers
    
    return new_val

def handle_arabic_decimal_input(text_field):
    """
    معالجة الفواصل العشرية العربية وتحويلها إلى نقطة عشرية
    
    يقوم بتحويل:
    - الفاصلة العربية (،) إلى نقطة (.)
    - حرف الزين (ز) إلى نقطة (.)
    - الأرقام العربية إلى أرقام إنجليزية
    
    Args:
        text_field: حقل النص المراد معالجته
        
    Returns:
        bool: True إذا تم تغيير القيمة، False إذا لم يتم تغييرها
    """
    if not text_field or not text_field.value:
        return False
    
    original_value = text_field.value
    new_value = original_value
    
    # تحويل الفواصل العشرية العربية إلى نقطة
    new_value = new_value.replace('،', '.')  # الفاصلة العربية
    new_value = new_value.replace('ز', '.')  # حرف الزين (يُستخدم أحياناً كفاصلة عشرية)
    
    # تحويل الأرقام العربية إلى أرقام إنجليزية
    arabic_digits = {
        '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4', 
        '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
    }
    
    for arabic_digit, english_digit in arabic_digits.items():
        new_value = new_value.replace(arabic_digit, english_digit)
    
    # تحديث القيمة إذا تغيرت
    if new_value != original_value:
        text_field.value = new_value
        return True
    
    return False


def normalize_numeric_input(value):
    """
    تطبيع المدخلات الرقمية (بدون تعديل حقل النص مباشرة)
    
    يقوم بتحويل الفواصل العشرية العربية والأرقام العربية إلى الصيغة الإنجليزية
    
    Args:
        value (str): القيمة المراد تطبيعها
        
    Returns:
        str: القيمة بعد التطبيع
    """
    if not value:
        return ""
    
    # تحويل إلى نص للتأكد
    normalized = str(value).strip()
    
    # تحويل الفواصل العشرية العربية إلى نقطة
    normalized = normalized.replace('،', '.')  # الفاصلة العربية
    normalized = normalized.replace('ز', '.')  # حرف الزين
    
    # تحويل الأرقام العربية إلى أرقام إنجليزية
    arabic_digits = {
        '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4', 
        '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
    }
    
    for arabic_digit, english_digit in arabic_digits.items():
        normalized = normalized.replace(arabic_digit, english_digit)
    
    return normalized