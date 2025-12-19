# مصنع السويفي - نظام الإدارة

تطبيق سطح مكتب لإدارة الفواتير والموظفين والحضور في مصنع السويفي.

## هيكل المشروع

```
alswaife/
├── src/                    # الكود المصدري
│   ├── main.py            # نقطة الدخول الرئيسية
│   ├── version.py         # معلومات الإصدار
│   ├── views/             # واجهات المستخدم
│   │   ├── dashboard_view.py
│   │   ├── invoice_view.py
│   │   ├── attendance_view.py
│   │   ├── blocks_view.py
│   │   └── purchases_view.py
│   └── utils/             # الأدوات المساعدة
│       ├── db_utils.py
│       ├── excel_utils.py
│       ├── path_utils.py
│       ├── attendance_utils.py
│       ├── blocks_utils.py
│       └── purchases_utils.py
├── data/                  # ملفات البيانات
│   ├── products.json
│   ├── employees.json
│   └── invoice.db
├── assets/                # الموارد
│   └── icon.ico
├── docs/                  # الوثائق
│   └── README.md
├── build/                 # ملفات البناء
│   ├── build_exe.bat
│   ├── AlSawifeFactory.spec
│   └── installer.iss
├── tests/                 # الاختبارات
├── dist/                  # الملفات المبنية
└── requirements.txt       # متطلبات Python
```

## كيفية التشغيل

### للتطوير:
```bash
cd src
python main.py
```

### للإنتاج:
- Windows: انقر مرتين على `AlSawifeFactory.exe`

## المميزات

- إدارة الفواتير مع تصدير إلى Excel
- تتبع الحضور والإنصراف
- إدارة البلوكات والمشتريات
- واجهة عربية مع دعم RTL
- قاعدة بيانات SQLite محلية

## المتطلبات

- Python 3.8+
- Flet
- XlsxWriter
- Tkinter (مدمج مع Python)

## البناء

```bash
# تثبيت المتطلبات
pip install -r requirements.txt

# بناء التطبيق
python -m PyInstaller AlSawifeFactory.spec
```
