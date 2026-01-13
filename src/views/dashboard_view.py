import asyncio
import flet as ft
import os
import threading
from views.invoice_view import InvoiceView
from views.attendance_view import AttendanceView
from views.blocks_view import BlocksView
from views.purchases_view import PurchasesView
from views.inventory_add_view import InventoryAddView
from views.inventory_disburse_view import InventoryDisburseView
from views.slides_add_view import SlidesAddView
from views.reports_view import ReportsView
from views.payments_view import PaymentsView
from utils.update_utils import check_for_updates, download_update, install_update
from utils.invoice_utils import save_invoice, update_client_ledger
from utils.log_utils import log_error, log_exception
from utils.dialog_utils import DialogManager


def save_callback(filepath, op_num, client, driver, date_str, phone, items):
    """
    دالة رد الاتصال لحفظ بيانات الفاتورة إلى Excel.

    Args:
        filepath (str): المسار لحفظ ملف Excel
        op_num (str): رقم العملية/الفاتورة
        client (str): اسم العميل
        driver (str): اسم السائق
        date_str (str): سلسلة التاريخ
        phone (str): رقم الهاتف
        items (list): قائمة عناصر الفاتورة
    """
    # Extract only the first 8 elements for saving to Excel (excluding length_before and discount)
    items_for_excel = []
    for item in items:
        # Take only the first 8 elements: description, block, thickness, material, count, length, height, price
        item_excel = tuple(item[:8]) if len(item) >= 8 else item
        items_for_excel.append(item_excel)

    # Save the invoice
    save_invoice(
        filepath, op_num, client, driver, items_for_excel, date_str=date_str, phone=phone
    )

    # Skip ledger update for "ايراد" clients
    if "ايراد" in client:
        return

    # Create/update client ledger
    try:
        # Calculate total amount from items
        total_amount = 0
        invoice_items_details = []

        # Process items to get details for the ledger
        for item in items_for_excel:
            try:
                desc = item[0] or ""
                block = item[1] or ""
                thickness = item[2] or ""
                material = item[3] or ""
                count = int(float(item[4]))
                length = float(item[5])
                height = float(item[6])
                price_val = float(item[7])

                # Calculate area and total for this item
                area = count * length * height
                total = area * price_val

                # Add to total amount
                total_amount += total

                # Store item details for the ledger
                invoice_items_details.append((desc, material, thickness, area, total))
            except (ValueError, IndexError):
                continue

        # Get the client folder path (parent of the invoice folder)
        client_folder = os.path.dirname(os.path.dirname(filepath))

        # Update or create the client's ledger
        success, error = update_client_ledger(
            client_folder,
            client,
            date_str,
            op_num,
            total_amount,
            driver,
            invoice_items_details,
        )

        if not success:
            log_error(f"Could not update client ledger: {error}")
    except Exception as e:
        log_exception(f"Error updating client ledger: {e}")

class DashboardView:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "مصنع السويفي"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.vertical_alignment = ft.MainAxisAlignment.CENTER
        self.page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        
        # Main container for the dashboard
        self.main_container = ft.Container(
            content=self.build_menu(),
            alignment=ft.alignment.center,
            expand=True
        )

    def build_ui(self):
        """Build the main dashboard UI"""
        self.main_container = ft.Column(
            controls=[
                ft.Container(
                    content=ft.Text("مصنع السويفي", size=32, weight=ft.FontWeight.BOLD),
                    alignment=ft.alignment.center,
                    padding=20
                ),
                ft.GridView(
                    controls=[
                        self.create_menu_card("إدارة الفواتير", ft.Icons.RECEIPT_LONG, self.open_invoices, ft.Colors.BLUE_700),
                        self.create_menu_card("إدارة الدفعات", ft.Icons.PAYMENTS, self.open_payments, ft.Colors.GREEN_700),
                        self.create_menu_card("الحضور والإنصراف", ft.Icons.PERSON, self.open_attendance, ft.Colors.LIME_700),
                        self.create_menu_card("إضافة بلوكات", ft.Icons.VIEW_IN_AR, self.open_blocks, ft.Colors.AMBER_700),
                        self.create_menu_card("مشتري", ft.Icons.SHOPPING_CART, self.open_purchases, ft.Colors.CYAN_700),
                        self.create_menu_card("المخزون", ft.Icons.INVENTORY, self.open_inventory, ft.Colors.DEEP_PURPLE_700),
                        self.create_menu_card("إضافة شرائح", ft.Icons.ADD, self.open_slides_add, ft.Colors.PINK_700),
                        self.create_menu_card("التقارير", ft.Icons.ASSESSMENT, self.open_reports, ft.Colors.TEAL_700),
                        self.create_menu_card("تحديث", ft.Icons.SYSTEM_UPDATE, self.open_update, ft.Colors.ORANGE_700),
                        self.create_menu_card("مزامنة", ft.Icons.SYNC, self.open_sync, ft.Colors.LIGHT_BLUE_700),
                        self.create_menu_card("عنا", ft.Icons.INFO, self.show_about_dialog, ft.Colors.PURPLE_700),
                    ],
                    runs_count=2,
                    max_extent=200,
                    spacing=20,
                    run_spacing=20,
                    padding=20,
                )
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=20,
            expand=True
        )

    def build_menu(self):
        return ft.Column(
            controls=[
                ft.Text(
                    "مصنع السويفي", 
                    size=50, 
                    weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_200,
                    animate_opacity=1000,
                ),
                ft.Container(height=50),
                # Create card-based menu grid
                ft.GridView(
                    controls=[
                        self.create_menu_card("إدارة الفواتير", ft.Icons.RECEIPT_LONG, self.open_invoices, ft.Colors.BLUE_700),
                        self.create_menu_card("إدارة الدفعات", ft.Icons.PAYMENTS, self.open_payments, ft.Colors.GREEN_700),
                        self.create_menu_card("الحضور والإنصراف", ft.Icons.PERSON, self.open_attendance, ft.Colors.LIME_700),
                        self.create_menu_card("إضافة بلوكات", ft.Icons.VIEW_IN_AR, self.open_blocks, ft.Colors.AMBER_700),
                        self.create_menu_card("مشتري", ft.Icons.SHOPPING_CART, self.open_purchases, ft.Colors.CYAN_700),
                        self.create_menu_card("المخزون", ft.Icons.INVENTORY, self.open_inventory, ft.Colors.DEEP_PURPLE_700),
                        self.create_menu_card("إضافة شرائح", ft.Icons.ADD, self.open_slides_add, ft.Colors.PINK_700),
                        self.create_menu_card("التقارير", ft.Icons.ASSESSMENT, self.open_reports, ft.Colors.TEAL_700),
                        self.create_menu_card("تحديث", ft.Icons.SYSTEM_UPDATE, self.open_update, ft.Colors.ORANGE_700),
                        self.create_menu_card("مزامنة", ft.Icons.SYNC, self.open_sync, ft.Colors.LIGHT_BLUE_700),
                        self.create_menu_card("عنا", ft.Icons.INFO, self.show_about_dialog, ft.Colors.PURPLE_700),
                    ],
                    runs_count=2,
                    max_extent=200,
                    spacing=20,
                    run_spacing=20,
                    padding=20,
                )
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=20,
            expand=True
        )

    def create_menu_card(self, text, icon, on_click, color):
        return ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Icon(icon, size=50, color=ft.Colors.WHITE),
                        ft.Text(text, size=18, weight=ft.FontWeight.W_600, text_align=ft.TextAlign.CENTER),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=15,
                ),
                padding=20,
                alignment=ft.alignment.center,
                bgcolor=color,
                border_radius=15,
                ink=True,
                on_click=on_click if on_click else lambda e: self.show_placeholder(text),
                animate=ft.Animation(300, ft.AnimationCurve.EASE_OUT),
            ),
            elevation=5,
        )

    def show_placeholder(self, feature):
        message = f" الخاصية {feature} قيد التطوير" if feature else "هذه الخاصية قيد التطوير"
        DialogManager.show_info_dialog(self.page, message, title="تنبيه")

    def show_about_dialog(self, e):
        """Show about dialog with developer information"""
        def close_dlg(e):
            DialogManager.close_dialog(self.page, dlg)

        dlg = ft.AlertDialog(
            modal=True,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # Logo/Icon
                        ft.Container(
                            content=ft.Icon(
                                ft.Icons.FACTORY,
                                size=80,
                                color=ft.Colors.BLUE_400
                            ),
                            alignment=ft.alignment.center,
                            padding=20,
                        ),
                        
                        # App Name
                        ft.Text(
                            "مصنع جرانيت السويفي",
                            size=24,
                            weight=ft.FontWeight.BOLD,
                            color=ft.Colors.BLUE_200,
                            text_align=ft.TextAlign.CENTER,
                        ),
                        
                        # Version
                        ft.Text(
                            "الإصدار 1.0",
                            size=14,
                            color=ft.Colors.GREY_400,
                            text_align=ft.TextAlign.CENTER,
                        ),
                        
                        ft.Divider(height=30, color=ft.Colors.GREY_700),
                        
                        # Developer Section
                        ft.Container(
                            content=ft.Column(
                                controls=[
                                    ft.Row(
                                        controls=[
                                            ft.Icon(ft.Icons.CODE, color=ft.Colors.GREEN_400, size=20),
                                            ft.Text("تطوير وبرمجة", size=14, color=ft.Colors.GREY_400),
                                        ],
                                        alignment=ft.MainAxisAlignment.CENTER,
                                        spacing=10,
                                    ),
                                    ft.Text(
                                        "محمود حسين",
                                        size=18,
                                        weight=ft.FontWeight.W_600,
                                        color=ft.Colors.WHITE,
                                        text_align=ft.TextAlign.CENTER,
                                    ),
                                ],
                                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                spacing=5,
                            ),
                            padding=10,
                        ),
                        
                        ft.Divider(height=20, color=ft.Colors.GREY_700),
                        
                        # Contact Info
                        ft.Container(
                            content=ft.Column(
                                controls=[
                                    ft.Row(
                                        controls=[
                                            ft.Icon(ft.Icons.PHONE, color=ft.Colors.BLUE_300, size=18),
                                            ft.Text("01126422405", size=14, color=ft.Colors.GREY_300),
                                        ],
                                        alignment=ft.MainAxisAlignment.CENTER,
                                        spacing=10,
                                    ),
                                    ft.Row(
                                        controls=[
                                            ft.Icon(ft.Icons.EMAIL, color=ft.Colors.RED_300, size=18),
                                            ft.Text("mh20192004@gmail.com", size=14, color=ft.Colors.GREY_300),
                                        ],
                                        alignment=ft.MainAxisAlignment.CENTER,
                                        spacing=10,
                                    ),
                                ],
                                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                spacing=10,
                            ),
                            padding=10,
                        ),
                        
                        ft.Container(height=10),
                        
                        # Copyright
                        ft.Text(
                            "© 2026 جميع الحقوق محفوظة",
                            size=12,
                            color=ft.Colors.GREY_500,
                            text_align=ft.TextAlign.CENTER,
                        ),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=5,
                ),
                padding=20,
                width=350,
            ),
            actions=[
                ft.TextButton(
                    "إغلاق",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.BLUE_300)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.CENTER,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=20),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def open_reports(self, e):
        """Open the enhanced reports view"""
        self.page.clean()
        reports_view = ReportsView(self.page, on_back=self.show)
        reports_view.build_ui()

    def open_update(self, e):
        """Open update dialog to check and download updates"""
        
        # Show checking progress
        progress_dlg = DialogManager.show_loading_dialog(self.page, "جاري التحقق...")
        
        def check():
            try:
                update_available, current_ver, latest_ver, download_url = check_for_updates()
                
                DialogManager.close_dialog(self.page, progress_dlg)
                
                if update_available and download_url:
                    self.show_update_available_dialog(current_ver, latest_ver, download_url)
                else:
                    self.show_no_update_dialog(current_ver, latest_ver)
            except Exception as ex:
                DialogManager.close_dialog(self.page, progress_dlg)
                self.show_update_error(f"فشل في التحقق من التحديثات: {str(ex)}")
        
        thread = threading.Thread(target=check)
        thread.start()

    def show_update_available_dialog(self, current_ver, latest_ver, download_url):
        """Show dialog when update is available"""
        def close_dlg(e):
            DialogManager.close_dialog(self.page, dlg)
        
        def start_download(e):
            close_dlg(e)
            self.download_and_install_update(download_url)
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.NEW_RELEASES, color=ft.Colors.ORANGE_400, size=24),
                    ft.Text("تحديث متاح!", weight=ft.FontWeight.BOLD, color=ft.Colors.ORANGE_300, size=16, rtl=True),
                ],
                spacing=8,
                rtl=True,
            ),
            content=ft.Column(
                controls=[
                    ft.Row(
                        controls=[
                            ft.Text("الحالي:", size=13, color=ft.Colors.GREY_400, rtl=True),
                            ft.Text(current_ver, size=13, color=ft.Colors.WHITE, weight=ft.FontWeight.BOLD),
                            ft.Text("←", size=13, color=ft.Colors.GREY_500),
                            ft.Text("الجديد:", size=13, color=ft.Colors.GREY_400, rtl=True),
                            ft.Text(latest_ver, size=13, color=ft.Colors.GREEN_400, weight=ft.FontWeight.BOLD),
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=5,
                        rtl=True,
                    ),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                tight=True,
            ),
            actions=[
                ft.ElevatedButton(
                    "تحميل",
                    icon=ft.Icons.DOWNLOAD,
                    bgcolor=ft.Colors.ORANGE_700,
                    color=ft.Colors.WHITE,
                    on_click=start_download,
                ),
                ft.TextButton(
                    "لاحقاً",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=10),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def show_no_update_dialog(self, current_ver, latest_ver):
        """Show dialog when no update is available"""
        DialogManager.show_success_dialog(
            self.page, 
            f"أنت تستخدم أحدث إصدار ({current_ver})", 
            title="لا يوجد تحديث"
        )

    def download_and_install_update(self, download_url):
        """Download and install the update"""
        
        # Cancel flag
        self.download_cancelled = False
        
        def cancel_download(e):
            self.download_cancelled = True
            status_text.value = "جاري الإلغاء..."
            cancel_btn.disabled = True
            self.page.update()
        
        # Progress bar and text
        progress_bar = ft.ProgressBar(width=200, color=ft.Colors.ORANGE_400, bgcolor=ft.Colors.GREY_700)
        progress_text = ft.Text("0%", size=12, color=ft.Colors.WHITE)
        status_text = ft.Text("جاري التحميل...", size=14, color=ft.Colors.WHITE)
        cancel_btn = ft.TextButton(
            "إلغاء",
            icon=ft.Icons.CANCEL,
            on_click=cancel_download,
            style=ft.ButtonStyle(color=ft.Colors.RED_400)
        )
        
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Column(
                controls=[
                    ft.Icon(ft.Icons.DOWNLOAD, size=35, color=ft.Colors.ORANGE_400),
                    status_text,
                    progress_bar,
                    progress_text,
                    cancel_btn,
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=8,
                tight=True,
            ),
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=10),
        )
        self.page.overlay.append(progress_dlg)
        progress_dlg.open = True
        self.page.update()
        
        def update_progress(percent):
            progress_bar.value = percent / 100
            progress_text.value = f"{int(percent)}%"
            self.page.update()
        
        def check_cancelled():
            return self.download_cancelled
        
        def download():
            try:
                setup_path = download_update(download_url, update_progress, check_cancelled)
                
                if self.download_cancelled:
                    DialogManager.close_dialog(self.page, progress_dlg)
                    self.show_download_cancelled_dialog()
                    return
                
                if setup_path:
                    status_text.value = "جاري تشغيل المثبت..."
                    cancel_btn.visible = False
                    self.page.update()
                    
                    if install_update(setup_path):
                        DialogManager.close_dialog(self.page, progress_dlg)
                        self.show_install_success_dialog()
                    else:
                        DialogManager.close_dialog(self.page, progress_dlg)
                        self.show_update_error("فشل في تشغيل المثبت")
                else:
                    DialogManager.close_dialog(self.page, progress_dlg)
                    if not self.download_cancelled:
                        self.show_update_error("فشل في تحميل التحديث")
            except Exception as ex:
                DialogManager.close_dialog(self.page, progress_dlg)
                self.show_update_error(f"خطأ: {str(ex)}")
        
        thread = threading.Thread(target=download)
        thread.start()
    
    def show_download_cancelled_dialog(self):
        """Show dialog when download is cancelled"""
        DialogManager.show_warning_dialog(self.page, "تم إلغاء تحميل التحديث", title="تم الإلغاء")

    def show_install_success_dialog(self):
        """Show dialog after installer starts"""
        def close_dlg(e):
            DialogManager.close_dialog(self.page, dlg)
        
        def close_app(e):
            self.page.window.close()
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=24),
                    ft.Text("تم بدء التثبيت", weight=ft.FontWeight.BOLD, color=ft.Colors.GREEN_300, size=16),
                ],
                spacing=8,
            ),
            content=ft.Text(
                "يُنصح بإغلاق البرنامج لإكمال التحديث",
                size=13,
                color=ft.Colors.GREY_400,
            ),
            actions=[
                ft.ElevatedButton(
                    "إغلاق",
                    icon=ft.Icons.EXIT_TO_APP,
                    bgcolor=ft.Colors.RED_700,
                    color=ft.Colors.WHITE,
                    on_click=close_app,
                ),
                ft.TextButton(
                    "استمرار",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=10),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def show_update_error(self, error_msg):
        """Show update error dialog"""
        DialogManager.show_error_dialog(self.page, error_msg, title="خطأ في التحديث")

    def open_sync(self, e):
        """Open sync dialog - search for devices and compare"""
        from utils.sync_utils import get_local_ip, discover_devices, CompareServer, COMPARE_PORT
        
        local_ip = get_local_ip()
        self.discovered_devices = []
        self.compare_server = None
        
        # عناصر الواجهة
        devices_list = ft.Column(spacing=5, scroll=ft.ScrollMode.AUTO)
        status_text = ft.Text("جاري البحث عن الأجهزة...", size=13, color=ft.Colors.ORANGE_300)
        search_progress = ft.ProgressRing(width=20, height=20, color=ft.Colors.CYAN_400, stroke_width=2)
        refresh_btn = ft.IconButton(
            icon=ft.Icons.REFRESH,
            icon_color=ft.Colors.CYAN_400,
            tooltip="إعادة البحث",
            visible=False,
        )
        
        def close_dlg(e):
            # إيقاف خادم المقارنة عند الإغلاق
            if self.compare_server:
                self.compare_server.stop()
            DialogManager.close_dialog(self.page, dlg)
        
        def on_device_click(device_ip):
            """عند اختيار جهاز - بدء المقارنة"""
            close_dlg(None)
            self._perform_compare(device_ip)
        
        def create_device_item(ip):
            """إنشاء عنصر جهاز في القائمة"""
            return ft.Container(
                content=ft.Row(
                    controls=[
                        ft.Icon(ft.Icons.COMPUTER, color=ft.Colors.CYAN_400, size=24),
                        ft.Column(
                            controls=[
                                ft.Text(ip, size=14, color=ft.Colors.WHITE, weight=ft.FontWeight.W_500),
                                ft.Text("جهاز متاح للمزامنة", size=11, color=ft.Colors.GREY_400),
                            ],
                            spacing=2,
                            expand=True,
                        ),
                        ft.Icon(ft.Icons.ARROW_FORWARD_IOS, color=ft.Colors.GREY_500, size=16),
                    ],
                    alignment=ft.MainAxisAlignment.START,
                    spacing=15,
                ),
                bgcolor=ft.Colors.GREY_800,
                border_radius=10,
                padding=15,
                ink=True,
                on_click=lambda e, ip=ip: on_device_click(ip),
            )
        
        def update_devices_list(devices):
            """تحديث قائمة الأجهزة"""
            devices_list.controls.clear()
            if devices:
                for ip in devices:
                    devices_list.controls.append(create_device_item(ip))
                status_text.value = f"تم العثور على {len(devices)} جهاز"
                status_text.color = ft.Colors.GREEN_300
            else:
                devices_list.controls.append(
                    ft.Container(
                        content=ft.Column(
                            controls=[
                                ft.Icon(ft.Icons.DEVICES_OTHER, color=ft.Colors.GREY_600, size=40),
                                ft.Text("لم يتم العثور على أجهزة", size=13, color=ft.Colors.GREY_500),
                                ft.Text("تأكد من تشغيل التطبيق على الجهاز الآخر", size=11, color=ft.Colors.GREY_600),
                            ],
                            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                            spacing=5,
                        ),
                        padding=20,
                        alignment=ft.alignment.center,
                    )
                )
                status_text.value = "لم يتم العثور على أجهزة"
                status_text.color = ft.Colors.ORANGE_300
            
            search_progress.visible = False
            refresh_btn.visible = True
            self.page.update()
        
        def search_devices():
            """البحث عن الأجهزة في الشبكة"""
            search_progress.visible = True
            refresh_btn.visible = False
            status_text.value = "جاري البحث عن الأجهزة..."
            status_text.color = ft.Colors.ORANGE_300
            devices_list.controls.clear()
            devices_list.controls.append(
                ft.Container(
                    content=ft.Column(
                        controls=[
                            ft.ProgressRing(width=30, height=30, color=ft.Colors.CYAN_400, stroke_width=3),
                            ft.Text("جاري فحص الشبكة...", size=12, color=ft.Colors.GREY_400),
                        ],
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=10,
                    ),
                    padding=30,
                    alignment=ft.alignment.center,
                )
            )
            self.page.update()
            
            def do_search():
                devices = discover_devices(timeout=3)
                self.discovered_devices = devices
                update_devices_list(devices)
            
            thread = threading.Thread(target=do_search)
            thread.daemon = True
            thread.start()
        
        def on_refresh(e):
            search_devices()
        
        refresh_btn.on_click = on_refresh
        
        # بدء خادم المقارنة للسماح للأجهزة الأخرى بالاتصال
        self.compare_server = CompareServer()
        self.compare_server.start()
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.SYNC, color=ft.Colors.LIGHT_BLUE_400, size=28),
                    ft.Text("مزامنة البيانات", weight=ft.FontWeight.BOLD, color=ft.Colors.LIGHT_BLUE_200, size=18),
                ],
                spacing=10,
            ),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # معلومات الجهاز الحالي
                        ft.Container(
                            content=ft.Row(
                                controls=[
                                    ft.Icon(ft.Icons.WIFI, color=ft.Colors.CYAN_400, size=18),
                                    ft.Text(f"عنوان IP الخاص بك: {local_ip}", size=14, color=ft.Colors.CYAN_300),
                                ],
                                alignment=ft.MainAxisAlignment.CENTER,
                                spacing=8,
                            ),
                            bgcolor=ft.Colors.GREY_800,
                            border_radius=10,
                            padding=10,
                        ),
                        ft.Container(height=10),
                        # شريط الحالة
                        ft.Row(
                            controls=[
                                search_progress,
                                status_text,
                                ft.Container(expand=True),
                                refresh_btn,
                            ],
                            alignment=ft.MainAxisAlignment.START,
                            spacing=10,
                        ),
                        ft.Divider(color=ft.Colors.GREY_700),
                        # قائمة الأجهزة
                        ft.Container(
                            content=devices_list,
                            height=250,
                            border=ft.border.all(1, ft.Colors.GREY_700),
                            border_radius=10,
                            padding=10,
                        ),
                        ft.Container(height=5),
                        ft.Text(
                            "اختر جهازاً للمقارنة وإرسال الفروقات",
                            size=11,
                            color=ft.Colors.GREY_500,
                            text_align=ft.TextAlign.CENTER,
                        ),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=5,
                ),
                padding=10,
                width=380,
            ),
            actions=[
                ft.TextButton(
                    "إغلاق",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.CENTER,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()
        
        # بدء البحث تلقائياً
        search_devices()

    def _perform_compare(self, target_ip):
        """تنفيذ عملية المقارنة وعرض الفروقات"""
        from utils.sync_utils import CompareClient
        
        # عرض نافذة التحميل
        loading_dlg = DialogManager.show_loading_dialog(self.page, "جاري المقارنة...")
        
        client = CompareClient()
        
        def on_compare_complete(differences, remote_ip):
            DialogManager.close_dialog(self.page, loading_dlg)
            self._show_differences_dialog(differences, remote_ip)
        
        def on_error(error):
            DialogManager.close_dialog(self.page, loading_dlg)
            self._show_sync_result(f"خطأ: {error}", False)
        
        client.on_compare_complete = on_compare_complete
        client.on_error = on_error
        
        client.get_remote_files_info(target_ip)

    def _show_differences_dialog(self, differences, target_ip):
        """عرض نافذة الفروقات مع إمكانية الإرسال"""
        from utils.sync_utils import CompareClient
        from datetime import datetime
        
        if not differences:
            self._show_sync_result("لا توجد فروقات - البيانات متطابقة!", True)
            return
        
        # قائمة الملفات المحددة للإرسال
        selected_files = set()
        
        def format_size(size):
            if size < 1024:
                return f"{size} B"
            elif size < 1024 * 1024:
                return f"{size / 1024:.1f} KB"
            else:
                return f"{size / (1024 * 1024):.1f} MB"
        
        def format_time(timestamp):
            if timestamp == 0:
                return "-"
            return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
        
        def get_status_color(status):
            if status == 'local_only':
                return ft.Colors.BLUE_400
            elif status == 'remote_only':
                return ft.Colors.GREEN_400
            elif status == 'local_newer':
                return ft.Colors.ORANGE_400
            else:
                return ft.Colors.PURPLE_400
        
        def get_status_icon(status):
            if status == 'local_only':
                return ft.Icons.ADD_CIRCLE
            elif status == 'remote_only':
                return ft.Icons.DOWNLOAD
            elif status == 'local_newer':
                return ft.Icons.ARROW_UPWARD
            else:
                return ft.Icons.ARROW_DOWNWARD
        
        # إنشاء عناصر القائمة
        list_items = []
        checkboxes = {}
        
        def on_checkbox_change(e, file_path):
            if e.control.value:
                selected_files.add(file_path)
            else:
                selected_files.discard(file_path)
            update_send_button()
        
        def update_send_button():
            send_btn.disabled = len(selected_files) == 0
            send_btn.text = f"إرسال ({len(selected_files)})" if selected_files else "إرسال"
            self.page.update()
        
        def select_all(e):
            for path, cb in checkboxes.items():
                # فقط الملفات المحلية يمكن إرسالها
                diff = next((d for d in differences if d['path'] == path), None)
                if diff and diff['status'] in ['local_only', 'local_newer']:
                    cb.value = True
                    selected_files.add(path)
            update_send_button()
        
        def deselect_all(e):
            for cb in checkboxes.values():
                cb.value = False
            selected_files.clear()
            update_send_button()
        
        for diff in differences:
            # فقط الملفات المحلية أو الأحدث محلياً يمكن إرسالها
            can_send = diff['status'] in ['local_only', 'local_newer']
            
            cb = ft.Checkbox(
                value=False,
                disabled=not can_send,
                on_change=lambda e, p=diff['path']: on_checkbox_change(e, p),
            )
            checkboxes[diff['path']] = cb
            
            item = ft.Container(
                content=ft.Row(
                    controls=[
                        cb,
                        ft.Icon(get_status_icon(diff['status']), color=get_status_color(diff['status']), size=20),
                        ft.Column(
                            controls=[
                                ft.Text(diff['path'], size=12, color=ft.Colors.WHITE, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                                ft.Text(diff['status_text'], size=10, color=get_status_color(diff['status'])),
                            ],
                            spacing=2,
                            expand=True,
                        ),
                        ft.Column(
                            controls=[
                                ft.Text(f"محلي: {format_size(diff['local_size'])}", size=9, color=ft.Colors.GREY_400),
                                ft.Text(f"بعيد: {format_size(diff['remote_size'])}", size=9, color=ft.Colors.GREY_400),
                            ],
                            spacing=2,
                            horizontal_alignment=ft.CrossAxisAlignment.END,
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.START,
                    spacing=10,
                ),
                bgcolor=ft.Colors.GREY_800,
                border_radius=8,
                padding=10,
                margin=ft.margin.only(bottom=5),
            )
            list_items.append(item)
        
        def close_dlg(e):
            DialogManager.close_dialog(self.page, dlg)
        
        def send_selected(e):
            if not selected_files:
                return
            close_dlg(e)
            self._send_selected_files(target_ip, list(selected_files))
        
        send_btn = ft.ElevatedButton(
            "إرسال",
            icon=ft.Icons.SEND,
            bgcolor=ft.Colors.ORANGE_700,
            color=ft.Colors.WHITE,
            disabled=True,
            on_click=send_selected,
        )
        
        # إحصائيات
        local_only = len([d for d in differences if d['status'] == 'local_only'])
        remote_only = len([d for d in differences if d['status'] == 'remote_only'])
        local_newer = len([d for d in differences if d['status'] == 'local_newer'])
        remote_newer = len([d for d in differences if d['status'] == 'remote_newer'])
        
        stats_row = ft.Row(
            controls=[
                ft.Container(
                    content=ft.Text(f"محلي فقط: {local_only}", size=10, color=ft.Colors.BLUE_300),
                    bgcolor=ft.Colors.BLUE_900,
                    border_radius=5,
                    padding=5,
                ),
                ft.Container(
                    content=ft.Text(f"بعيد فقط: {remote_only}", size=10, color=ft.Colors.GREEN_300),
                    bgcolor=ft.Colors.GREEN_900,
                    border_radius=5,
                    padding=5,
                ),
                ft.Container(
                    content=ft.Text(f"محلي أحدث: {local_newer}", size=10, color=ft.Colors.ORANGE_300),
                    bgcolor=ft.Colors.ORANGE_900,
                    border_radius=5,
                    padding=5,
                ),
                ft.Container(
                    content=ft.Text(f"بعيد أحدث: {remote_newer}", size=10, color=ft.Colors.PURPLE_300),
                    bgcolor=ft.Colors.PURPLE_900,
                    border_radius=5,
                    padding=5,
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=5,
            wrap=True,
        )
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.COMPARE_ARROWS, color=ft.Colors.ORANGE_400, size=24),
                    ft.Text(f"الفروقات ({len(differences)} ملف)", weight=ft.FontWeight.BOLD, color=ft.Colors.ORANGE_200, size=16),
                ],
                spacing=10,
            ),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        stats_row,
                        ft.Divider(color=ft.Colors.GREY_700),
                        ft.Row(
                            controls=[
                                ft.TextButton("تحديد الكل", on_click=select_all, style=ft.ButtonStyle(color=ft.Colors.CYAN_300)),
                                ft.TextButton("إلغاء التحديد", on_click=deselect_all, style=ft.ButtonStyle(color=ft.Colors.GREY_400)),
                            ],
                            alignment=ft.MainAxisAlignment.CENTER,
                        ),
                        ft.Container(
                            content=ft.Column(
                                controls=list_items,
                                scroll=ft.ScrollMode.AUTO,
                                spacing=0,
                            ),
                            height=300,
                            border=ft.border.all(1, ft.Colors.GREY_700),
                            border_radius=10,
                            padding=10,
                        ),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=10,
                width=450,
            ),
            actions=[
                send_btn,
                ft.TextButton(
                    "إغلاق",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def _send_selected_files(self, target_ip, file_paths):
        """إرسال الملفات المحددة"""
        from utils.sync_utils import CompareClient
        
        progress_bar = ft.ProgressBar(width=280, value=0, color=ft.Colors.ORANGE_400, bgcolor=ft.Colors.GREY_700)
        status_text = ft.Text("جاري تجهيز الملفات...", size=14, color=ft.Colors.WHITE)
        progress_text = ft.Text("0%", size=12, color=ft.Colors.GREY_400)
        
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Icon(ft.Icons.UPLOAD, size=40, color=ft.Colors.ORANGE_400),
                        ft.Container(height=10),
                        status_text,
                        progress_bar,
                        progress_text,
                        ft.Text(f"إرسال {len(file_paths)} ملف", size=11, color=ft.Colors.GREY_500),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=30,
                width=320,
            ),
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(progress_dlg)
        progress_dlg.open = True
        self.page.update()
        
        client = CompareClient()
        
        def on_progress(percent):
            progress_bar.value = percent / 100
            progress_text.value = f"{int(percent)}%"
            if percent < 30:
                status_text.value = "جاري ضغط الملفات..."
            else:
                status_text.value = "جاري إرسال الملفات..."
            self.page.update()
        
        def on_complete(success, message):
            DialogManager.close_dialog(self.page, progress_dlg)
            self._show_sync_result(message, success)
        
        def on_error(error):
            DialogManager.close_dialog(self.page, progress_dlg)
            self._show_sync_result(error, False)
        
        client.on_send_progress = on_progress
        client.on_send_complete = on_complete
        client.on_error = on_error
        
        client.send_selected_files(target_ip, file_paths)

    def _show_sync_result(self, message, success):
        """Show sync result dialog"""
        if success:
            DialogManager.show_success_dialog(self.page, message, title="نجاح")
        else:
            DialogManager.show_error_dialog(self.page, message, title="خطأ")

    def open_invoices(self, e):
        # Clear page and load InvoiceView directly without animation
        self.page.clean()
        
        # Use save_callback from dashboard_view module
        app = InvoiceView(self.page, save_callback)
        app.build_ui()

    def open_payments(self, e):
        """Open payments management view"""
        self.page.clean()
        payments_view = PaymentsView(self.page, on_back=self.show)
        payments_view.build_ui()

    def open_attendance(self, e):
        # Clear page and load AttendanceView directly without animation
        self.page.clean()
        
        # Store save_callback for later use
        if hasattr(self, 'save_callback'):
            setattr(self.page, '_save_callback', self.save_callback)
        
        app = AttendanceView(self.page)
        app.build_ui()

    def open_blocks(self, e):
        # Clear page and load BlocksView directly without animation
        self.page.clean()
        blocks_view = BlocksView(self.page, on_back=self.go_back)
        blocks_view.build_ui()
        self.page.update()

    def open_purchases(self, e):
        # Clear page and load PurchasesView directly without animation
        self.page.clean()
        purchases_view = PurchasesView(self.page, on_back=self.go_back)
        purchases_view.build_ui()
        self.page.update()

    def open_inventory(self, e):
        """Open inventory dialog with options to add or disburse - card style"""
        def close_dlg(e):
            DialogManager.close_dialog(self.page, dlg)
        
        def open_add(e):
            close_dlg(e)
            self.open_inventory_add(None)
        
        def open_disburse(e):
            close_dlg(e)
            self.open_inventory_disburse(None)
        
        # Card for adding to inventory
        add_card = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Icon(ft.Icons.ADD_SHOPPING_CART, size=40, color=ft.Colors.WHITE),
                        ft.Text("إضافة للمخزون", size=15, weight=ft.FontWeight.W_600, text_align=ft.TextAlign.CENTER),
                        ft.Text("إضافة أصناف جديدة", size=11, color=ft.Colors.GREY_400, text_align=ft.TextAlign.CENTER),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=8,
                    tight=True,
                ),
                padding=15,
                alignment=ft.alignment.center,
                bgcolor=ft.Colors.GREEN_700,
                border_radius=15,
                ink=True,
                on_click=open_add,
                width=150,
                height=130,
            ),
            elevation=8,
        )
        
        # Card for disbursing from inventory
        disburse_card = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Icon(ft.Icons.REMOVE_SHOPPING_CART, size=40, color=ft.Colors.WHITE),
                        ft.Text("صرف من المخزون", size=15, weight=ft.FontWeight.W_600, text_align=ft.TextAlign.CENTER),
                        ft.Text("صرف أصناف موجودة", size=11, color=ft.Colors.GREY_400, text_align=ft.TextAlign.CENTER),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=8,
                    tight=True,
                ),
                padding=15,
                alignment=ft.alignment.center,
                bgcolor=ft.Colors.RED_700,
                border_radius=15,
                ink=True,
                on_click=open_disburse,
                width=150,
                height=130,
            ),
            elevation=8,
        )
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.INVENTORY, color=ft.Colors.DEEP_PURPLE_400, size=28),
                    ft.Text("المخزون", weight=ft.FontWeight.BOLD, color=ft.Colors.DEEP_PURPLE_200, size=18),
                ],
                spacing=10,
            ),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text("اختر العملية المطلوبة:", size=14, color=ft.Colors.GREY_400),
                        ft.Container(height=15),
                        ft.Row(
                            controls=[add_card, disburse_card],
                            alignment=ft.MainAxisAlignment.CENTER,
                            spacing=20,
                        ),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    tight=True,
                ),
                padding=10,
            ),
            actions=[
                ft.TextButton(
                    "إغلاق",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.CENTER,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def open_inventory_add(self, e):
        """Open add inventory dialog"""

        # Close any open dialogs first
        self.page.overlay.clear()
        self.page.update()
        # Store reference to self for back navigation
        self.page._dashboard_ref = self
        # Clear page and load InventoryAddView directly
        self.page.clean()
        inventory_view = InventoryAddView(self.page, on_back=self.go_back_to_inventory)
        inventory_view.build_ui()
        self.page.update()


    def open_inventory_disburse(self, e):
        """Open disburse inventory dialog"""

        # Close any open dialogs first
        self.page.overlay.clear()
        self.page.update()
        # Store reference to self for back navigation
        self.page._dashboard_ref = self
        # Clear page and load InventoryDisburseView directly
        self.page.clean()
        inventory_view = InventoryDisburseView(self.page, on_back=self.go_back_to_inventory)
        inventory_view.build_ui()
        self.page.update()


    def open_slides_add(self, e):
        """Open add slides inventory dialog"""

        # Close any open dialogs first
        self.page.overlay.clear()
        self.page.update()
        # Store reference to self for back navigation
        self.page._dashboard_ref = self
        # Clear page and load SlidesAddView directly
        self.page.clean()
        slides_view = SlidesAddView(self.page, on_back=self.go_back_to_inventory)
        slides_view.build_ui()
        self.page.update()


    def go_back_to_inventory(self):
        """Go back to the main dashboard"""

        # Completely clear all overlays to prevent accumulation
        self.page.overlay.clear()
        self.page.update()
        # Show the main dashboard
        self.show(getattr(self, 'save_callback', None))
    def go_back(self):
        self.reset_ui()
        self.page.add(self.main_container)
        self.main_container.opacity = 1
        self.page.update()

    def reset_ui(self):
        self.page.clean()
        self.page.appbar = None
        self.page.floating_action_button = None
        self.page.title = "مصنع السويفي"
        self.page.update()

    def show(self, callback=None):
        if callback:
            self.save_callback = callback
        self.reset_ui()
        self.page.add(self.main_container)