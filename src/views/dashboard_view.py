import flet as ft
import os
from views.invoice_view import InvoiceView
from views.attendance_view import AttendanceView
from views.blocks_view import BlocksView
from views.purchases_view import PurchasesView
from views.inventory_add_view import InventoryAddView
from views.inventory_disburse_view import InventoryDisburseView
from views.slides_add_view import SlidesAddView
from utils.path_utils import resource_path
from utils.reports_utils import parse_user_request_with_ai, execute_report
from utils.update_utils import check_for_updates, download_update, install_update

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
                        self.create_menu_card("الحضور والإنصراف", ft.Icons.PERSON, self.open_attendance, ft.Colors.GREEN_700),
                        self.create_menu_card("إضافة بلوكات", ft.Icons.VIEW_IN_AR, self.open_blocks, ft.Colors.AMBER_700),
                        self.create_menu_card("مشتري", ft.Icons.SHOPPING_CART, self.open_purchases, ft.Colors.CYAN_700),
                        self.create_menu_card("إضافة للمخزون", ft.Icons.ADD_SHOPPING_CART, self.open_inventory_add, ft.Colors.GREEN_700),
                        self.create_menu_card("صرف من المخزون", ft.Icons.REMOVE_SHOPPING_CART, self.open_inventory_disburse, ft.Colors.RED_700),
                        self.create_menu_card("إضافة شرائح", ft.Icons.ADD, self.open_slides_add, ft.Colors.BLUE_700),
                        self.create_menu_card("التقارير", ft.Icons.ASSESSMENT, self.open_reports, ft.Colors.TEAL_700),
                        self.create_menu_card("تحديث", ft.Icons.SYSTEM_UPDATE, self.open_update, ft.Colors.ORANGE_700),
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
                        self.create_menu_card("الحضور والإنصراف", ft.Icons.PERSON, self.open_attendance, ft.Colors.GREEN_700),
                        self.create_menu_card("إضافة بلوكات", ft.Icons.VIEW_IN_AR, self.open_blocks, ft.Colors.AMBER_700),
                        self.create_menu_card("مشتري", ft.Icons.SHOPPING_CART, self.open_purchases, ft.Colors.CYAN_700),
                        self.create_menu_card("إضافة للمخزون", ft.Icons.ADD_SHOPPING_CART, self.open_inventory_add, ft.Colors.GREEN_700),
                        self.create_menu_card("صرف من المخزون", ft.Icons.REMOVE_SHOPPING_CART, self.open_inventory_disburse, ft.Colors.RED_700),
                        self.create_menu_card("إضافة شرائح", ft.Icons.ADD, self.open_slides_add, ft.Colors.BLUE_700),
                        self.create_menu_card("التقارير", ft.Icons.ASSESSMENT, self.open_reports, ft.Colors.TEAL_700),
                        self.create_menu_card("تحديث", ft.Icons.SYSTEM_UPDATE, self.open_update, ft.Colors.ORANGE_700),
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
        
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            title=ft.Text("تنبيه"),
            content=ft.Text(message, rtl=True),
            actions=[
                ft.TextButton("حسناً", on_click=close_dlg)
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def show_about_dialog(self, e):
        """Show about dialog with developer information"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
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
        """Open reports dialog with AI-powered natural language input"""
        
        # Text field for user request
        request_field = ft.TextField(
            label="اكتب طلبك هنا",
            hint_text="مثال: أعطني تقرير الإيرادات من 1/10 حتى 31/10",
            multiline=True,
            min_lines=2,
            max_lines=4,
            width=400,
            border_radius=10,
            filled=True,
            bgcolor=ft.Colors.GREY_800,
            border_color=ft.Colors.GREY_600,
            focused_border_color=ft.Colors.TEAL_400,
            rtl=True,
        )
        
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        def submit_request(e):
            user_request = request_field.value.strip()
            if not user_request:
                return
            close_dlg(None)
            self.process_report_request(user_request)
        
        def quick_report(report_type):
            close_dlg(None)
            # Create a simple query for quick reports
            query = {"report_type": report_type}
            self.execute_report_with_progress(query)
        
        # Quick report buttons
        quick_buttons = ft.Row(
            controls=[
                ft.ElevatedButton(
                    "الإيرادات",
                    icon=ft.Icons.TRENDING_UP,
                    bgcolor=ft.Colors.GREEN_700,
                    color=ft.Colors.WHITE,
                    on_click=lambda e: quick_report("income"),
                ),
                ft.ElevatedButton(
                    "المصروفات",
                    icon=ft.Icons.TRENDING_DOWN,
                    bgcolor=ft.Colors.RED_700,
                    color=ft.Colors.WHITE,
                    on_click=lambda e: quick_report("expenses"),
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=10,
        )
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.AUTO_AWESOME, color=ft.Colors.TEAL_400, size=28),
                    ft.Text("التقارير الذكية", weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_200),
                ],
                spacing=10,
            ),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text(
                            "اكتب طلبك بالعربية وسيتم تحليله تلقائياً",
                            size=14,
                            color=ft.Colors.GREY_400,
                            text_align=ft.TextAlign.CENTER,
                        ),
                        ft.Container(height=10),
                        request_field,
                        ft.Container(height=15),
                        ft.Text("أو اختر تقرير سريع:", size=12, color=ft.Colors.GREY_500),
                        quick_buttons,
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=5,
                ),
                padding=10,
                width=420,
            ),
            actions=[
                ft.ElevatedButton(
                    "إنشاء التقرير",
                    icon=ft.Icons.PLAY_ARROW,
                    bgcolor=ft.Colors.TEAL_700,
                    color=ft.Colors.WHITE,
                    on_click=submit_request,
                ),
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

    def process_report_request(self, user_request: str):
        """Process user's natural language request"""
        # Show progress
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.ProgressRing(width=50, height=50, color=ft.Colors.TEAL_400),
                        ft.Container(height=20),
                        ft.Text("جاري تحليل الطلب...", size=16, color=ft.Colors.WHITE),
                        ft.Text("يتم استخدام الذكاء الاصطناعي", size=12, color=ft.Colors.GREY_400),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=30,
                width=300,
            ),
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(progress_dlg)
        progress_dlg.open = True
        self.page.update()
        
        import threading
        def process():
            try:
                # Parse request with AI
                query = parse_user_request_with_ai(user_request)
                print(f"[DEBUG] Parsed query: {query}")
                
                progress_dlg.open = False
                self.page.update()
                
                # Execute report
                self.execute_report_with_progress(query)
            except Exception as ex:
                progress_dlg.open = False
                self.page.update()
                self.show_report_error(f"فشل في تحليل الطلب: {str(ex)}")
        
        thread = threading.Thread(target=process)
        thread.start()

    def execute_report_with_progress(self, query: dict):
        """Execute report with progress indicator"""
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.ProgressRing(width=50, height=50, color=ft.Colors.TEAL_400),
                        ft.Container(height=20),
                        ft.Text("جاري إنشاء التقرير...", size=16, color=ft.Colors.WHITE),
                        ft.Text("يتم قراءة وتحليل البيانات", size=12, color=ft.Colors.GREY_400),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=30,
                width=300,
            ),
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(progress_dlg)
        progress_dlg.open = True
        self.page.update()
        
        import threading
        def generate():
            try:
                documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
                result = execute_report(query, documents_path)
                
                progress_dlg.open = False
                self.page.update()
                
                if result:
                    self.show_report_success(result)
                else:
                    self.show_report_error("لا توجد بيانات متاحة للتقرير المطلوب")
            except Exception as ex:
                progress_dlg.open = False
                self.page.update()
                self.show_report_error(f"فشل في إنشاء التقرير: {str(ex)}")
        
        thread = threading.Thread(target=generate)
        thread.start()

    def show_report_success(self, filepath):
        """Show success dialog with option to open report"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        def open_file(e):
            close_dlg(e)
            try:
                os.startfile(filepath)
            except:
                pass
        
        def open_folder(e):
            close_dlg(e)
            try:
                os.startfile(os.path.dirname(filepath))
            except:
                pass
        
        dlg = ft.AlertDialog(
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=28),
                    ft.Text("تم إنشاء التقرير بنجاح", weight=ft.FontWeight.BOLD, color=ft.Colors.GREEN_300),
                ],
                spacing=10,
            ),
            content=ft.Text(f"تم حفظ التقرير في:\n{os.path.basename(filepath)}", size=14, rtl=True),
            actions=[
                ft.TextButton("فتح التقرير", on_click=open_file, style=ft.ButtonStyle(color=ft.Colors.GREEN_300)),
                ft.TextButton("فتح المجلد", on_click=open_folder, style=ft.ButtonStyle(color=ft.Colors.BLUE_300)),
                ft.TextButton("إغلاق", on_click=close_dlg, style=ft.ButtonStyle(color=ft.Colors.GREY_400)),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def show_report_error(self, error_msg):
        """Show error dialog"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.ERROR, color=ft.Colors.RED_400, size=28),
                    ft.Text("خطأ", weight=ft.FontWeight.BOLD, color=ft.Colors.RED_300),
                ],
                spacing=10,
            ),
            content=ft.Text(error_msg, size=14, rtl=True),
            actions=[
                ft.TextButton("إغلاق", on_click=close_dlg, style=ft.ButtonStyle(color=ft.Colors.GREY_400)),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def open_update(self, e):
        """Open update dialog to check and download updates"""
        import threading
        
        # Show checking progress
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Column(
                controls=[
                    ft.ProgressRing(width=40, height=40, color=ft.Colors.ORANGE_400),
                    ft.Text("جاري التحقق...", size=14, color=ft.Colors.WHITE),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=10,
                tight=True,
            ),
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=10),
        )
        self.page.overlay.append(progress_dlg)
        progress_dlg.open = True
        self.page.update()
        
        def check():
            try:
                update_available, current_ver, latest_ver, download_url = check_for_updates()
                
                progress_dlg.open = False
                self.page.update()
                
                if update_available and download_url:
                    self.show_update_available_dialog(current_ver, latest_ver, download_url)
                else:
                    self.show_no_update_dialog(current_ver, latest_ver)
            except Exception as ex:
                progress_dlg.open = False
                self.page.update()
                self.show_update_error(f"فشل في التحقق من التحديثات: {str(ex)}")
        
        thread = threading.Thread(target=check)
        thread.start()

    def show_update_available_dialog(self, current_ver, latest_ver, download_url):
        """Show dialog when update is available"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
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
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=24),
                    ft.Text("لا يوجد تحديث", weight=ft.FontWeight.BOLD, color=ft.Colors.GREEN_300, size=16, rtl=True),
                ],
                spacing=8,
                rtl=True,
            ),
            content=ft.Row(
                controls=[
                    ft.Text("الإصدار:", size=13, color=ft.Colors.GREY_400, rtl=True),
                    ft.Text(current_ver, size=13, color=ft.Colors.GREEN_400, weight=ft.FontWeight.BOLD),
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=5,
                rtl=True,
            ),
            actions=[
                ft.TextButton(
                    "حسناً",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREEN_300)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.CENTER,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=10),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def download_and_install_update(self, download_url):
        """Download and install the update"""
        import threading
        
        # Progress bar and text
        progress_bar = ft.ProgressBar(width=200, color=ft.Colors.ORANGE_400, bgcolor=ft.Colors.GREY_700)
        progress_text = ft.Text("0%", size=12, color=ft.Colors.WHITE)
        status_text = ft.Text("جاري التحميل...", size=14, color=ft.Colors.WHITE)
        
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Column(
                controls=[
                    ft.Icon(ft.Icons.DOWNLOAD, size=35, color=ft.Colors.ORANGE_400),
                    status_text,
                    progress_bar,
                    progress_text,
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
        
        def download():
            try:
                setup_path = download_update(download_url, update_progress)
                
                if setup_path:
                    status_text.value = "جاري تشغيل المثبت..."
                    self.page.update()
                    
                    if install_update(setup_path):
                        progress_dlg.open = False
                        self.page.update()
                        self.show_install_success_dialog()
                    else:
                        progress_dlg.open = False
                        self.page.update()
                        self.show_update_error("فشل في تشغيل المثبت")
                else:
                    progress_dlg.open = False
                    self.page.update()
                    self.show_update_error("فشل في تحميل التحديث")
            except Exception as ex:
                progress_dlg.open = False
                self.page.update()
                self.show_update_error(f"خطأ: {str(ex)}")
        
        thread = threading.Thread(target=download)
        thread.start()

    def show_install_success_dialog(self):
        """Show dialog after installer starts"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
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
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.ERROR, color=ft.Colors.RED_400, size=28),
                    ft.Text("خطأ في التحديث", weight=ft.FontWeight.BOLD, color=ft.Colors.RED_300),
                ],
                spacing=10,
            ),
            content=ft.Text(error_msg, size=14, rtl=True, color=ft.Colors.WHITE),
            actions=[
                ft.TextButton("إغلاق", on_click=close_dlg, style=ft.ButtonStyle(color=ft.Colors.GREY_400)),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def open_invoices(self, e):
        # Clear page and load InvoiceView directly without animation
        self.page.clean()
        
        if hasattr(self, 'save_callback'):
            app = InvoiceView(self.page, self.save_callback)
            app.build_ui()
        else:
            self.page.add(ft.Text("Error: Save callback not found"))

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

    def open_inventory_add(self, e):
        """Open add inventory dialog"""

        # Close any open dialogs first
        for overlay in self.page.overlay:
            if hasattr(overlay, 'open') and overlay.open:
                overlay.open = False
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
        for overlay in self.page.overlay:
            if hasattr(overlay, 'open') and overlay.open:
                overlay.open = False
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
        for overlay in self.page.overlay:
            if hasattr(overlay, 'open') and overlay.open:
                overlay.open = False
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

    def show(self, save_callback):
        self.save_callback = save_callback
        self.reset_ui()
        self.page.add(self.main_container)