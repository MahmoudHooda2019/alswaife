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
from utils.invoice_utils import save_invoice, update_client_ledger
from utils.log_utils import log_error, log_exception


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
                        self.create_menu_card("الحضور والإنصراف", ft.Icons.PERSON, self.open_attendance, ft.Colors.GREEN_700),
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
                        self.create_menu_card("الحضور والإنصراف", ft.Icons.PERSON, self.open_attendance, ft.Colors.GREEN_700),
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
        
        # Quick report buttons - Row 1: Income & Expenses
        quick_buttons_row1 = ft.Row(
            controls=[
                ft.ElevatedButton(
                    "الإيرادات",
                    icon=ft.Icons.TRENDING_UP,
                    bgcolor=ft.Colors.GREEN_700,
                    color=ft.Colors.WHITE,
                    width=120,
                    on_click=lambda e: quick_report("income"),
                ),
                ft.ElevatedButton(
                    "المصروفات",
                    icon=ft.Icons.TRENDING_DOWN,
                    bgcolor=ft.Colors.RED_700,
                    color=ft.Colors.WHITE,
                    width=120,
                    on_click=lambda e: quick_report("expenses"),
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=10,
        )
        
        # Quick report buttons - Row 2: Inventory
        quick_buttons_row2 = ft.Row(
            controls=[
                ft.ElevatedButton(
                    "المخزون",
                    icon=ft.Icons.INVENTORY,
                    bgcolor=ft.Colors.INDIGO_700,
                    color=ft.Colors.WHITE,
                    width=120,
                    on_click=lambda e: quick_report("inventory"),
                ),
                ft.ElevatedButton(
                    "الحضور",
                    icon=ft.Icons.PERSON,
                    bgcolor=ft.Colors.BLUE_700,
                    color=ft.Colors.WHITE,
                    width=120,
                    on_click=lambda e: quick_report("attendance"),
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=10,
        )
        
        # Quick report buttons - Row 3: Blocks & Slides
        quick_buttons_row3 = ft.Row(
            controls=[
                ft.ElevatedButton(
                    "البلوكات",
                    icon=ft.Icons.VIEW_IN_AR,
                    bgcolor=ft.Colors.AMBER_700,
                    color=ft.Colors.WHITE,
                    width=120,
                    on_click=lambda e: quick_report("blocks"),
                ),
                ft.ElevatedButton(
                    "الشرائح",
                    icon=ft.Icons.LAYERS,
                    bgcolor=ft.Colors.PURPLE_700,
                    color=ft.Colors.WHITE,
                    width=120,
                    on_click=lambda e: quick_report("slides"),
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=10,
        )
        
        # Machine number dropdown for machine production report
        machine_dropdown = ft.Dropdown(
            label="رقم الماكينة",
            width=100,
            options=[
                ft.dropdown.Option("1", "1"),
                ft.dropdown.Option("2", "2"),
                ft.dropdown.Option("3", "3"),
                ft.dropdown.Option("4", "4"),
                ft.dropdown.Option("5", "5"),
            ],
            value="1",
            border_radius=10,
            filled=True,
            bgcolor=ft.Colors.GREY_800,
        )
        
        def machine_report(e):
            close_dlg(None)
            query = {"report_type": "machine_production", "machine_number": machine_dropdown.value}
            self.execute_report_with_progress(query)
        
        # Quick report buttons - Row 4: Machine Production
        quick_buttons_row4 = ft.Row(
            controls=[
                ft.ElevatedButton(
                    "إنتاج الماكينة",
                    icon=ft.Icons.PRECISION_MANUFACTURING,
                    bgcolor=ft.Colors.CYAN_700,
                    color=ft.Colors.WHITE,
                    width=140,
                    on_click=machine_report,
                ),
                machine_dropdown,
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
                        quick_buttons_row1,
                        quick_buttons_row2,
                        quick_buttons_row3,
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=8,
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

    def open_sync(self, e):
        """Open sync dialog for LAN data synchronization"""
        from utils.sync_utils import get_local_ip
        
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        local_ip = get_local_ip()
        
        # Card for sending data
        send_card = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Icon(ft.Icons.UPLOAD, size=45, color=ft.Colors.WHITE),
                        ft.Text("إرسال البيانات", size=16, weight=ft.FontWeight.W_600, text_align=ft.TextAlign.CENTER),
                        ft.Text("إرسال لجهاز آخر", size=12, color=ft.Colors.GREY_400, text_align=ft.TextAlign.CENTER),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=20,
                alignment=ft.alignment.center,
                bgcolor=ft.Colors.BLUE_700,
                border_radius=15,
                ink=True,
                on_click=lambda e: self._open_send_dialog(close_dlg),
                width=160,
                height=150,
            ),
            elevation=8,
        )
        
        # Card for receiving data
        receive_card = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Icon(ft.Icons.DOWNLOAD, size=45, color=ft.Colors.WHITE),
                        ft.Text("استقبال البيانات", size=16, weight=ft.FontWeight.W_600, text_align=ft.TextAlign.CENTER),
                        ft.Text("استقبال من جهاز آخر", size=12, color=ft.Colors.GREY_400, text_align=ft.TextAlign.CENTER),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=20,
                alignment=ft.alignment.center,
                bgcolor=ft.Colors.GREEN_700,
                border_radius=15,
                ink=True,
                on_click=lambda e: self._open_receive_dialog(close_dlg),
                width=160,
                height=150,
            ),
            elevation=8,
        )
        
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
                        ft.Container(height=15),
                        ft.Text("اختر العملية:", size=14, color=ft.Colors.GREY_400),
                        ft.Container(height=10),
                        ft.Row(
                            controls=[send_card, receive_card],
                            alignment=ft.MainAxisAlignment.CENTER,
                            spacing=20,
                        ),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
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

    def _open_send_dialog(self, close_parent):
        """فتح نافذة إرسال البيانات"""
        close_parent(None)
        
        ip_field = ft.TextField(
            label="عنوان IP للجهاز المستقبل",
            hint_text="مثال: 192.168.1.100",
            value="192.168.",
            width=280,
            border_radius=10,
            filled=True,
            bgcolor=ft.Colors.GREY_800,
            border_color=ft.Colors.GREY_600,
            focused_border_color=ft.Colors.BLUE_400,
            prefix_icon=ft.Icons.COMPUTER,
            rtl=False,
            text_align=ft.TextAlign.LEFT,
        )
        
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        def start_send(e):
            target_ip = ip_field.value.strip()
            if not target_ip:
                ip_field.error_text = "يرجى إدخال عنوان IP"
                self.page.update()
                return
            close_dlg(e)
            self._perform_send(target_ip)
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.UPLOAD, color=ft.Colors.BLUE_400, size=24),
                    ft.Text("إرسال البيانات", weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_200, size=16),
                ],
                spacing=10,
            ),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text("تأكد من أن الجهاز الآخر في وضع الاستقبال", size=13, color=ft.Colors.ORANGE_300),
                        ft.Container(height=15),
                        ip_field,
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=5,
                ),
                padding=10,
                width=320,
            ),
            actions=[
                ft.ElevatedButton(
                    "إرسال",
                    icon=ft.Icons.SEND,
                    bgcolor=ft.Colors.BLUE_700,
                    color=ft.Colors.WHITE,
                    on_click=start_send,
                ),
                ft.TextButton(
                    "إلغاء",
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

    def _open_receive_dialog(self, close_parent):
        """فتح نافذة استقبال البيانات"""
        from utils.sync_utils import get_local_ip, SyncServer
        
        close_parent(None)
        
        local_ip = get_local_ip()
        self.sync_server = SyncServer()
        
        progress_bar = ft.ProgressBar(width=280, value=0, color=ft.Colors.GREEN_400, bgcolor=ft.Colors.GREY_700)
        status_text = ft.Text("في انتظار الاتصال...", size=14, color=ft.Colors.WHITE)
        progress_text = ft.Text("0%", size=12, color=ft.Colors.GREY_400)
        
        def close_dlg(e):
            self.sync_server.stop()
            dlg.open = False
            self.page.update()
        
        def on_progress(percent):
            progress_bar.value = percent / 100
            progress_text.value = f"{int(percent)}%"
            status_text.value = "جاري استقبال البيانات..."
            self.page.update()
        
        def on_complete(success, message):
            self.sync_server.stop()
            dlg.open = False
            self.page.update()
            self._show_sync_result(message, success)
        
        def on_error(error):
            self.sync_server.stop()
            dlg.open = False
            self.page.update()
            self._show_sync_result(f"خطأ: {error}", False)
        
        self.sync_server.on_progress = on_progress
        self.sync_server.on_complete = on_complete
        self.sync_server.on_error = on_error
        
        success, result = self.sync_server.start()
        
        if not success:
            self._show_sync_result(f"فشل في بدء الخادم: {result}", False)
            return
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.DOWNLOAD, color=ft.Colors.GREEN_400, size=24),
                    ft.Text("استقبال البيانات", weight=ft.FontWeight.BOLD, color=ft.Colors.GREEN_200, size=16),
                ],
                spacing=10,
            ),
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Container(
                            content=ft.Column(
                                controls=[
                                    ft.Text("أدخل هذا العنوان في الجهاز المرسل:", size=13, color=ft.Colors.GREY_400),
                                    ft.Container(
                                        content=ft.Text(local_ip, size=24, weight=ft.FontWeight.BOLD, color=ft.Colors.CYAN_300, selectable=True),
                                        bgcolor=ft.Colors.GREY_800,
                                        border_radius=10,
                                        padding=15,
                                        alignment=ft.alignment.center,
                                    ),
                                ],
                                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                spacing=10,
                            ),
                            padding=10,
                        ),
                        ft.Divider(color=ft.Colors.GREY_700),
                        ft.ProgressRing(width=30, height=30, color=ft.Colors.GREEN_400, stroke_width=3),
                        status_text,
                        progress_bar,
                        progress_text,
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=10,
                width=320,
            ),
            actions=[
                ft.TextButton(
                    "إلغاء",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.RED_300)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.CENTER,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def _perform_send(self, target_ip):
        """تنفيذ عملية الإرسال"""
        from utils.sync_utils import SyncClient
        
        progress_bar = ft.ProgressBar(width=280, value=0, color=ft.Colors.BLUE_400, bgcolor=ft.Colors.GREY_700)
        status_text = ft.Text("جاري تجهيز البيانات...", size=14, color=ft.Colors.WHITE)
        progress_text = ft.Text("0%", size=12, color=ft.Colors.GREY_400)
        
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Icon(ft.Icons.UPLOAD, size=40, color=ft.Colors.BLUE_400),
                        ft.Container(height=10),
                        status_text,
                        progress_bar,
                        progress_text,
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
        
        client = SyncClient()
        
        def on_progress(percent):
            progress_bar.value = percent / 100
            progress_text.value = f"{int(percent)}%"
            if percent < 30:
                status_text.value = "جاري ضغط البيانات..."
            else:
                status_text.value = "جاري إرسال البيانات..."
            self.page.update()
        
        def on_complete(success, message):
            progress_dlg.open = False
            self.page.update()
            self._show_sync_result(message, success)
        
        def on_error(error):
            progress_dlg.open = False
            self.page.update()
            self._show_sync_result(error, False)
        
        client.on_progress = on_progress
        client.on_complete = on_complete
        client.on_error = on_error
        
        client.send_data(target_ip)
    
    def _show_sync_result(self, message, success):
        """Show sync result dialog"""
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        icon = ft.Icons.CHECK_CIRCLE if success else ft.Icons.ERROR
        color = ft.Colors.GREEN_400 if success else ft.Colors.RED_400
        title_color = ft.Colors.GREEN_300 if success else ft.Colors.RED_300
        title_text = "نجاح" if success else "خطأ"
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(icon, color=color, size=28),
                    ft.Text(title_text, weight=ft.FontWeight.BOLD, color=title_color, size=16),
                ],
                spacing=10,
            ),
            content=ft.Text(message, size=14, color=ft.Colors.WHITE, rtl=True),
            actions=[
                ft.TextButton(
                    "حسناً",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.LIGHT_BLUE_300)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.CENTER,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def open_invoices(self, e):
        # Clear page and load InvoiceView directly without animation
        self.page.clean()
        
        # Use save_callback from dashboard_view module
        app = InvoiceView(self.page, save_callback)
        app.build_ui()

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
            dlg.open = False
            self.page.update()
        
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

    def show(self, callback=None):
        if callback:
            self.save_callback = callback
        self.reset_ui()
        self.page.add(self.main_container)