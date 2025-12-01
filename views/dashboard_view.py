import flet as ft
import os
from views.invoice_view import InvoiceView
from views.attendance_view import AttendanceView
from utils.path_utils import resource_path

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
                        self.create_menu_card("المخازن", ft.Icons.INVENTORY_2, None, ft.Colors.ORANGE_700),
                        self.create_menu_card("العملاء", ft.Icons.PEOPLE, None, ft.Colors.PURPLE_700),
                        self.create_menu_card("الإعدادات", ft.Icons.SETTINGS, None, ft.Colors.RED_700),
                        self.create_menu_card("التقارير", ft.Icons.ASSESSMENT, None, ft.Colors.TEAL_700),
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
        message = f"خاصية {feature} قيد التطوير" if feature else "هذه الخاصية قيد التطوير"
        
        def close_dlg(e):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            title=ft.Text("تنبيه"),
            content=ft.Text(message),
            actions=[
                ft.TextButton("حسناً", on_click=close_dlg)
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()

    def open_invoices(self, e):
        # Animate out
        self.main_container.opacity = 0
        self.page.update()
        
        # Wait for animation (simulated)
        import time
        time.sleep(0.3)
        
        # Clear page and load InvoiceView
        self.page.clean()
        
        
        if hasattr(self, 'save_callback'):
             app = InvoiceView(self.page, self.save_callback)
             # Add a back button to the app
             back_btn = ft.IconButton(
                 icon=ft.Icons.ARROW_BACK, 
                 on_click=lambda _: self.go_back(),
                 tooltip="العودة للقائمة الرئيسية"
             )
             self.page.add(ft.Row([back_btn]))
             app.build_ui()
        else:
             self.page.add(ft.Text("Error: Save callback not found"))

    def open_attendance(self, e):
        # Animate out
        self.main_container.opacity = 0
        self.page.update()
        
        # Wait for animation (simulated)
        import time
        time.sleep(0.3)
        
        # Clear page and load AttendanceView
        self.page.clean()
        
        # Store save_callback for later use
        if hasattr(self, 'save_callback'):
            setattr(self.page, '_save_callback', self.save_callback)
        
        app = AttendanceView(self.page)
        app.build_ui()


    def reset_ui(self):
        self.page.clean()
        self.page.appbar = None
        self.page.floating_action_button = None
        self.page.title = "مصنع السويفي"
        self.page.update()

    def go_back(self):
        self.reset_ui()
        self.page.add(self.main_container)
        self.main_container.opacity = 1
        self.page.update()

    def show(self, save_callback):
        self.save_callback = save_callback
        self.reset_ui()
        self.page.add(self.main_container)