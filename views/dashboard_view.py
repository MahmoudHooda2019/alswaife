import flet as ft
import os
from views.invoice_view import InvoiceView
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
                self.create_menu_button("إدارة الفواتير", ft.Icons.RECEIPT_LONG, self.open_invoices),
                self.create_menu_button("الحضور والإنصراف", ft.Icons.PERSON, None),
                self.create_menu_button("المخازن", ft.Icons.INVENTORY_2, None),
                self.create_menu_button("العملاء", ft.Icons.PEOPLE, None),
                self.create_menu_button("الإعدادات", ft.Icons.SETTINGS, None),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=20
        )

    def create_menu_button(self, text, icon, on_click):
        return ft.Container(
            content=ft.Row(
                controls=[
                    ft.Icon(icon, size=30, color=ft.Colors.WHITE),
                    ft.Text(text, size=20, weight=ft.FontWeight.W_500)
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=15
            ),
            width=300,
            height=60,
            bgcolor=ft.Colors.BLUE_GREY_800,
            border_radius=10,
            ink=True,
            on_click=on_click if on_click else lambda e: self.show_placeholder(text),
            animate=ft.Animation(300, ft.AnimationCurve.EASE_OUT),
            padding=10
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
