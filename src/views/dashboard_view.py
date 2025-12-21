import flet as ft
import os
from views.invoice_view import InvoiceView
from views.attendance_view import AttendanceView
from views.blocks_view import BlocksView
from views.purchases_view import PurchasesView
from views.inventory_view import InventoryAddView
from views.inventory_disburse_view import InventoryDisburseView
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
                        self.create_menu_card("البلوكات", ft.Icons.VIEW_IN_AR, self.open_blocks, ft.Colors.AMBER_700),
                        self.create_menu_card("المشتريات", ft.Icons.SHOPPING_CART, self.open_purchases, ft.Colors.CYAN_700),
                        # Removed the single inventory card and added two separate cards
                        self.create_menu_card("إضافة للمخزون", ft.Icons.ADD_SHOPPING_CART, self.open_inventory_add, ft.Colors.GREEN_700),
                        self.create_menu_card("صرف من المخزون", ft.Icons.REMOVE_SHOPPING_CART, self.open_inventory_disburse, ft.Colors.RED_700),
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
                        self.create_menu_card("البلوكات", ft.Icons.VIEW_IN_AR, self.open_blocks, ft.Colors.AMBER_700),
                        self.create_menu_card("المشتريات", ft.Icons.SHOPPING_CART, self.open_purchases, ft.Colors.CYAN_700),
                        # Removed the single inventory card and added two separate cards
                        self.create_menu_card("إضافة للمخزون", ft.Icons.ADD_SHOPPING_CART, self.open_inventory_add, ft.Colors.GREEN_700),
                        self.create_menu_card("صرف من المخزون", ft.Icons.REMOVE_SHOPPING_CART, self.open_inventory_disburse, ft.Colors.RED_700),
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