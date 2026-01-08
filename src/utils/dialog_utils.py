import flet as ft
import asyncio

class DialogManager:
    """
    Utility class for managing dialogs in the application.
    Ensures consistent styling and safe overlay management.
    """
    
    @staticmethod
    def show_success_dialog(page: ft.Page, message: str, title: str = "نجاح", on_dismiss=None):
        """Show a success dialog with green styling"""
        DialogManager._show_basic_dialog(
            page, 
            title, 
            message, 
            ft.Icons.CHECK_CIRCLE, 
            ft.Colors.GREEN_400, 
            ft.Colors.GREEN_300,
            on_dismiss
        )

    @staticmethod
    def show_error_dialog(page: ft.Page, message: str, title: str = "خطأ", on_dismiss=None):
        """Show an error dialog with red styling"""
        DialogManager._show_basic_dialog(
            page, 
            title, 
            message, 
            ft.Icons.ERROR, 
            ft.Colors.RED_400, 
            ft.Colors.RED_300,
            on_dismiss
        )

    @staticmethod
    def show_warning_dialog(page: ft.Page, message: str, title: str = "تنبيه", on_dismiss=None):
        """Show a warning dialog with orange styling"""
        DialogManager._show_basic_dialog(
            page, 
            title, 
            message, 
            ft.Icons.WARNING_AMBER_ROUNDED, 
            ft.Colors.ORANGE_400, 
            ft.Colors.ORANGE_300,
            on_dismiss
        )

    @staticmethod
    def show_info_dialog(page: ft.Page, message: str, title: str = "معلومات", on_dismiss=None):
        """Show an info dialog with blue styling"""
        DialogManager._show_basic_dialog(
            page, 
            title, 
            message, 
            ft.Icons.INFO, 
            ft.Colors.BLUE_400, 
            ft.Colors.BLUE_300,
            on_dismiss
        )

    @staticmethod
    def show_confirm_dialog(page: ft.Page, message: str, on_confirm, title: str = "تأكيد", confirm_text: str = "نعم", cancel_text: str = "لا"):
        """Show a confirmation dialog"""
        DialogManager._cleanup_overlay(page)
        
        def close_dlg(e):
            DialogManager.close_dialog(page, dlg)

        def on_confirm_click(e):
            close_dlg(e)
            if on_confirm:
                on_confirm()

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.HELP_OUTLINE, color=ft.Colors.CYAN_400, size=28),
                    ft.Text(title, weight=ft.FontWeight.BOLD, color=ft.Colors.CYAN_300),
                ],
                spacing=10,
                rtl=True
            ),
            content=ft.Text(message, size=14, color=ft.Colors.WHITE, rtl=True),
            actions=[
                ft.ElevatedButton(
                    confirm_text,
                    bgcolor=ft.Colors.RED_700,
                    color=ft.Colors.WHITE,
                    on_click=on_confirm_click
                ),
                ft.TextButton(
                    cancel_text,
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.GREY_400)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        
        page.overlay.append(dlg)
        dlg.open = True
        page.update()

    @staticmethod
    def show_loading_dialog(page: ft.Page, message: str = "جاري التحميل..."):
        """Show a loading dialog and return it so it can be closed later"""
        DialogManager._cleanup_overlay(page)
        
        dlg = ft.AlertDialog(
            modal=True,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.ProgressRing(width=40, height=40, color=ft.Colors.ORANGE_400, stroke_width=4),
                        ft.Container(height=10),
                        ft.Text(message, size=14, color=ft.Colors.WHITE),
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                    tight=True
                ),
                padding=30,
                width=200,
            ),
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        page.overlay.append(dlg)
        dlg.open = True
        page.update()
        return dlg

    @staticmethod
    def close_dialog(page: ft.Page, dlg: ft.AlertDialog):
        """Safely close a dialog"""
        try:
            dlg.open = False
            page.update()
            # Note: We do not remove from overlay here to avoid crashes during event handling.
            # Cleanup happens when opening the next dialog.
        except Exception:
            pass

    @staticmethod
    def _cleanup_overlay(page: ft.Page):
        """Clean up closed dialogs from overlay to prevent memory leaks"""
        try:
            # Remove any AlertDialogs that are not open
            # We iterate over a copy of the list to safely remove items
            for control in list(page.overlay):
                if isinstance(control, ft.AlertDialog) and not control.open:
                    page.overlay.remove(control)
        except Exception:
            pass

    @staticmethod
    def _show_basic_dialog(page: ft.Page, title: str, message: str, icon: str, icon_color: str, title_color: str, on_dismiss=None):
        """Internal method to build and show a basic dialog"""
        DialogManager._cleanup_overlay(page)
        
        def close_dlg(e):
            DialogManager.close_dialog(page, dlg)
            if on_dismiss:
                on_dismiss()

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(icon, color=icon_color, size=28),
                    ft.Text(title, weight=ft.FontWeight.BOLD, color=title_color, size=16),
                ],
                spacing=10,
                rtl=True
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
        
        page.overlay.append(dlg)
        dlg.open = True
        page.update()

    @staticmethod
    def show_custom_dialog(page: ft.Page, title: str, content: ft.Control, actions: list, icon: str = None, icon_color: str = None, title_color: str = ft.Colors.WHITE):
        """Show a custom dialog with consistent styling"""
        DialogManager._cleanup_overlay(page)
        
        title_controls = []
        if icon:
            title_controls.append(ft.Icon(icon, color=icon_color, size=28))
        
        title_controls.append(ft.Text(title, weight=ft.FontWeight.BOLD, color=title_color, size=16))
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=title_controls,
                spacing=10,
                rtl=True
            ),
            content=content,
            actions=actions,
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
            shape=ft.RoundedRectangleBorder(radius=15),
        )
        
        page.overlay.append(dlg)
        dlg.open = True
        page.update()
        return dlg
