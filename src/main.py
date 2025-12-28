"""
مصنع السويفي - نظام الإدارة
تطبيق رسومي لإدارة المصنع والفواتير.
"""

import sys
import os
import flet as ft
from utils.path_utils import resource_path
from utils.log_utils import log_error, log_exception

# Version information
try:
    from version import __version__
except ImportError:
    __version__ = "1.0"

# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Set locale for Arabic support
import locale

try:
    # Try to set Arabic locale
    locale.setlocale(locale.LC_ALL, "ar_SA.UTF-8")
except:
    pass

from views.dashboard_view import DashboardView


def main(page: ft.Page):
    """نقطة الدخول الرئيسية للتطبيق."""
    try:
        # Configure window properties
        page.title = f"مصنع السويفي - الإصدار {__version__}"

        # Set window to maximized
        page.window.maximized = True

        # Set app icon
        icon_path = resource_path(os.path.join("assets", "icon.ico"))
        page.window.icon = icon_path

        # Create and show the Dashboard
        dashboard = DashboardView(page)
        dashboard.show()

    except Exception as e:
        # Log the error
        log_exception(f"Application error: {e}")

        # Show error dialog
        error_msg = f"حدث خطأ غير متوقع:\n{str(e)}"

        def close_dlg(e):
            dlg.open = False
            page.update()

        dlg = ft.AlertDialog(
            title=ft.Text("خطأ في التطبيق"),
            content=ft.Text(error_msg, rtl=True),
            actions=[ft.TextButton("موافق", on_click=close_dlg)],
        )
        page.overlay.append(dlg)
        dlg.open = True
        page.update()
        sys.exit(1)


if __name__ == "__main__":
    # Run as desktop app with assets [directory]
    ft.app(target=main, view=ft.AppView.FLET_APP, assets_dir="assets")