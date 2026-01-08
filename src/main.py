"""
مصنع السويفي - نظام الإدارة
تطبيق رسومي لإدارة المصنع والفواتير.
"""

import sys
import os
import flet as ft
from utils.utils import resource_path
from utils.log_utils import log_exception

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

        # Lazy import to catch errors during view loading
        from views.dashboard_view import DashboardView

        # Create and show the Dashboard
        dashboard = DashboardView(page)
        dashboard.show()

    except Exception as e:
        # Log the error
        log_exception(f"Application error: {e}")

        # Show error dialog
        error_msg = f"حدث خطأ غير متوقع:\n{str(e)}"
        
        page.dialog = ft.AlertDialog(
            title=ft.Text("خطأ في التطبيق"),
            content=ft.Text(error_msg, rtl=True),
            actions=[ft.TextButton("موافق", on_click=lambda _: page.window_close())],
        )
        page.dialog.open = True
        page.update()


if __name__ == "__main__":
    try:
        # Run as desktop app with assets directory
        ft.app(target=main, view=ft.AppView.FLET_APP, assets_dir="assets")
    except Exception as e:
        # Fallback error handling for crashes before UI starts
        try:
            log_exception(f"Fatal startup error: {e}")
        except:
            pass
            
        # Write to crash file on desktop for visibility
        try:
            import traceback
            from datetime import datetime
            desktop = os.path.join(os.path.expanduser("~"), "Documents")
            crash_file = os.path.join(desktop, "AL_SWAIFE_CRASH.txt")
            with open(crash_file, "w", encoding="utf-8") as f:
                f.write(f"Crash Time: {datetime.now()}\n")
                f.write("-" * 50 + "\n")
                f.write(f"Error: {str(e)}\n")
                f.write("-" * 50 + "\n")
                f.write(traceback.format_exc())
        except:
            pass
        
        # Re-raise to show in console if available
        raise