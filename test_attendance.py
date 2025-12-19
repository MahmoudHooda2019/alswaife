import sys
import os
# Add the project root to the path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

import flet as ft

# Now we can import the attendance view
from src.views.attendance_view import AttendanceView

def main(page: ft.Page):
    page.title = "Test Attendance View"
    page.rtl = True
    page.theme_mode = ft.ThemeMode.DARK
    
    # Create attendance view
    attendance_view = AttendanceView(page)
    attendance_view.build_ui()

if __name__ == "__main__":
    ft.app(target=main)