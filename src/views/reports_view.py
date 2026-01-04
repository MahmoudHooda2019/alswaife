"""
Reports View - Enhanced reports interface with RecyclerView and CardView design
Allows users to select sections, sub-sections, and date ranges
"""

import flet as ft
import os
from datetime import datetime
from typing import Dict, List, Optional, Callable
from utils.reports_utils import execute_report


# Define all report sections with their sub-options
REPORT_SECTIONS_DATA = [
    {
        "id": "blocks",
        "name": "البلوكات المنشورة",
        "icon": ft.Icons.VIEW_IN_AR,
        "color": ft.Colors.AMBER_700,
        "sub_options": [
            {"id": "blocks_published", "name": "البلوكات المنشورة كاملة", "description": "البلوكات التي تحتوي على A و B"},
        ]
    },
    {
        "id": "clients",
        "name": "ديون العملاء",
        "icon": ft.Icons.PERSON,
        "color": ft.Colors.RED_700,
        "sub_options": [
            {"id": "clients_debts", "name": "العملاء المدينين", "description": "العملاء الذين عليهم ديون فقط"},
        ]
    },
    {
        "id": "machines",
        "name": "إنتاج الماكينات",
        "icon": ft.Icons.PRECISION_MANUFACTURING,
        "color": ft.Colors.CYAN_700,
        "sub_options": [
            {"id": "machine_production_1", "name": "ماكينة 1", "description": "إنتاج الماكينة رقم 1"},
            {"id": "machine_production_2", "name": "ماكينة 2", "description": "إنتاج الماكينة رقم 2"},
            {"id": "machine_production_3", "name": "ماكينة 3", "description": "إنتاج الماكينة رقم 3"},
        ]
    },
    {
        "id": "income_expenses",
        "name": "الإيرادات والمصروفات",
        "icon": ft.Icons.ACCOUNT_BALANCE_WALLET,
        "color": ft.Colors.GREEN_700,
        "sub_options": [
            {"id": "income", "name": "الإيرادات فقط", "description": "تقرير جميع الإيرادات"},
            {"id": "expenses", "name": "المصروفات فقط", "description": "تقرير جميع المصروفات"},
            {"id": "income_expenses_both", "name": "الإيرادات والمصروفات معاً", "description": "تقرير شامل"},
        ]
    },
    {
        "id": "inventory",
        "name": "مخزون الأدوات",
        "icon": ft.Icons.INVENTORY_2,
        "color": ft.Colors.INDIGO_700,
        "sub_options": [
            {"id": "inventory_consumption", "name": "استهلاك الأدوات", "description": "المصروفات من أذون الصرف لكل صنف"},
        ]
    },
]


class ReportsView:
    """Main reports view with RecyclerView-like design"""
    
    def __init__(self, page: ft.Page, on_back: Optional[Callable] = None):
        self.page = page
        self.on_back = on_back
        self.page.title = "مصنع السويفي - التقارير"
        self.page.rtl = True
        self.page.theme_mode = ft.ThemeMode.DARK
        
        # State management
        self.selected_reports: Dict[str, List[str]] = {}  # section_id -> [sub_option_ids]
        self.expanded_sections: List[str] = []
        
        # UI components
        self.date_from_field: Optional[ft.TextField] = None
        self.date_to_field: Optional[ft.TextField] = None
        self.sections_container: Optional[ft.Column] = None
        self.selected_summary: Optional[ft.Text] = None
        self.section_cards: Dict[str, dict] = {}  # Store card references
        
    def build_ui(self):
        """Build the reports UI"""
        # AppBar
        app_bar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK,
                on_click=self.go_back,
                tooltip="العودة",
            ),
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.ASSESSMENT, size=24, color=ft.Colors.TEAL_200),
                    ft.Text(
                        "التقارير",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                        color=ft.Colors.TEAL_200,
                    ),
                ],
                spacing=10,
            ),
            actions=[
                ft.IconButton(
                    icon=ft.Icons.SELECT_ALL,
                    on_click=self._select_all,
                    tooltip="تحديد الكل",
                ),
                ft.IconButton(
                    icon=ft.Icons.DESELECT,
                    on_click=self._deselect_all,
                    tooltip="إلغاء التحديد",
                ),
            ],
            bgcolor=ft.Colors.GREY_900,
        )
        
        # Date range section
        self.date_from_field = ft.TextField(
            label="من تاريخ",
            hint_text="DD/MM/YYYY",
            width=200,
            border_radius=10,
            filled=True,
            bgcolor=ft.Colors.GREY_800,
            border_color=ft.Colors.TEAL_700,
            focused_border_color=ft.Colors.TEAL_400,
            prefix_icon=ft.Icons.CALENDAR_TODAY,
            text_style=ft.TextStyle(size=15),
        )
        
        self.date_to_field = ft.TextField(
            label="إلى تاريخ",
            hint_text="DD/MM/YYYY",
            value=datetime.now().strftime("%d/%m/%Y"),
            width=200,
            border_radius=10,
            filled=True,
            bgcolor=ft.Colors.GREY_800,
            border_color=ft.Colors.TEAL_700,
            focused_border_color=ft.Colors.TEAL_400,
            prefix_icon=ft.Icons.CALENDAR_TODAY,
            text_style=ft.TextStyle(size=15),
        )
        
        date_section = ft.Container(
            content=ft.Column(
                controls=[
                    ft.Row(
                        controls=[
                            ft.Icon(ft.Icons.DATE_RANGE, color=ft.Colors.TEAL_400, size=20),
                            ft.Text("نطاق التاريخ", size=16, weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_300),
                        ],
                        spacing=10,
                    ),
                    ft.Row(
                        controls=[
                            self.date_from_field,
                            ft.Text("←", size=20, color=ft.Colors.GREY_500),
                            self.date_to_field,
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=15,
                    ),
                ],
                spacing=15,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=20,
            bgcolor=ft.Colors.GREY_900,
            border_radius=15,
            border=ft.border.all(1, ft.Colors.GREY_700),
            margin=ft.margin.only(bottom=15, left=10, right=10),
        )
        
        # Sections header
        sections_header = ft.Container(
            content=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CATEGORY, color=ft.Colors.TEAL_400, size=20),
                    ft.Text("اختر الأقسام", size=16, weight=ft.FontWeight.BOLD, color=ft.Colors.TEAL_300),
                ],
                spacing=10,
            ),
            padding=ft.padding.only(right=15, bottom=10),
        )
        
        # Build section cards
        self.sections_container = ft.Column(
            controls=self._build_section_cards(),
            spacing=8,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )
        
        # Selected summary
        self.selected_summary = ft.Text(
            "لم يتم تحديد أي تقارير",
            size=14,
            color=ft.Colors.GREY_500,
            text_align=ft.TextAlign.CENTER,
        )
        
        # Generate button
        generate_btn = ft.ElevatedButton(
            text="إنشاء التقارير",
            icon=ft.Icons.PLAY_ARROW,
            bgcolor=ft.Colors.TEAL_700,
            color=ft.Colors.WHITE,
            width=250,
            height=50,
            on_click=self._generate_reports,
            style=ft.ButtonStyle(
                shape=ft.RoundedRectangleBorder(radius=12),
            ),
        )
        
        # Bottom section
        bottom_section = ft.Container(
            content=ft.Column(
                controls=[
                    self.selected_summary,
                    ft.Container(height=10),
                    generate_btn,
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=5,
            ),
            padding=20,
            bgcolor=ft.Colors.GREY_900,
            border_radius=ft.border_radius.only(top_left=20, top_right=20),
            border=ft.border.only(top=ft.BorderSide(1, ft.Colors.GREY_700)),
        )
        
        # Main layout
        self.page.appbar = app_bar
        
        main_content = ft.Column(
            controls=[
                date_section,
                sections_header,
                self.sections_container,
            ],
            spacing=10,
            expand=True,
        )
        
        self.page.add(
            ft.Container(
                content=ft.Column(
                    controls=[
                        main_content,
                        bottom_section,
                    ],
                    spacing=0,
                ),
                expand=True,
            )
        )
        
        self.page.update()
    
    def _build_section_cards(self) -> List[ft.Control]:
        """Build all section cards"""
        cards = []
        for section_data in REPORT_SECTIONS_DATA:
            card = self._create_section_card(section_data)
            cards.append(card)
        return cards
    
    def _create_section_card(self, section_data: dict) -> ft.Card:
        """Create a single section card with expandable sub-options"""
        section_id = section_data["id"]
        
        # Initialize selection state
        if section_id not in self.selected_reports:
            self.selected_reports[section_id] = []
        
        # Create sub-options checkboxes
        sub_options_controls = []
        for opt in section_data["sub_options"]:
            cb = ft.Checkbox(
                label=opt["name"],
                value=opt["id"] in self.selected_reports.get(section_id, []),
                on_change=lambda e, sid=section_id, oid=opt["id"]: self._on_sub_option_change(e, sid, oid),
            )
            sub_options_controls.append(
                ft.Container(
                    content=ft.Column(
                        controls=[
                            cb,
                            ft.Text(opt["description"], size=11, color=ft.Colors.GREY_500),
                        ],
                        spacing=2,
                    ),
                    padding=ft.padding.only(right=20, bottom=8),
                )
            )
        
        sub_options_column = ft.Column(
            controls=sub_options_controls,
            spacing=5,
        )
        
        sub_options_container = ft.Container(
            content=sub_options_column,
            padding=ft.padding.all(15),
            bgcolor=ft.Colors.GREY_900,
            border_radius=ft.border_radius.only(bottom_left=12, bottom_right=12),
            visible=section_id in self.expanded_sections,
        )
        
        # Main checkbox for section
        main_checkbox = ft.Checkbox(
            value=len(self.selected_reports.get(section_id, [])) > 0,
            on_change=lambda e, sid=section_id: self._on_section_checkbox_change(e, sid),
            scale=1.2,
        )
        
        # Expand button
        expand_btn = ft.IconButton(
            icon=ft.Icons.EXPAND_LESS if section_id in self.expanded_sections else ft.Icons.EXPAND_MORE,
            icon_color=ft.Colors.WHITE,
            on_click=lambda e, sid=section_id: self._toggle_section(sid),
            tooltip="عرض/إخفاء الخيارات",
        )
        
        # Store references for later updates
        self.section_cards[section_id] = {
            "main_checkbox": main_checkbox,
            "expand_btn": expand_btn,
            "sub_options_container": sub_options_container,
            "sub_options_column": sub_options_column,
            "data": section_data,
        }
        
        return ft.Card(
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        # Header
                        ft.Container(
                            content=ft.Row(
                                controls=[
                                    main_checkbox,
                                    ft.Icon(section_data["icon"], size=28, color=ft.Colors.WHITE),
                                    ft.Text(
                                        section_data["name"],
                                        size=18,
                                        weight=ft.FontWeight.BOLD,
                                        color=ft.Colors.WHITE,
                                        expand=True,
                                    ),
                                    expand_btn,
                                ],
                                alignment=ft.MainAxisAlignment.START,
                                spacing=12,
                            ),
                            padding=ft.padding.symmetric(horizontal=15, vertical=12),
                            bgcolor=section_data["color"],
                            border_radius=ft.border_radius.only(
                                top_left=12, top_right=12,
                                bottom_left=0 if section_id in self.expanded_sections else 12,
                                bottom_right=0 if section_id in self.expanded_sections else 12,
                            ),
                            on_click=lambda e, sid=section_id: self._toggle_section(sid),
                        ),
                        # Sub-options
                        sub_options_container,
                    ],
                    spacing=0,
                ),
                border_radius=12,
            ),
            elevation=6,
            margin=ft.margin.symmetric(vertical=4, horizontal=10),
        )

    def _toggle_section(self, section_id: str):
        """Toggle section expansion"""
        if section_id in self.expanded_sections:
            self.expanded_sections.remove(section_id)
        else:
            self.expanded_sections.append(section_id)
        
        # Update UI
        if section_id in self.section_cards:
            card_data = self.section_cards[section_id]
            is_expanded = section_id in self.expanded_sections
            card_data["sub_options_container"].visible = is_expanded
            card_data["expand_btn"].icon = ft.Icons.EXPAND_LESS if is_expanded else ft.Icons.EXPAND_MORE
        
        self.page.update()
    
    def _on_section_checkbox_change(self, e, section_id: str):
        """Handle main section checkbox change"""
        if e.control.value:
            # Select all sub-options
            section_data = self.section_cards[section_id]["data"]
            self.selected_reports[section_id] = [opt["id"] for opt in section_data["sub_options"]]
            # Expand section
            if section_id not in self.expanded_sections:
                self._toggle_section(section_id)
        else:
            # Deselect all sub-options
            self.selected_reports[section_id] = []
        
        # Update sub-option checkboxes
        self._update_sub_options_ui(section_id)
        self._update_summary()
        self.page.update()
    
    def _on_sub_option_change(self, e, section_id: str, option_id: str):
        """Handle sub-option checkbox change"""
        if e.control.value:
            if option_id not in self.selected_reports[section_id]:
                self.selected_reports[section_id].append(option_id)
        else:
            if option_id in self.selected_reports[section_id]:
                self.selected_reports[section_id].remove(option_id)
        
        # Update main checkbox
        if section_id in self.section_cards:
            main_cb = self.section_cards[section_id]["main_checkbox"]
            main_cb.value = len(self.selected_reports[section_id]) > 0
        
        self._update_summary()
        self.page.update()
    
    def _update_sub_options_ui(self, section_id: str):
        """Update sub-options checkboxes UI"""
        if section_id not in self.section_cards:
            return
        
        card_data = self.section_cards[section_id]
        sub_column = card_data["sub_options_column"]
        selected = self.selected_reports.get(section_id, [])
        
        for container in sub_column.controls:
            if isinstance(container, ft.Container):
                col = container.content
                if isinstance(col, ft.Column) and col.controls:
                    cb = col.controls[0]
                    if isinstance(cb, ft.Checkbox):
                        # Find the option id from the label
                        for opt in card_data["data"]["sub_options"]:
                            if opt["name"] == cb.label:
                                cb.value = opt["id"] in selected
                                break
    
    def _update_summary(self):
        """Update the selected reports summary"""
        total_selected = sum(len(opts) for opts in self.selected_reports.values())
        
        if total_selected == 0:
            self.selected_summary.value = "لم يتم تحديد أي تقارير"
            self.selected_summary.color = ft.Colors.GREY_500
        else:
            self.selected_summary.value = f"تم تحديد {total_selected} تقرير"
            self.selected_summary.color = ft.Colors.TEAL_300
    
    def _select_all(self, e):
        """Select all sections and sub-options"""
        for section_data in REPORT_SECTIONS_DATA:
            section_id = section_data["id"]
            self.selected_reports[section_id] = [opt["id"] for opt in section_data["sub_options"]]
            if section_id in self.section_cards:
                self.section_cards[section_id]["main_checkbox"].value = True
                self._update_sub_options_ui(section_id)
        
        self._update_summary()
        self.page.update()
    
    def _deselect_all(self, e):
        """Deselect all sections"""
        for section_id in self.selected_reports:
            self.selected_reports[section_id] = []
            if section_id in self.section_cards:
                self.section_cards[section_id]["main_checkbox"].value = False
                self._update_sub_options_ui(section_id)
        
        self._update_summary()
        self.page.update()
    
    def go_back(self, e):
        """Navigate back"""
        if self.on_back:
            self.on_back()
    
    def _generate_reports(self, e):
        """Generate selected reports"""
        # Collect selected reports
        selected_reports = []
        for section_id, options in self.selected_reports.items():
            selected_reports.extend(options)
        
        if not selected_reports:
            self._show_dialog(
                "تنبيه",
                "الرجاء تحديد تقرير واحد على الأقل",
                ft.Colors.ORANGE_400
            )
            return
        
        # Show progress dialog
        self._show_progress_dialog(selected_reports)
    
    def _show_progress_dialog(self, selected_reports: List[str]):
        """Show progress dialog while generating reports"""
        progress_text = ft.Text("جاري إنشاء التقارير...", size=16, color=ft.Colors.WHITE)
        progress_bar = ft.ProgressBar(width=300, color=ft.Colors.TEAL_400, bgcolor=ft.Colors.GREY_700)
        current_report = ft.Text("", size=14, color=ft.Colors.GREY_400)
        
        progress_dlg = ft.AlertDialog(
            modal=True,
            content=ft.Container(
                content=ft.Column(
                    controls=[
                        ft.ProgressRing(width=50, height=50, color=ft.Colors.TEAL_400),
                        ft.Container(height=20),
                        progress_text,
                        ft.Container(height=10),
                        progress_bar,
                        current_report,
                    ],
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=10,
                ),
                padding=30,
                width=350,
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
                generated_files = []
                total = len(selected_reports)
                
                for i, report_id in enumerate(selected_reports):
                    # Update progress
                    progress = (i + 1) / total
                    progress_bar.value = progress
                    current_report.value = f"({i + 1}/{total}) جاري إنشاء: {self._get_report_name(report_id)}"
                    self.page.update()
                    
                    # Build query based on report type
                    query = self._build_query(report_id)
                    
                    # Execute report
                    result = execute_report(query, documents_path)
                    if result:
                        generated_files.append(result)
                
                progress_dlg.open = False
                self.page.update()
                
                if generated_files:
                    self._show_success_dialog(generated_files)
                else:
                    self._show_dialog(
                        "تنبيه",
                        "لا توجد بيانات متاحة للتقارير المحددة",
                        ft.Colors.ORANGE_400
                    )
                    
            except Exception as ex:
                progress_dlg.open = False
                self.page.update()
                self._show_dialog(
                    "خطأ",
                    f"فشل في إنشاء التقارير: {str(ex)}",
                    ft.Colors.RED_400
                )
        
        thread = threading.Thread(target=generate)
        thread.start()
    
    def _build_query(self, report_id: str) -> Dict:
        """Build query dictionary for a report"""
        query = {
            "report_type": report_id,
            "date_from": self.date_from_field.value if self.date_from_field.value else None,
            "date_to": self.date_to_field.value if self.date_to_field.value else None,
        }
        
        # Handle machine production reports
        if report_id.startswith("machine_production_"):
            machine_num = report_id.split("_")[-1]
            query["report_type"] = "machine_production"
            query["machine_number"] = machine_num
        
        return query
    
    def _get_report_name(self, report_id: str) -> str:
        """Get display name for a report ID"""
        for section_data in REPORT_SECTIONS_DATA:
            for opt in section_data["sub_options"]:
                if opt["id"] == report_id:
                    return opt["name"]
        return report_id
    
    def _show_dialog(self, title: str, message: str, title_color=ft.Colors.BLUE_300):
        """Show a styled dialog"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()
        
        dlg = ft.AlertDialog(
            title=ft.Text(title, color=title_color, weight=ft.FontWeight.BOLD),
            content=ft.Text(message, size=16, rtl=True),
            actions=[
                ft.TextButton(
                    "حسناً",
                    on_click=close_dlg,
                    style=ft.ButtonStyle(color=ft.Colors.BLUE_300)
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
        )
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()
    
    def _show_success_dialog(self, filepaths: List[str]):
        """Show success dialog with generated files"""
        def close_dlg(e=None):
            dlg.open = False
            self.page.update()
        
        def open_folder(e=None):
            close_dlg()
            try:
                if filepaths:
                    folder = os.path.dirname(filepaths[0])
                    os.startfile(folder)
            except:
                pass
        
        def open_first_file(e=None):
            close_dlg()
            try:
                if filepaths:
                    os.startfile(filepaths[0])
            except:
                pass
        
        # Build file list
        file_list = ft.Column(
            controls=[
                ft.Container(
                    content=ft.Row(
                        controls=[
                            ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=16),
                            ft.Text(
                                os.path.basename(fp),
                                size=12,
                                color=ft.Colors.GREY_300,
                            ),
                        ],
                        spacing=8,
                    ),
                    padding=5,
                )
                for fp in filepaths[:5]
            ],
            spacing=5,
        )
        
        if len(filepaths) > 5:
            file_list.controls.append(
                ft.Text(f"... و {len(filepaths) - 5} ملفات أخرى", size=12, color=ft.Colors.GREY_500)
            )
        
        dlg = ft.AlertDialog(
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_400, size=28),
                    ft.Text(
                        f"تم إنشاء {len(filepaths)} تقرير بنجاح",
                        color=ft.Colors.GREEN_300,
                        weight=ft.FontWeight.BOLD,
                        size=16,
                    ),
                ],
                spacing=10,
            ),
            content=ft.Container(
                content=file_list,
                padding=10,
                height=200,
            ),
            actions=[
                ft.TextButton(
                    "فتح الملف الأول",
                    on_click=open_first_file,
                    icon=ft.Icons.FILE_OPEN,
                    style=ft.ButtonStyle(color=ft.Colors.GREEN_300)
                ),
                ft.TextButton(
                    "فتح المجلد",
                    on_click=open_folder,
                    icon=ft.Icons.FOLDER_OPEN,
                    style=ft.ButtonStyle(color=ft.Colors.BLUE_300)
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
