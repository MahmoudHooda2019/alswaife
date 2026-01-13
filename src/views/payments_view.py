"""
Payments View - Client Payment Management Interface
"""

import flet as ft
import os
from datetime import datetime

from utils.utils import resource_path, get_current_date
from utils.log_utils import log_error, log_exception
from utils.dialog_utils import DialogManager
from utils.payments_utils import (
    add_payment, get_client_payments, get_client_balance,
    delete_payment, update_payment, get_all_clients_with_balance,
    export_client_statement, update_client_statement
)


class PaymentsView:
    def __init__(self, page: ft.Page, on_back=None):
        self.page = page
        self.on_back = on_back
        
        # Database path
        self.documents_path = os.path.join(os.path.expanduser("~"), "Documents", "alswaife")
        if not os.path.exists(self.documents_path):
            os.makedirs(self.documents_path)
        self.db_path = os.path.join(self.documents_path, 'invoice.db')
        
        # Invoices path for client list
        self.invoices_root = os.path.join(self.documents_path, 'الفواتير')
        
        # Current selected client
        self.selected_client = None
        self.client_balance = 0.0
        
        # UI Components
        self.client_dropdown = None
        self.balance_text = None
        self.payments_table = None
        self.main_container = None
        
        # Payment form fields
        self.date_field = None
        self.amount_field = None
        self.notes_field = None
        
    def build_ui(self):
        """Build the payments management UI with AppBar"""
        
        # Load clients
        clients = self.load_clients()
        
        # Balance text for AppBar
        self.balance_text = ft.Text(
            "",
            size=16,
            weight=ft.FontWeight.BOLD,
            color=ft.Colors.WHITE,
        )
        
        # Create AppBar with title, balance, and save action
        app_bar = ft.AppBar(
            leading=ft.IconButton(
                icon=ft.Icons.ARROW_BACK,
                on_click=self.go_back,
                tooltip="العودة"
            ),
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.PAYMENTS, size=24),
                    ft.Text(
                        "إدارة الدفعات",
                        size=20,
                        weight=ft.FontWeight.BOLD,
                    ),
                    ft.Container(width=20),  # Spacer
                    self.balance_text,
                ],
                spacing=10
            ),
            actions=[
                ft.IconButton(
                    icon=ft.Icons.SAVE,
                    on_click=self.save_payment,
                    tooltip="حفظ الدفعة",
                ),
            ],
            bgcolor=ft.Colors.GREY_900,
        )
        
        # Set the AppBar
        self.page.appbar = app_bar
        
        # Client selection dropdown
        self.client_dropdown = ft.Dropdown(
            label="اختر العميل",
            width=300,
            options=[ft.dropdown.Option(c) for c in clients],
            on_change=self.on_client_selected,
        )
        
        # Payment form fields (disabled until client is selected)
        self.date_field = ft.TextField(
            label="تاريخ الدفع",
            value=get_current_date('%d/%m/%Y'),
            width=200,
            disabled=True,
        )
        
        self.amount_field = ft.TextField(
            label="المبلغ",
            width=200,
            keyboard_type=ft.KeyboardType.NUMBER,
            input_filter=ft.InputFilter(regex_string=r"^[0-9]*\.?[0-9]*$"),
            disabled=True,
        )
        
        self.notes_field = ft.TextField(
            label="ملاحظات",
            width=300,
            multiline=True,
            max_lines=2,
            disabled=True,
        )
        

        
        # Payments table
        self.payments_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("التاريخ", weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("النوع", weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("المبلغ", weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("رقم الفاتورة", weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("ملاحظات", weight=ft.FontWeight.BOLD)),
                ft.DataColumn(ft.Text("إجراءات", weight=ft.FontWeight.BOLD)),
            ],
            rows=[],
            border=ft.border.all(1, ft.Colors.GREY_700),
            border_radius=10,
            heading_row_color=ft.Colors.BLUE_GREY_900,
            data_row_color={"": ft.Colors.GREY_900, "hovered": ft.Colors.GREY_800},
        )
        
        # Main content with simplified design
        main_content = ft.Column(
            controls=[
                # Client selection and form section
                ft.Container(
                    content=ft.Column(
                        controls=[
                            # Client dropdown
                            ft.Row(
                                controls=[self.client_dropdown],
                                alignment=ft.MainAxisAlignment.CENTER,
                            ),
                            
                            # Payment form fields
                            ft.Row(
                                controls=[
                                    self.date_field,
                                    self.amount_field,
                                    self.notes_field,
                                ],
                                alignment=ft.MainAxisAlignment.CENTER,
                                spacing=20,
                            ),
                        ],
                        spacing=20,
                    ),
                    padding=20,
                ),
                
                # Payments table section
                ft.Container(
                    content=ft.Column(
                        controls=[
                            ft.Text("سجل الدفعات", size=18, weight=ft.FontWeight.BOLD),
                            ft.Container(
                                content=self.payments_table,
                                border_radius=10,
                                padding=10,
                            ),
                        ],
                        scroll=ft.ScrollMode.AUTO,
                    ),
                    expand=True,
                    padding=20,
                ),
            ],
            expand=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
        
        self.page.clean()
        self.page.add(main_content)
        self.page.update()
    
    def load_clients(self) -> list:
        """Load client names from invoices directory, excluding revenue clients"""
        clients = []
        if os.path.exists(self.invoices_root):
            try:
                for item in os.listdir(self.invoices_root):
                    if os.path.isdir(os.path.join(self.invoices_root, item)):
                        # استثناء عملاء الإيرادات
                        if "ايراد" not in item and "إيراد" not in item:
                            clients.append(item)
            except OSError as e:
                log_error(f"Error loading clients: {e}")
        return sorted(clients)
    
    def on_client_selected(self, e):
        """Handle client selection"""
        self.selected_client = e.control.value
        
        # Enable payment fields when client is selected
        if self.selected_client:
            self.date_field.disabled = False
            self.amount_field.disabled = False
            self.notes_field.disabled = False
            
            # Update notes field based on client balance
            self.client_balance = get_client_balance(self.db_path, self.selected_client)
            if self.client_balance > 0:
                # Client owes money - default to "سداد"
                self.notes_field.value = "سداد"
            else:
                # Client has credit or zero balance - default to "مدفوع مقدم"
                self.notes_field.value = "مدفوع مقدم"
        else:
            self.date_field.disabled = True
            self.amount_field.disabled = True
            self.notes_field.disabled = True
            self.notes_field.value = ""
        
        self.refresh_payments()
        self.page.update()
    
    def save_payment(self, e):
        """Save payment from the form fields"""
        if not self.selected_client:
            DialogManager.show_warning_dialog(self.page, "يرجى اختيار العميل أولاً")
            return
        
        try:
            amount = float(self.amount_field.value or 0)
            if amount <= 0:
                DialogManager.show_warning_dialog(self.page, "يرجى إدخال مبلغ صحيح")
                return
            
            # Determine payment type based on notes
            notes = self.notes_field.value or ""
            if "مدفوع مقدم" in notes or "دفعة مقدمة" in notes:
                payment_type = "دفعة مقدمة"
            else:
                payment_type = "سداد"
            
            # For payments, amount should be negative (reduces debt)
            actual_amount = -amount
            
            success = add_payment(
                db_path=self.db_path,
                client_name=self.selected_client,
                payment_date=self.date_field.value,
                amount=actual_amount,
                payment_type=payment_type,
                notes=notes
            )
            
            if success:
                # Update client statement (كشف حساب.xlsx)
                client_folder = os.path.join(self.invoices_root, self.selected_client)
                update_client_statement(self.db_path, self.selected_client, client_folder)
                
                # Clear form fields
                self.amount_field.value = ""
                
                # Update notes based on new balance
                new_balance = get_client_balance(self.db_path, self.selected_client)
                if new_balance > 0:
                    self.notes_field.value = "سداد"
                else:
                    self.notes_field.value = "مدفوع مقدم"
                
                self.refresh_payments()
                DialogManager.show_success_dialog(
                    self.page, 
                    f"تم إضافة الدفعة بنجاح\n• تم تحديث كشف حساب العميل\n• تم إضافة الدفعة لسجل الإيرادات\nالمبلغ: {amount:,.0f} جنيه"
                )
            else:
                DialogManager.show_error_dialog(self.page, "فشل في إضافة الدفعة")
                
        except ValueError:
            DialogManager.show_warning_dialog(self.page, "يرجى إدخال مبلغ صحيح")
    
    def refresh_payments(self):
        """Refresh the payments table for selected client"""
        if not self.selected_client:
            # Reset balance display when no client selected
            self.balance_text.value = ""
            self.payments_table.rows.clear()
            self.page.update()
            return
        
        # Get balance
        self.client_balance = get_client_balance(self.db_path, self.selected_client)
        
        # Update balance display in AppBar
        if self.client_balance > 0:
            self.balance_text.value = f"الرصيد: {abs(self.client_balance):,.0f} جنيه (مدين)"
            self.balance_text.color = ft.Colors.RED_400
        elif self.client_balance < 0:
            self.balance_text.value = f"الرصيد: {abs(self.client_balance):,.0f} جنيه (دائن)"
            self.balance_text.color = ft.Colors.GREEN_400
        else:
            self.balance_text.value = "الرصيد: 0 جنيه"
            self.balance_text.color = ft.Colors.WHITE
        
        # Get payments
        payments = get_client_payments(self.db_path, self.selected_client)
        
        # Update table
        self.payments_table.rows.clear()
        
        for payment in payments:
            amount = payment['amount']
            amount_color = ft.Colors.RED_400 if amount > 0 else ft.Colors.GREEN_400
            amount_text = f"+{amount:,.0f}" if amount > 0 else f"{amount:,.0f}"
            
            row = ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(payment['date'])),
                    ft.DataCell(ft.Text(payment['type'])),
                    ft.DataCell(ft.Text(amount_text, color=amount_color, weight=ft.FontWeight.BOLD)),
                    ft.DataCell(ft.Text(payment['invoice_number'] or "-")),
                    ft.DataCell(ft.Text(payment['notes'] or "-", max_lines=1)),
                    ft.DataCell(
                        ft.IconButton(
                            icon=ft.Icons.DELETE,
                            icon_color=ft.Colors.RED_400,
                            tooltip="حذف",
                            data=payment['id'],
                            on_click=self.confirm_delete_payment,
                        )
                    ),
                ]
            )
            self.payments_table.rows.append(row)
        
        self.page.update()
    
    
    def confirm_delete_payment(self, e):
        """Confirm before deleting a payment"""
        payment_id = e.control.data
        
        def close_dlg(e):
            DialogManager.close_dialog(self.page, dlg)
        
        def do_delete(e):
            success = delete_payment(self.db_path, payment_id)
            
            if success:
                # Update client statement (كشف حساب.xlsx)
                client_folder = os.path.join(self.invoices_root, self.selected_client)
                update_client_statement(self.db_path, self.selected_client, client_folder)
                
                self.refresh_payments()
                DialogManager.show_success_dialog(
                    self.page, 
                    "تم حذف الدفعة بنجاح\n• تم تحديث كشف حساب العميل\n• تم حذف الدفعة من سجل الإيرادات"
                )
            else:
                DialogManager.show_error_dialog(self.page, "فشل في حذف الدفعة")
            
            close_dlg(e)
        
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                controls=[
                    ft.Icon(ft.Icons.WARNING, color=ft.Colors.ORANGE_400),
                    ft.Text("تأكيد الحذف", weight=ft.FontWeight.BOLD),
                ],
                spacing=10,
            ),
            content=ft.Text("هل أنت متأكد من حذف هذه الدفعة؟\n\nسيتم حذفها من:\n• كشف حساب العميل\n• سجل الإيرادات"),
            actions=[
                ft.ElevatedButton(
                    "حذف",
                    icon=ft.Icons.DELETE,
                    bgcolor=ft.Colors.RED_700,
                    color=ft.Colors.WHITE,
                    on_click=do_delete,
                ),
                ft.TextButton(
                    "إلغاء",
                    on_click=close_dlg,
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            bgcolor=ft.Colors.GREY_900,
        )
        
        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()
    
    def go_back(self, e):
        """Go back to dashboard"""
        if self.on_back:
            self.on_back()
