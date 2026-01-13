"""
Payments Utilities for Client Payment Management
This module provides functions to manage client payments and balances.
"""

import sqlite3
import os
from typing import List, Tuple, Optional
from datetime import datetime

from utils.log_utils import log_error, log_exception
from utils.db_utils import get_db_connection


def ensure_payments_table(db_path: str):
    """Ensure the payments table exists in the database."""
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    client_name TEXT NOT NULL,
                    payment_date TEXT NOT NULL,
                    amount REAL NOT NULL,
                    payment_type TEXT DEFAULT 'سداد',
                    invoice_number TEXT,
                    notes TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_payments_client 
                ON payments (client_name)
                """
            )
    except Exception as e:
        log_error(f"Error creating payments table: {e}")


def add_payment(db_path: str, client_name: str, payment_date: str, amount: float,
                payment_type: str = "سداد", invoice_number: str = "", notes: str = "") -> bool:
    """
    Add a new payment record for a client and record it in income file.
    
    Args:
        db_path: Path to the SQLite database
        client_name: Name of the client
        payment_date: Date of payment (DD/MM/YYYY format)
        amount: Payment amount (positive for payments, negative for debts)
        payment_type: Type of payment (سداد, دفعة مقدمة, فاتورة)
        invoice_number: Related invoice number (optional)
        notes: Additional notes
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        ensure_payments_table(db_path)
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO payments (client_name, payment_date, amount, payment_type, invoice_number, notes)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (client_name, payment_date, amount, payment_type, invoice_number, notes))
            
            # If this is a payment (negative amount), add to income file
            if amount < 0 and payment_type in ["سداد", "دفعة مقدمة"]:
                from utils.purchases_utils import add_payment_to_income_file
                
                # Add to income file (use positive amount for income)
                add_payment_to_income_file(
                    client_name=client_name,
                    payment_date=payment_date,
                    amount=abs(amount),  # Convert to positive for income
                    notes=f"{payment_type} - {notes}" if notes else payment_type
                )
            
            return True
    except Exception as e:
        log_error(f"Error adding payment: {e}")
        return False


def get_client_payments(db_path: str, client_name: str) -> List[dict]:
    """
    Get all payments for a specific client.
    
    Args:
        db_path: Path to the SQLite database
        client_name: Name of the client
        
    Returns:
        List of payment dictionaries
    """
    try:
        ensure_payments_table(db_path)
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, payment_date, amount, payment_type, invoice_number, notes, created_at
                FROM payments
                WHERE client_name = ?
                ORDER BY created_at DESC
            ''', (client_name,))
            
            rows = cursor.fetchall()
            payments = []
            for row in rows:
                payments.append({
                    'id': row[0],
                    'date': row[1],
                    'amount': row[2],
                    'type': row[3],
                    'invoice_number': row[4],
                    'notes': row[5],
                    'created_at': row[6]
                })
            return payments
    except Exception as e:
        log_error(f"Error getting client payments: {e}")
        return []


def get_client_balance(db_path: str, client_name: str) -> float:
    """
    Calculate the current balance for a client.
    Positive = client owes money (debt)
    Negative = client has credit (advance payment)
    
    Args:
        db_path: Path to the SQLite database
        client_name: Name of the client
        
    Returns:
        float: Current balance
    """
    try:
        ensure_payments_table(db_path)
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT COALESCE(SUM(amount), 0)
                FROM payments
                WHERE client_name = ?
            ''', (client_name,))
            
            result = cursor.fetchone()
            return result[0] if result else 0.0
    except Exception as e:
        log_error(f"Error calculating client balance: {e}")
        return 0.0


def delete_payment(db_path: str, payment_id: int) -> bool:
    """
    Delete a payment record and remove it from income file.
    
    Args:
        db_path: Path to the SQLite database
        payment_id: ID of the payment to delete
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # First get the payment details before deleting
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT client_name, payment_date, amount, payment_type
                FROM payments WHERE id = ?
            ''', (payment_id,))
            payment_data = cursor.fetchone()
            
            if not payment_data:
                return False
            
            client_name, payment_date, amount, payment_type = payment_data
            
            # Delete from database
            cursor.execute('DELETE FROM payments WHERE id = ?', (payment_id,))
            
            # If this was a payment (negative amount), remove from income file
            if amount < 0 and payment_type in ["سداد", "دفعة مقدمة"]:
                from utils.purchases_utils import remove_payment_from_income_file
                
                # Remove from income file
                remove_payment_from_income_file(
                    client_name=client_name,
                    payment_date=payment_date
                )
            
            return True
    except Exception as e:
        log_error(f"Error deleting payment: {e}")
        return False


def update_payment(db_path: str, payment_id: int, payment_date: str = None,
                   amount: float = None, payment_type: str = None, notes: str = None) -> bool:
    """
    Update an existing payment record.
    
    Args:
        db_path: Path to the SQLite database
        payment_id: ID of the payment to update
        payment_date: New date (optional)
        amount: New amount (optional)
        payment_type: New type (optional)
        notes: New notes (optional)
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        updates = []
        values = []
        
        if payment_date is not None:
            updates.append("payment_date = ?")
            values.append(payment_date)
        if amount is not None:
            updates.append("amount = ?")
            values.append(amount)
        if payment_type is not None:
            updates.append("payment_type = ?")
            values.append(payment_type)
        if notes is not None:
            updates.append("notes = ?")
            values.append(notes)
        
        if not updates:
            return True
        
        values.append(payment_id)
        
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(f'''
                UPDATE payments SET {", ".join(updates)}
                WHERE id = ?
            ''', values)
            return True
    except Exception as e:
        log_error(f"Error updating payment: {e}")
        return False


def get_all_clients_with_balance(db_path: str) -> List[dict]:
    """
    Get all clients that have payment records with their balances.
    
    Args:
        db_path: Path to the SQLite database
        
    Returns:
        List of client dictionaries with name and balance
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT client_name, SUM(amount) as balance
                FROM payments
                GROUP BY client_name
                ORDER BY client_name
            ''')
            
            rows = cursor.fetchall()
            clients = []
            for row in rows:
                clients.append({
                    'name': row[0],
                    'balance': row[1]
                })
            return clients
    except Exception as e:
        log_error(f"Error getting clients with balance: {e}")
        return []


def add_invoice_to_payments(db_path: str, client_name: str, invoice_number: str,
                            invoice_date: str, total_amount: float) -> bool:
    """
    Add an invoice as a debt entry in the payments table.
    
    Args:
        db_path: Path to the SQLite database
        client_name: Name of the client
        invoice_number: Invoice number
        invoice_date: Date of the invoice
        total_amount: Total amount of the invoice (will be added as positive/debt)
        
    Returns:
        bool: True if successful, False otherwise
    """
    return add_payment(
        db_path=db_path,
        client_name=client_name,
        payment_date=invoice_date,
        amount=total_amount,  # Positive = debt
        payment_type="فاتورة",
        invoice_number=invoice_number,
        notes=f"فاتورة رقم {invoice_number}"
    )


def remove_invoice_from_payments(db_path: str, client_name: str, invoice_number: str) -> bool:
    """
    Remove an invoice entry from the payments table.
    
    Args:
        db_path: Path to the SQLite database
        client_name: Name of the client
        invoice_number: Invoice number to remove
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                DELETE FROM payments 
                WHERE client_name = ? AND invoice_number = ? AND payment_type = 'فاتورة'
            ''', (client_name, invoice_number))
            return True
    except Exception as e:
        log_error(f"Error removing invoice from payments: {e}")
        return False


def export_client_statement(db_path: str, client_name: str, output_path: str) -> bool:
    """
    Export/Update client statement to Excel file with side-by-side tables for invoices and payments.
    This is the main function used to create/update the client ledger.
    
    Args:
        db_path: Path to the SQLite database
        client_name: Name of the client
        output_path: Path to save the Excel file
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        import xlsxwriter
        import os
        
        # Delete existing file if exists (to recreate with new data)
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except PermissionError:
                log_error(f"Cannot update ledger - file is open: {output_path}")
                return False
        
        payments = get_client_payments(db_path, client_name)
        
        # Separate invoices and payments
        invoices = [p for p in payments if p['type'] == 'فاتورة']
        actual_payments = [p for p in payments if p['type'] != 'فاتورة']
        
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet("كشف حساب")
        worksheet.right_to_left()
        
        # Formats with improved colors
        title_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 18, 'bg_color': '#1F4E78', 'font_color': '#FFFFFF'
        })
        
        header_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 12, 'bg_color': '#4472C4', 'font_color': '#FFFFFF'
        })
        
        section_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 14, 'bg_color': '#B4C7E7', 'font_color': '#1F4E78'
        })
        
        cell_fmt = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 11
        })
        
        money_fmt = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 11, 'num_format': '#,##0'
        })
        
        area_fmt = workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 11, 'num_format': '0.00'
        })
        
        # Improved balance format with better green color
        balance_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 14, 'bg_color': '#548235', 'font_color': '#FFFFFF'
        })
        
        total_fmt = workbook.add_format({
            'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
            'font_name': 'Arial', 'font_size': 12, 'bg_color': '#D9E2F3', 'font_color': '#1F4E78'
        })
        
        # ==================
        # Title Row (Row 0)
        # ==================
        worksheet.merge_range('A1:F1', f"كشف حساب العميل / {client_name}", title_fmt)
        worksheet.set_row(0, 35)
        
        # Balance display - spans two cells for better alignment
        worksheet.merge_range('G1:H1', "إجمالي المديونية:", balance_fmt)
        
        # ==================
        # Invoices Table (Columns A-F, starting row 1)
        # ==================
        worksheet.merge_range(1, 0, 1, 5, "جدول الفواتير", section_fmt)
        worksheet.set_row(1, 25)
        
        # Invoices headers (Row 2)
        inv_headers = ["رقم الفاتورة", "اسم السائق", "تاريخ التحميل", "النوع", "الكمية م²", "إجمالي السعر"]
        for col, header in enumerate(inv_headers):
            worksheet.write(2, col, header, header_fmt)
        worksheet.set_row(2, 25)
        
        # Invoices data (starting Row 3)
        invoices_start_row = 3
        inv_row = 3
        for inv in reversed(invoices):  # Oldest first
            inv_details = get_invoice_details_with_items(db_path, inv['invoice_number'])
            
            if inv_details and inv_details.get('items'):
                # Get invoice items for detailed breakdown
                items = inv_details['items']
                first_item = True
                
                for item in items:
                    # For first item, show all invoice details
                    if first_item:
                        worksheet.write(inv_row, 0, inv['invoice_number'] or "", cell_fmt)
                        worksheet.write(inv_row, 1, inv_details.get('driver', '') if inv_details else "", cell_fmt)
                        worksheet.write(inv_row, 2, inv['date'], cell_fmt)
                        first_item = False
                    else:
                        # For subsequent items, merge cells to show they belong to same invoice
                        worksheet.write(inv_row, 0, "", cell_fmt)
                        worksheet.write(inv_row, 1, "", cell_fmt)
                        worksheet.write(inv_row, 2, "", cell_fmt)
                    
                    # Item details
                    material = item.get('material', '')
                    thickness = item.get('thickness', '')
                    type_text = f"{material} - {thickness}" if material and thickness else (material or thickness or "")
                    
                    worksheet.write(inv_row, 3, type_text, cell_fmt)
                    worksheet.write(inv_row, 4, item.get('area', 0), area_fmt)
                    worksheet.write(inv_row, 5, item.get('total', 0), money_fmt)
                    inv_row += 1
                    
                # After all items, merge the invoice number, driver, and date cells
                if len(items) > 1:
                    start_row = inv_row - len(items)
                    end_row = inv_row - 1
                    
                    # Merge invoice number column
                    worksheet.merge_range(start_row, 0, end_row, 0, inv['invoice_number'] or "", cell_fmt)
                    # Merge driver column
                    worksheet.merge_range(start_row, 1, end_row, 1, inv_details.get('driver', ''), cell_fmt)
                    # Merge date column
                    worksheet.merge_range(start_row, 2, end_row, 2, inv['date'], cell_fmt)
            else:
                # Fallback for invoices without detailed items
                worksheet.write(inv_row, 0, inv['invoice_number'] or "", cell_fmt)
                worksheet.write(inv_row, 1, inv_details.get('driver', '') if inv_details else "", cell_fmt)
                worksheet.write(inv_row, 2, inv['date'], cell_fmt)
                worksheet.write(inv_row, 3, "", cell_fmt)
                worksheet.write(inv_row, 4, 0, area_fmt)
                worksheet.write(inv_row, 5, abs(inv['amount']), money_fmt)
                inv_row += 1
        
        # Invoices total row
        invoices_total_row = inv_row
        if invoices:
            worksheet.merge_range(inv_row, 0, inv_row, 4, "المجموع", total_fmt)
            worksheet.write_formula(inv_row, 5, f"=SUM(F{invoices_start_row+1}:F{inv_row})", total_fmt)
        else:
            worksheet.merge_range(inv_row, 0, inv_row, 4, "المجموع", total_fmt)
            worksheet.write(inv_row, 5, 0, total_fmt)
        
        # ==================
        # Payments Table (Columns G-I, starting row 1) - No gap column
        # ==================
        pay_col_start = 6  # Column G (removed gap)
        
        worksheet.merge_range(1, pay_col_start, 1, pay_col_start + 2, "جدول المدفوعات", section_fmt)
        
        # Payments headers (Row 2)
        pay_headers = ["تاريخ الدفعة", "المبلغ", "الملاحظات"]
        for i, header in enumerate(pay_headers):
            worksheet.write(2, pay_col_start + i, header, header_fmt)
        
        # Payments data (starting Row 3)
        payments_start_row = 3
        pay_row = 3
        for pay in reversed(actual_payments):  # Oldest first
            worksheet.write(pay_row, pay_col_start, pay['date'], cell_fmt)
            worksheet.write(pay_row, pay_col_start + 1, abs(pay['amount']), money_fmt)
            worksheet.write(pay_row, pay_col_start + 2, pay['notes'] or pay['type'], cell_fmt)
            pay_row += 1
        
        # Payments total row
        payments_total_row = pay_row
        if actual_payments:
            worksheet.write(pay_row, pay_col_start, "المجموع", total_fmt)
            worksheet.write_formula(pay_row, pay_col_start + 1, f"=SUM(H{payments_start_row+1}:H{pay_row})", total_fmt)
            worksheet.write(pay_row, pay_col_start + 2, "", total_fmt)
        else:
            worksheet.write(pay_row, pay_col_start, "المجموع", total_fmt)
            worksheet.write(pay_row, pay_col_start + 1, 0, total_fmt)
            worksheet.write(pay_row, pay_col_start + 2, "", total_fmt)
        
        # Balance Formula - now in column I (adjusted for no gap)
        balance_formula = f"=F{invoices_total_row+1}-H{payments_total_row+1}"
        worksheet.write_formula('I1', balance_formula, balance_fmt)
        
        # Column widths - adjusted for no gap
        worksheet.set_column(0, 0, 12)  # رقم الفاتورة
        worksheet.set_column(1, 1, 12)  # اسم السائق
        worksheet.set_column(2, 2, 14)  # تاريخ التحميل
        worksheet.set_column(3, 3, 18)  # النوع
        worksheet.set_column(4, 4, 12)  # الكمية م²
        worksheet.set_column(5, 5, 14)  # إجمالي السعر
        worksheet.set_column(6, 6, 14)  # تاريخ الدفعة
        worksheet.set_column(7, 7, 12)  # المبلغ
        worksheet.set_column(8, 8, 18)  # الملاحظات / المديونية
        
        workbook.close()
        return True
        
    except Exception as e:
        log_error(f"Error exporting client statement: {e}")
        return False


def update_client_statement(db_path: str, client_name: str, client_folder: str) -> bool:
    """
    Update the client's ledger file (كشف حساب.xlsx) with latest data.
    Called automatically after saving invoices or adding payments.
    
    Args:
        db_path: Path to the SQLite database
        client_name: Name of the client
        client_folder: Path to the client's folder
        
    Returns:
        bool: True if successful, False otherwise
    """
    import os
    output_path = os.path.join(client_folder, "كشف حساب.xlsx")
    return export_client_statement(db_path, client_name, output_path)


def get_invoice_details_with_items(db_path: str, invoice_number: str) -> Optional[dict]:
    """
    Get invoice details with individual items from the invoices table.
    
    Args:
        db_path: Path to the SQLite database
        invoice_number: Invoice number
        
    Returns:
        Dictionary with invoice details and items list or None
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            # Get invoice header
            cursor.execute('''
                SELECT driver_name, total_amount
                FROM invoices
                WHERE invoice_number = ?
            ''', (invoice_number,))
            inv_row = cursor.fetchone()
            
            if not inv_row:
                return None
            
            # Get individual items with their details
            cursor.execute('''
                SELECT material, thickness, count, length, height, price, area, (area * price) as total
                FROM invoice_items
                WHERE invoice_number = ?
                ORDER BY id
            ''', (invoice_number,))
            items_rows = cursor.fetchall()
            
            items = []
            for item_row in items_rows:
                items.append({
                    'material': item_row[0] or "",
                    'thickness': item_row[1] or "",
                    'count': item_row[2] or 0,
                    'length': item_row[3] or 0,
                    'height': item_row[4] or 0,
                    'price': item_row[5] or 0,
                    'area': item_row[6] or 0,
                    'total': item_row[7] or 0
                })
            
            return {
                'driver': inv_row[0] or "",
                'total_amount': inv_row[1] or 0,
                'items': items
            }
    except Exception as e:
        log_error(f"Error getting invoice details with items: {e}")
        return None


def get_invoice_details(db_path: str, invoice_number: str) -> Optional[dict]:
    """
    Get invoice details from the invoices table.
    
    Args:
        db_path: Path to the SQLite database
        invoice_number: Invoice number
        
    Returns:
        Dictionary with invoice details or None
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            # Get invoice header
            cursor.execute('''
                SELECT driver_name, total_amount
                FROM invoices
                WHERE invoice_number = ?
            ''', (invoice_number,))
            inv_row = cursor.fetchone()
            
            if not inv_row:
                return None
            
            # Get total area and type (material - thickness) from invoice items
            cursor.execute('''
                SELECT SUM(area), GROUP_CONCAT(DISTINCT (material || ' - ' || thickness))
                FROM invoice_items
                WHERE invoice_number = ?
            ''', (invoice_number,))
            items_row = cursor.fetchone()
            
            return {
                'driver': inv_row[0] or "",
                'total': inv_row[1] or 0,
                'area': items_row[0] if items_row and items_row[0] else 0,
                'type': items_row[1] if items_row and items_row[1] else ""
            }
    except Exception as e:
        log_error(f"Error getting invoice details: {e}")
        return None
