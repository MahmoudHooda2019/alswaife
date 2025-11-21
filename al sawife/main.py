#!/usr/bin/env python3
"""
Invoice Creator Application
A GUI application for creating invoices and exporting them to Excel format.
"""

import sys
import os
import traceback
from tkinter import messagebox
import customtkinter as ctk

# Add the project root to the Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from ui import InvoiceUI
from excel_utils import save_invoice


def save_callback(filepath, op_num, client, driver, date_str, phone, items):
    """
    Callback function to save invoice data to Excel.
    
    Args:
        filepath (str): Path to save the Excel file
        op_num (str): Operation/invoice number
        client (str): Client name
        driver (str): Driver name
        date_str (str): Date string
        phone (str): Phone number
        items (list): List of invoice items
    """
    save_invoice(filepath, op_num, client, driver, items, date_str=date_str, phone=phone)


def main():
    """Main application entry point."""
    try:
        app = InvoiceUI(save_callback)
        app.mainloop()
    except Exception as e:
        # Log the full traceback for debugging
        error_msg = f"An unexpected error occurred:\n{str(e)}\n\n"
        error_msg += "Details:\n" + traceback.format_exc()
        
        # Try to show a message box, but fall back to console if GUI fails
        try:
            root = ctk.CTk()
            root.withdraw()  # Hide the main window
            messagebox.showerror("Application Error", error_msg)
            root.destroy()
        except:
            print(error_msg)
        
        sys.exit(1)


if __name__ == "__main__":
    main()