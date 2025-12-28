import sqlite3
from typing import Optional
from contextlib import contextmanager
import os

from utils.log_utils import log_error, log_exception


@contextmanager
def get_db_connection(db_path: str):
    """Context manager for database connections."""
    # Ensure directory exists
    db_dir = os.path.dirname(db_path)
    if db_dir and not os.path.exists(db_dir):
        os.makedirs(db_dir)
    
    conn = sqlite3.connect(db_path)
    try:
        yield conn
    except Exception as e:
        conn.rollback()
        raise e
    else:
        conn.commit()
    finally:
        conn.close()


def init_db(db_path: str):
    """Create DB and settings table if not exists."""
    # Ensure directory exists
    db_dir = os.path.dirname(db_path)
    if db_dir and not os.path.exists(db_dir):
        os.makedirs(db_dir)
        
    with get_db_connection(db_path) as conn:
        cur = conn.cursor()
        
        # Create settings table for general app settings
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """
        )
        
        # Create counters table
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS counters (
                key TEXT PRIMARY KEY,
                value INTEGER
            )
            """
        )
        
        # Ensure default invoice counter exists
        cur.execute("SELECT value FROM counters WHERE key=?", ("invoice",))
        row = cur.fetchone()
        if row is None:
            cur.execute("INSERT INTO counters(key, value) VALUES(?, ?)", ("invoice", 1))
        
        # Create invoices table to store invoice data
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT UNIQUE NOT NULL,
                client_name TEXT,
                driver_name TEXT,
                phone TEXT,
                date TEXT,
                file_path TEXT,
                total_amount REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        
        # Create invoice_items table to store individual items
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS invoice_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_number TEXT NOT NULL,
                description TEXT,
                block TEXT,
                thickness TEXT,
                material TEXT,
                count REAL,
                length REAL,
                height REAL,
                price REAL,
                area REAL,
                total REAL,
                length_before REAL,  -- Added for length calculation
                discount REAL,       -- Added for length calculation
                FOREIGN KEY (invoice_number) REFERENCES invoices (invoice_number) ON DELETE CASCADE
            )
            """
        )
        
        # Check if length_before and discount columns exist, if not, add them
        # This handles schema migration for existing databases
        try:
            cur.execute("ALTER TABLE invoice_items ADD COLUMN length_before REAL")
        except sqlite3.OperationalError:
            # Column already exists
            pass
        
        try:
            cur.execute("ALTER TABLE invoice_items ADD COLUMN discount REAL")
        except sqlite3.OperationalError:
            # Column already exists
            pass


def get_counter(db_path: str, key: str = "invoice") -> int:
    """
    Get the next available invoice number by finding the maximum existing invoice number.
    """
    try:
        with get_db_connection(db_path) as conn:
            cur = conn.cursor()
            
            # First, try to get the maximum invoice number from the invoices table
            cur.execute("SELECT MAX(CAST(invoice_number AS INTEGER)) FROM invoices WHERE invoice_number GLOB '[0-9]*'")
            row = cur.fetchone()
            max_invoice = int(row[0]) if row and row[0] else 0
            
            # Also check the counter table
            cur.execute("SELECT value FROM counters WHERE key=?", (key,))
            counter_row = cur.fetchone()
            counter_value = int(counter_row[0]) if counter_row else 1
            
            # Return the higher value + 1 (or just the higher value if it's from counter)
            # The next invoice number should be max(max_invoice + 1, counter_value)
            next_value = max(max_invoice + 1, counter_value)
            
            return next_value
    except Exception as e:
        return 1


def increment_counter(db_path: str, key: str = "invoice") -> int:
    """
    Increment the counter and return the new value.
    Takes into account the maximum existing invoice number.
    """
    try:
        with get_db_connection(db_path) as conn:
            cur = conn.cursor()
            
            # First, get the maximum invoice number from the invoices table
            cur.execute("SELECT MAX(CAST(invoice_number AS INTEGER)) FROM invoices WHERE invoice_number GLOB '[0-9]*'")
            row = cur.fetchone()
            max_invoice = int(row[0]) if row and row[0] else 0
            
            # Get current counter value
            cur.execute("SELECT value FROM counters WHERE key=?", (key,))
            counter_row = cur.fetchone()
            current_counter = int(counter_row[0]) if counter_row else 1
            
            # New value should be max of (max_invoice, current_counter) + 1
            new_value = max(max_invoice, current_counter) + 1
            
            # Update the counter in database
            if counter_row:
                cur.execute("UPDATE counters SET value=? WHERE key=?", (new_value, key))
            else:
                cur.execute("INSERT INTO counters(key, value) VALUES(?, ?)", (key, new_value))
            
            return new_value
    except Exception as e:
        log_error(f"Error incrementing counter: {e}")
        return 1


def get_zoom_level(db_path: str) -> float:
    """
    Get the saved zoom level from database.
    
    Args:
        db_path: Path to the SQLite database
        
    Returns:
        float: Saved zoom level (default: 1.0)
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute('SELECT value FROM settings WHERE key = ?', ('zoom_level',))
            result = cursor.fetchone()
            
            if result:
                return float(result[0])
            return 1.0  # Default zoom level
    except Exception as e:
        return 1.0


def set_zoom_level(db_path: str, zoom_level: float) -> None:
    """
    Save zoom level to database.
    
    Args:
        db_path: Path to the SQLite database
        zoom_level: Zoom level to save
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            # Insert or update zoom_level
            cursor.execute('''
                INSERT OR REPLACE INTO settings (key, value)
                VALUES (?, ?)
            ''', ('zoom_level', str(zoom_level)))
    except Exception as e:
        pass


def save_invoice_to_db(db_path: str, invoice_number: str, client_name: str, driver_name: str, 
                        phone: str, date: str, file_path: str, items: list, total_amount: float = 0):
    """
    Save invoice data to database.
    
    Args:
        db_path: Path to the SQLite database
        invoice_number: Invoice number
        client_name: Client name
        driver_name: Driver name
        phone: Phone number
        date: Date of invoice
        file_path: Path where the Excel file is saved
        items: List of invoice items
        total_amount: Total amount of the invoice
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            # Calculate total amount if not provided
            if total_amount == 0:
                for item in items:
                    try:
                        count = int(float(item[4])) if item[4] else 0  # Convert to int (count should be whole number)
                        length = float(item[5]) if item[5] else 0
                        height = float(item[6]) if item[6] else 0
                        price = float(item[7]) if item[7] else 0
                        area = count * length * height
                        total_amount += area * price
                    except (ValueError, IndexError):
                        continue
            
            # Insert or update invoice record
            cursor.execute('''
                INSERT OR REPLACE INTO invoices 
                (invoice_number, client_name, driver_name, phone, date, file_path, total_amount, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ''', (invoice_number, client_name, driver_name, phone, date, file_path, total_amount))
            
            # Delete existing items for this invoice
            cursor.execute('DELETE FROM invoice_items WHERE invoice_number = ?', (invoice_number,))
            
            # Insert new items
            for item in items:
                try:
                    description = item[0] if len(item) > 0 else ""
                    block = item[1] if len(item) > 1 else ""
                    thickness = item[2] if len(item) > 2 else ""
                    material = item[3] if len(item) > 3 else ""
                    count = int(float(item[4])) if item[4] else 0  # Convert to int, not float
                    length = float(item[5]) if item[5] else 0
                    height = float(item[6]) if item[6] else 0
                    price = float(item[7]) if item[7] else 0
                    # Extract length_before and discount if available (new format)
                    length_before = float(item[8]) if len(item) > 8 else 0  # New field
                    discount = float(item[9]) if len(item) > 9 else 0      # New field
                    area = count * length * height
                    total = area * price
                    
                    cursor.execute('''
                        INSERT INTO invoice_items 
                        (invoice_number, description, block, thickness, material, count, length, height, price, area, total, length_before, discount)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (invoice_number, description, block, thickness, material, count, length, height, price, area, total, length_before, discount))
                except (ValueError, IndexError):
                    continue
            
    except Exception as e:
        log_exception(f"Error saving invoice to database: {e}")
        raise e


def load_invoice_from_db(db_path: str, invoice_number: str) -> Optional[dict]:
    """
    Load invoice data from database.
    
    Args:
        db_path: Path to the SQLite database
        invoice_number: Invoice number to load
        
    Returns:
        Dictionary containing invoice data, or None if not found
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            # Get invoice details
            cursor.execute('''
                SELECT invoice_number, client_name, driver_name, phone, date, file_path, total_amount
                FROM invoices WHERE invoice_number = ?
            ''', (invoice_number,))
            invoice_row = cursor.fetchone()
            
            if not invoice_row:
                return None
            
            # Get invoice items
            cursor.execute('''
                SELECT description, block, thickness, material, count, length, height, price, length_before, discount
                FROM invoice_items WHERE invoice_number = ?
            ''', (invoice_number,))
            items = cursor.fetchall()
            
            return {
                "invoice_number": invoice_row[0],
                "client_name": invoice_row[1],
                "driver_name": invoice_row[2],
                "phone": invoice_row[3],
                "date": invoice_row[4],
                "file_path": invoice_row[5],
                "total_amount": invoice_row[6],
                "items": items
            }
            
    except Exception as e:
        log_exception(f"Error loading invoice from database: {e}")
        return None


def invoice_exists(db_path: str, invoice_number: str) -> bool:
    """
    Check if an invoice exists in the database.
    
    Args:
        db_path: Path to the SQLite database
        invoice_number: Invoice number to check
        
    Returns:
        True if invoice exists, False otherwise
    """
    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute('SELECT 1 FROM invoices WHERE invoice_number = ?', (invoice_number,))
            result = cursor.fetchone()
            
            return result is not None
            
    except Exception as e:
        log_exception(f"Error checking if invoice exists: {e}")
        return False