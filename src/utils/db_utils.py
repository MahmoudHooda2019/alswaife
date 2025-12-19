import sqlite3
from typing import Optional
from contextlib import contextmanager
import os


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


def get_counter(db_path: str, key: str = "invoice") -> int:
    try:
        with get_db_connection(db_path) as conn:
            cur = conn.cursor()
            cur.execute("SELECT value FROM counters WHERE key=?", (key,))
            row = cur.fetchone()
            return int(row[0]) if row else 1
    except Exception as e:
        return 1


def increment_counter(db_path: str, key: str = "invoice") -> int:
    try:
        with get_db_connection(db_path) as conn:
            cur = conn.cursor()
            
            cur.execute("SELECT value FROM counters WHERE key=?", (key,))
            row = cur.fetchone()
            if row:
                new_value = int(row[0]) + 1
                cur.execute("UPDATE counters SET value=? WHERE key=?", (new_value, key))
            else:
                new_value = 2
                cur.execute("INSERT OR REPLACE INTO counters(key, value) VALUES(?, ?)", (key, new_value))
            
            return new_value
    except Exception as e:
        # Fallback to in-memory counter
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