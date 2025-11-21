import sqlite3
from typing import Optional
from contextlib import contextmanager


@contextmanager
def get_db_connection(db_path: str):
    """Context manager for database connections."""
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
    """Create DB and counters table if not exists."""
    with get_db_connection(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS counters (
                key TEXT PRIMARY KEY,
                value INTEGER
            )
            """
        )
        # ensure default invoice counter exists
        cur.execute("SELECT value FROM counters WHERE key=?", ("invoice",))
        row = cur.fetchone()
        if row is None:
            cur.execute("INSERT INTO counters(key, value) VALUES(?, ?)", ("invoice", 1))


def get_counter(db_path: str, key: str = "invoice") -> int:
    with get_db_connection(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT value FROM counters WHERE key=?", (key,))
        row = cur.fetchone()
        return int(row[0]) if row else 1


def increment_counter(db_path: str, key: str = "invoice") -> int:
    with get_db_connection(db_path) as conn:
        cur = conn.cursor()
        # use transaction
        cur.execute("SELECT value FROM counters WHERE key=?", (key,))
        row = cur.fetchone()
        if row:
            new_value = int(row[0]) + 1
            cur.execute("UPDATE counters SET value=? WHERE key=?", (new_value, key))
        else:
            new_value = 2
            cur.execute("INSERT OR REPLACE INTO counters(key, value) VALUES(?, ?)", (key, new_value))
        return new_value