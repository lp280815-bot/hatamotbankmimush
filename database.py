# -*- coding: utf-8 -*-
"""
מודול מסד נתונים SQLite לניהול ספקים, מיילים והגדרות
"""

import sqlite3
import json
from contextlib import contextmanager
from typing import Dict, List, Optional

DB_FILE = "app_database.db"

@contextmanager
def get_db_connection():
    """Context manager for database connections"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def init_database():
    """Initialize database with required tables"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        
        # טבלת ספקים
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS suppliers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_num TEXT UNIQUE NOT NULL,
                account_name TEXT NOT NULL,
                email TEXT,
                phone TEXT,
                notes TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # טבלת מיפוי שמות לספקים (VLOOKUP)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS name_mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                search_term TEXT UNIQUE NOT NULL,
                supplier_account TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (supplier_account) REFERENCES suppliers(account_num)
            )
        """)
        
        # טבלת מיפוי סכומים לספקים (VLOOKUP)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS amount_mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                amount REAL UNIQUE NOT NULL,
                supplier_account TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (supplier_account) REFERENCES suppliers(account_num)
            )
        """)
        
        # טבלת הגדרות
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # טבלת לוג העברות שנשלחו
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS email_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                supplier_account TEXT NOT NULL,
                recipient_email TEXT NOT NULL,
                subject TEXT NOT NULL,
                body TEXT NOT NULL,
                status TEXT NOT NULL,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (supplier_account) REFERENCES suppliers(account_num)
            )
        """)

# -------------------- ספקים --------------------
def get_supplier(account_num: str) -> Optional[Dict]:
    """Get supplier by account number"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM suppliers WHERE account_num = ?", (account_num,))
        row = cursor.fetchone()
        return dict(row) if row else None

def save_supplier(account_num: str, account_name: str, email: str = None, phone: str = None, notes: str = None):
    """Save or update supplier"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO suppliers (account_num, account_name, email, phone, notes)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(account_num) DO UPDATE SET
                account_name = excluded.account_name,
                email = excluded.email,
                phone = excluded.phone,
                notes = excluded.notes,
                updated_at = CURRENT_TIMESTAMP
        """, (account_num, account_name, email, phone, notes))

def get_all_suppliers() -> List[Dict]:
    """Get all suppliers"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM suppliers ORDER BY account_name")
        return [dict(row) for row in cursor.fetchall()]

# -------------------- מיפוי שמות --------------------
def save_name_mapping(search_term: str, supplier_account: str):
    """Save name to supplier mapping"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO name_mappings (search_term, supplier_account)
            VALUES (?, ?)
            ON CONFLICT(search_term) DO UPDATE SET
                supplier_account = excluded.supplier_account
        """, (search_term, supplier_account))

def get_name_mappings() -> Dict[str, str]:
    """Get all name mappings as dict"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT search_term, supplier_account FROM name_mappings")
        return {row['search_term']: row['supplier_account'] for row in cursor.fetchall()}

# -------------------- מיפוי סכומים --------------------
def save_amount_mapping(amount: float, supplier_account: str):
    """Save amount to supplier mapping"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO amount_mappings (amount, supplier_account)
            VALUES (?, ?)
            ON CONFLICT(amount) DO UPDATE SET
                supplier_account = excluded.supplier_account
        """, (round(amount, 2), supplier_account))

def get_amount_mappings() -> Dict[float, str]:
    """Get all amount mappings as dict"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT amount, supplier_account FROM amount_mappings")
        return {row['amount']: row['supplier_account'] for row in cursor.fetchall()}

# -------------------- הגדרות --------------------
def save_setting(key: str, value: str):
    """Save application setting"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO settings (key, value)
            VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET
                value = excluded.value,
                updated_at = CURRENT_TIMESTAMP
        """, (key, value))

def get_setting(key: str, default: str = None) -> Optional[str]:
    """Get application setting"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cursor.fetchone()
        return row['value'] if row else default

# -------------------- לוג מיילים --------------------
def log_email(supplier_account: str, recipient_email: str, subject: str, body: str, status: str):
    """Log sent email"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO email_log (supplier_account, recipient_email, subject, body, status)
            VALUES (?, ?, ?, ?, ?)
        """, (supplier_account, recipient_email, subject, body, status))

def get_email_logs(limit: int = 100) -> List[Dict]:
    """Get email logs"""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT e.*, s.account_name 
            FROM email_log e
            LEFT JOIN suppliers s ON e.supplier_account = s.account_num
            ORDER BY e.sent_at DESC
            LIMIT ?
        """, (limit,))
        return [dict(row) for row in cursor.fetchall()]

# -------------------- מיגרציה מ-JSON --------------------
def migrate_from_json(json_file: str = "rules_store.json"):
    """Migrate data from JSON file to database"""
    import os
    if not os.path.exists(json_file):
        return
    
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Migrate name mappings
        for search_term, supplier_account in data.get('name_map', {}).items():
            save_name_mapping(search_term, supplier_account)
        
        # Migrate amount mappings
        for amount_str, supplier_account in data.get('amount_map', {}).items():
            try:
                amount = float(amount_str)
                save_amount_mapping(amount, supplier_account)
            except ValueError:
                pass
        
        print(f"Migration from {json_file} completed successfully")
    except Exception as e:
        print(f"Migration error: {e}")
