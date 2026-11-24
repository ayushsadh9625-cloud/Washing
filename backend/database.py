import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
DB_DIR = os.path.join(ROOT_DIR, "database")
DB_PATH = os.path.join(DB_DIR, "avkg_factory.db")

def get_db_connection():
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Parties Table (Har user ki apni parties)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS parties (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            name TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(username, name)
        )
    """)

    # Stock Entries Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS stock_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            party_id INTEGER NOT NULL,
            entry_date TEXT NOT NULL,
            lot_number TEXT NOT NULL,
            size TEXT NOT NULL,
            colour TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            remarks TEXT,
            batch_number INTEGER DEFAULT 1,
            is_exported INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (party_id) REFERENCES parties (id)
        )
    """)

    # Excel Batches Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS excel_batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            party_id INTEGER NOT NULL,
            batch_number INTEGER NOT NULL,
            file_path TEXT NOT NULL,
            total_entries INTEGER NOT NULL,
            total_pieces INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (party_id) REFERENCES parties (id)
        )
    """)

    conn.commit()
    conn.close()