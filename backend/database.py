import sqlite3
import os

# Database file ka path set karna
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_DIR = os.path.join(BASE_DIR, "database")
os.makedirs(DB_DIR, exist_ok=True)
DB_PATH = os.path.join(DB_DIR, "avkg_factory.db")

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    # Row factory lagane se data dictionary / column name ke roop mein milta hai
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Parties Table - Saari parties ki list store karne ke liye
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS parties (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # 2. Stock Entries Table - Har party ki stock entry (Size, Colour, Lot, Qty)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS stock_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
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
        FOREIGN KEY (party_id) REFERENCES parties (id) ON DELETE CASCADE
    )
    """)

    # 3. Excel Batches Table - Jo 10-entry batches export honge unka record
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS excel_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        party_id INTEGER NOT NULL,
        batch_number INTEGER NOT NULL,
        file_path TEXT NOT NULL,
        total_entries INTEGER DEFAULT 10,
        total_pieces INTEGER NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (party_id) REFERENCES parties (id) ON DELETE CASCADE
    )
    """)

    conn.commit()
    conn.close()
    print("Database & Tables successfully initialize ho gayi hain!")