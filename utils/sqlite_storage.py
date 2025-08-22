import sqlite3
from typing import List, Tuple

DB_PATH = 'emails.db'

def init_db(db_path: str = DB_PATH):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    
    # First create the table if it doesn't exist
    c.execute('''
        CREATE TABLE IF NOT EXISTS emails (
            subject TEXT,
            content TEXT,
            received TEXT,
            stream TEXT,
            person_name TEXT DEFAULT "",
            PRIMARY KEY (subject, content, received)
        )
    ''')
    
    # Now check if person_name column exists, if not add it
    c.execute("PRAGMA table_info(emails)")
    columns = [column[1] for column in c.fetchall()]
    
    if 'person_name' not in columns:
        c.execute('ALTER TABLE emails ADD COLUMN person_name TEXT DEFAULT ""')
    
    conn.commit()
    conn.close()

def insert_email(subject: str, content: str, received: str, stream: str = None, person_name: str = "", db_path: str = DB_PATH):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    try:
        c.execute('''
            INSERT OR IGNORE INTO emails (subject, content, received, stream, person_name)
            VALUES (?, ?, ?, ?, ?)
        ''', (subject, content, received, stream, person_name))
        conn.commit()
    finally:
        conn.close()

def insert_emails_bulk(emails: List[Tuple[str, str, str, str, str]], db_path: str = DB_PATH):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    try:
        c.executemany('''
            INSERT OR IGNORE INTO emails (subject, content, received, stream, person_name)
            VALUES (?, ?, ?, ?, ?)
        ''', emails)
        conn.commit()
    finally:
        conn.close()

def get_all_emails(db_path: str = DB_PATH) -> List[Tuple[str, str, str, str, str]]:
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('SELECT subject, content, received, stream, person_name FROM emails')
    results = c.fetchall()
    conn.close()
    return results 