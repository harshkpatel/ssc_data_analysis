import sqlite3
from typing import List, Tuple

DB_PATH = 'emails.db'

def init_db(db_path: str = DB_PATH):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS emails (
            subject TEXT,
            content TEXT,
            received TEXT,
            stream TEXT,
            PRIMARY KEY (subject, content, received)
        )
    ''')
    conn.commit()
    conn.close()

def insert_email(subject: str, content: str, received: str, stream: str = None, db_path: str = DB_PATH):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    try:
        c.execute('''
            INSERT OR IGNORE INTO emails (subject, content, received, stream)
            VALUES (?, ?, ?, ?)
        ''', (subject, content, received, stream))
        conn.commit()
    finally:
        conn.close()

def insert_emails_bulk(emails: List[Tuple[str, str, str, str]], db_path: str = DB_PATH):
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    try:
        c.executemany('''
            INSERT OR IGNORE INTO emails (subject, content, received, stream)
            VALUES (?, ?, ?, ?)
        ''', emails)
        conn.commit()
    finally:
        conn.close()

def get_all_emails(db_path: str = DB_PATH) -> List[Tuple[str, str, str, str]]:
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute('SELECT subject, content, received, stream FROM emails')
    results = c.fetchall()
    conn.close()
    return results 