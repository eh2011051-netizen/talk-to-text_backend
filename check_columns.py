import sqlite3
import os

db_path = os.path.join(os.getcwd(), 'instance', 'database.db')
if not os.path.exists(db_path):
    db_path = 'database.db'

print(f"Checking database at: {db_path}")

if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(user)")
    columns = [row[1] for row in cursor.fetchall()]
    print(f"User table columns: {columns}")
    
    needed = ['status', 'last_seen', 'image', 'privacy_last_seen', 'privacy_profile_photo', 'privacy_about', 'privacy_read_receipts']
    missing = [c for c in needed if c not in columns]
    print(f"Missing columns: {missing}")
    conn.close()
else:
    print("Database not found")
