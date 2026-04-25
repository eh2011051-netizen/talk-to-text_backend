
import sqlite3
import os

db_path = os.path.join(os.getcwd(), 'instance', 'database.db')
if not os.path.exists(db_path):
    # Try another path
    db_path = os.path.join(os.getcwd(), 'database.db')

print(f"Checking database at: {db_path}")

if os.path.exists(db_path):
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check current columns in friendship
        cursor.execute("PRAGMA table_info(friendship)")
        friendship_columns = [row[1] for row in cursor.fetchall()]
        print(f"Current columns in friendship: {friendship_columns}")
        
        friendship_cols_to_add = [
            ('is_pinned', 'BOOLEAN DEFAULT 0'),
            ('is_muted', 'BOOLEAN DEFAULT 0'),
            ('is_archived', 'BOOLEAN DEFAULT 0'),
            ('is_favourite', 'BOOLEAN DEFAULT 0'),
            ('is_deleted', 'BOOLEAN DEFAULT 0')
        ]
        
        for col_name, col_type in friendship_cols_to_add:
            if col_name not in friendship_columns:
                print(f"Adding column {col_name} to friendship...")
                cursor.execute(f"ALTER TABLE friendship ADD COLUMN {col_name} {col_type}")
        
        # Check current columns in user
        cursor.execute("PRAGMA table_info(user)")
        user_columns = [row[1] for row in cursor.fetchall()]
        print(f"Current columns in user: {user_columns}")
        
        user_cols_to_add = [
            ('status', "VARCHAR(20) DEFAULT 'offline'"),
            ('last_seen', "DATETIME DEFAULT CURRENT_TIMESTAMP"),
            ('image', "TEXT"),
            ('privacy_last_seen', "VARCHAR(20) DEFAULT 'everyone'"),
            ('privacy_profile_photo', "VARCHAR(20) DEFAULT 'everyone'"),
            ('privacy_about', "VARCHAR(20) DEFAULT 'everyone'"),
            ('privacy_read_receipts', "BOOLEAN DEFAULT 1")
        ]
        
        for col_name, col_type in user_cols_to_add:
            if col_name not in user_columns:
                print(f"Adding column {col_name} to user...")
                cursor.execute(f"ALTER TABLE user ADD COLUMN {col_name} {col_type}")
        
        # Check current columns in message
        cursor.execute("PRAGMA table_info(message)")
        message_columns = [row[1] for row in cursor.fetchall()]
        print(f"Current columns in message: {message_columns}")

        message_cols_to_add = [
            ('reaction', 'VARCHAR(32)'),
        ]

        for col_name, col_type in message_cols_to_add:
            if col_name not in message_columns:
                print(f"Adding column {col_name} to message...")
                cursor.execute(f"ALTER TABLE message ADD COLUMN {col_name} {col_type}")

        conn.commit()
        
        # Final Verification
        print("\n--- Final Verification ---")
        cursor.execute("PRAGMA table_info(user)")
        final_user_columns = [row[1] for row in cursor.fetchall()]
        
        needed = ['status', 'last_seen', 'image', 'privacy_last_seen', 'privacy_profile_photo', 'privacy_about', 'privacy_read_receipts']
        missing = [c for c in needed if c not in final_user_columns]
        
        if not missing:
            print("All required User table columns are present.")
        else:
            print(f"FAILED: Missing columns in User table: {missing}")
            
        conn.close()
        print("Database migration successful!")
    except Exception as e:
        print(f"Error during migration: {e}")
else:
    print("Database file not found.")
