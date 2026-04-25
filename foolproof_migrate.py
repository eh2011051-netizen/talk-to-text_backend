
import sqlite3
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def migrate():
    # Attempt to find the database
    possible_paths = [
        os.path.join(os.getcwd(), 'instance', 'database.db'),
        os.path.join(os.getcwd(), 'database.db'),
        'database.db',
        '/app/instance/database.db',
        '/app/database.db'
    ]
    
    db_path = None
    for path in possible_paths:
        if os.path.exists(path):
            db_path = path
            break
            
    if not db_path:
        logger.error("Database file not found in any expected location.")
        return

    logger.info(f"Using database at: {db_path}")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    def add_column_if_missing(table, column, type_def):
        cursor.execute(f"PRAGMA table_info({table})")
        columns = [row[1] for row in cursor.fetchall()]
        if column not in columns:
            logger.info(f"Adding column {column} to {table}...")
            try:
                cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {type_def}")
            except Exception as e:
                logger.error(f"Failed to add {column} to {table}: {e}")
        else:
            logger.debug(f"Column {column} already exists in {table}.")

    def create_table_if_missing(table_name, create_sql):
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
        if not cursor.fetchone():
            logger.info(f"Creating missing table {table_name}...")
            try:
                cursor.execute(create_sql)
            except Exception as e:
                logger.error(f"Failed to create table {table_name}: {e}")
        else:
            logger.debug(f"Table {table_name} already exists.")

    try:
        # 1. HARDEN USER TABLE
        user_cols = [
            ('status', "VARCHAR(20) DEFAULT 'offline'"),
            ('last_seen', "DATETIME DEFAULT CURRENT_TIMESTAMP"),
            ('image', "TEXT"),
            ('privacy_last_seen', "VARCHAR(20) DEFAULT 'everyone'"),
            ('privacy_profile_photo', "VARCHAR(20) DEFAULT 'everyone'"),
            ('privacy_about', "VARCHAR(20) DEFAULT 'everyone'"),
            ('privacy_read_receipts', "BOOLEAN DEFAULT 1"),
            ('reset_token', "TEXT"),
            ('reset_token_expiry', "DATETIME"),
            ('phone_number', "TEXT"),
            ('email_otp', "TEXT"),
            ('phone_otp', "TEXT"),
            ('is_verified', "BOOLEAN DEFAULT 0"),
            ('otp_expiry', "DATETIME"),
            ('email_verified', "BOOLEAN DEFAULT 0"),
            ('email_token', "TEXT"),
            ('email_token_expires', "DATETIME"),
            ('phone_verified', "BOOLEAN DEFAULT 0"),
            ('phone_otp_expires', "DATETIME"),
            ('email_otp_expires', "DATETIME"),
            ('bio', "TEXT"),
            ('password_updated_at', "DATETIME"),
            ('two_factor_pin_hash', "VARCHAR(128)"),
            ('two_factor_enabled', "BOOLEAN DEFAULT 0"),
            ('is_active', "BOOLEAN DEFAULT 1"),
            ('deactivated_at', "DATETIME")
        ]
        for col, dtype in user_cols:
            add_column_if_missing('user', col, dtype)

        # 2. HARDEN MEETINGS TABLE
        meeting_cols = [
            ('source', "VARCHAR(50) DEFAULT 'upload'"),
            ('participant_mapping', "TEXT DEFAULT '{}'"),
            ('filepath', "VARCHAR(500)"),
            ('transcript', "TEXT"),
            ('duration', "FLOAT DEFAULT 0.0"),
            ('participants_count', "INTEGER DEFAULT 0"),
            ('is_favorite', "BOOLEAN DEFAULT 0"),
            ('language', "VARCHAR(10) DEFAULT 'en'"),
            ('transcript_language', "VARCHAR(10) DEFAULT 'en'"),
            ('has_transcription', "BOOLEAN DEFAULT 0"),
            ('has_notes', "BOOLEAN DEFAULT 0"),
            ('processing_steps', "TEXT DEFAULT '[]'"),
            ('current_step_progress', "INTEGER DEFAULT 0")
        ]
        for col, dtype in meeting_cols:
            add_column_if_missing('meetings', col, dtype)

        # 3. HARDEN FRIENDSHIP TABLE
        friendship_cols = [
            ('is_pinned', "BOOLEAN DEFAULT 0"),
            ('is_muted', "BOOLEAN DEFAULT 0"),
            ('is_archived', "BOOLEAN DEFAULT 0"),
            ('is_favourite', "BOOLEAN DEFAULT 0"),
            ('is_deleted', "BOOLEAN DEFAULT 0"),
            ('is_blocked', "BOOLEAN DEFAULT 0"),
            ('blocked_by_id', "INTEGER REFERENCES user(id)")
        ]
        for col, dtype in friendship_cols:
            add_column_if_missing('friendship', col, dtype)

        # 4. HARDEN MESSAGE TABLE
        message_cols = [
            ('broadcast_id', "INTEGER REFERENCES broadcast_list(id)"),
            ('is_deleted', "BOOLEAN DEFAULT 0"),
            ('deleted_for', "TEXT DEFAULT '[]'"),
            ('is_starred_by', "TEXT DEFAULT '[]'"),
            ('reaction', "VARCHAR(32)")
        ]
        for col, dtype in message_cols:
            add_column_if_missing('message', col, dtype)

        # 5. CREATE MISSING TABLES
        create_table_if_missing('activity', """
            CREATE TABLE activity (
                id INTEGER PRIMARY KEY,
                user_id INTEGER NOT NULL REFERENCES user(id),
                type VARCHAR(50) NOT NULL,
                title VARCHAR(200) NOT NULL,
                description TEXT,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                meeting_id INTEGER REFERENCES meetings(id),
                activity_metadata TEXT DEFAULT '{}'
            )
        """)

        create_table_if_missing('user_metrics', """
            CREATE TABLE user_metrics (
                id INTEGER PRIMARY KEY,
                user_id INTEGER NOT NULL REFERENCES user(id),
                date DATE NOT NULL,
                uploads_count INTEGER DEFAULT 0,
                processing_time_total INTEGER DEFAULT 0,
                exports_count INTEGER DEFAULT 0,
                active_minutes INTEGER DEFAULT 0,
                languages_used TEXT DEFAULT '[]'
            )
        """)

        # Ensure call_log exists
        create_table_if_missing('call_log', """
            CREATE TABLE call_log (
                id INTEGER PRIMARY KEY,
                user_id INTEGER NOT NULL REFERENCES user(id),
                other_user_id INTEGER REFERENCES user(id),
                type VARCHAR(20) NOT NULL,
                is_video BOOLEAN DEFAULT 0,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                duration INTEGER DEFAULT 0
            )
        """)

        conn.commit()
        logger.info("Migration completed successfully.")

    except Exception as e:
        logger.error(f"Migration error: {e}")
        conn.rollback()
    finally:
        conn.close()

if __name__ == "__main__":
    migrate()
