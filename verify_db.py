from app import app, db
from sqlalchemy import inspect

def check_and_create_tables():
    with app.app_context():
        inspector = inspect(db.engine)
        tables = inspector.get_table_names()
        print(f"Existing tables: {tables}")
        
        required_tables = [
            'group', 'group_member', 'group_invite', 
            'broadcast_list', 'broadcast_recipient', 'call_log'
        ]
        
        missing = [t for t in required_tables if t not in tables]
        
        if missing:
            print(f"Missing tables: {missing}")
            print("Creating missing tables...")
            db.create_all()
            print("Tables created successfully.")
        else:
            print("All required tables exist.")

if __name__ == "__main__":
    check_and_create_tables()
