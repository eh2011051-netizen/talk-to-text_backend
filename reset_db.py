import os
from app import app, db

def reset_database():
    """
    Completely wipes the database and recreates all tables.
    Use this to start with a fresh, empty database.
    """
    print("=" * 60)
    print("DATABASE RESET TOOL")
    print("=" * 60)
    print("WARNING: This will delete ALL users, meetings, messages, and all other data.")
    print("This action CANNOT be undone.")
    
    confirm = input("\nAre you sure you want to proceed? (y/n): ")
    if confirm.lower() != 'y':
        print("\nOperation cancelled.")
        return

    with app.app_context():
        # Get the database URI to show what we're deleting
        db_uri = app.config.get('SQLALCHEMY_DATABASE_URI', 'Unknown')
        print(f"\nTarget Database URI: {db_uri}")
        
        try:
            print("Dropping all tables...")
            db.drop_all()
            
            print("Creating all tables from scratch...")
            db.create_all()
            
            print("\nSUCCESS: Everything has been removed.")
            print("You can now restart your server and register a fresh account.")
            
            # If it's SQLite, show the file path for manual verification
            if db_uri.startswith('sqlite:///'):
                db_path = db_uri.replace('sqlite:///', '')
                # Reconstruct absolute path if it was relative to instance
                if not os.path.isabs(db_path):
                    # In app.py, it's often in 'instance' folder
                    pass 
                print(f"Local SQLite file reset: {db_path}")
                
        except Exception as e:
            print(f"\nERROR during reset: {e}")
            print("Make sure no other processes (like the Flask server) are using the database file.")

if __name__ == "__main__":
    reset_database()
