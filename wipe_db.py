
from app import app, db
import os

def wipe_database():
    with app.app_context():
        print("Wiping all tables...")
        db.drop_all()
        db.create_all()
        print("Database wiped and tables recreated successfully.")

if __name__ == "__main__":
    wipe_database()