import os
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Mock objects to match app.py requirements
class MockMeeting:
    def __init__(self):
        self.id = 999
        self.title = "Unicode Test Meeting 🌍"
        self.filename = "test_unicode.mp4"
        self.upload_date = datetime.now()
        # Test text in various languages
        self.test_text = (
            "English: Hello World\n"
            "Arabic: مرحبا بالعالم\n"
            "Hindi: नमस्ते दुनिया\n"
            "Chinese: 你好，世界\n"
            "Urdu: ہیلو ورلڈ"
        )
        self.transcription = json.dumps({
            "optimized": self.test_text
        })
        self.notes = json.dumps({
            "summary": "This meeting tests Unicode support across languages: " + self.test_text,
            "key_points": ["Point 1: " + self.test_text],
            "action_items": ["Action 1: " + self.test_text],
            "decisions": ["Decision 1: " + self.test_text],
            "sentiment": "Positive 😃"
        })

# Import the functions from app.py
# Note: In a real environment, we'd need to make sure imports work correctly.
# For this script to run standalone, I'll copy the logic or assume it's available.
# Since I'm testing the logic I just wrote, I'll use the functions from the ACTUAL app.py if possible.

# For the sake of this verification, I will try to import them if app.py allows it
import sys
sys.path.append('.')
try:
    from app import create_enhanced_docx, create_enhanced_pdf, fix_text_direction
    
    meeting = MockMeeting()
    
    # Test fix_text_direction directly
    print("Testing fix_text_direction...")
    arabic_text = "مرحبا بالعالم"
    reshaped = fix_text_direction(arabic_text)
    print(f"Original: {arabic_text}")
    print(f"Reshaped: {reshaped}")
    
    # Ensure outputs Dir exists
    os.makedirs("outputs", exist_ok=True)
    
    docx_path = "outputs/unicode_test.docx"
    pdf_path = "outputs/unicode_test.pdf"
    
    print("Generating DOCX with Unicode...")
    create_enhanced_docx(meeting, docx_path)
    print(f"DOCX created at {docx_path}")
    
    print("Generating PDF with Unicode...")
    create_enhanced_pdf(meeting, pdf_path)
    print(f"PDF created at {pdf_path}")
    
    print("\nVerification successful if no errors occurred above.")
    print(f"Check the files in {os.path.abspath('outputs')}")

except ImportError as e:
    print(f"Import error: {e}. Make sure you are in the correct directory.")
except Exception as e:
    print(f"An error occurred during verification: {e}")
    import traceback
    traceback.print_exc()
