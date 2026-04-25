import os
import yt_dlp
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_extraction(url, cookies_path=None):
    logger.info(f"Testing extraction for: {url}")
    
    ydl_opts = {
        "quiet": False,
        "no_warnings": False,
        "nocheckcertificate": True,
        "youtube_include_dash_manifest": False,
        "youtube_include_hls_manifest": False,
        "extractor_args": {
            "youtube": {
                "player_client": ["ios", "mweb", "android"],
                "skip": ["dash", "hls", "translated_subs"],
            }
        },
    }
    
    if cookies_path and os.path.exists(cookies_path):
        logger.info(f"✅ Found cookies file at: {cookies_path}")
        ydl_opts["cookiefile"] = cookies_path
    else:
        logger.warning("⚠️ No cookies file provided/found. Testing without cookies.")

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            # We use extract_info with download=False to just test if we CAN get the metadata
            info = ydl.extract_info(url, download=False)
            logger.info(f"🚀 SUCCESS! Successfully extracted metadata for: {info.get('title')}")
            return True
    except Exception as e:
        logger.error(f"❌ FAILURE: {str(e)}")
        if "Sign in to confirm you’re not a bot" in str(e):
            logger.error("💡 HINT: This is a bot detection error. Your cookies.txt is either missing, expired, or being ignored.")
        return False

if __name__ == "__main__":
    import sys
    test_url = sys.argv[1] if len(sys.argv) > 1 else "https://www.youtube.com/watch?v=b-Pn0yXL9y8"
    
    # Try current directory cookies.txt
    current_cookies = os.path.join(os.getcwd(), "cookies.txt")
    
    print("\n" + "="*50)
    print("YOUTUBE EXTRACTION DIAGNOSTIC TOOL")
    print("="*50 + "\n")
    
    test_extraction(test_url, current_cookies)
    
    print("\n" + "="*50)
    print("INSTRUCTIONS:")
    print("1. If test failed, export NEW cookies.txt from your browser.")
    print("2. Ensure cookies.txt is in the same folder as this script.")
    print("3. Run this script again: python test_cookies.py")
    print("4. Once it passes LOCALLY, upload it to your server.")
    print("="*50 + "\n")
