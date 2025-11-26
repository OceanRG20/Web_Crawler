# Global crawler behavior
USER_AGENT = "Mozilla/5.0 (MVP crawler; contact: you@example.com)"
TIMEOUT = 25
RETRIES = 2
SLEEP = 1.0
MAX_PAGES_PER_SITE = 12

# Which internal pages to follow during discovery
DISCOVERY_KEYWORDS = [
    "about","company","history","services","industries","capabilities",
    "equipment","careers","jobs","contact","facility","markets"
]

# Evidence saving
SAVE_HTML = False   # if True, saves raw HTML under evidence/<domain>/pages
SAVE_META = True    # saves evidence/<domain>/meta.json

# --- Google Sheets: env-driven toggle ---
import os
from dotenv import load_dotenv
load_dotenv()

SHEETS_ENABLED = os.getenv("SHEETS_ENABLED", "false").lower() == "true"
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
WORKSHEET_NAME  = os.getenv("WORKSHEET_NAME", "Output")
SERVICE_ACCOUNT_JSON = os.getenv("SERVICE_ACCOUNT_JSON", "service_account.json")
FAIL_ON_SHEETS_ERROR = os.getenv("FAIL_ON_SHEETS_ERROR", "false").lower() == "true"
