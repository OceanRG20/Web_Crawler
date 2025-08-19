# Global settings for the manual CSV-only crawler

USER_AGENT = "Mozilla/5.0 (MVP crawler manual; contact: you@example.com)"
TIMEOUT = 25
RETRIES = 2
SLEEP = 1.0
MAX_PAGES_PER_SITE = 12

# Discovery keywords to follow on each domain
DISCOVERY_KEYWORDS = [
    "about","company","history","services","industries","capabilities",
    "equipment","careers","jobs","contact","facility","markets"
]

# Evidence saving
SAVE_HTML = False   # set True to save raw HTML to evidence/{domain}/pages
SAVE_META = True    # save evidence/{domain}/meta.json (discovered URLs, errors)


# --- Google Sheets: env-driven toggle ---
import os
from dotenv import load_dotenv
load_dotenv()

SHEETS_ENABLED = os.getenv("SHEETS_ENABLED", "false").lower() == "true"
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
WORKSHEET_NAME  = os.getenv("WORKSHEET_NAME", "Output")
SERVICE_ACCOUNT_JSON = os.getenv("SERVICE_ACCOUNT_JSON", "service_account.json")

# Fail the run if Sheets breaks? default: False (keep crawling + CSV)
import os
FAIL_ON_SHEETS_ERROR = os.getenv("FAIL_ON_SHEETS_ERROR", "false").lower() == "true"