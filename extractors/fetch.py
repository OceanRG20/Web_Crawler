import os
import time
import urllib.parse
import requests
from config import USER_AGENT, TIMEOUT, RETRIES, SLEEP, SAVE_HTML

HEADERS = {"User-Agent": USER_AGENT}

def http_get(url: str) -> str:
    """
    Simple fetcher with retries; returns HTML text or ''.
    """
    for attempt in range(1, RETRIES + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
            if r.status_code == 200 and "text/html" in (r.headers.get("Content-Type", "")).lower():
                return r.text
        except Exception:
            pass
        time.sleep(0.6 * attempt)
    return ""

def absurl(base: str, href: str) -> str:
    return urllib.parse.urljoin(base, href or "")

def normalize_domain(u: str) -> str:
    if not u.startswith("http"):
        u = "https://" + u
    netloc = urllib.parse.urlparse(u).netloc.lower()
    return netloc[4:] if netloc.startswith("www.") else netloc

def maybe_save_html(domain: str, url: str, html: str):
    if not SAVE_HTML or not html:
        return
    safe_dir = os.path.join("evidence", domain, "pages")
    os.makedirs(safe_dir, exist_ok=True)
    path = urllib.parse.urlparse(url).path or "index"
    name = path.strip("/").replace("/", "_") or "index"
    fp = os.path.join(safe_dir, f"{name}.html")
    try:
        with open(fp, "w", encoding="utf-8") as f:
            f.write(html)
    except Exception:
        pass
