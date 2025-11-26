import re
from bs4 import BeautifulSoup

def _strip_html(s: str) -> str:
    return re.sub(r"<.*?>", "", s or "", flags=re.S).strip()

def company_name(html: str, text: str) -> str:
    soup = BeautifulSoup(html or "", "lxml")
    title = (soup.title.get_text().strip() if soup.title else "") or ""
    # remove leading "Home | " or trailing " | Company"
    title = re.sub(r"^\s*(Home|Welcome)\s*\|\s*", "", title, flags=re.I)
    title = re.sub(r"\s*\|\s*(Home|Welcome|Official Site).*$", "", title, flags=re.I)
    if title:
        return title[:160]
    h1 = soup.find("h1")
    if h1:
        return h1.get_text(strip=True)[:160]
    return ""

# --- list-like capture helpers ---
_LIST_SEP = re.compile(r"[•\u2022\|\n;,]+")
_INDUSTRY_WORDS = r"(aerospace|medical|automotive|defense|electronics|energy|industrial|pharma|semiconductor|consumer|commercial)"
_SERVICE_WORDS  = r"(mold|tool(ing)?|machin|cnc|edm|grind|design|engineering|metrology|polish|wire edm|sinker edm|inspection|cleanroom)"

def _extract_after_heading(text: str, heading_re: str, vocab_re: str, limit=150):
    """
    Look for a heading like 'Industries' / 'Capabilities' / 'Services',
    take the next ~150 chars, split by bullets/commas, and keep items that match vocab.
    """
    m = re.search(heading_re + r"[:\-]?\s*([^\.\|]{20,200})", text, re.I)
    items = []
    if m:
        chunk = m.group(1)
        for part in _LIST_SEP.split(chunk):
            p = part.strip()
            if re.search(vocab_re, p, re.I) and 2 <= len(p) <= 80:
                items.append(p)
    # fallback: scan full text for vocab words separated by bullets/commas
    if not items:
        candidates = [p.strip() for p in _LIST_SEP.split(text) if 2 <= len(p) <= 80]
        items = [p for p in candidates if re.search(vocab_re, p, re.I)]
    # de-dup while preserving order
    seen, out = set(), []
    for it in items:
        low = it.lower()
        if low not in seen:
            out.append(it); seen.add(low)
        if len(out) >= 20: break
    return ", ".join(out)

def industries(url: str, text: str) -> str:
    return _extract_after_heading(text, r"(industries|markets|applications|we serve)", _INDUSTRY_WORDS)

def services(url: str, text: str) -> str:
    return _extract_after_heading(text, r"(services|capabilities|what we do|our process|capability)", _SERVICE_WORDS)

def facility_sqft(text: str) -> str:
    m = re.search(r"(\d{3,7})\s*(sq\.?\s*ft|square\s*feet|ft²)", text or "", re.I)
    return m.group(1) if m else ""

def employees(text: str) -> str:
    m = re.search(r"(\d{2,5})\s+(employees|team members|staff|people)", text or "", re.I)
    return m.group(1) if m else ""
