import re
from bs4 import BeautifulSoup

def clean_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "lxml")
    for s in soup(["script","style","noscript"]): s.extract()
    import re as _re
    return _re.sub(r"\s+", " ", soup.get_text(" ", strip=True))

def find_phone(text: str) -> str:
    m = re.search(r"\(?\b\d{3}\)?[-.\s]*\d{3}[-.\s]*\d{4}\b", text or "")
    return m.group(0) if m else ""

def find_address_us(text: str) -> str:
    m = re.search(r"\d{2,6}\s+[A-Za-z0-9 .'-]+,\s*[A-Za-z .'-]+,\s*[A-Z]{2}\s+\d{5}(?:-\d{4})?", text or "")
    return m.group(0) if m else ""
