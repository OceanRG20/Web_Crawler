import re
from bs4 import BeautifulSoup

def clean_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "lxml")
    for s in soup(["script","style","noscript"]): s.extract()
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True))

def find_phone(text: str) -> str:
    """
    Capture common US phone formats, prefer the first match.
    """
    m = re.search(r"\+?1?[-.\s]*\(?\b\d{3}\)?[-.\s]*\d{3}[-.\s]*\d{4}\b", text or "")
    return m.group(0) if m else ""

def find_address_us(text: str) -> str:
    """
    Return US-looking address substring such as:
      '205 2nd Ave NW, Bertha MN 56437'
      '205 2nd Ave NW, Bertha, MN 56437'
      '205 2nd Ave NW, Bertha MN 56437, USA'
    """
    pattern = (
        r"\d{2,6}\s+[A-Za-z0-9 .'-]+"     # street number + name
        r",\s*[A-Za-z .'-]+"              # city (after first comma)
        r",?\s*[A-Z]{2}\s+"               # optional comma + state
        r"\d{5}(?:-\d{4})?"               # zip
        r"(?:,\s*(?:USA|United States))?" # optional country
    )
    m = re.search(pattern, text or "", re.I)
    return m.group(0) if m else ""
