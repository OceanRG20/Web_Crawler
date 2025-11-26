import re
from bs4 import BeautifulSoup

# ---------- Text helpers ----------
def clean_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "lxml")
    for s in soup(["script", "style", "noscript"]):
        s.extract()
    # collapse whitespace but keep spaces between blocks
    return re.sub(r"[ \t]+", " ", soup.get_text(" ", strip=True))

# ---------- Phone detection ----------
# Accepts many variants (with/without +1, (), ., -, spaces), ignores extensions.
_PHONE_RX = re.compile(
    r"""(?xi)
    (?:\+?1[\s\.\-]*)?               # optional +1
    \(?\s*(\d{3})\s*\)?[\s\.\-]*     # area
    (\d{3})[\s\.\-]*                 # prefix
    (\d{4})                          # line
    (?:\s*(?:x|ext\.?|extension)\s*\d+)?  # optional extension (ignored)
    """
)

def _format_phone_triplet(a: str, p: str, l: str) -> str:
    return f"({a}) {p}-{l}"

def find_phone(text: str) -> str:
    """
    Return first US phone in canonical form '(AAA) PPP-LLLL', or '' if none.
    """
    if not text:
        return ""
    m = _PHONE_RX.search(text)
    if not m:
        return ""
    return _format_phone_triplet(m.group(1), m.group(2), m.group(3))

# ---------- Address detection ----------
_ADDR_RX = re.compile(
    r"""(?xi)
    \b
    (\d{2,6}\s+[A-Za-z0-9 .'\-]+)     # street (group 1)
    ,\s*([A-Za-z .'\-]+)              # city   (group 2)
    ,?\s*([A-Z]{2})\s+                # state  (group 3)
    (\d{5}(?:-\d{4})?)                # ZIP or ZIP+4 (group 4)
    (?:\s*,\s*(?:USA|United States))? # optional country
    \b
    """,
)

def find_address_us(text: str) -> str:
    """
    Return a US-looking address substring if present, else ''.
    """
    m = _ADDR_RX.search(text or "")
    return m.group(0) if m else ""

def split_address(addr: str):
    """
    Split a matched address into (street, city, state, zip). Returns empty strings if no match.
    """
    m = _ADDR_RX.search(addr or "")
    if not m:
        return "", "", "", ""
    street, city, state, zipc = m.group(1), m.group(2), m.group(3), m.group(4)
    return street.strip(), city.strip(), state.strip(), zipc.strip()
