import re
from bs4 import BeautifulSoup

def company_name(html: str, text: str) -> str:
    title = BeautifulSoup(html or "", "lxml").title
    if title:
        return title.get_text().strip()[:160]
    h1 = re.search(r"<h1[^>]*>(.*?)</h1>", html or "", re.I | re.S)
    if h1:
        import re as _re
        return _re.sub(r"<.*?>", "", h1.group(1)).strip()[:160]
    return ""

def industries(url: str, text: str) -> str:
    if _looks_like_industries(url, text):
        parts = re.split(r"[•\u2022,\n;|]", text or "")
        vals = [p.strip() for p in parts if re.search(r"(aerospace|medical|automotive|defense|electronics|energy|industrial|pharma|semiconductor)", p, re.I)]
        if vals: return ", ".join(sorted(set(vals))[:12])
    return ""

def services(url: str, text: str) -> str:
    if _looks_like_services(url, text):
        parts = re.split(r"[•\u2022,\n;|]", text or "")
        vals = [p.strip() for p in parts if re.search(r"(mold|tool(ing)?|machin|cnc|edm|grind|design|engineering|metrology|polish|wire edm|sinker edm)", p, re.I)]
        if vals: return ", ".join(sorted(set(vals))[:20])
    return ""

def facility_sqft(text: str) -> str:
    m = re.search(r"(\d{3,7})\s*(sq\.?\s*ft|square\s*feet|ft²)", text or "", re.I)
    return m.group(1) if m else ""

def employees(text: str) -> str:
    m = re.search(r"(\d{2,5})\s+(employees|team members|staff|people)", text or "", re.I)
    return m.group(1) if m else ""

def _looks_like_industries(url: str, text: str) -> bool:
    return bool(re.search(r"(industries|markets|applications|we serve)", url or "", re.I) or
                re.search(r"(industries|markets|applications)", text or "", re.I))

def _looks_like_services(url: str, text: str) -> bool:
    return bool(re.search(r"(services|capabilities|what we do)", url or "", re.I) or
                re.search(r"(services|capabilities)", text or "", re.I))
