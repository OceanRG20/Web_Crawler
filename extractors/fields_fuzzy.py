import re
from datetime import datetime

_CURRENT_YEAR = datetime.utcnow().year

def year_established(text: str, html: str = "") -> tuple[str, str]:
    """Return ('1987 (exact)' or '1990 (estimated)', snippet) or ('','')."""
    if not text:
        return "", ""
    # Exact year in safe range
    for m in re.finditer(r"\b(18\d{2}|19\d{2}|20[0-2]\d)\b", text):
        y = int(m.group(1))
        if 1850 <= y <= _CURRENT_YEAR:
            return f"{y} (exact)", m.group(0)
    # Relative: "30+ years"
    m = re.search(r"(\d{1,3})\s*\+?\s*years", text, re.I)
    if m:
        years = int(m.group(1))
        est = max(1850, _CURRENT_YEAR - years)
        return f"{est} (estimated)", m.group(0)
    return "", ""

def owner_and_status(text: str) -> tuple[str, str]:
    if not text:
        return "", ""
    owner = ""
    m = re.search(r"Founded by ([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)", text)
    if m:
        owner = m.group(1)
    t = text.lower()
    status = "Retired" if "retired" in t else ("Active" if "still works" in t or "active" in t else "")
    return owner, status
