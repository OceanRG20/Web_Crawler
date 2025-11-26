import re
from datetime import datetime

_CURRENT_YEAR = datetime.utcnow().year

def year_established(text: str, html: str = "") -> tuple[str, str]:
    """Return ('1987 (exact)' or '1990 (estimated)', snippet) or ('','')."""
    if not text:
        return "", ""
    for m in re.finditer(r"\b(18\d{2}|19\d{2}|20[0-2]\d)\b", text):
        y = int(m.group(1))
        if 1850 <= y <= _CURRENT_YEAR:
            return f"{y} (exact)", m.group(0)
    m = re.search(r"(\d{1,3})\s*\+?\s*years", text, re.I)
    if m:
        years = int(m.group(1))
        est = max(1850, _CURRENT_YEAR - years)
        return f"{est} (estimated)", m.group(0)
    return "", ""

# --- Ownership parsing ---
# We try hard to detect: owner name, private/family text, and "status" hint (Active/Retired/etc.)
_OWNER_NAME = r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,2})"

def owner_and_status(text: str) -> tuple[str, str, bool]:
    """
    Returns: (owner_name, ownership_text, is_family_business)
    Examples of matches:
      'Founded by John Smith in 1987'
      'Owner: Jane Doe'
      'Privately owned'
      'We are a family-owned business'
    """
    if not text:
        return "", "", False
    t = text.lower()

    # Family business flag
    family = any(kw in t for kw in ["family-owned", "family owned", "family business"])

    # Ownership text (high level)
    ownership_bits = []
    if "privately owned" in t or "privately-held" in t or "privately held" in t:
        ownership_bits.append("Privately owned")
    if family:
        ownership_bits.append("Family-owned")

    # Try to find a likely owner / founder name
    owner = ""
    m = re.search(rf"Founded by\s+{_OWNER_NAME}", text)
    if m:
        owner = m.group(1)
        if "retired" in t:
            ownership_bits.append("Founder (retired)")
    else:
        m = re.search(rf"Owner[:\s]+\s*{_OWNER_NAME}", text, re.I)
        if m:
            owner = m.group(1)
        else:
            m = re.search(rf"(?:CEO|President|Founder)\s+{_OWNER_NAME}", text)
            if m:
                owner = m.group(1)

    # Status
    status = ""
    if "retired" in t:
        status = "Retired"
    elif "still works" in t or "active" in t:
        status = "Active"

    # Build ownership text
    own_text = ", ".join([p for p in ownership_bits if p]) or ("Privately owned" if "private" in t else "")
    if not own_text and owner:
        own_text = "Owner identified"

    return owner, own_text, family
