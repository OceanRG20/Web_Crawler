import re

_JOB_TITLES = [
    r"CNC\s*Machinist", r"Machinist", r"Quality\s*Engineer", r"Engineer",
    r"Toolmaker", r"Operator", r"Production\s*Manager",
    r"Manufacturing\s*Engineer", r"Process\s*Engineer", r"Quality\s*Technician",
]

def looks_like_jobs_url(url: str) -> bool:
    u = (url or "").lower()
    return any(k in u for k in ("careers","jobs","join-our-team","employment"))

def extract_jobs(url: str, text: str) -> str:
    if not looks_like_jobs_url(url):
        return ""
    titles = set()
    for patt in _JOB_TITLES:
        for m in re.finditer(rf"\b{patt}\b", text or "", re.I):
            titles.add(m.group(0))
    return "; ".join(sorted(titles))
