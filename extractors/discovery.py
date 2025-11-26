from bs4 import BeautifulSoup
from extractors.fetch import http_get, absurl
from config import MAX_PAGES_PER_SITE, DISCOVERY_KEYWORDS

def discover(start_url: str):
    html = http_get(start_url)
    urls = [start_url]
    if not html:
        return urls

    soup = BeautifulSoup(html, "lxml")
    base_netloc = _netloc(start_url)
    seen = set(urls)

    for a in soup.select("a[href]"):
        href = (a.get("href") or "").strip()
        url = absurl(start_url, href)
        if _netloc(url) != base_netloc:
            continue
        low = url.lower()
        if any(f"/{k}" in low or low.endswith(f"/{k}") for k in DISCOVERY_KEYWORDS):
            if url not in seen:
                urls.append(url); seen.add(url)
                if len(urls) >= MAX_PAGES_PER_SITE:
                    break
    return urls

def _netloc(u: str) -> str:
    import urllib.parse
    return urllib.parse.urlparse(u).netloc.lower()
