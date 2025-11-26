import os, re, yaml

_DICTS = {"equipment": [], "target_phrases": [], "disqualifiers": []}

def _load():
    path = os.path.join("config", "dictionaries.yaml")
    if not os.path.isfile(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    for k in _DICTS:
        vals = data.get(k, []) or []
        _DICTS[k] = [str(v).strip() for v in vals if str(v).strip()]

_load()

def detect_equipment(text: str) -> set:
    hits = set(); txt = text or ""
    for brand in _DICTS["equipment"]:
        if re.search(rf"\b{re.escape(brand)}\b", txt, re.I):
            hits.add(brand.title())
    return hits

def detect_phrases(text: str):
    t_hits, d_hits = set(), set()
    txt = text or ""
    for p in _DICTS["target_phrases"]:
        if re.search(p, txt, re.I): t_hits.add(p)
    for p in _DICTS["disqualifiers"]:
        if re.search(p, txt, re.I): d_hits.add(p)
    return t_hits, d_hits
