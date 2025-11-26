#!/usr/bin/env python3
# CSV + raw_export.txt (deduplicated, organized raw blocks) — NO Google Sheets

import os, re, time, yaml, argparse
from datetime import datetime
from difflib import SequenceMatcher

from config import (
    SLEEP, SAVE_HTML, SAVE_META,
)
from utils.csv_utils import write_csv

# -------- extractors --------
from extractors.fetch import http_get, normalize_domain, maybe_save_html
from extractors.discovery import discover
from extractors.parse_text import clean_text, find_phone, find_address_us
from extractors.fields_core import company_name, industries, services, facility_sqft, employees
from extractors.fields_fuzzy import year_established, owner_and_status
from extractors.signals import detect_equipment, detect_phrases
from extractors.jobs import extract_jobs


# =========================
# CLI / IO utilities
# =========================
def _parse_args():
    p = argparse.ArgumentParser(description="Crawler → CSV + raw TXT (deduped)")
    p.add_argument("--source", default="", help="Label for Source column (e.g., AMBA, ThomasNet, Manual)")
    return p.parse_args()

def _ensure_dirs():
    os.makedirs("output", exist_ok=True)
    os.makedirs(os.path.join("output", "snapshots"), exist_ok=True)
    if SAVE_HTML or SAVE_META:
        os.makedirs("evidence", exist_ok=True)

def _read_urls():
    path = os.path.join("config", "urls.txt")
    pairs = []
    with open(path, "r", encoding="utf-8") as f:
        for ln in f:
            ln = ln.strip()
            if not ln or ln.startswith("#"):
                continue
            if "|" in ln:
                url, src = [x.strip() for x in ln.split("|", 1)]
            else:
                url, src = ln, ""
            pairs.append((url, src))
    return pairs

def _read_schema():
    path = os.path.join("config", "csv_schema.yaml")
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    cols = data.get("columns", [])
    if not cols:
        raise SystemExit("csv_schema.yaml has no 'columns' list.")
    return cols

def _write_csv_backup(rows, cols):
    out_path = os.path.join("output", "output.csv")
    write_csv(out_path, cols, rows)
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    snap_path = os.path.join("output", "snapshots", f"output_{ts}.csv")
    try:
        write_csv(snap_path, cols, rows)
    except Exception:
        pass
    print(f"[OK] CSV saved: {out_path} (snapshot created).")


# =========================
# Helper Normalizers
# =========================
_ADDR_PARTS_RE = re.compile(
    r"(?P<street>\d{2,6}\s+[A-Za-z0-9 .'-]+),\s*"
    r"(?P<city>[A-Za-z .'-]+),?\s*"
    r"(?P<state>[A-Z]{2})\s+"
    r"(?P<zip>\d{5}(?:-\d{4})?)"
    r"(?:,\s*(?:USA|United States))?",
    re.I
)

def split_address(text: str):
    m = _ADDR_PARTS_RE.search(text or "")
    if not m:
        return "","","",""
    return m.group("street"), m.group("city"), m.group("state"), m.group("zip")

def _normalize_phone_str(p: str) -> str:
    if not p:
        return ""
    digits = re.sub(r"\D", "", p)
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    return p.strip()

def _normalize_zip(z: str) -> str:
    if not z:
        return ""
    s = re.sub(r"[^\d \-]", "", z)
    m = re.search(r"\b(\d{5})[- ]?(\d{4})\b", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"\b(\d{5})\b", s)
    if m:
        return m.group(1)
    m = re.search(r"\b(\d{4})\b", s)
    if m:
        return m.group(1).rjust(5, "0")
    return ""


# =========================
# Founding year / employees
# =========================
_YEAR_RE = re.compile(r"\b(since|est\.?|founded|established)\s*(in\s*)?(\d{4})\b", re.I)
_EMP_RE  = re.compile(r"\b(\d{1,4})\s*(employees|staff|team members|associates)\b", re.I)

def years_of_operation_from_text(text: str) -> str:
    m = _YEAR_RE.search(text or "")
    if not m: 
        return ""
    yr = int(m.group(3))
    now = datetime.utcnow().year
    if 1850 <= yr <= now:
        return str(now - yr)
    return ""

def extract_employee_count(text: str) -> str:
    m = _EMP_RE.search(text or "")
    if not m:
        return ""
    return m.group(1)

def estimated_revenues(emp_str: str):
    try:
        n = int(re.sub(r"[^\d]", "", emp_str or ""))  # $200k / employee
        return f"${n*200_000:,} (est)"
    except Exception:
        return ""


# =========================
# Target Status
# =========================
def target_status(equip_hits, target_hits, disq_hits, had_error):
    if had_error:
        return "X"
    # weight positive signals stronger than generic equipment
    strong_targets = {"lsr", "liquid silicone rubber", "medical device", "class 101", "implantable", "drug delivery"}
    has_strong = any(t.lower() in strong_targets for t in target_hits)
    if has_strong or equip_hits:
        if disq_hits:
            # if both, lean unknown → C?
            return "C?"
        return "CY"
    if disq_hits:
        return "CN"
    return "C?"


# =========================
# De-duplication utilities for RAW text
# =========================
_NOISE_PAT = re.compile(
    r"(privacy|terms|cookies|copyright|all rights reserved|subscribe|"
    r"login|sign in|create account|follow us|menu|home|sitemap|"
    r"©\s?\d{4}|^\d{1,2}:\d{2}\s?(am|pm)\b)",
    re.I
)

def _clean_lines_to_sentences(text: str) -> list:
    """Split into sentences or short lines; drop obvious noise."""
    # normalize whitespace
    t = re.sub(r"\s+", " ", text or " ").strip()
    # slice into pseudo-sentences (keep things readable)
    # split on . ; : | bullets, while preserving important commas
    parts = re.split(r"(?<=[\.\!\?])\s+|[\|\u2022•]+", t)
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # kill short crumbs & pure noise
        if len(p) < 25:
            # allow short lines only if they look like contact facts
            if not re.search(r"\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}|\b[A-Z]{2}\s?\d{5}\b", p):
                continue
        if _NOISE_PAT.search(p):
            continue
        out.append(p)
    return out

def _is_similar(a: str, b: str, thresh: float = 0.92) -> bool:
    return SequenceMatcher(None, a.lower(), b.lower()).ratio() >= thresh

def dedupe_sentences(sentences: list, similarity=0.92, max_sentences=120):
    """Keep first occurrence; drop exact & near-duplicate sentences."""
    kept = []
    for s in sentences:
        if any(_is_similar(s, k, similarity) for k in kept):
            continue
        kept.append(s)
        if len(kept) >= max_sentences:
            break
    return kept

def build_raw_block(rec: dict, raw_text: str, row_idx: int) -> str:
    """Client-style block: header lines + deduped, readable text."""
    lines = []
    lines.append(f"Row {row_idx}:  {rec.get('Company Name','')}".rstrip())
    lines.append(rec.get("Public Website Homepage URL",""))
    if rec.get("Domain"):
        lines.append(rec["Domain"])

    # Address
    street = rec.get("Street Address","").strip()
    city = rec.get("City","").strip()
    state = rec.get("State","").strip()
    zc = rec.get("Zipcode","").strip()
    if street:
        lines.append(street)
    city_line = " ".join([x for x in [f"{city}," if city else "", state, zc] if x]).strip()
    if city_line:
        lines.append(city_line)

    # Phone
    if rec.get("Phone"):
        lines.append(f"Phone:  {rec['Phone']}")

    # Deduplicate & organize raw text
    if raw_text:
        sentences = _clean_lines_to_sentences(raw_text)
        sentences = dedupe_sentences(sentences, similarity=0.93, max_sentences=140)
        if sentences:
            lines.append("")
            lines.extend(sentences)

    return "\n".join(lines).rstrip() + "\n\n"


# =========================
# Per-company pipeline
# =========================
def process_company(home_url: str, cols: list, source_hint: str = ""):
    """Return (rec, raw_concat_text) after crawling & text merging."""
    start_url = home_url if home_url.startswith("http") else f"https://{home_url}"
    domain = normalize_domain(start_url)

    print(f"[INFO] Processing: {start_url}  (Source: {source_hint or '—'})")
    pages = discover(start_url)
    if start_url not in pages:
        pages = [start_url] + pages

    equip_hits, target_hits, disq_hits = set(), set(), set()
    had_error = False
    raw_text_parts = []  # collect raw site text

    now_ts = datetime.now().strftime("%m/%d/%Y %H:%M")
    rec = {c: "" for c in cols}
    rec.update({
        "Company Name": "",
        "Target Status": "",
        "Public Website Homepage URL": start_url,
        "Domain": domain,
        "Source": source_hint or "",
        "Street Address": "", "City": "", "State": "", "Zipcode": "",
        "Phone": "",
        "Industries served": "",
        "Products and services offered": "",
        "Specific references from text search": "",
        "Square footage (facility)": "",
        "Number of employees": "",
        "Estimated Revenues": "",
        "Years of operation": "",
        "Ownership": "",
        "Equipment": "",
        "CNC 3-axis": "", "CNC 5-axis": "", "Spares/Repairs": "", "Family business": "",
        "Jobs": "",
        "Source URLs": "|".join(pages),
        "First Seen": now_ts,
        "Last Seen": now_ts,
        "Last Update": now_ts,
        "Notes (Approach/Contacts/Info)": "",
    })

    for idx, url in enumerate(pages, start=1):
        time.sleep(SLEEP)
        html = http_get(url)
        if not html:
            had_error = True
            continue

        if SAVE_HTML:
            maybe_save_html(domain, url, html)

        text = clean_text(html)

        # accumulate raw text
        raw_text_parts.append(text)

        # Company / phone
        if not rec["Company Name"]:
            rec["Company Name"] = company_name(html, text)
        if not rec["Phone"]:
            rec["Phone"] = _normalize_phone_str(find_phone(text))

        # Address
        if not rec["Street Address"]:
            addr_full = find_address_us(text)
            if addr_full:
                st, city, state, zc = split_address(addr_full)
                rec["Street Address"] = st
                rec["City"] = city
                rec["State"] = state
                rec["Zipcode"] = _normalize_zip(zc)

        # Industries / services
        if not rec["Industries served"]:
            rec["Industries served"] = industries(url, text)
        if not rec["Products and services offered"]:
            rec["Products and services offered"] = services(url, text)

        # Year evidence → years of operation (fallback from fuzzy extractor)
        if not rec["Years of operation"]:
            # prefer fuzzy function if it returns a year string
            y_val, _snip = year_established(text, html)
            yrs = ""
            if y_val:
                # y_val might be a year string like "1985"
                try:
                    y = int(re.search(r"\b(18\d{2}|19\d{2}|20[0-2]\d)\b", y_val).group(0))
                    yrs = str(datetime.utcnow().year - y)
                except Exception:
                    pass
            if not yrs:
                yrs = years_of_operation_from_text(text)
            rec["Years of operation"] = yrs

        # Employees + revenue
        if not rec["Number of employees"]:
            # prefer dedicated extractor; then regex fallback
            emp = employees(text) or extract_employee_count(text)
            rec["Number of employees"] = emp
            rec["Estimated Revenues"] = estimated_revenues(emp)

        # Ownership / family
        if not rec["Ownership"] or not rec["Family business"]:
            owner, own_text, is_family = owner_and_status(text)
            parts = []
            if own_text: parts.append(own_text)
            if owner: parts.append(f"Owner: {owner}")
            if parts and not rec["Ownership"]:
                rec["Ownership"] = ", ".join(parts)
            if is_family and not rec["Family business"]:
                rec["Family business"] = "Y"

        # Signals
        e_hits = detect_equipment(text)
        equip_hits |= e_hits
        t_hits, d_hits = detect_phrases(text)
        target_hits |= t_hits
        disq_hits   |= d_hits

        # Jobs
        if not rec["Jobs"]:
            jobs = extract_jobs(url, text)
            if jobs:
                rec["Jobs"] = jobs

        # Binary flags inferred
        t_lower = (text or "").lower()
        if not rec["CNC 3-axis"] and re.search(r"\b(3[-\s]?axis|3 axis)\b", t_lower):
            rec["CNC 3-axis"] = "Y"
        if not rec["CNC 5-axis"] and re.search(r"\b(5[-\s]?axis|5 axis)\b", t_lower):
            rec["CNC 5-axis"] = "Y"
        if not rec["Spares/Repairs"] and re.search(r"\b(spares|spare parts|repair|repairs|maintenance)\b", t_lower):
            rec["Spares/Repairs"] = "Y"

    # Summaries
    rec["Equipment"] = ", ".join(sorted(equip_hits))
    rec["Specific references from text search"] = ", ".join(sorted(target_hits))
    rec["Target Status"] = target_status(equip_hits, target_hits, disq_hits, had_error)  # <-- always set

    # raw text (for dedupe/organize)
    raw_concat = " ".join(raw_text_parts)
    return rec, raw_concat


# =========================
# Main
# =========================
def main():
    args = _parse_args()
    _ensure_dirs()
    cols = _read_schema()
    url_pairs = _read_urls()

    # clear raw_export.txt at the start
    raw_fp = os.path.join("output", "raw_export.txt")
    open(raw_fp, "w", encoding="utf-8").close()

    rows = []
    for i, (u, src_from_file) in enumerate(url_pairs, start=1):
        source_label = args.source or src_from_file or ""
        try:
            rec, raw_txt = process_company(u, cols, source_hint=source_label)
        except Exception as e:
            domain = normalize_domain(u if u.startswith("http") else f"https://{u}")
            now_ts = datetime.now().strftime("%m/%d/%Y %H:%M")
            rec, raw_txt = ({c: "" for c in cols}, "")
            rec.update({
                "Public Website Homepage URL": u,
                "Domain": domain,
                "Source": source_label,
                "Target Status": "X",
                "First Seen": now_ts, "Last Seen": now_ts, "Last Update": now_ts
            })
            print(f"[WARN] Error while processing {u}: {e}")

        rows.append(rec)

        # Append client-style block (deduped/raw organized)
        block = build_raw_block(rec, raw_txt, i)
        with open(raw_fp, "a", encoding="utf-8") as f:
            f.write(block)

    _write_csv_backup(rows, cols)
    print("[OK] CSV + raw TXT written.")


if __name__ == "__main__":
    main()
