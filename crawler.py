#!/usr/bin/env python3
import os, re, time, yaml, argparse
from datetime import datetime

from config import (
    SLEEP, SAVE_HTML, SAVE_META,
    SHEETS_ENABLED, GOOGLE_SHEET_ID, WORKSHEET_NAME, SERVICE_ACCOUNT_JSON,
    FAIL_ON_SHEETS_ERROR
)
from utils.csv_utils import write_csv
if SHEETS_ENABLED:
    from utils.sheet_utils import open_sheet, upsert_record

# ----------------- extractors -----------------
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
    p = argparse.ArgumentParser(description="Lightweight web crawler → CSV + combined TXT")
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
# Helpers
# =========================
# Split US address into parts (Street, City, State, Zip)
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

def detect_yesno_flags(text: str):
    t = (text or "").lower()
    cnc3   = "Y" if ("3 axis" in t or "3-axis" in t or "3axis" in t) else ""
    cnc5   = "Y" if ("5 axis" in t or "5-axis" in t or "5axis" in t) else ""
    spares = "Y" if ("spares" in t or "spare parts" in t or "repair" in t or "repairs" in t) else ""
    family = "Y" if ("family owned" in t or "family-owned" in t or "family business" in t) else ""
    return cnc3, cnc5, spares, family

def years_of_operation_from_evidence(year_str: str):
    if not year_str:
        return ""
    m = re.search(r"\b(18\d{2}|19\d{2}|20[0-2]\d)\b", year_str)
    if m:
        yr = int(m.group(1))
        now = datetime.utcnow().year
        if 1850 <= yr <= now:
            return str(now - yr)
    return ""

def estimated_revenues(emp_str: str):
    try:
        n = int(re.sub(r"[^\d]", "", emp_str or ""))
        return f"${n*200_000:,} (est)"
    except Exception:
        return ""

def target_status(equip_hits, target_hits, disq_hits, had_error):
    if had_error:
        return "X"
    if target_hits or equip_hits:
        return "CY"
    if disq_hits:
        return "CN"
    return "C?"

def _normalize_phone_str(p: str) -> str:
    if not p:
        return ""
    digits = re.sub(r"\D", "", p)
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
    m = re.search(r"\b(\d{4})\b", s)    # left pad rare 4‑digit captures
    if m:
        return m.group(1).rjust(5, "0")
    return ""

# =========================
# TXT Export Helper (combined)
# =========================
def _write_combined_txt(all_blocks: list):
    out_path = os.path.join("output", "all_rows.txt")
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n\n".join(all_blocks))
        print(f"[OK] Combined TXT saved: {out_path}")
    except Exception as e:
        print(f"[WARN] Could not write combined TXT: {e}")

# =========================
# Per-company pipeline
# =========================
def process_company(home_url: str, cols: list, source_hint: str = "") -> dict:
    start_url = home_url if home_url.startswith("http") else f"https://{home_url}"
    domain = normalize_domain(start_url)

    print(f"[INFO] Processing: {start_url}  (Source: {source_hint or '—'})")
    pages = discover(start_url)
    # guarantee home page first if discovery didn't include it
    if start_url not in pages:
        pages = [start_url] + pages

    equip_hits, target_hits, disq_hits = set(), set(), set()
    had_error = False

    now_ts = datetime.now().strftime("%m/%d/%Y %H:%M")
    rec = {c: "" for c in cols}
    rec.update({
        "Company Name": "",
        "Target Status": "C",
        "Public Website Homepage URL": start_url,
        "Domain": domain,
        "Source": source_hint,
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

        print(f"[INFO]   [{idx}/{len(pages)}] Fetched OK: {url}")

        if SAVE_HTML:
            maybe_save_html(domain, url, html)

        text = clean_text(html)

        # Core identity/contacts
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

        # Year evidence → years of operation
        if not rec["Years of operation"]:
            val, _snip = year_established(text, html)
            if val:
                rec["Years of operation"] = years_of_operation_from_evidence(val)

        # Facility & employees & revenue
        if not rec["Square footage (facility)"]:
            rec["Square footage (facility)"] = facility_sqft(text)
        if not rec["Number of employees"]:
            rec["Number of employees"] = employees(text)
            rec["Estimated Revenues"] = estimated_revenues(rec["Number of employees"])

        # Ownership / family
        if not rec["Ownership"] or not rec["Family business"]:
            owner, own_text, is_family = owner_and_status(text)
            if owner or own_text:
                parts = []
                if own_text: parts.append(own_text)
                if owner: parts.append(f"Owner: {owner}")
                rec["Ownership"] = ", ".join(parts)
            if is_family and not rec["Family business"]:
                rec["Family business"] = "Y"

        # Signals → target/disqualifiers
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

        # Binary flags (also inferred from generic text)
        cnc3, cnc5, spares, family = detect_yesno_flags(text)
        rec["CNC 3-axis"]      = rec["CNC 3-axis"] or cnc3
        rec["CNC 5-axis"]      = rec["CNC 5-axis"] or cnc5
        rec["Spares/Repairs"]  = rec["Spares/Repairs"] or spares
        rec["Family business"] = rec["Family business"] or family

    rec["Equipment"] = ", ".join(sorted(equip_hits))
    rec["Specific references from text search"] = ", ".join(sorted(target_hits))
    rec["Target Status"] = target_status(equip_hits, target_hits, disq_hits, had_error)

    print(f"[INFO] Done: {start_url} → Status={rec['Target Status']}")
    return rec

# =========================
# Main
# =========================
def main():
    args = _parse_args()
    _ensure_dirs()
    cols = _read_schema()
    url_pairs = _read_urls()

    ws = None
    if SHEETS_ENABLED:
        try:
            ws = open_sheet(GOOGLE_SHEET_ID, WORKSHEET_NAME, SERVICE_ACCOUNT_JSON)
        except Exception as e:
            print(f"[WARN] Google Sheets unavailable: {e}")
            if FAIL_ON_SHEETS_ERROR:
                raise

    rows = []
    all_blocks = []
    for idx, (u, src_from_file) in enumerate(url_pairs, start=1):
        source_label = args.source or src_from_file
        try:
            rec = process_company(u, cols, source_hint=source_label)
        except Exception as e:
            # fail-soft row with status X
            domain = normalize_domain(u if u.startswith("http") else f"https://{u}")
            now_ts = datetime.now().strftime("%m/%d/%Y %H:%M")
            rec = {c: "" for c in cols}
            rec.update({
                "Public Website Homepage URL": u,
                "Domain": domain,
                "Source": source_label,
                "Target Status": "X",
                "First Seen": now_ts, "Last Seen": now_ts, "Last Update": now_ts
            })
            print(f"[WARN] Error while processing {u}: {e}")

        rows.append(rec)

        if ws is not None:
            try:
                upsert_record(ws, rec)
            except Exception as e:
                print(f"[WARN] Sheets upsert failed for {u}: {e}")
                if FAIL_ON_SHEETS_ERROR:
                    raise

        # Build a clean block for the combined TXT (human-readable, not raw HTML)
        block = [
            f"Row {idx}: {rec.get('Company Name','')}",
            rec.get("Public Website Homepage URL",""),
            f"Address: {rec.get('Street Address','')} {rec.get('City','')} {rec.get('State','')} {rec.get('Zipcode','')}".strip(),
            f"Phone: {rec.get('Phone','')}",
            f"Industries: {rec.get('Industries served','')}",
            f"Services: {rec.get('Products and services offered','')}",
            f"Keywords: {rec.get('Specific references from text search','')}",
            f"Equipment: {rec.get('Equipment','')}",
            f"CNC 3-axis: {rec.get('CNC 3-axis','')}, CNC 5-axis: {rec.get('CNC 5-axis','')}",
            f"Spares/Repairs: {rec.get('Spares/Repairs','')}, Family Business: {rec.get('Family business','')}",
            f"Target Status: {rec.get('Target Status','')}",
            "-"*40
        ]
        all_blocks.append("\n".join(block))

    _write_csv_backup(rows, cols)
    _write_combined_txt(all_blocks)

if __name__ == "__main__":
    main()
