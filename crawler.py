#!/usr/bin/env python3
import os
import re
import time
import yaml
from datetime import datetime

# --- config & I/O helpers ---
from config import (
    SLEEP, SAVE_HTML, SAVE_META, MAX_PAGES_PER_SITE,
    SHEETS_ENABLED, GOOGLE_SHEET_ID, WORKSHEET_NAME, SERVICE_ACCOUNT_JSON, FAIL_ON_SHEETS_ERROR
)

from utils.csv_utils import write_csv
if SHEETS_ENABLED:
    from utils.sheet_utils import open_sheet, upsert_record

# --- extractors ---
from extractors.fetch import http_get, normalize_domain, maybe_save_html
from extractors.discovery import discover
from extractors.parse_text import clean_text, find_phone, find_address_us
from extractors.fields_core import company_name, industries, services, facility_sqft, employees
from extractors.fields_fuzzy import year_established, owner_and_status
from extractors.signals import detect_equipment, detect_phrases
from extractors.jobs import extract_jobs


# =========================
# IO utilities
# =========================
def _ensure_dirs():
    os.makedirs("output", exist_ok=True)
    os.makedirs(os.path.join("output", "snapshots"), exist_ok=True)
    if SAVE_HTML or SAVE_META:
        os.makedirs("evidence", exist_ok=True)

def _read_urls():
    path = os.path.join("config", "urls.txt")
    with open(path, "r", encoding="utf-8") as f:
        return [ln.strip() for ln in f if ln.strip() and not ln.strip().startswith("#")]

def _read_schema():
    path = os.path.join("config", "csv_schema.yaml")
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    cols = data.get("columns", [])
    if not cols:
        raise SystemExit("csv_schema.yaml has no 'columns' list.")
    return cols

def _write_csv_backup(rows, cols):
    # main file
    out_path = os.path.join("output", "output.csv")
    write_csv(out_path, cols, rows)
    # snapshot
    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    snap_path = os.path.join("output", "snapshots", f"output_{ts}.csv")
    try:
        write_csv(snap_path, cols, rows)
    except Exception:
        pass
    print(f"[OK] CSV saved: {out_path} (snapshot created).")

# =========================
# Google Sheets: Safe Open
# =========================
def _try_open_sheet():
    # Skip if Google Sheets integration is disabled
    if not SHEETS_ENABLED:
        return None

    try:
        # Attempt to open the target Google Sheet
        return open_sheet(GOOGLE_SHEET_ID, WORKSHEET_NAME, SERVICE_ACCOUNT_JSON)

    except Exception as e:
        # Log warning if opening fails
        print(f"[WARN] Google Sheets unavailable: {e}")

        # If strict mode is enabled, stop execution
        if 'FAIL_ON_SHEETS_ERROR' in globals() and FAIL_ON_SHEETS_ERROR:
            raise

        # Otherwise, ignore Sheets errors and continue with CSV only
        return None


# =========================
# Google Sheets: Safe Upsert
# =========================
def _try_upsert(ws, rec):
    # Skip if no worksheet available
    if ws is None:
        return

    try:
        # Attempt to insert/update record in Google Sheets
        upsert_record(ws, rec)

    except Exception as e:
        # Log warning if upsert fails
        print(f"[WARN] Sheets upsert failed for {rec.get('Public Website Homepage URL','?')}: {e}")

        # If strict mode is enabled, stop execution
        if 'FAIL_ON_SHEETS_ERROR' in globals() and FAIL_ON_SHEETS_ERROR:
            raise


# =========================
# Helpers for client fields
# =========================
_ADDR_PARTS_RE = re.compile(
    r"(?P<street>\d{2,6}\s+[A-Za-z0-9 .'-]+),\s*(?P<city>[A-Za-z .'-]+),\s*(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)"
)

def split_address(text: str):
    m = _ADDR_PARTS_RE.search(text or "")
    if not m:
        return "","","",""
    return m.group("street"), m.group("city"), m.group("state"), m.group("zip")

def detect_yesno_flags(text: str):
    t = (text or "").lower()
    cnc3   = "Y" if ("3 axis" in t or "3-axis" in t) else ""
    cnc5   = "Y" if ("5 axis" in t or "5-axis" in t) else ""
    spares = "Y" if ("spares" in t or "repair" in t or "repairs" in t) else ""
    family = "Y" if ("family owned" in t or "family-owned" in t or "family business" in t) else ""
    return cnc3, cnc5, spares, family

def years_of_operation(year_str: str):
    """Convert '1987 (exact)' or '1990 (estimated)' to a count of years."""
    if not year_str:
        return ""
    m = re.search(r"\b(18\d{2}|19\d{2}|20[0-2]\d)\b", year_str)
    if m:
        yr = int(m.group(1))
        now = datetime.utcnow().year
        if 1850 <= yr <= now:
            return str(now - yr)
    # fallback: if the string already contains a number like '40 (estimated)'
    m2 = re.search(r"\b(\d{1,3})\b", year_str)
    return m2.group(1) if m2 else ""

def estimated_revenues(emp_str: str):
    """Employees × $200,000 (est)."""
    try:
        n = int(re.sub(r"[^\d]", "", emp_str or ""))
        return f"${n*200_000:,} (est)"
    except Exception:
        return ""

def target_status(equip_hits, target_hits, disq_hits, had_error):
    """
    Status codes:
      C  = Added to DB (default)
      CY = Candidate Yes (signals found)
      CN = Candidate No (disqualifiers)
      C? = Unclear after crawl
      X  = Crawl error/partial
    """
    if had_error:
        return "X"
    if target_hits or equip_hits:
        return "CY"
    if disq_hits:
        return "CN"
    return "C?"


# =========================
# Per-company pipeline
# =========================
def process_company(home_url: str, cols: list, source_hint: str = "") -> dict:
    start_url = home_url if home_url.startswith("http") else f"https://{home_url}"
    domain = normalize_domain(start_url)
    today = datetime.utcnow().strftime("%Y-%m-%d")

    pages = discover(start_url)
    errors = []
    equip_hits, target_hits, disq_hits = set(), set(), set()
    had_error = False

    # Initialize record with all expected columns
    rec = {c: "" for c in cols}
    rec.update({
        "Company Name": "",
        "Target Status": "C",
        "Public Website Homepage URL": start_url,
        "Domain": domain,
        "Source": source_hint,
        "Street Address": "", "City": "", "State": "", "Zipcode": "",
        "Phone": "",
        "Industries served": "", "Products and services offered": "",
        "Specific references from text search": "",
        "Square footage (facility)": "",
        "Number of employees": "",
        "Estimated Revenues": "",
        "Years of operation": "",
        "Ownership": "",
        "Equipment": "",
        "CNC 3-axis": "", "CNC 5-axis": "", "Spares/Repairs": "", "Family business": "",
        "Jobs": "",
        "Year Evidence URL": "", "Year Evidence Snippet": "",
        "Owner Evidence URL": "",
        "Source URLs": "|".join(pages),
        "First Seen": today, "Last Seen": today,
        "Errors": "",
    })

    for url in pages:
        time.sleep(SLEEP)
        html = http_get(url)
        if not html:
            errors.append(f"no_html:{url}")
            had_error = True
            continue

        if SAVE_HTML:
            maybe_save_html(domain, url, html)

        text = clean_text(html)

        # Identity & contacts
        if not rec["Company Name"]:
            rec["Company Name"] = company_name(html, text)
        if not rec["Phone"]:
            rec["Phone"] = find_phone(text)

        # Address (split)
        if not rec["Street Address"]:
            addr_full = find_address_us(text)
            if addr_full:
                st, city, state, zc = split_address(addr_full)
                rec["Street Address"], rec["City"], rec["State"], rec["Zipcode"] = st, city, state, zc

        # Industries / services
        if not rec["Industries served"]:
            rec["Industries served"] = industries(url, text)
        if not rec["Products and services offered"]:
            rec["Products and services offered"] = services(url, text)

        # Years of operation (from year_established)
        if not rec["Years of operation"]:
            val, snip = year_established(text, html)  # e.g., "1990 (estimated)" or "1987 (exact)"
            if val:
                rec["Years of operation"] = years_of_operation(val)
                rec["Year Evidence URL"] = url
                rec["Year Evidence Snippet"] = snip

        # Ownership
        if not rec["Ownership"]:
            owner, status = owner_and_status(text)
            if owner or status:
                combined = f"{owner} ({status})".strip().strip("() ")
                rec["Ownership"] = combined
                rec["Owner Evidence URL"] = url

        # Facility & employees (plus revenue est)
        if not rec["Square footage (facility)"]:
            rec["Square footage (facility)"] = facility_sqft(text)
        if not rec["Number of employees"]:
            rec["Number of employees"] = employees(text)
            rec["Estimated Revenues"] = estimated_revenues(rec["Number of employees"])

        # Signals: equipment & key phrases / disqualifiers
        equip_hits |= detect_equipment(text)
        t_hits, d_hits = detect_phrases(text)
        target_hits |= t_hits
        disq_hits   |= d_hits

        # Jobs (careers)
        if not rec["Jobs"]:
            jobs = extract_jobs(url, text)
            if jobs:
                rec["Jobs"] = jobs

        # Yes/No flags
        cnc3, cnc5, spares, family = detect_yesno_flags(text)
        rec["CNC 3-axis"]      = rec["CNC 3-axis"] or cnc3
        rec["CNC 5-axis"]      = rec["CNC 5-axis"] or cnc5
        rec["Spares/Repairs"]  = rec["Spares/Repairs"] or spares
        rec["Family business"] = rec["Family business"] or family

    # Finalize signals & status
    rec["Equipment"] = ", ".join(sorted(equip_hits))
    rec["Specific references from text search"] = ", ".join(sorted(target_hits))
    rec["Target Status"] = target_status(equip_hits, target_hits, disq_hits, had_error)
    rec["Errors"] = ";".join(errors)

    return rec


# =========================
# Main
# =========================
def main():
    _ensure_dirs()
    urls = _read_urls()
    cols = _read_schema()

    # Open the sheet, but don't crash if it fails
    ws = _try_open_sheet()

    rows = []
    for u in urls:
        rec = process_company(u, cols)
        rows.append(rec)
        # Try to upsert; ignore failure so crawl continues
        _try_upsert(ws, rec)

    # Always write CSV backup, regardless of Sheets status
    _write_csv_backup(rows, cols)

    if ws is None and SHEETS_ENABLED:
        print("[INFO] Run completed with CSV only (Sheets unavailable).")
    elif SHEETS_ENABLED:
        print("[OK] CSV written and Google Sheet updated.")
    else:
        print("[OK] CSV written (Sheets disabled).")



if __name__ == "__main__":
    main()
