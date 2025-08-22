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

# ----- extractors -----
from extractors.fetch import http_get, normalize_domain, maybe_save_html
from extractors.discovery import discover
from extractors.parse_text import (
    clean_text, find_phone, find_address_us, split_address
)
from extractors.fields_core import company_name, industries, services, facility_sqft, employees
from extractors.fields_fuzzy import year_established, owner_and_status
from extractors.signals import detect_equipment, detect_phrases
from extractors.jobs import extract_jobs

# =========================
# CLI / IO utilities
# =========================
TS_FORMAT = "%Y-%m-%d %H:%M:%S"   # Excel-friendly

def _parse_args():
    p = argparse.ArgumentParser(description="Lightweight site crawler -> CSV (+ optional Sheets upsert)")
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
def _normalize_phone_str(p: str) -> str:
    """Ensure '(AAA) PPP-LLLL' or ''."""
    if not p:
        return ""
    digits = re.sub(r"\D", "", p)
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    return p.strip()

def _normalize_zip(z: str) -> str:
    """Normalize to 5-digit or ZIP+4 (#####-####)."""
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

def _years_of_operation_from_evidence(year_str: str):
    if not year_str:
        return ""
    m = re.search(r"\b(18\d{2}|19\d{2}|20[0-2]\d)\b", year_str)
    if m:
        yr = int(m.group(1))
        now = datetime.utcnow().year
        if 1850 <= yr <= now:
            return str(now - yr)
    m2 = re.search(r"\b(\d{1,3})\b", year_str)
    return m2.group(1) if m2 else ""

def _estimated_revenues_from_employees(emp_str: str):
    try:
        n = int(re.sub(r"[^\d]", "", emp_str or ""))
        return f"${n*200_000:,} (est)"
    except Exception:
        return ""

def _target_status(equip_hits, target_hits, disq_hits, had_error):
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

    # discover without unsupported kwargs
    pages = discover(start_url)
    if start_url not in pages:
        pages = [start_url] + pages

    equip_hits, target_hits, disq_hits = set(), set(), set()
    had_error = False

    ts_now = datetime.now().strftime(TS_FORMAT)
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
        "First Seen": ts_now,
        "Last Seen": ts_now,
        "Last Update": ts_now,
        "Notes (Approach/Contacts/Info)": "",
    })

    print(f"[INFO] Processing: {start_url}  (Source: {source_hint or '—'})")
    for idx, url in enumerate(pages, start=1):
        time.sleep(SLEEP)
        try:
            html = http_get(url)
            if not html:
                had_error = True
                print(f"[WARN]   [{idx}/{len(pages)}] No HTML fetched: {url}")
                continue
            if SAVE_HTML:
                maybe_save_html(domain, url, html)

            print(f"[INFO]   [{idx}/{len(pages)}] Fetched OK: {url}")
            text = clean_text(html)

            # Company & phone
            if not rec["Company Name"]:
                rec["Company Name"] = company_name(html, text)
            if not rec["Phone"]:
                rec["Phone"] = find_phone(text)
            if rec.get("Phone"):
                rec["Phone"] = _normalize_phone_str(rec["Phone"])

            # Address
            if not rec["Street Address"]:
                addr_full = find_address_us(text)
                if addr_full:
                    st, city, state, zc = split_address(addr_full)
                    rec["Street Address"], rec["City"], rec["State"], rec["Zipcode"] = (
                        st, city, state, _normalize_zip(zc)
                    )

            # Industries / services
            if not rec["Industries served"]:
                rec["Industries served"] = industries(url, text)
            if not rec["Products and services offered"]:
                rec["Products and services offered"] = services(url, text)

            # Years of operation
            if not rec["Years of operation"]:
                val, snip = year_established(text, html)
                if val:
                    rec["Years of operation"] = _years_of_operation_from_evidence(val)
                    rec["Year Evidence URL"] = url
                    rec["Year Evidence Snippet"] = snip

            # Ownership + family flag
            if not rec["Ownership"] or not rec["Family business"]:
                owner, own_text, is_family = owner_and_status(text)
                if owner or own_text or is_family:
                    parts = []
                    if own_text: parts.append(own_text)
                    if owner: parts.append(f"Owner: {owner}")
                    if parts:
                        rec["Ownership"] = ", ".join(parts)
                        rec["Owner Evidence URL"] = url
                    if is_family and not rec["Family business"]:
                        rec["Family business"] = "Y"

            # Facility & employees + revenue est
            if not rec["Square footage (facility)"]:
                rec["Square footage (facility)"] = facility_sqft(text)
            if not rec["Number of employees"]:
                rec["Number of employees"] = employees(text)
                rec["Estimated Revenues"] = _estimated_revenues_from_employees(rec["Number of employees"])

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

        except Exception as e:
            had_error = True
            print(f"[WARN]   [{idx}/{len(pages)}] Error while parsing {url}: {e}")

    rec["Equipment"] = ", ".join(sorted(equip_hits))
    rec["Specific references from text search"] = ", ".join(sorted(target_hits))
    rec["Target Status"] = _target_status(equip_hits, target_hits, disq_hits, had_error)

    print(f"[INFO] Done: {start_url} → Status={rec['Target Status']}")
    return rec

# =========================
# Google Sheets (best-effort)
# =========================
def _try_open_sheet():
    if not SHEETS_ENABLED:
        return None
    try:
        return open_sheet(GOOGLE_SHEET_ID, WORKSHEET_NAME, SERVICE_ACCOUNT_JSON)
    except Exception as e:
        print(f"[WARN] Google Sheets unavailable: {e}")
        if FAIL_ON_SHEETS_ERROR:
            raise
        return None

def _try_upsert(ws, rec):
    if ws is None:
        return
    try:
        upsert_record(ws, rec)
    except Exception as e:
        print(f"[WARN] Sheets upsert failed for {rec.get('Public Website Homepage URL','?')}: {e}")
        if FAIL_ON_SHEETS_ERROR:
            raise

# =========================
# Main
# =========================
def main():
    args = _parse_args()
    _ensure_dirs()
    cols = _read_schema()
    url_pairs = _read_urls()

    ws = _try_open_sheet()

    rows = []
    for u, src_from_file in url_pairs:
        source_label = args.source or src_from_file
        try:
            print(f"[INFO] Processing: {u}  (Source: {source_label or '—'})")
            rec = process_company(u, cols, source_hint=source_label)
        except Exception as e:
            # Hard fail for this URL — emit minimal record so CSV still shows it
            print(f"[WARN] Error while processing {u}: {e}")
            domain = normalize_domain(u if u.startswith("http") else f"https://{u}")
            ts = datetime.now().strftime(TS_FORMAT)
            rec = {c: "" for c in cols}
            rec.update({
                "Public Website Homepage URL": u,
                "Domain": domain,
                "Source": source_label,
                "Target Status": "X",
                "First Seen": ts, "Last Seen": ts, "Last Update": ts,
            })

        rows.append(rec)
        _try_upsert(ws, rec)
        time.sleep(SLEEP)

    _write_csv_backup(rows, cols)

    if ws is None:
        print("[INFO] Run completed with CSV only (Sheets unavailable).")
    else:
        print("[INFO] Run completed with CSV + Sheets upserts.")

if __name__ == "__main__":
    main()
