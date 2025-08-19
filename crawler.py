#!/usr/bin/env python3
import os
import csv
import json
import time
from datetime import datetime

import yaml

from config import (
    SLEEP, SAVE_HTML, SAVE_META, MAX_PAGES_PER_SITE
)

from extractors.fetch import http_get, normalize_domain, maybe_save_html
from extractors.discovery import discover
from extractors.parse_text import clean_text, find_phone, find_address_us
from extractors.fields_core import company_name, industries, services, facility_sqft, employees
from extractors.fields_fuzzy import year_established, owner_and_status
from extractors.signals import detect_equipment, detect_phrases
from extractors.jobs import extract_jobs


# ---------- paths & IO ----------
def _ensure_dirs():
    os.makedirs("output", exist_ok=True)
    os.makedirs(os.path.join("output", "snapshots"), exist_ok=True)
    if SAVE_HTML or SAVE_META:
        os.makedirs("evidence", exist_ok=True)
    os.makedirs("logs", exist_ok=True)

def _read_urls():
    path = os.path.join("config", "urls.txt")
    urls = []
    with open(path, "r", encoding="utf-8") as f:
        for ln in f:
            ln = ln.strip()
            if not ln or ln.startswith("#"):
                continue
            urls.append(ln)
    return urls

def _read_schema():
    path = os.path.join("config", "csv_schema.yaml")
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    cols = data.get("columns", [])
    if not cols:
        raise SystemExit("csv_schema.yaml has no 'columns' list.")
    return cols

def _write_csv(rows, cols):
    out_path = os.path.join("output", "output.csv")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in cols})

    ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    snap = os.path.join("output", "snapshots", f"output_{ts}.csv")
    try:
        with open(snap, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            for r in rows:
                w.writerow({c: r.get(c, "") for c in cols})
    except Exception:
        pass

    print(f"[OK] wrote {len(rows)} rows -> {out_path} (snapshot saved)")

def _save_meta(domain, pages, errors):
    if not SAVE_META:
        return
    ddir = os.path.join("evidence", domain)
    os.makedirs(ddir, exist_ok=True)
    payload = {
        "domain": domain,
        "discovered_urls": pages,
        "max_pages": MAX_PAGES_PER_SITE,
        "errors": errors,
        "timestamp_utc": datetime.utcnow().isoformat(),
    }
    fp = os.path.join(ddir, "meta.json")
    try:
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
    except Exception:
        pass


# ---------- per-company ----------
def _process_company(url_or_domain: str, columns: list) -> dict:
    start_url = url_or_domain if url_or_domain.startswith("http") else f"https://{url_or_domain}"
    rid = normalize_domain(start_url)
    today = datetime.utcnow().strftime("%Y-%m-%d")

    pages = discover(start_url)
    errors = []
    equip_hits, target_hits, disq_hits = set(), set(), set()

    row = {c: "" for c in columns}
    row.update({
        "record_id": rid,
        "URL": start_url,
        "Source URLs": "|".join(pages),
        "First Seen": today,
        "Last Seen": today,
    })

    for url in pages:
        time.sleep(SLEEP)
        html = http_get(url)
        if not html:
            errors.append(f"no_html:{url}")
            continue

        if SAVE_HTML:
            maybe_save_html(rid, url, html)

        text = clean_text(html)

        # identity
        if not row.get("Name"):
            row["Name"] = company_name(html, text)

        # contacts
        if not row.get("Phone"):
            row["Phone"] = find_phone(text)
        if not row.get("Address"):
            row["Address"] = find_address_us(text)

        # fuzzy year
        if not row.get("Year Established"):
            val, snip = year_established(text, html)
            if val:
                row["Year Established"] = val
                row["Year Evidence URL"] = url
                row["Year Evidence Snippet"] = snip

        # owner & status
        if (not row.get("Owner")) or (not row.get("Owner Status")):
            owner, status = owner_and_status(text)
            if owner and not row.get("Owner"):
                row["Owner"] = owner
                row["Owner Evidence URL"] = url
            if status and not row.get("Owner Status"):
                row["Owner Status"] = status

        # facility & employees
        if not row.get("Facility Size"):
            row["Facility Size"] = facility_sqft(text)
        if not row.get("Employees"):
            row["Employees"] = employees(text)

        # industries & services
        if not row.get("Industries"):
            row["Industries"] = industries(url, text)
        if not row.get("Products/Services"):
            row["Products/Services"] = services(url, text)

        # signals
        equip_hits |= detect_equipment(text)
        t_hits, d_hits = detect_phrases(text)
        target_hits |= t_hits
        disq_hits |= d_hits

        # jobs
        if not row.get("Jobs"):
            jobs = extract_jobs(url, text)
            if jobs:
                row["Jobs"] = jobs

    # finalize
    row["Equipment"] = ", ".join(sorted(equip_hits))
    row["Target Phrases"] = ", ".join(sorted(target_hits))
    row["Disqualifiers"] = ", ".join(sorted(disq_hits))

    # simple score
    score = 0
    if equip_hits: score += 40
    if target_hits: score += 25
    if disq_hits: score -= 40
    if row.get("Year Established"): score += 5
    if row.get("Jobs"): score += 5
    row["Score"] = score
    row["Qualified"] = "Yes" if score >= 50 else "No"

    row["Errors"] = ";".join(errors)
    _save_meta(rid, pages, errors)
    return row


# ---------- main ----------
def main():
    _ensure_dirs()
    urls = _read_urls()
    cols = _read_schema()
    rows = [_process_company(u, cols) for u in urls]
    _write_csv(rows, cols)

if __name__ == "__main__":
    main()
