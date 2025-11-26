import csv
import json
import re
from pathlib import Path
import pgeocode
import sys

# Resolve project root (…/Web_Crawler), even when run from utils/
PROJECT_ROOT = Path(__file__).resolve().parents[1]

# === CONFIG (static paths relative to project root) ===
ZIP_FILE = PROJECT_ROOT / "config" / "zips.txt"
CSV_OUT = PROJECT_ROOT / "exports" / "zip_cities.csv"
JSON_OUT = PROJECT_ROOT / "exports" / "zip_cities.json"

def normalize_zip(z):
    """Normalize ZIP or ZIP+4 into a 5-digit string with leading zeros."""
    if not z:
        return ""
    s = re.sub(r"[^\d]", "", str(z).strip())  # digits only
    if len(s) >= 5:
        s = s[:5]
    return s.zfill(5) if s.isdigit() and len(s) == 5 else ""

def read_zip_list(path: Path):
    """Read zips from txt file (one per line)."""
    if not path.exists():
        raise FileNotFoundError(f"ZIP input file not found: {path}")
    zips = []
    for line in path.read_text(encoding="utf-8").splitlines():
        z = normalize_zip(line)
        if z:
            zips.append(z)
    return zips

def lookup_city_state(zips):
    nomi = pgeocode.Nominatim("us")
    rows = []
    for z in zips:
        rec = nomi.query_postal_code(z)
        if rec is None or rec.place_name is None:
            rows.append({"zip": z, "city": "", "state": "", "found": "false"})
        else:
            rows.append({
                "zip": z,
                "city": rec.place_name or "",
                "state": rec.state_code or "",
                "found": "true",
            })
    return rows

def write_csv(path: Path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    cols = ["zip", "city", "state", "found"]
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in cols})

def write_json(path: Path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)

def main():
    try:
        print(f"Reading ZIPs from: {ZIP_FILE}")
        zips = read_zip_list(ZIP_FILE)
    except FileNotFoundError as e:
        print(f"❌ {e}")
        print("👉 Create the file and try again. Example contents:\n  54630\n  60193\n  06455\n")
        sys.exit(1)

    rows = lookup_city_state(zips)
    write_csv(CSV_OUT, rows)
    write_json(JSON_OUT, rows)
    print(f"✅ Processed {len(rows)} ZIP codes")
    print(f"CSV → {CSV_OUT}")
    print(f"JSON → {JSON_OUT}")

if __name__ == "__main__":
    main()
