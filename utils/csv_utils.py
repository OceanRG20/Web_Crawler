import csv
import os

def write_csv(file_path, fieldnames, rows):
    """Write rows to CSV file, overwriting if exists."""
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    with open(file_path, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

def append_csv(file_path, fieldnames, rows):
    """Append rows to CSV, create file if not exists."""
    file_exists = os.path.isfile(file_path)
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    with open(file_path, mode="a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)

def read_csv(file_path):
    """Read CSV file into list of dicts."""
    if not os.path.exists(file_path):
        return []
    with open(file_path, mode="r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader)
