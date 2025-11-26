import csv
import os

def write_csv(file_path, fieldnames, rows):
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    with open(file_path, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow({c: r.get(c, "") for c in fieldnames})