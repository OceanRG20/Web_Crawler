# Web Crawler – AMBA/Custom Sources

A lightweight Python crawler for extracting company data (name, phone, address, industries, services, employees, etc.) from target websites.  
It outputs results to **CSV** and (optionally) to **Google Sheets**.

---

## 📂 Project Structure

├── crawler.py # Main runner
├── config/
│ ├── urls.txt # Input list of websites
│ ├── csv_schema.yaml # Defines CSV columns
│ └── service_account.json# (Optional) Google Sheets service account
├── extractors/ # HTML/text parsers
├── utils/ # CSV + Sheets utilities
├── output/ # Results (CSV + snapshots)
│ └── snapshots/ # Timestamped backups
└── evidence/ # (Optional) Saved HTML/metadata

## ⚙️ Requirements

- Python **3.8+**
- Install dependencies:

```bash
pip install -r requirements.txt
```
