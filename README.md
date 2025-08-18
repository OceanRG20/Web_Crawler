# MVP Crawler → Google Sheets (3 URLs)

## Setup

1. Create a Google Sheet and copy its ID.
2. Create a Service Account in Google Cloud, download the JSON, save as `service_account.json`.
3. Share the Sheet with the service account email (Editor).
4. Fill `.env`:
   - GOOGLE_SHEET_ID=...
   - WORKSHEET_NAME=Output
   - SERVICE_ACCOUNT_JSON=service_account.json
5. Install deps:
