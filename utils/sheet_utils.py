import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Columns the crawler manages. Notes columns live to the right and are untouched.
MANAGED_COLS = [
    "Company Name","Target Status","Public Website Homepage URL","Domain","Source",
    "Street Address","City","State","Zipcode","Phone",
    "Industries served","Products and services offered","Specific references from text search",
    "Square footage (facility)","Number of employees","Estimated Revenues","Years of operation",
    "Ownership","Equipment","CNC 3-axis","CNC 5-axis","Spares/Repairs","Family business","Jobs",
    "Year Evidence URL","Year Evidence Snippet","Owner Evidence URL",
    "Source URLs","First Seen","Last Seen","Errors"
]

def _client(json_path: str):
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
    return gspread.authorize(creds)

def open_sheet(sheet_id: str, worksheet_name: str, json_path: str):
    gc = _client(json_path)
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=2000, cols=80)
    # Ensure header
    headers = ws.row_values(1)
    if not headers:
        ws.update("A1", [MANAGED_COLS + ["Notes: Approach strategy","Notes: Contacts","Notes: Collected info"]])
    else:
        missing = [c for c in MANAGED_COLS if c not in headers]
        if missing:
            ws.update("A1", [headers + missing])
    return ws

def _header_map(ws):
    headers = ws.row_values(1)
    return {h: i+1 for i, h in enumerate(headers)}

def _index_by_url(ws):
    hm = _header_map(ws)
    if "Public Website Homepage URL" not in hm: return {}
    col = hm["Public Website Homepage URL"]
    vals = ws.col_values(col)
    out = {}
    for r, v in enumerate(vals, start=1):
        if r == 1: continue
        key = (v or "").strip().lower()
        if key: out[key] = r
    return out

def upsert_record(ws, rec: dict):
    hm = _header_map(ws)
    idx = _index_by_url(ws)

    key = (rec.get("Public Website Homepage URL","") or "").strip().lower()
    if not key:
        raise ValueError("Missing Public Website Homepage URL")

    now = datetime.utcnow().strftime("%Y-%m-%d")
    rec.setdefault("First Seen", now)
    rec["Last Seen"] = now

    rownum = idx.get(key)
    if rownum:
        updates = []
        for col in MANAGED_COLS:
            if col not in hm: continue
            a1 = gspread.utils.rowcol_to_a1(rownum, hm[col])
            updates.append({"range": a1, "values": [[rec.get(col, "")]]})
        if updates:
            ws.batch_update([{"range": u["range"], "values": u["values"]} for u in updates])
        return rownum

    # Append new row with current headers (notes columns left blank)
    headers = ws.row_values(1)
    row = [rec.get(h, "") if h in MANAGED_COLS else "" for h in headers]
    ws.append_row(row, value_input_option="RAW")
    return ws.row_count
