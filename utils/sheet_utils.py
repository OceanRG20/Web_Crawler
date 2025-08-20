import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Columns the crawler manages — NO 'Errors' anymore.
MANAGED_COLS = [
    "Company Name","Target Status","Public Website Homepage URL","Domain","Source",
    "Street Address","City","State","Zipcode","Phone",
    "Industries served","Products and services offered","Specific references from text search",
    "Square footage (facility)","Number of employees","Estimated Revenues","Years of operation",
    "Ownership","Equipment","CNC 3-axis","CNC 5-axis","Spares/Repairs","Family business","Jobs",
    "Year Evidence URL","Year Evidence Snippet","Owner Evidence URL",
    "Source URLs","First Seen","Last Seen","Last Update"
]

NOTES_COL = "Notes (Approach/Contacts/Info)"

def _client(json_path: str):
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
    return gspread.authorize(creds)

def _sanitize_sheet_id(sheet_id: str) -> str:
    if not sheet_id:
        return sheet_id
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_id)
    return m.group(1) if m else sheet_id

def open_sheet(sheet_id: str, worksheet_name: str, json_path: str):
    gc = _client(json_path)
    sh = gc.open_by_key(_sanitize_sheet_id(sheet_id))
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=4000, cols=120)

    headers = ws.row_values(1)
    if not headers:
        ws.update("A1", [MANAGED_COLS + [NOTES_COL]])
    else:
        missing = [c for c in MANAGED_COLS if c not in headers]
        if NOTES_COL not in headers:
            missing.append(NOTES_COL)
        if missing:
            ws.update("A1", [headers + missing])
    return ws

def _header_map(ws):
    headers = ws.row_values(1)
    return {h: i+1 for i, h in enumerate(headers)}

def _index_by_url(ws):
    hm = _header_map(ws)
    if "Public Website Homepage URL" not in hm:
        return {}
    col = hm["Public Website Homepage URL"]
    vals = ws.col_values(col)
    out = {}
    for r, v in enumerate(vals, start=1):
        if r == 1: continue
        key = (v or "").strip().lower()
        if key: out[key] = r
    return out

def _row_dict(ws, rownum: int) -> dict:
    headers = ws.row_values(1)
    values = ws.row_values(rownum)
    values += [""] * (len(headers) - len(values))
    return {h: values[i] if i < len(values) else "" for i, h in enumerate(headers)}

def upsert_record(ws, rec: dict):
    """
    Merge by Homepage URL.
    - Set First Seen on insert
    - Always refresh Last Seen
    - Touch Last Update only if any managed column (except First/Last/Update) changed
    - Never write the Notes column (human-only)
    """
    hm = _header_map(ws)
    idx = _index_by_url(ws)

    key = (rec.get("Public Website Homepage URL","") or "").strip().lower()
    if not key:
        raise ValueError("Missing Public Website Homepage URL")

    today = datetime.utcnow().strftime("%Y-%m-%d")
    rec.setdefault("First Seen", today)
    rec["Last Seen"] = today

    rownum = idx.get(key)
    if rownum:
        current = _row_dict(ws, rownum)
        changed = False
        compare_cols = [c for c in MANAGED_COLS if c not in ("First Seen","Last Seen","Last Update")]

        updates = []
        for col in MANAGED_COLS:
            if col not in hm: 
                continue
            new_val = rec.get(col, "")
            if col in compare_cols and (current.get(col,"") or "") != (new_val or ""):
                changed = True
            a1 = gspread.utils.rowcol_to_a1(rownum, hm[col])
            updates.append({"range": a1, "values": [[new_val]]})

        if "Last Update" in hm:
            a1_upd = gspread.utils.rowcol_to_a1(rownum, hm["Last Update"])
            updates.append({"range": a1_upd, "values": [[today if changed else current.get("Last Update","")]]})

        if updates:
            ws.batch_update([{"range": u["range"], "values": u["values"]} for u in updates])
        return rownum

    # Append new row with current headers (Notes left blank)
    headers = ws.row_values(1)
    row = [rec.get(h, "") if h in MANAGED_COLS else "" for h in headers]
    ws.append_row(row, value_input_option="RAW")
    return ws.row_count
