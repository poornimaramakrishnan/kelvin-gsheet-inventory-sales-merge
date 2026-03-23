"""
Full Reset — Clears merge results for a fresh re-run.
=======================================================
Author:  Poornima Ramakrishnan
Contact: poornima2489@gmail.com
=======================================================

Clears:
  • Manager Column Y indicators
  • Staff "Matched row" tabs (reset to clean headers)
  • Staff "conflict or unavail" tabs (reset to clean headers)

Environment variables:
  GOOGLE_CREDENTIALS_JSON  — base64-encoded token JSON
  MANAGER_SHEET_ID         — Manager sheet ID
  STAFF_SHEETS_JSON        — JSON array of staff sheets
"""

import os
import sys
import json
import time
import base64
import tempfile
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

MANAGER_SHEET_ID = os.environ.get("MANAGER_SHEET_ID", "")
STAFF_SHEETS = json.loads(os.environ.get("STAFF_SHEETS_JSON", "[]"))
STAFF_IDS = [s["id"] for s in STAFF_SHEETS]

CLEAN_HEADERS = [['Name', 'Procure date', 'Buy price', 'Qty', 'Status', ' ',
                   'Deal Date', 'Staff', '', '', '', '', '', '', '', '', 'Note']]


def get_credentials():
    creds_b64 = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_b64:
        print("ERROR: GOOGLE_CREDENTIALS_JSON not set.")
        sys.exit(1)
    token_json = base64.b64decode(creds_b64).decode("utf-8")
    token_path = os.path.join(tempfile.gettempdir(), "gsheets_token.json")
    with open(token_path, "w") as f:
        f.write(token_json)
    creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print("ERROR: Token invalid and cannot be refreshed.")
            sys.exit(1)
    return creds


creds = get_credentials()
gc = gspread.authorize(creds)

# 1. Clear Manager Column Y
print("Clearing Manager Sheet1 Column Y...")
mgr_ws = gc.open_by_key(MANAGER_SHEET_ID).worksheet("Sheet1")
all_vals = mgr_ws.get_all_values()
cells = []
for i, row in enumerate(all_vals[1:], start=2):
    if len(row) > 24 and row[24].strip():
        cells.append(gspread.Cell(row=i, col=25, value=""))
if cells:
    mgr_ws.update_cells(cells)
    print(f"  Cleared {len(cells)} Column Y cells.")
time.sleep(1.1)

# 2. Reset Matched row / conflict tabs
for sid in STAFF_IDS:
    ss = gc.open_by_key(sid)
    for tab_name in ["Matched row", "conflict or unavail"]:
        ws = ss.worksheet(tab_name)
        ws.clear()
        time.sleep(0.5)
        ws.update('A1:Q1', CLEAN_HEADERS)
        time.sleep(0.5)
        print(f"  {ss.title} → '{tab_name}': reset with clean headers.")

print("\n✅ Full reset complete!")
