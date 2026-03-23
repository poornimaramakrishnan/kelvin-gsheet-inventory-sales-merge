"""
Restore Staff — Restores staff "data" tabs from backup JSON files.
===================================================================
Author:  Poornima Ramakrishnan
Contact: poornima2489@gmail.com
===================================================================

Reads the most recent backup from the backup/ directory and restores
the "data" tab of each staff sheet.

Environment variables:
  GOOGLE_CREDENTIALS_JSON  — base64-encoded token JSON
  STAFF_SHEETS_JSON        — JSON array of staff sheets
"""

import os
import sys
import json
import time
import base64
import tempfile
import glob
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(SCRIPT_DIR, "..", "backup")

STAFF_SHEETS = json.loads(os.environ.get("STAFF_SHEETS_JSON", "[]"))


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


def find_latest_backup():
    """Find the most recent backup subdirectory."""
    if not os.path.exists(BACKUP_DIR):
        print(f"ERROR: Backup directory not found: {BACKUP_DIR}")
        sys.exit(1)

    subdirs = sorted(glob.glob(os.path.join(BACKUP_DIR, "run_*")), reverse=True)
    if not subdirs:
        print("ERROR: No backup runs found.")
        sys.exit(1)

    return subdirs[0]


creds = get_credentials()
gc = gspread.authorize(creds)
latest = find_latest_backup()
print(f"Using backup: {latest}")

for staff_info in STAFF_SHEETS:
    label = staff_info["label"]
    sid = staff_info["id"]

    backup_path = os.path.join(latest, f"{label}.json")
    if not os.path.exists(backup_path):
        print(f"  ⚠️  No backup file for '{label}', skipping.")
        continue

    with open(backup_path, "r", encoding="utf-8") as f:
        backup = json.load(f)

    if "data" not in backup:
        print(f"  ⚠️  No 'data' tab in backup for '{label}', skipping.")
        continue

    original_data = backup["data"]["data"]
    ss = gc.open_by_key(sid)
    ws = ss.worksheet("data")

    ws.clear()
    time.sleep(0.5)

    if original_data:
        num_rows = len(original_data)
        num_cols = max(len(row) for row in original_data)
        end_col = chr(64 + num_cols) if num_cols <= 26 else 'R'
        rng = f"A1:{end_col}{num_rows}"
        padded = [row + [''] * (num_cols - len(row)) for row in original_data]
        ws.update(padded, rng, value_input_option="RAW")
        time.sleep(1)
        print(f"  ✅ '{label}' → 'data': restored {num_rows - 1} data rows ({num_cols} cols)")

print("\n✅ Staff data tabs restored from backup!")
