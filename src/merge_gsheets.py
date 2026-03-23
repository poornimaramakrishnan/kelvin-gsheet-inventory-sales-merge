"""
Google Sheets Merge Script — GitHub Actions Edition
=====================================================
Author:  Poornima Ramakrishnan
Contact: poornima2489@gmail.com
=====================================================
Designed to run headless via GitHub Actions with credentials
supplied through environment variables / repository secrets.

Fully self-contained with built-in:
  ✅ Automatic pre-merge backup (timestamped JSON snapshots)
  ✅ Plain-English row-by-row logging of every decision
  ✅ Exhaustive post-merge validation (cell-by-cell integrity check)
  ✅ Automatic rollback if validation fails
  ✅ Structured JSON report output for email notifications

Rules:
  • Staff rows are only processed if Column H (Staff) has been filled in.
  • A Manager row is eligible for matching only if G (Deal Date), H (Staff),
    AND Y (Indicator) are ALL empty.
  • Matching is done by Order ID (col Q) + SKU (col F).
  • First-come-first-served: Staff 1 is processed before Staff 2.
  • Matched staff rows → "Matched row" tab.
  • Unmatched / problematic staff rows → "conflict or unavail" tab.
  • Processed rows are deleted from the staff "data" tab.
  • Indicator "updated by <staff label>" written to Manager Column Y.

Environment variables (set via GitHub Secrets):
  GOOGLE_CREDENTIALS_JSON  — base64-encoded Google OAuth2 token JSON
  MANAGER_SHEET_ID         — Google Sheet ID for the Manager sheet
  STAFF_SHEETS_JSON        — JSON array: [{"id": "...", "label": "..."},...]
"""

import os
import sys
import json
import time
import base64
import tempfile
from datetime import datetime
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# ---------------------------------------------------------------------------
# Configuration — from environment variables
# ---------------------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

MANAGER_SHEET_ID = os.environ.get("MANAGER_SHEET_ID", "")
STAFF_SHEETS = json.loads(os.environ.get("STAFF_SHEETS_JSON", "[]"))

ALL_SHEETS = [
    {"id": MANAGER_SHEET_ID, "name": "manager"},
] + [{"id": s["id"], "name": s["label"]} for s in STAFF_SHEETS]

# Column indices (0-based)
COL_A_NAME = 0
COL_B_PROCURE_DATE = 1
COL_C_BUY_PRICE = 2
COL_D_QTY = 3
COL_E_STATUS = 4
COL_F_SKU = 5
COL_G_DEAL_DATE = 6
COL_H_STAFF = 7
COL_Q_ORDER_ID = 16
COL_Y_FLAG = 24

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(SCRIPT_DIR, "..", "backup")

# ---------------------------------------------------------------------------
# Structured report collector
# ---------------------------------------------------------------------------
report_data = {
    "run_timestamp": "",
    "status": "unknown",
    "manager_sheet_id": "",
    "staff_sheets": [],
    "phases": [],
    "staff_results": [],
    "validation": {"passed": False, "checks": [], "errors": []},
    "logs": [],
}


def report_log(msg, level="info"):
    """Append a log entry to the structured report."""
    report_data["logs"].append({"time": datetime.now().isoformat(), "level": level, "msg": msg})


# ---------------------------------------------------------------------------
# Logging helpers
# ---------------------------------------------------------------------------
def banner(title):
    print(f"\n{'=' * 76}")
    print(f"  {title}")
    print(f"{'=' * 76}")
    report_log(f"═══ {title}", "section")


def log(msg, indent=1):
    prefix = "    " * indent
    print(f"{prefix}{msg}")
    report_log(msg)


def log_blank():
    print()


# ---------------------------------------------------------------------------
# Auth — headless via environment variable
# ---------------------------------------------------------------------------
def get_credentials():
    """Load credentials from GOOGLE_CREDENTIALS_JSON env var (base64-encoded token JSON)."""
    creds_b64 = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_b64:
        print("ERROR: GOOGLE_CREDENTIALS_JSON environment variable is not set.")
        print("Set it as a GitHub repository secret containing your base64-encoded token.json.")
        sys.exit(1)

    try:
        token_json = base64.b64decode(creds_b64).decode("utf-8")
    except Exception as e:
        print(f"ERROR: Failed to decode GOOGLE_CREDENTIALS_JSON: {e}")
        sys.exit(1)

    # Write to a temp file for gspread
    token_path = os.path.join(tempfile.gettempdir(), "gsheets_token.json")
    with open(token_path, "w") as f:
        f.write(token_json)

    creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            # Update the token in case it was refreshed
            with open(token_path, "w") as f:
                f.write(creds.to_json())
            # Also update env var so next run uses refreshed token
            refreshed_b64 = base64.b64encode(creds.to_json().encode()).decode()
            report_data["refreshed_token_b64"] = refreshed_b64
        else:
            print("ERROR: Token is invalid and cannot be refreshed.")
            print("You need to re-generate your token.json locally and update the secret.")
            sys.exit(1)

    return creds


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def safe_get(row, idx, default=""):
    return row[idx].strip() if idx < len(row) and row[idx].strip() else default


def api_sleep():
    time.sleep(1.1)


def normalize_row(row, length):
    padded = list(row) + [''] * max(0, length - len(row))
    return [str(c).strip() for c in padded[:length]]


def describe_name(row):
    name = safe_get(row, COL_A_NAME)
    if not name:
        return "(unnamed product)"
    return name[:60] + "..." if len(name) > 60 else name


def describe_row_details(row):
    name = safe_get(row, COL_A_NAME) or "(empty)"
    procure = safe_get(row, COL_B_PROCURE_DATE) or "(empty)"
    price = safe_get(row, COL_C_BUY_PRICE) or "(empty)"
    qty = safe_get(row, COL_D_QTY) or "(empty)"
    status = safe_get(row, COL_E_STATUS) or "(empty)"
    sku = safe_get(row, COL_F_SKU) or "(empty)"
    deal_date = safe_get(row, COL_G_DEAL_DATE) or "(empty)"
    staff = safe_get(row, COL_H_STAFF) or "(empty)"
    oid = safe_get(row, COL_Q_ORDER_ID) or "(empty)"
    lines = [
        f"Product Name : {name}",
        f"SKU          : {sku}",
        f"Order ID     : {oid}",
        f"Procure Date : {procure}",
        f"Buy Price    : {price}",
        f"Qty          : {qty}",
        f"Status       : {status}",
        f"Deal Date (G): {deal_date}",
        f"Staff (H)    : {staff}",
    ]
    return lines


# ---------------------------------------------------------------------------
# PHASE 0 — AUTOMATIC BACKUP
# ---------------------------------------------------------------------------
def phase0_backup(gc):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_subdir = os.path.join(BACKUP_DIR, f"run_{timestamp}")
    os.makedirs(backup_subdir, exist_ok=True)

    banner("PHASE 0 — Taking a safety backup of ALL sheets before we touch anything")
    log("Saving a snapshot of every sheet so we can undo everything if needed.")
    log_blank()

    phase_info = {"phase": 0, "title": "Backup", "sheets_backed_up": []}

    all_backups = {}
    for sheet_info in ALL_SHEETS:
        ss = gc.open_by_key(sheet_info["id"])
        backup = {}
        for ws in ss.worksheets():
            all_vals = ws.get_all_values()
            backup[ws.title] = {
                "data": all_vals,
                "row_count": ws.row_count,
                "col_count": ws.col_count,
            }
            log(f"📸 Saved '{sheet_info['name']}' tab '{ws.title}' — {len(all_vals)} rows backed up.")
            phase_info["sheets_backed_up"].append(
                {"sheet": sheet_info["name"], "tab": ws.title, "rows": len(all_vals)}
            )

        filepath = os.path.join(backup_subdir, f"{sheet_info['name']}.json")
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(backup, f, ensure_ascii=False, indent=2)
        all_backups[sheet_info["id"]] = backup

    log_blank()
    log(f"✅ All backups saved to: {backup_subdir}")
    report_data["phases"].append(phase_info)
    return all_backups, backup_subdir


# ---------------------------------------------------------------------------
# PHASE 1 — PROCESSING ORDER
# ---------------------------------------------------------------------------
def phase1_prescan(gc):
    banner("PHASE 1 — Processing order (first-come-first-served)")
    log("Staff sheets will be processed in this order:")
    for i, s in enumerate(STAFF_SHEETS, 1):
        log(f"  {i}. {s['label']}")
    log_blank()
    log("The first staff sheet to match a Manager row wins.")
    report_data["phases"].append({"phase": 1, "title": "Processing Order",
                                   "order": [s["label"] for s in STAFF_SHEETS]})
    return set()


# ---------------------------------------------------------------------------
# PHASE 2 — MERGE
# ---------------------------------------------------------------------------
def build_manager_index(manager_data):
    index = {}
    eligible_count = 0
    ineligible_count = 0

    for i, row in enumerate(manager_data):
        row_num = i + 2
        oid = safe_get(row, COL_Q_ORDER_ID)
        sku = safe_get(row, COL_F_SKU)
        g_val = safe_get(row, COL_G_DEAL_DATE)
        h_val = safe_get(row, COL_H_STAFF)
        y_val = safe_get(row, COL_Y_FLAG)
        mgr_name = describe_name(row)
        if not oid or not sku:
            continue
        key = (oid, sku)
        if key not in index:
            index[key] = []

        is_eligible = (g_val == "" and h_val == "" and y_val == "")
        index[key].append({
            "row_num": row_num,
            "eligible": is_eligible,
            "g_val": g_val,
            "h_val": h_val,
            "y_val": y_val,
            "name": mgr_name,
        })
        if is_eligible:
            eligible_count += 1
        else:
            ineligible_count += 1

    log(f"Manager index built: {eligible_count} rows available for matching, "
        f"{ineligible_count} already occupied (G, H, or Y filled).")
    return index


def find_and_claim_match(index, order_id, sku):
    key = (order_id, sku)
    if key not in index:
        return None, None, None
    for entry in index[key]:
        if entry["eligible"]:
            entry["eligible"] = False
            return entry["row_num"], entry["name"], None
    reasons = []
    for entry in index[key]:
        parts = []
        if entry["g_val"]:
            parts.append(f"G(Deal Date)='{entry['g_val']}'")
        if entry["h_val"]:
            parts.append(f"H(Staff)='{entry['h_val']}'")
        if entry["y_val"]:
            parts.append(f"Y(Indicator)='{entry['y_val']}'")
        reasons.append(f"Manager Row {entry['row_num']} '{entry['name']}': {', '.join(parts)}")
    return None, None, reasons


def phase2_merge(gc, cross_staff_conflicts):
    banner("PHASE 2 — Merging staff data into the Manager sheet")

    log("Loading the Manager sheet (Sheet1)...")
    manager_ss = gc.open_by_key(MANAGER_SHEET_ID)
    manager_ws = manager_ss.worksheet("Sheet1")
    manager_all = manager_ws.get_all_values()
    manager_data = manager_all[1:]
    log(f"The Manager sheet has {len(manager_data)} data rows and {len(manager_all[0])} columns.")
    log_blank()

    log("Building a lookup of Manager rows to find matches...")
    manager_index = build_manager_index(manager_data)

    merge_record = {
        "manager_flags": [],
        "staff_actions": {},
    }

    for staff_idx, staff_info in enumerate(STAFF_SHEETS, 1):
        sid = staff_info["id"]
        label = staff_info["label"]

        banner(f"Processing Staff {staff_idx}: '{label}'")

        staff_ss = gc.open_by_key(sid)
        staff_data_ws = staff_ss.worksheet("data")
        staff_matched_ws = staff_ss.worksheet("Matched row")
        staff_conflict_ws = staff_ss.worksheet("conflict or unavail")

        staff_all = staff_data_ws.get_all_values()
        staff_headers = staff_all[0]
        staff_data = staff_all[1:]
        num_cols = len(staff_headers) if staff_headers else 17
        log(f"This staff sheet has {len(staff_data)} data rows.")
        log_blank()

        matched_rows = []
        conflict_rows = []
        conflict_reasons = []
        manager_updates = []
        rows_to_delete = []
        skipped_count = 0

        # For report
        staff_report = {
            "label": label,
            "total_rows": len(staff_data),
            "matched": [],
            "conflicts": [],
            "skipped": 0,
        }

        for i, row in enumerate(staff_data):
            row_num = i + 2
            name = describe_name(row)
            h_val = safe_get(row, COL_H_STAFF)

            if not h_val:
                skipped_count += 1
                continue

            oid = safe_get(row, COL_Q_ORDER_ID)
            sku = safe_get(row, COL_F_SKU)

            log(f"── Row {row_num}: {name} ──")
            log(f"Staff row details:", 2)
            for detail in describe_row_details(row):
                log(f"  {detail}", 2)
            log_blank()

            if not oid or not sku:
                reason = "Order ID or SKU is missing — cannot look up on Manager sheet."
                log(f"❌ {reason}", 2)
                log(f"→ Moving to 'conflict or unavail'.", 2)
                conflict_rows.append(row)
                conflict_reasons.append(reason)
                rows_to_delete.append(row_num)
                staff_report["conflicts"].append({
                    "row": row_num, "name": name, "sku": sku or "(empty)",
                    "order_id": oid or "(empty)", "reason": reason,
                })
                log_blank()
                continue

            match_row_num, mgr_name, fail_reasons = find_and_claim_match(manager_index, oid, sku)

            if match_row_num is not None:
                flag_text = f"updated by {label}"
                log(f"✅ MATCH FOUND!", 2)
                log(f"Staff '{name}' (Order ID: {oid}, SKU: {sku})", 2)
                log(f"→ Manager Row {match_row_num} '{mgr_name}'", 2)
                log_blank()
                log(f"Actions:", 2)
                log(f"  1. Write '{flag_text}' → Manager Row {match_row_num}, Col Y", 2)
                log(f"  2. Copy to 'Matched row' tab", 2)
                log(f"  3. Delete from 'data' tab", 2)
                manager_updates.append((match_row_num, flag_text))
                matched_rows.append(row)
                rows_to_delete.append(row_num)
                staff_report["matched"].append({
                    "staff_row": row_num, "name": name, "sku": sku,
                    "order_id": oid, "mgr_row": match_row_num, "mgr_name": mgr_name,
                    "staff_name": h_val,
                })
            else:
                if fail_reasons:
                    log(f"❌ NO MATCH — all Manager rows for this key are taken:", 2)
                    for fr in fail_reasons:
                        log(f"  • {fr}", 2)
                    reason = "All matching Manager rows are already occupied."
                else:
                    log(f"❌ NO MATCH — no Manager row with Order ID '{oid}' + SKU '{sku}'.", 2)
                    reason = f"No Manager row with Order ID '{oid}' + SKU '{sku}'."
                log_blank()
                log(f"→ Moving to 'conflict or unavail'.", 2)
                conflict_rows.append(row)
                conflict_reasons.append(reason)
                rows_to_delete.append(row_num)
                staff_report["conflicts"].append({
                    "row": row_num, "name": name, "sku": sku,
                    "order_id": oid, "reason": reason,
                })
            log_blank()

        staff_report["skipped"] = skipped_count
        report_data["staff_results"].append(staff_report)

        # Summary
        banner(f"Summary for '{label}'")
        log(f"Total rows:      {len(staff_data)}")
        log(f"Staff filled H:  {len(matched_rows) + len(conflict_rows)}")
        log(f"  ✅ Matched:    {len(matched_rows)}")
        log(f"  ❌ Conflicts:  {len(conflict_rows)}")
        log(f"  ⏭  Skipped:   {skipped_count}")
        log_blank()

        merge_record["staff_actions"][sid] = {
            "label": label,
            "original_data_count": len(staff_data),
            "num_cols": num_cols,
            "matched_rows": [row[:num_cols] for row in matched_rows],
            "conflict_rows": [row[:num_cols] for row in conflict_rows],
            "skipped_count": skipped_count,
        }

        # ---- Apply writes ----
        if manager_updates:
            log(f"Writing Manager Column Y for {len(manager_updates)} matched row(s)...")
            cells = [gspread.Cell(row=r, col=COL_Y_FLAG + 1, value=f) for r, f in manager_updates]
            manager_ws.update_cells(cells)
            merge_record["manager_flags"].extend(manager_updates)
            api_sleep()
            log("✅ Manager Column Y updated.")

        if matched_rows:
            trimmed = [row[:num_cols] for row in matched_rows]
            existing = staff_matched_ws.get_all_values()
            start = len(existing) + 1
            end = start + len(trimmed) - 1
            end_col = chr(64 + num_cols) if num_cols <= 26 else 'R'
            rng = f"A{start}:{end_col}{end}"
            staff_matched_ws.update(trimmed, rng, value_input_option="RAW")
            api_sleep()
            log(f"✅ {len(matched_rows)} row(s) → 'Matched row' tab ({rng}).")

        if conflict_rows:
            trimmed = [row[:num_cols] for row in conflict_rows]
            existing = staff_conflict_ws.get_all_values()
            start = len(existing) + 1
            end = start + len(trimmed) - 1
            end_col = chr(64 + num_cols) if num_cols <= 26 else 'R'
            rng = f"A{start}:{end_col}{end}"
            staff_conflict_ws.update(trimmed, rng, value_input_option="RAW")
            api_sleep()
            log(f"✅ {len(conflict_rows)} row(s) → 'conflict or unavail' tab ({rng}).")

        if rows_to_delete:
            log(f"Deleting {len(rows_to_delete)} row(s) from 'data' tab (bottom-up)...")
            rows_to_delete.sort(reverse=True)
            for rn in rows_to_delete:
                staff_data_ws.delete_rows(rn)
                api_sleep()
            remaining = len(staff_data) - len(rows_to_delete)
            log(f"✅ 'data' tab: {remaining} rows remaining (was {len(staff_data)}).")

    return merge_record


# ---------------------------------------------------------------------------
# PHASE 3 — EXHAUSTIVE VALIDATION
# ---------------------------------------------------------------------------
def phase3_validate(gc, merge_record):
    banner("PHASE 3 — Verifying everything by re-reading the live sheets")
    log("Reading all sheets back from Google to verify every row landed correctly.")

    errors = []
    checks = []

    # 3A. Manager Column Y
    log_blank()
    log("[Check 1] Verifying Manager Column Y indicators...")
    mgr_ws = gc.open_by_key(MANAGER_SHEET_ID).worksheet("Sheet1")
    mgr_live = mgr_ws.get_all_values()

    for row_num, expected_flag in merge_record["manager_flags"]:
        if row_num > len(mgr_live):
            errors.append(f"Manager row {row_num} missing (sheet has {len(mgr_live)} rows)")
            continue
        live_row = mgr_live[row_num - 1]
        actual_y = safe_get(live_row, COL_Y_FLAG)
        if actual_y != expected_flag:
            errors.append(f"Manager Row {row_num} Col Y: expected '{expected_flag}', found '{actual_y}'")

    if not any("Manager" in e for e in errors):
        msg = f"✅ All {len(merge_record['manager_flags'])} Manager Column Y indicators correct."
        log(msg)
        checks.append(msg)

    # 3B. Per-staff validation
    for sid, actions in merge_record["staff_actions"].items():
        label = actions["label"]
        num_cols = actions["num_cols"]
        original_count = actions["original_data_count"]
        matched = actions["matched_rows"]
        conflicts = actions["conflict_rows"]

        log_blank()
        log(f"[Check 2] Verifying '{label}'...")

        staff_ss = gc.open_by_key(sid)

        data_ws = staff_ss.worksheet("data")
        data_live = data_ws.get_all_values()
        remaining = len(data_live) - 1

        expected_remaining = original_count - len(matched) - len(conflicts)
        if remaining != expected_remaining:
            errors.append(
                f"[{label}] Row count mismatch: expected {expected_remaining} remaining, found {remaining}."
            )
        else:
            msg = f"✅ [{label}] Row conservation: {original_count} = {len(matched)} matched + {len(conflicts)} conflict + {remaining} remaining."
            log(msg)
            checks.append(msg)

        if matched:
            matched_ws = staff_ss.worksheet("Matched row")
            matched_live = matched_ws.get_all_values()[1:]
            for i, expected_row in enumerate(matched):
                norm_exp = normalize_row(expected_row, num_cols)
                found = any(normalize_row(lr, num_cols) == norm_exp for lr in matched_live)
                if not found:
                    oid = safe_get(expected_row, COL_Q_ORDER_ID)
                    sku = safe_get(expected_row, COL_F_SKU)
                    errors.append(f"[{label}] Matched row #{i+1} (OID={oid}, SKU={sku}) NOT in 'Matched row' tab!")
            if not any(f"[{label}] Matched" in e for e in errors):
                msg = f"✅ [{label}] All {len(matched)} matched row(s) confirmed in 'Matched row' tab."
                log(msg)
                checks.append(msg)

        if conflicts:
            conflict_ws = staff_ss.worksheet("conflict or unavail")
            conflict_live = conflict_ws.get_all_values()[1:]
            for i, expected_row in enumerate(conflicts):
                norm_exp = normalize_row(expected_row, num_cols)
                found = any(normalize_row(lr, num_cols) == norm_exp for lr in conflict_live)
                if not found:
                    oid = safe_get(expected_row, COL_Q_ORDER_ID)
                    sku = safe_get(expected_row, COL_F_SKU)
                    errors.append(f"[{label}] Conflict row #{i+1} (OID={oid}, SKU={sku}) NOT in 'conflict or unavail' tab!")
            if not any(f"[{label}] Conflict" in e for e in errors):
                msg = f"✅ [{label}] All {len(conflicts)} conflict row(s) confirmed in 'conflict or unavail' tab."
                log(msg)
                checks.append(msg)

        # Orphan check
        data_remaining = data_live[1:]
        all_processed = matched + conflicts
        orphans = 0
        for proc_row in all_processed:
            norm_proc = normalize_row(proc_row, num_cols)
            for data_row in data_remaining:
                if normalize_row(data_row, num_cols) == norm_proc:
                    oid = safe_get(proc_row, COL_Q_ORDER_ID)
                    sku = safe_get(proc_row, COL_F_SKU)
                    errors.append(f"[{label}] ORPHAN: (OID={oid}, SKU={sku}) still in 'data' tab!")
                    orphans += 1
                    break
        if orphans == 0:
            msg = f"✅ [{label}] No orphaned rows in 'data' tab."
            log(msg)
            checks.append(msg)

    # Final verdict
    log_blank()
    report_data["validation"]["checks"] = checks
    report_data["validation"]["errors"] = errors

    if errors:
        log(f"❌ VALIDATION FAILED — {len(errors)} problem(s) found:")
        for e in errors:
            log(f"   • {e}")
        report_data["validation"]["passed"] = False
        return False, errors
    else:
        log(f"✅ ALL CHECKS PASSED — every row is in the right place.")
        report_data["validation"]["passed"] = True
        return True, []


# ---------------------------------------------------------------------------
# PHASE 4 — AUTOMATIC ROLLBACK
# ---------------------------------------------------------------------------
def phase4_rollback(gc, backup_subdir):
    banner("PHASE 4 — AUTOMATIC ROLLBACK (undoing all changes)")
    log("Restoring every sheet to pre-run state from backup.")
    log_blank()

    for sheet_info in ALL_SHEETS:
        backup_path = os.path.join(backup_subdir, f"{sheet_info['name']}.json")
        if not os.path.exists(backup_path):
            log(f"❌ Backup file missing for '{sheet_info['name']}'!")
            continue

        with open(backup_path, "r", encoding="utf-8") as f:
            backup = json.load(f)

        ss = gc.open_by_key(sheet_info["id"])

        for tab_name, tab_data in backup.items():
            ws = ss.worksheet(tab_name)
            original = tab_data["data"]
            ws.clear()
            api_sleep()
            if original:
                num_rows = len(original)
                num_cols = max(len(r) for r in original)
                end_col = chr(64 + num_cols) if num_cols <= 26 else 'R'
                rng = f"A1:{end_col}{num_rows}"
                padded = [r + [''] * (num_cols - len(r)) for r in original]
                ws.update(padded, rng, value_input_option="RAW")
                api_sleep()
            log(f"✅ Restored '{sheet_info['name']}' tab '{tab_name}'.")

    log_blank()
    log("✅ ROLLBACK COMPLETE — all sheets restored.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    run_start = datetime.now()
    report_data["run_timestamp"] = run_start.isoformat()
    report_data["manager_sheet_id"] = MANAGER_SHEET_ID
    report_data["staff_sheets"] = [{"id": s["id"], "label": s["label"]} for s in STAFF_SHEETS]

    banner("Google Sheets Merge Script — GitHub Actions Edition")
    log("Author: Poornima Ramakrishnan (poornima2489@gmail.com)")
    log_blank()
    log("This script will:")
    log("  1. Back up all sheets")
    log("  2. Process staff rows where Column H is filled")
    log("  3. Match by Order ID + SKU to Manager sheet")
    log("  4. Move matched → 'Matched row', conflicts → 'conflict or unavail'")
    log("  5. Write indicators to Manager Column Y")
    log("  6. Delete processed rows from staff 'data' tab")
    log("  7. Re-read & validate every cell")
    log("  8. Auto-rollback if validation fails")
    log_blank()

    if not MANAGER_SHEET_ID:
        print("ERROR: MANAGER_SHEET_ID environment variable not set.")
        sys.exit(1)
    if not STAFF_SHEETS:
        print("ERROR: STAFF_SHEETS_JSON environment variable not set or empty.")
        sys.exit(1)

    creds = get_credentials()
    gc = gspread.authorize(creds)

    # PHASE 0
    pre_backup, backup_subdir = phase0_backup(gc)

    # PHASE 1
    cross_staff_conflicts = phase1_prescan(gc)

    # PHASE 2
    merge_record = phase2_merge(gc, cross_staff_conflicts)

    # PHASE 3
    passed, errors = phase3_validate(gc, merge_record)

    run_end = datetime.now()
    report_data["duration_seconds"] = (run_end - run_start).total_seconds()

    if not passed:
        phase4_rollback(gc, backup_subdir)
        report_data["status"] = "ROLLED_BACK"
        banner("❌ MERGE ABORTED & ROLLED BACK")
        log("All sheets restored to original state.")
    else:
        report_data["status"] = "SUCCESS"
        banner("✅ MERGE COMPLETED SUCCESSFULLY")
        log("All staff rows processed, all validations passed.")
        log(f"Backup preserved at: {backup_subdir}")

    # Write structured report JSON for the report generator
    report_path = os.path.join(SCRIPT_DIR, "..", "merge_report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report_data, f, ensure_ascii=False, indent=2)
    log(f"Report data saved to: {report_path}")

    if not passed:
        sys.exit(1)


if __name__ == "__main__":
    main()
