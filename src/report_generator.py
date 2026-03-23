"""
Report Generator — Converts merge_report.json into a beautiful HTML email.
============================================================================
Author:  Poornima Ramakrishnan
Contact: poornima2489@gmail.com
============================================================================

Reads the structured JSON output from merge_gsheets.py and renders it
through a Jinja2 HTML template into a self-contained email-ready HTML file.
"""

import os
import sys
import json
from datetime import datetime
from jinja2 import Environment, FileSystemLoader

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SCRIPT_DIR)
TEMPLATE_DIR = os.path.join(ROOT_DIR, "templates")
REPORT_JSON = os.path.join(ROOT_DIR, "merge_report.json")
OUTPUT_HTML = os.path.join(ROOT_DIR, "merge_report.html")


def format_duration(seconds):
    """Human-friendly duration string."""
    if seconds < 60:
        return f"{seconds:.0f}s"
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    return f"{minutes}m {secs}s"


def generate_report():
    # Load report data
    if not os.path.exists(REPORT_JSON):
        print(f"ERROR: {REPORT_JSON} not found. Run merge_gsheets.py first.")
        sys.exit(1)

    with open(REPORT_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Compute aggregates
    total_matched = 0
    total_conflicts = 0
    total_skipped = 0
    for sr in data.get("staff_results", []):
        total_matched += len(sr.get("matched", []))
        total_conflicts += len(sr.get("conflicts", []))
        total_skipped += sr.get("skipped", 0)
    total_processed = total_matched + total_conflicts

    # Determine hero styling
    status = data.get("status", "unknown")
    if status == "SUCCESS":
        hero_class = "success"
        hero_icon = "✅"
        hero_title = "Merge Completed Successfully"
        hero_subtitle = f"{total_matched} rows matched · {total_conflicts} conflicts · All validations passed"
    elif status == "ROLLED_BACK":
        hero_class = "rolled-back"
        hero_icon = "⚠️"
        hero_title = "Merge Rolled Back"
        hero_subtitle = "Validation failed — all sheets restored to original state"
    else:
        hero_class = "failure"
        hero_icon = "❌"
        hero_title = "Merge Failed"
        hero_subtitle = "An error occurred during the merge process"

    # Format date
    try:
        dt = datetime.fromisoformat(data["run_timestamp"])
        run_date = dt.strftime("%b %d, %Y at %I:%M %p")
    except Exception:
        run_date = data.get("run_timestamp", "Unknown")

    duration = format_duration(data.get("duration_seconds", 0))

    # Render template
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=False)
    template = env.get_template("report.html")

    html = template.render(
        hero_class=hero_class,
        hero_icon=hero_icon,
        hero_title=hero_title,
        hero_subtitle=hero_subtitle,
        run_date=run_date,
        duration=duration,
        num_staff=len(data.get("staff_results", [])),
        total_processed=total_processed,
        total_matched=total_matched,
        total_conflicts=total_conflicts,
        total_skipped=total_skipped,
        staff_results=data.get("staff_results", []),
        validation_checks=data.get("validation", {}).get("checks", []),
        validation_errors=data.get("validation", {}).get("errors", []),
    )

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ HTML report generated: {OUTPUT_HTML}")
    print(f"   Status: {hero_title}")
    print(f"   Matched: {total_matched} | Conflicts: {total_conflicts} | Skipped: {total_skipped}")
    return OUTPUT_HTML


if __name__ == "__main__":
    generate_report()
