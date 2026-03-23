"""
Email Sender — Sends the HTML merge report via Gmail SMTP.
============================================================
Author:  Poornima Ramakrishnan
Contact: poornima2489@gmail.com
============================================================

Uses Gmail App Password authentication (not regular password).
Set these environment variables (GitHub Secrets):
  GMAIL_USER         — your Gmail address (e.g. user@gmail.com)
  GMAIL_APP_PASSWORD — 16-char app password from Google Account settings
  REPORT_RECIPIENTS  — comma-separated email addresses to send the report to
"""

import os
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime


def send_report():
    gmail_user = os.environ.get("GMAIL_USER", "")
    gmail_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    recipients_str = os.environ.get("REPORT_RECIPIENTS", "")

    if not gmail_user or not gmail_password:
        print("⚠️  GMAIL_USER or GMAIL_APP_PASSWORD not set. Skipping email.")
        print("   Set these as GitHub repository secrets to enable email reports.")
        return False

    if not recipients_str:
        print("⚠️  REPORT_RECIPIENTS not set. Skipping email.")
        return False

    recipients = [r.strip() for r in recipients_str.split(",") if r.strip()]

    # Read the HTML report
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(script_dir)
    report_path = os.path.join(root_dir, "merge_report.html")

    if not os.path.exists(report_path):
        print(f"ERROR: {report_path} not found. Run report_generator.py first.")
        return False

    with open(report_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # Determine subject line from report
    import json
    json_path = os.path.join(root_dir, "merge_report.json")
    subject = "Sheet Merge Report"
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        status = data.get("status", "")
        ts = datetime.now().strftime("%b %d, %Y %I:%M %p")
        if status == "SUCCESS":
            total_m = sum(len(s.get("matched", [])) for s in data.get("staff_results", []))
            total_c = sum(len(s.get("conflicts", [])) for s in data.get("staff_results", []))
            subject = f"✅ Sheet Merge Complete — {total_m} matched, {total_c} conflicts ({ts})"
        elif status == "ROLLED_BACK":
            subject = f"⚠️ Sheet Merge Rolled Back — Validation Failed ({ts})"
        else:
            subject = f"❌ Sheet Merge Failed ({ts})"

    # Build email
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"SheetMerge Bot <{gmail_user}>"
    msg["To"] = ", ".join(recipients)

    # Plain-text fallback
    plain_text = f"""\
Sheet Merge Report
==================
Status: {data.get('status', 'Unknown')}
Date: {ts}

This email contains an HTML report. Please view it in an HTML-capable email client.

---
Powered by SheetMerge
Built by Poornima Ramakrishnan (poornima2489@gmail.com)
"""

    msg.attach(MIMEText(plain_text, "plain"))
    msg.attach(MIMEText(html_content, "html"))

    # Send via Gmail SMTP
    try:
        print(f"📧 Sending report to: {', '.join(recipients)}...")
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail_user, gmail_password)
            server.sendmail(gmail_user, recipients, msg.as_string())
        print(f"✅ Email sent successfully!")
        return True
    except smtplib.SMTPAuthenticationError:
        print("❌ Gmail authentication failed.")
        print("   Make sure GMAIL_APP_PASSWORD is a valid 16-character App Password.")
        print("   Generate one at: https://myaccount.google.com/apppasswords")
        return False
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
        return False


if __name__ == "__main__":
    success = send_report()
    sys.exit(0 if success else 1)
