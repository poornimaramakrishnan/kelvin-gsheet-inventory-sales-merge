# 📊 SheetMerge — Google Sheets Automated Merge

<div align="center">

![Python](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white)
![GitHub Actions](https://img.shields.io/badge/GitHub_Actions-Automated-2088FF?logo=github-actions&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google_Sheets-API_v4-34A853?logo=google-sheets&logoColor=white)
![License](https://img.shields.io/badge/License-Private-red)

**Merge staff Google Sheets into a Manager sheet with one click from your browser.**
No Python install needed. Beautiful HTML email report after every run.

</div>

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🔄 **One-Click Merge** | Run from GitHub's UI — just click "Run workflow" |
| 📊 **Smart Matching** | Matches by Order ID + SKU with eligibility checks |
| 🛡️ **Foolproof Safety** | Auto-backup → Merge → Validate → Rollback if anything fails |
| 📧 **Email Reports** | Beautiful HTML report emailed after every run |
| 📎 **Artifact Storage** | Reports & backups saved as GitHub Action artifacts |
| 🔑 **Secure** | Credentials stored as encrypted GitHub Secrets |
| ⏰ **Schedulable** | Optional cron schedule for automatic daily runs |

---

## 🏗️ How It Works

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│  Staff Sheet  │     │  Staff Sheet  │     │   Manager    │
│     #1        │     │     #2        │     │    Sheet     │
└──────┬───────┘     └──────┬───────┘     └──────┬───────┘
       │                     │                     │
       └─────────┬───────────┘                     │
                 ▼                                 ▼
        ┌────────────────────────────────────────────┐
        │           SheetMerge Engine                 │
        │  ┌─────────┐ ┌─────────┐ ┌──────────────┐ │
        │  │ Phase 0  │ │ Phase 2 │ │   Phase 3    │ │
        │  │ Backup   │→│ Merge   │→│  Validate    │ │
        │  └─────────┘ └─────────┘ └──────┬───────┘ │
        │                                  │ fail?   │
        │                          ┌───────▼──────┐  │
        │                          │   Phase 4    │  │
        │                          │  Rollback    │  │
        │                          └──────────────┘  │
        └────────────────────────────────────────────┘
                           │
                           ▼
                 ┌───────────────────┐
                 │  📧 HTML Report   │
                 │  via Gmail SMTP   │
                 └───────────────────┘
```

---

## 🚀 Setup Guide

### Step 1: Fork / Clone This Repo

```bash
git clone https://github.com/YOUR_USERNAME/sheetmerge.git
```

### Step 2: Generate Your Google OAuth Token

You need a `token.json` file. Generate it locally once:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project → Enable **Google Sheets API** and **Google Drive API**
3. Create **OAuth 2.0 Client ID** (Desktop app type)
4. Download the `credentials.json` file
5. Run locally to generate `token.json`:
   ```bash
   pip install gspread google-auth google-auth-oauthlib
   python -c "
   from google_auth_oauthlib.flow import InstalledAppFlow
   flow = InstalledAppFlow.from_client_secrets_file('credentials.json', [
       'https://www.googleapis.com/auth/spreadsheets',
       'https://www.googleapis.com/auth/drive'
   ])
   creds = flow.run_local_server(port=0)
   with open('token.json', 'w') as f:
       f.write(creds.to_json())
   print('✅ token.json created!')
   "
   ```
6. Base64-encode the token for GitHub Secrets:
   ```bash
   # Linux/Mac:
   base64 -w 0 token.json

   # Windows PowerShell:
   [Convert]::ToBase64String([IO.File]::ReadAllBytes("token.json"))
   ```

### Step 3: Configure GitHub Secrets

Go to **Repository → Settings → Secrets and variables → Actions** and add:

| Secret Name | Value |
|------------|-------|
| `GOOGLE_CREDENTIALS_JSON` | Base64-encoded content of `token.json` |
| `GMAIL_USER` | Your Gmail address (e.g. `you@gmail.com`) |
| `GMAIL_APP_PASSWORD` | Gmail App Password ([generate here](https://myaccount.google.com/apppasswords)) |
| `REPORT_RECIPIENTS` | *(optional)* Default email recipients, comma-separated |

### Step 4: Run It! 🎉

1. Go to **Actions** tab in your repo
2. Click **"🔄 Sheet Merge"** workflow
3. Click **"Run workflow"**
4. Fill in:
   - **Manager Sheet ID** — your Google Sheet ID
   - **Staff Sheets JSON** — `[{"id":"SHEET_ID","label":"Staff Name"},...]`
   - **Email recipients** — comma-separated addresses
5. Click **"Run workflow"** ✅

---

## 📁 Project Structure

```
sheetmerge/
├── .github/
│   └── workflows/
│       └── merge.yml           # GitHub Actions workflow
├── src/
│   ├── merge_gsheets.py        # 🔄 Core merge engine (4-phase)
│   ├── report_generator.py     # 📊 HTML report builder
│   ├── send_report.py          # 📧 Gmail SMTP sender
│   ├── full_reset.py           # 🔄 Reset merge results
│   └── restore_staff.py        # ♻️ Restore staff data from backup
├── templates/
│   └── report.html             # 🎨 Jinja2 HTML email template
├── .gitignore
├── requirements.txt
└── README.md
```

---

## 📧 Email Report Preview

The HTML report includes:

- ✅/❌ Status banner with gradient header
- 📊 Aggregate stats cards (matched, conflicts, skipped)
- 📋 Per-staff breakdown tables
- 🔍 Validation checklist
- 🕐 Run timestamp and duration
- 📎 Also saved as GitHub Actions artifact for 30 days

---

## 🔧 Column Mapping

| Column | Index | Purpose |
|--------|-------|---------|
| A | 0 | Product Name |
| B | 1 | Procure Date |
| C | 2 | Buy Price |
| D | 3 | Quantity |
| E | 4 | Status |
| F | 5 | **SKU** (matching key) |
| G | 6 | Deal Date (eligibility check) |
| H | 7 | Staff Name (eligibility check) |
| Q | 16 | **Order ID** (matching key) |
| Y | 24 | Indicator flag ("updated by ...") |

**Eligibility Rule:** A Manager row is eligible only if **G**, **H**, and **Y** are **all empty**.

---

## 🔒 Security

- ✅ No credentials stored in the repository
- ✅ All secrets encrypted via GitHub's secret store
- ✅ Token is refreshed automatically on each run
- ✅ Temporary files cleaned up after each workflow run

---

## 👩‍💻 Author

**Poornima Ramakrishnan**
📧 [poornima2489@gmail.com](mailto:poornima2489@gmail.com)

---

<div align="center">
<sub>Built with ❤️ using Python, gspread, and GitHub Actions</sub>
</div>
