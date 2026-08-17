# 8888 Complaint Monitoring and Reporting System

Google Apps Script web app for DOLE HRDS. Uses **two Google Sheets** (`Users`, `Complaints`) as the database and **Google Drive** for document storage.

## Files to copy into Apps Script

| File | Purpose |
|---|---|
| `Code.gs` | Server logic, login, workflow, AI extraction, reports |
| `Index.html` | Login, dashboard, and all modules |

## Setup

1. Open [script.google.com](https://script.google.com) and create a project.
2. Paste `Code.gs`.
3. Add an HTML file named exactly **Index** (the editor will list it as Index.html). Paste the contents of this repo’s `Index.html` into that file. Do not name it `Index.html.html`.
4. In **Services**, add **Drive API**.
5. Run `setupSystem`. Authorize Sheets, Drive, and Docs.
6. Deploy → **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
7. After any code change, open **Deploy → Manage deployments → pencil → New version → Deploy**. The `/exec` link does not update until you publish a new version.
8. Open the web app URL. If the page is blank, hard-refresh (`Ctrl+F5`) or try an incognito window.

`setupSystem` creates:

- Spreadsheet **8888 Complaint Monitoring Database** with sheets **Users** and **Complaints**
- Drive folder **8888 Complaint Documents** (one subfolder per ticket number)

## Default accounts

| Username | Password | Role |
|---|---|---|
| `admin` | `Admin@8888` | Administrator |
| `encoder` | `Encoder@8888` | Encoder |
| `viewer` | `Viewer@8888` | Viewer |

Change these passwords after first login.

## Workflow

Central Office referral → Receiving / encoding → RD signed Action Slip → Endorsement to Field Office / Unit → 24-hour reply monitoring → Forward action taken to HRDS → Monthly reports.

The 8888 Reference/Ticket Number is the primary identifier. HRDS records final action taken only (no Central Office submission tracking).
