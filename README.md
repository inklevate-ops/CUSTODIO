# 8888 Complaint Monitoring and Reporting System

Google Apps Script web app for DOLE HRDS. Uses **two Google Sheets** (`Users`, `Complaints`) as the database and **Google Drive** for document storage.

## Files to copy into Apps Script

| File | Purpose |
|---|---|
| `Code.gs` | Server logic, login, workflow, AI extraction, reports |
| `Index.html` | Login, dashboard, and all modules |

## Setup

1. Open [script.google.com](https://script.google.com) and create a project.
2. Paste `Code.gs` and `Index.html` (File name must be **Index**).
3. In **Services**, add **Drive API**.
4. Run `setupSystem`. Authorize Sheets, Drive, and Docs.
5. Deploy → **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Open the web app URL.

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
