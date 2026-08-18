# CUSTODIO

CUSTODIO SUPLICO FIREWORKS — a Google Apps Script ERP for managing clients/shows,
products & inventory, employees & payroll, loans, payments, expenses, and reporting
for a fireworks business.

## Architecture

```
Google Sheet  --->  Apps Script (Code.gs)  --->  Index.html (frontend)
```

The frontend is a single-page app served by `doGet()` and talks to the backend
exclusively through `google.script.run` calls defined in `Code.gs`. There is no
external server, database, or build step — Google Sheets is the database.

## Files

- `Code.gs` — all backend logic: authentication, CRUD for every module, business
  calculations (tax, discounts, crew pay, payroll, inventory sync), and the sheet
  helper utilities.
- `Index.html` — the entire frontend UI (login screen, sidebar, dashboard, and every
  module) as inline HTML/CSS/JS.
- `appsscript.json` — the Apps Script project manifest (web app deployment settings,
  timezone, runtime).

## Setup

1. Create a new Google Sheet. This will act as the database.
2. In the Sheet, open **Extensions → Apps Script**.
3. Replace the default `Code.gs` with the contents of this repo's `Code.gs`, and add
   an HTML file named `Index` with the contents of this repo's `Index.html`.
   (If you use [`clasp`](https://github.com/google/clasp), you can push this repo
   directly — just make sure the Apps Script project is bound to your Sheet.)
4. In the Apps Script editor, run the `setupSystem` function once (select it from the
   function dropdown and click **Run**). This creates all required sheets/tabs and an
   initial admin account.
5. Deploy the project as a **Web App** (Deploy → New deployment → Web app):
   - Execute as: **Me**
   - Who has access: choose according to your needs (e.g. **Anyone with the link**)
6. Open the deployed web app URL and sign in with the admin account created by
   `setupSystem` (see `APP.INITIAL_ADMIN` in `Code.gs` for the default email and
   password — change the password after first login by updating the `Users` sheet).

## Modules

Dashboard, Calendar, Clients / Show workspace (Overview, Products, Crew, Expenses,
Payments, Profit), Products, Inventory (with Stock In / Stock Out), Loans, Employees
(with advances/loans), Payroll, Payments, Expenses, Reports, Notifications, Settings,
and Admin Profile.

Documents (file uploads) is not yet implemented — it's a placeholder for a future
phase since it requires Google Drive integration.
