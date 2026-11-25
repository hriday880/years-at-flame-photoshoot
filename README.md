# Years at FLAME — Photoshoot Registration

Yearbook photoshoot check-in form, admin photocode desk console, printable QR standee, and one-click `.xlsx` export for **Years at FLAME**.

## Architecture
- **Student Form (`/`)**: 5-step registration (`Details` -> `Batch` -> `Studies` -> `Student ID` -> `Review`) with searchable `MAJOR_MINORS` comboboxes and double-entry Student ID verification.
- **Admin Console (`/admin`)**: Basic-auth protected desk view with live metrics, search/batch/status filters, double-entry photocode assignment (with uniqueness + online edit lock), printable QR standee (`/admin/qr.svg`), and Excel download (`/admin/export.xlsx`).
- **Dual-Mode Storage**:
  - **Turso DB (Production)**: When `TURSO_DATABASE_URL` and `TURSO_AUTH_TOKEN` are set, persists registrations and an append-only `event_log` via Turso's Hrana v2 HTTP pipeline (works seamlessly on Vercel Serverless Functions).
  - **Local Offline Mode**: When Turso env vars are omitted, saves to `registrations.jsonl` and `registrations.xlsx` in the project directory.

## Local Run
```bash
./start_mac.command   # or start_windows.bat
```
Or manually:
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python AppUpdate.py
```

## Deploy to Vercel + Turso DB
1. Create a Turso DB and token with the Turso CLI:
   ```bash
   turso db create years-at-flame
   turso db show years-at-flame --url
   turso db tokens create years-at-flame
   ```
2. Set `TURSO_DATABASE_URL`, `TURSO_AUTH_TOKEN`, and `ADMIN_PASSWORD` in Vercel environment variables.
3. Deploy with `vercel deploy --prod`.
