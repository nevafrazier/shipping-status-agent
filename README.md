# Shipping Status Agent

Automates daily shipping status reports by scraping FedEx tracking links from an Excel file using Playwright.

## What it does
- Auto-detects the newest `.xlsx` in the configured folder
- Processes tabs matching `Failed_Delivery_W#`, `Delayed_Delivery_W#`, `Out_for_delivery_W#`
- Writes **Current Status** → column X and **Days in Transit** → column Y
- Skips rows already marked "Delivered"
- Retries failed rows 3x; writes "Check Manually" if still unreachable
- Saves progress every 100 rows so no data is lost if interrupted

## Requirements
- Python 3
- [Playwright](https://playwright.dev/python/) — `pip install playwright && playwright install`
- `openpyxl` — `pip install openpyxl`
- `python-dotenv` — `pip install python-dotenv`

## Setup

```bash
cp .env.example .env
# Edit .env and set SHIPPING_FOLDER to your report folder path
```

## Usage
Double-click `Run Tracking Report.command` or run:

```bash
python3 scrape_tracking.py
```

Drop your weekly `.xlsx` file in the configured folder — the script automatically picks up the newest one.

## Configuration

All configuration is managed through environment variables loaded from a `.env` file. Copy `.env.example` to `.env` to get started — `.env` is listed in `.gitignore` and will never be committed.

| Variable | Description | Default |
|---|---|---|
| `SHIPPING_FOLDER` | Path to the folder containing weekly `.xlsx` reports | `~/Desktop/Shipping Agent` |

## Notes
- Date format on the tracking site is `DD/MM/YYYY`
- Pickup day counts as day 0 for days-in-transit calculation
- "Returns" and "Return" badges are normalized to "Returned"

## Security Best Practices

This project follows security best practices appropriate for tools handling internal business data:

- **No hardcoded paths or credentials** — the report folder path is loaded from a `.env` file via `python-dotenv`, keeping machine-specific configuration out of source code
- **`.env` is gitignored** — sensitive configuration is never committed to version control; only `.env.example` (with placeholder values) is tracked
- **Excel files are gitignored** — the `.gitignore` excludes all `.xlsx` and `.xls` files, preventing internal business data from being accidentally pushed to a remote repository
- **Headless browser isolation** — Playwright runs Chromium in headless mode with no stored session data; no cookies or credentials are persisted between runs
- **No third-party auth required** — the script only accesses the public FedEx tracking page using URLs already present in the Excel file; no API keys or login credentials are needed
