# Shipping Status Agent

Automates daily shipping status reports by scraping FedEx tracking links from an Excel file using Playwright.

## What it does
- Auto-detects the newest `.xlsx` in the folder
- Processes tabs matching `Failed_Delivery_W#`, `Delayed_Delivery_W#`, `Out_for_delivery_W#`
- Writes **Current Status** → column X and **Days in Transit** → column Y
- Skips rows already marked "Delivered"
- Retries failed rows 3x; writes "Check Manually" if still unreachable
- Saves progress every 100 rows

## Requirements
- Python 3
- [Playwright](https://playwright.dev/python/) — `pip install playwright && playwright install`
- `openpyxl` — `pip install openpyxl`

## Usage
Double-click `Run Tracking Report.command` or run:

```bash
python3 scrape_tracking.py
```

Drop your weekly `.xlsx` file in the same folder — the script automatically picks up the newest one.

## Notes
- Date format on the tracking site is `DD/MM/YYYY`
- Pickup day counts as day 0 for days-in-transit calculation
- "Returns" and "Return" badges are normalized to "Returned"
