# Weather_Auto

Automatically refreshes an Excel workbook on a shared network drive every day at
**8:00 AM**, so the web APIs inside it pull fresh data without anyone opening it.

> **Setup:** the private file paths are not stored in this repo. Copy
> `config_local.example.py` to `config_local.py` and fill in your real workbook
> path and image-output folder. (`config_local.py` is git-ignored.)

## 🔑 The essentials (read this first)

Each day at 8:00 AM the job runs this sequence, all recorded in the `logs` folder:

> **back up the workbook → refresh all web data (wait for the full download) →
> export a picture of the CoreSummary table → save.**

The handful of things that actually matter:

1. **The PC must be ON and LOGGED IN at 8:00 AM.** A locked screen is fine; being
   signed out or shut down is not — the job simply won't run.
2. **The workbook must stay "trusted"** (you clicked *Enable Content* once). If
   that trust is ever reset, the web queries silently won't download.
3. **The refresh takes ~5 minutes** (a real download). The timeout is set to 10
   minutes in `config.py` to leave headroom; don't lower it.
4. **Judge success by the `Data` worksheet or the log, NOT the Power Query editor.**
   The editor's *"download did not complete"* preview message is misleading. A
   full row count (~1,500+, varies daily) means success.
5. **The shared picture** is saved to the image folder set in `config_local.py`
   as `Weather_latest.png` (plus a dated copy). Sending it to WhatsApp is manual.

## How it works (the simple version)

Excel can't refresh itself while it's closed, so three things work together:

1. **Windows Task Scheduler** — the "alarm clock". At 8:00 AM it runs `run_refresh.bat`.
2. **`run_refresh.bat`** — starts the project's private Python (in the `venv` folder).
3. **`refresh_weather.py`** — opens the workbook invisibly, refreshes all the web
   APIs, saves, and closes.

Each run **first backs up** the workbook to the `backups` folder (keeping the
last-known-good copy), **then** refreshes, and finally **exports a picture** of
the CoreSummary table (cells `G1:Q66`) to the image folder set in `config_local.py`.
It also writes a note to the `logs` folder so you can confirm it worked.

Backups older than 30 days are deleted automatically so the folder never fills
up. You can change that in `config.py` (`BACKUP_RETENTION_DAYS`).

## 📷 The picture (for sharing, e.g. WhatsApp)

After every refresh, the CoreSummary table is saved as an image in the folder set
by `OUTPUT_DIR` in `config_local.py`:

- **`Weather_latest.png`** — always the newest picture (grab this one to share).
- **`Weather_YYYY-MM-DD.png`** — a dated copy, kept for 30 days as history.

To change which cells become the picture, edit `IMAGE_SHEET` / `IMAGE_RANGE` in
`config.py`. To change *where* it's saved, edit `OUTPUT_DIR` in `config_local.py`.
Sending it anywhere (WhatsApp, email, etc.) is done by you.

## ✅ How to check the data actually refreshed

Look at the **Data worksheet tab** and confirm it is fully populated (it fills
down to around **row 1560–1575** — the exact total changes a little each day as
the live forecast data changes, which is normal).

**Do NOT judge it by the Power Query Editor.** The editor's preview pane can say
*"download did not complete"* even when the data loaded perfectly — that message
is about the editor's own preview cache, which only updates when you manually
click *Refresh Preview* inside the editor. It does not reflect what the daily
refresh actually loads into the worksheet.

Even easier: open the newest file in the `logs` folder. Every run records the
row count, e.g. `'Data' (sheet 'Data'): 1561 rows`. A full count (roughly
1,500+) means success; only a handful of rows means it was cut short.

## ⚠️ One important rule

Your PC must be **turned on and logged in** at 8:00 AM.
- Screen **locked** is fine. ✅
- **Logged out / signed out / shut down** will NOT work. ❌

## The files

| File | What it is |
|------|-----------|
| `refresh_weather.py` | The main program that refreshes Excel. |
| `config.py` | Settings you might change (timeouts, retention, image range). |
| `config_local.py` | Your private paths (git-ignored — not on GitHub). |
| `config_local.example.py` | Template to copy into `config_local.py` on setup. |
| `run_refresh.bat` | What Task Scheduler runs. |
| `requirements.txt` | The list of libraries this project needs. |
| `venv/` | The project's private copy of Python + libraries (don't edit). |
| `logs/` | A log of every run. |
| `backups/` | A dated copy of the workbook from before each refresh. |

The exported pictures (`Weather_latest.png` + dated copies) are saved to the
folder set by `OUTPUT_DIR` in `config_local.py`, not inside the project folder.

## Testing it yourself (optional)

Double-click `run_refresh.bat`. Excel will open in the background, refresh, and
close. Then open the newest file in the `logs` folder to see the result.

## If it ever stops working

1. Open the newest file in the `logs` folder and read the last few lines.
2. Check the row count it logged. If `Data` shows only a handful of rows (or a
   `TIMEOUT` warning appears), the download was cut short — increase
   `REFRESH_TIMEOUT_SECONDS` in `config.py`, or check the web source is up.
3. Check the shared drive is connected and the workbook path in `config_local.py` is correct.
4. Make sure the PC was on and logged in at 8:00 AM.
5. Make sure the workbook still has **Enable Content** trusted (a "Trusted
   Document"). If that trust is ever reset, the data connections won't run.
