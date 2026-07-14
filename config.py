"""
config.py
---------
All the settings you might want to change live here, in one place.
If the workbook ever moves, or you want the logs somewhere else, edit this file
only -- you never need to touch refresh_weather.py.

NOTE: the actual file paths (which are private) live in `config_local.py`, which
is NOT uploaded to GitHub. To set this project up on a machine, copy
`config_local.example.py` to `config_local.py` and fill in your real paths.
"""

from pathlib import Path

# --- Private, machine-specific paths (kept out of git) ---------------------
# These come from config_local.py so they never get published to a public repo.
try:
    from config_local import WORKBOOK_PATH, OUTPUT_DIR
except ImportError as exc:
    raise SystemExit(
        "Missing config_local.py -- copy config_local.example.py to "
        "config_local.py and fill in your real WORKBOOK_PATH and OUTPUT_DIR."
    ) from exc

WORKBOOK_PATH = str(WORKBOOK_PATH)   # full path to the Excel workbook to refresh
OUTPUT_DIR = Path(OUTPUT_DIR)        # folder where the exported images are written

# Folder where run logs are written (one line per run, plus errors).
# By default: a "logs" folder next to this script.
LOG_DIR = Path(__file__).resolve().parent / "logs"

# How long (in seconds) to wait for the web APIs to finish refreshing
# before giving up. A real download of this workbook takes ~4.5 minutes, so
# this is set to 10 minutes to leave comfortable headroom on a slow day.
REFRESH_TIMEOUT_SECONDS = 600

# --- Force-fresh downloads --------------------------------------------------
# Power Query keeps an on-disk cache of web responses under
# %LOCALAPPDATA%\Microsoft\Office\<ver>\PowerQuery\Cache. If it is NOT cleared,
# a daily refresh can silently serve yesterday's cached pages -- finishing in
# ~10 seconds with STALE data instead of doing a real ~5-minute download.
# Clearing it before each run forces Power Query to actually re-download.
# (This is a cache shared by all Power Query workbooks for this user; clearing
# it just makes them re-download next time -- harmless.)
CLEAR_POWERQUERY_CACHE = True

# --- Backups ---------------------------------------------------------------
# Every run copies the workbook here BEFORE refreshing, so you always have the
# last-known-good version. Default: a "backups" folder next to this script.
BACKUP_DIR = Path(__file__).resolve().parent / "backups"

# Backups older than this many days are deleted automatically, so the folder
# never fills up the disk. Set to 0 to keep backups forever.
BACKUP_RETENTION_DAYS = 30

# --- Image export ----------------------------------------------------------
# After each refresh, these cells are exported to a PNG image (e.g. to share
# on WhatsApp). Change the sheet or range here if you ever want a different area.
IMAGE_SHEET = "CoreSummary"
IMAGE_RANGE = "G1:Q66"

# (OUTPUT_DIR -- where the images are written -- is set above from config_local.py)

# "Weather_latest.png" is always overwritten with the newest image (easy to grab).
# A dated copy (e.g. Weather_2026-07-10.png) is also kept for history.
IMAGE_LATEST_NAME = "Weather_latest.png"

# Dated image copies older than this many days are deleted. 0 = keep forever.
IMAGE_RETENTION_DAYS = 30
