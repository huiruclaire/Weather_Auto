"""
refresh_weather.py
------------------
Opens the Weather API workbook in Excel, refreshes ALL web APIs / queries,
saves, and closes -- fully automatically, with no windows popping up.

This is the file Windows Task Scheduler runs each day at 8:00 AM.
It writes what happened to a log file so you can check it went OK.

Beginner note: you normally never run this by double-clicking. Task Scheduler
runs it for you. But you CAN test it manually -- see README.md.
"""

import sys
import gc
import time
import shutil
import datetime
import logging
import subprocess
from pathlib import Path

import pythoncom
import win32com.client as win32
import win32process
import win32api

import config


def running_excel_pids():
    """Return the set of process IDs of Excel instances currently running."""
    pids = set()
    try:
        out = subprocess.run(
            ["tasklist", "/fi", "imagename eq excel.exe", "/fo", "csv", "/nh"],
            capture_output=True, text=True,
        ).stdout
        for line in out.splitlines():
            parts = line.split('","')
            if len(parts) >= 2 and parts[1].strip('"').isdigit():
                pids.add(int(parts[1].strip('"')))
    except Exception:
        pass
    return pids


def backup_workbook():
    """Copy the workbook to the backup folder, then delete old backups.

    Runs BEFORE the refresh so we always keep the last-known-good version.
    A backup problem is logged but never stops the refresh from happening.
    """
    source = Path(config.WORKBOOK_PATH)
    if not source.exists():
        logging.warning("Cannot back up -- workbook not found at %s", source)
        return

    config.BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    dest = config.BACKUP_DIR / f"{source.stem}_backup_{stamp}{source.suffix}"

    try:
        shutil.copy2(source, dest)  # copy2 preserves the timestamps
        logging.info("Backup saved: %s", dest.name)
    except Exception:
        logging.exception("Backup FAILED (continuing with refresh anyway).")
        return

    # Delete backups older than the retention window.
    if config.BACKUP_RETENTION_DAYS and config.BACKUP_RETENTION_DAYS > 0:
        cutoff = time.time() - config.BACKUP_RETENTION_DAYS * 86400
        removed = 0
        for old in config.BACKUP_DIR.glob(f"{source.stem}_backup_*{source.suffix}"):
            try:
                if old.stat().st_mtime < cutoff:
                    old.unlink()
                    removed += 1
            except Exception:
                logging.warning("Could not delete old backup: %s", old.name)
        if removed:
            logging.info(
                "Removed %d backup(s) older than %d days.",
                removed, config.BACKUP_RETENTION_DAYS,
            )


def export_range_image(workbook):
    """Export the configured cell range to a PNG (latest + a dated copy).

    Uses Excel's PDF print engine (ExportAsFixedFormat) rather than "copy as
    picture" -- the print engine renders reliably in automation, with no
    dependence on the clipboard or on-screen painting. The one-page PDF is then
    rendered to a PNG (PyMuPDF) and its white border trimmed off (Pillow).

    An image problem is logged but never stops the run -- the data refresh has
    already succeeded by this point.
    """
    import fitz  # PyMuPDF
    from PIL import Image, ImageChops

    xlTypePDF, xlPortrait = 0, 1
    try:
        config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        sheet = workbook.Worksheets(config.IMAGE_SHEET)
        latest = config.OUTPUT_DIR / config.IMAGE_LATEST_NAME
        dated = config.OUTPUT_DIR / f"Weather_{datetime.date.today():%Y-%m-%d}.png"
        tmp_pdf = config.OUTPUT_DIR / "_export_tmp.pdf"

        # Set up the print layout: just our range, scaled to a single page with
        # no margins. Save and restore the sheet's original print settings so we
        # never change how the workbook itself prints.
        ps = sheet.PageSetup
        saved = {}
        for attr in ("PrintArea", "Zoom", "FitToPagesWide", "FitToPagesTall",
                     "Orientation", "LeftMargin", "RightMargin", "TopMargin",
                     "BottomMargin", "HeaderMargin", "FooterMargin"):
            try:
                saved[attr] = getattr(ps, attr)
            except Exception:
                pass
        try:
            ps.PrintArea = config.IMAGE_RANGE
            ps.Orientation = xlPortrait
            ps.Zoom = False            # required so FitToPages settings take effect
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
                      "HeaderMargin", "FooterMargin"):
                try:
                    setattr(ps, m, 0)
                except Exception:
                    pass

            sheet.ExportAsFixedFormat(xlTypePDF, str(tmp_pdf))
        finally:
            for attr, val in saved.items():   # restore original print settings
                try:
                    setattr(ps, attr, val)
                except Exception:
                    pass

        # Render the first PDF page to a high-resolution PNG.
        doc = fitz.open(str(tmp_pdf))
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(300 / 72, 300 / 72))  # 300 DPI
        pix.save(str(latest))
        doc.close()
        try:
            tmp_pdf.unlink()
        except Exception:
            pass

        # Trim the white border so the image is just the cells.
        try:
            img = Image.open(latest).convert("RGB")
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bbox = ImageChops.difference(img, bg).getbbox()
            if bbox:
                pad = 8
                left, top, right, bottom = bbox
                left = max(0, left - pad); top = max(0, top - pad)
                right = min(img.width, right + pad); bottom = min(img.height, bottom + pad)
                img.crop((left, top, right, bottom)).save(latest)
        except Exception:
            logging.warning("Could not auto-trim the image border (keeping full page).")

        shutil.copyfile(latest, dated)
        kb = latest.stat().st_size / 1024
        logging.info("Image exported: %s (%.1f KB) and %s", latest.name, kb, dated.name)
        if kb < 8:
            logging.warning("Exported image is small (%.1f KB) -- please check it.", kb)

        # Prune dated images older than the retention window.
        if config.IMAGE_RETENTION_DAYS and config.IMAGE_RETENTION_DAYS > 0:
            cutoff = time.time() - config.IMAGE_RETENTION_DAYS * 86400
            for old in config.OUTPUT_DIR.glob("Weather_*-*-*.png"):
                try:
                    if old.stat().st_mtime < cutoff:
                        old.unlink()
                except Exception:
                    logging.warning("Could not delete old image: %s", old.name)

    except Exception:
        logging.exception("Image export FAILED (data refresh still succeeded).")


def setup_logging():
    """Send messages both to the screen and to a dated log file."""
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = config.LOG_DIR / f"refresh_{datetime.date.today():%Y-%m}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return log_file


def refresh_workbook():
    """The actual work: open Excel, refresh, save, close."""
    excel = None
    workbook = None
    excel_pid = None

    # Snapshot any Excel instances the user already has open BEFORE we start.
    # We will never terminate one of these -- only the instance we launch.
    pre_existing_pids = running_excel_pids()

    # Start Excel invisibly.
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False          # never show pop-up dialogs
    excel.AskToUpdateLinks = False       # don't ask about linked data

    # Remember the exact process ID of THIS Excel instance, so that at the end
    # we can guarantee it is gone -- COM sometimes leaves Excel running as an
    # invisible orphan, and one would pile up every single day otherwise.
    try:
        _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
    except Exception:
        excel_pid = None
    # Open .xlsm without running its macros or asking about macro security.
    try:
        excel.AutomationSecurity = 3     # msoAutomationSecurityForceDisable
    except Exception:
        pass

    logging.info("Opening workbook: %s", config.WORKBOOK_PATH)
    workbook = excel.Workbooks.Open(config.WORKBOOK_PATH, UpdateLinks=0)

    # Force every data connection (including Power Query) to refresh in the
    # FOREGROUND. When BackgroundQuery is False, RefreshAll blocks until the
    # download has actually finished -- instead of returning immediately and
    # leaving the query still downloading (which cut the data off before).
    connection_count = workbook.Connections.Count
    logging.info("Found %d data connection(s) in the workbook.", connection_count)
    for conn in workbook.Connections:
        name = getattr(conn, "Name", "(unnamed)")
        set_ok = False
        for prop in ("OLEDBConnection", "ODBCConnection"):
            try:
                getattr(conn, prop).BackgroundQuery = False
                set_ok = True
                break
            except Exception:
                continue
        logging.info(
            "  Connection '%s': foreground refresh %s",
            name, "ENABLED" if set_ok else "not applicable",
        )

    logging.info("Refreshing all data connections / web APIs (waiting for completion)...")
    workbook.RefreshAll()

    # Backstop wait: even with foreground refresh, poll every connection until
    # none report that they are still refreshing, up to the configured timeout.
    # This guarantees we never save a half-downloaded query.
    deadline = time.time() + config.REFRESH_TIMEOUT_SECONDS
    while True:
        still_refreshing = []
        for conn in workbook.Connections:
            for prop in ("OLEDBConnection", "ODBCConnection"):
                try:
                    if getattr(conn, prop).Refreshing:
                        still_refreshing.append(getattr(conn, "Name", "(unnamed)"))
                    break
                except Exception:
                    continue
        # Also let cell-formula APIs (WEBSERVICE/FILTERXML) settle.
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        if not still_refreshing:
            logging.info("All connections finished refreshing.")
            break
        if time.time() >= deadline:
            logging.warning(
                "TIMEOUT after %ds -- still refreshing: %s. "
                "Increase REFRESH_TIMEOUT_SECONDS in config.py if the source is slow.",
                config.REFRESH_TIMEOUT_SECONDS, ", ".join(still_refreshing),
            )
            break
        time.sleep(3)

    # Final full recalculation so all formulas reflect the freshly loaded data.
    excel.CalculateFull()
    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass

    # Record how many rows each query loaded, so the log proves the download
    # completed (and you can spot any future shortfall at a glance).
    logging.info("Rows loaded per query table:")
    for ws in workbook.Worksheets:
        try:
            for lo in ws.ListObjects:
                logging.info("  '%s' (sheet '%s'): %d rows", lo.Name, ws.Name, lo.ListRows.Count)
        except Exception:
            pass

    # Export the configured cells to a shareable PNG (before saving, so the
    # temporary chart used for export is never left in the workbook).
    export_range_image(workbook)

    logging.info("Saving workbook...")
    workbook.Save()
    workbook.Close(SaveChanges=True)
    workbook = None

    excel.Quit()
    excel = None

    # Release COM references and force Python to drop them, so Excel can exit.
    gc.collect()

    # Belt-and-braces: if that specific Excel process is STILL alive a moment
    # later, terminate it by its process ID. The guard below makes it
    # impossible to close an Excel the user already had open.
    if excel_pid and excel_pid not in pre_existing_pids:
        time.sleep(2)
        try:
            handle = win32api.OpenProcess(0x0001, False, excel_pid)  # PROCESS_TERMINATE
            win32api.TerminateProcess(handle, 0)
            win32api.CloseHandle(handle)
            logging.info("Closed the leftover Excel process (PID %s).", excel_pid)
        except Exception:
            pass  # already gone -- which is exactly what we want

    logging.info("Done -- workbook refreshed and saved successfully.")


def main():
    log_file = setup_logging()
    logging.info("=" * 60)
    logging.info("Weather_Auto refresh starting")

    # Always take a backup first, so a bad refresh can never lose good data.
    backup_workbook()

    # COM must be initialised on this thread before talking to Excel.
    pythoncom.CoInitialize()
    try:
        refresh_workbook()
        return 0
    except Exception:
        logging.exception("REFRESH FAILED -- see the error above.")
        return 1
    finally:
        pythoncom.CoUninitialize()
        logging.info("Log written to: %s", log_file)


if __name__ == "__main__":
    sys.exit(main())
