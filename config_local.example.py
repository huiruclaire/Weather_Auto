"""
config_local.example.py
-----------------------
TEMPLATE for the private, machine-specific paths.

Setup: copy this file to `config_local.py` (same folder) and replace the paths
below with your real ones. `config_local.py` is git-ignored, so your real paths
stay off GitHub.
"""

from pathlib import Path

# Full path to the Excel workbook that should be refreshed.
WORKBOOK_PATH = r"C:\path\to\Your Workbook.xlsm"

# Folder where the exported images are written.
OUTPUT_DIR = Path(r"C:\path\to\image output folder")
