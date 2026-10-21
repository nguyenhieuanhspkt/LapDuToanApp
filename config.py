"""
Configuration module for LapDuToanApp.
Specifies paths to Data on OneDrive EVN and Source Code.
"""

import os
from pathlib import Path

# Base Paths
SOURCE_DIR = Path(__file__).resolve().parent

# App Data Directory on OneDrive EVN
DEFAULT_DATA_DIR = Path(r"D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Hiếu\2026_LAPDUTOAN_APP")
DATA_DIR = Path(os.environ.get("LAPDUTOAN_DATA_DIR", DEFAULT_DATA_DIR))

# Subdirectories
DATABASE_DIR = DATA_DIR / "database"
IMAGES_DIR = DATA_DIR / "images_evidence"
TEMPLATES_DIR = DATA_DIR / "templates"
OUTPUT_DIR = DATA_DIR / "output"

# File Paths
NORMS_DB_FILE = DATABASE_DIR / "norms_database.json"
PRICES_DB_FILE = DATABASE_DIR / "resource_prices.json"
TEMPLATE_XLSX = TEMPLATES_DIR / "DU_TOAN_PCCC_2026_LATEST.xlsx"

def check_environment():
    """Verify that all required data directories and files exist."""
    issues = []
    if not DATA_DIR.exists():
        issues.append(f"Data directory does not exist: {DATA_DIR}")
    if not NORMS_DB_FILE.exists():
        issues.append(f"Norms database missing: {NORMS_DB_FILE}")
    if not PRICES_DB_FILE.exists():
        issues.append(f"Resource prices missing: {PRICES_DB_FILE}")
    if not IMAGES_DIR.exists():
        issues.append(f"Images directory missing: {IMAGES_DIR}")
        
    return {
        "status": len(issues) == 0,
        "issues": issues,
        "data_dir": str(DATA_DIR),
        "source_dir": str(SOURCE_DIR)
    }

if __name__ == "__main__":
    result = check_environment()
    print("Environment Status:", "OK" if result["status"] else "FAILED")
    if result["issues"]:
        for issue in result["issues"]:
            print(" -", issue)
    else:
        print("All data directories and database files are ready.")
