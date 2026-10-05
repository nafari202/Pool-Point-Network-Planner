"""Run the full pool point planning pipeline.

    python run_pipeline.py            # use cached Census data if present
    python run_pipeline.py --refresh  # re-download the Census series from FRED

On Windows with Excel installed, the workbook is then finished by
scripts/excel_finalize.ps1, which adds PivotTables and caches formula values.
"""
import json
import shutil
import subprocess
import sys

from src import config
from src.pipeline import run


def finalize_workbook(path):
    script = config.ROOT / "scripts" / "excel_finalize.ps1"
    if sys.platform != "win32" or not shutil.which("powershell"):
        print("Skipping Excel finalize step (needs Windows and Excel). Formulas calculate when the file is opened.")
        return
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-Path", str(path)],
        capture_output=True, text=True, timeout=600)
    print(result.stdout.strip() or result.stderr.strip())


if __name__ == "__main__":
    results = run(refresh_data="--refresh" in sys.argv)
    print(json.dumps(results["summary"], indent=2))
    finalize_workbook(config.OUTPUT_DIR / "pool_point_operations_report.xlsx")
    print(f"\nOutputs written to {config.OUTPUT_DIR} and {config.IMAGE_DIR}")
