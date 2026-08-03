import os
import sys
import pandas as pd

# Mengimpor daftar LPSE dan fungsi scraper langsung dari app.py
from app import DAFTAR_LPSE, run_scraper, FILE_EXCEL_OUTPUT

class DummyLog:
    def info(self, msg):
        print(f"INFO: {msg}")
    def warning(self, msg):
        print(f"WARNING: {msg}")
    def success(self, msg):
        print(f"SUCCESS: {msg}")

def main():
    print("🤖 Memulai proses scraping otomatis LPSE...")
    print(f"📊 Total Portal LPSE yang akan discan: {len(DAFTAR_LPSE)} Portal")

    log = DummyLog()

    df_result = run_scraper(
        selected_lpse=DAFTAR_LPSE,
        target_years=[2026, 2025],
        max_pages=5,
        log_container=log
    )

    print("✨ Process Scraping Selesai!")

if __name__ == "__main__":
    main()