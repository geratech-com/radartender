import os
import pandas as pd
from app import DAFTAR_LPSE, FILE_EXCEL_OUTPUT, clean_df_master, run_scraper

class DummyLog:
    def info(self, msg):
        print(f"[INFO] {msg}")
    def warning(self, msg):
        print(f"[WARN] {msg}")
    def success(self, msg):
        print(f"[SUCCESS] {msg}")

if __name__ == "__main__":
    print("===================================================")
    print(" 🚀 MEMULAI UPDATE ROBOT FULL (104 LPSE NASIONAL)")
    print("===================================================")
    
    if os.path.exists(FILE_EXCEL_OUTPUT):
        print("\n🧹 Membersihkan Excel lokal...")
        try:
            df_old = pd.read_excel(FILE_EXCEL_OUTPUT)
            df_clean = clean_df_master(df_old)
            with pd.ExcelWriter(FILE_EXCEL_OUTPUT, engine="openpyxl") as writer:
                df_clean.to_excel(writer, index=False, sheet_name="Data LPSE Nasional")
        except Exception:
            pass

    log = DummyLog()
    target_years = [2026, 2025, 2024]
    max_pages = 10
    
    run_scraper(DAFTAR_LPSE, target_years, max_pages, log)
    
    print("\n===================================================")
    print(" 🎉 PROSES SCRAPING & SINKRONISASI SELESAI PERFECT!")
    print("===================================================")