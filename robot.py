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
    
    # 1. BERSINKRONISASI & SAPU BERSIH EXCEL LOKAL DULU
    if os.path.exists(FILE_EXCEL_OUTPUT):
        print("\n🧹 Membersihkan Excel lokal dari proyek fisik murni...")
        try:
            df_old = pd.read_excel(FILE_EXCEL_OUTPUT)
            total_awal = len(df_old)
            df_clean = clean_df_master(df_old)
            
            with pd.ExcelWriter(FILE_EXCEL_OUTPUT, engine="openpyxl") as writer:
                df_clean.to_excel(writer, index=False, sheet_name="Data LPSE Nasional")
            
            print(f"✅ Pembersihan Selesai! (Dari {total_awal} paket -> tersisa {len(df_clean)} paket sah 3 kategori).")
        except Exception as e:
            print(f"⚠️ Catatan pembersihan: {e}")

    # 2. JALANKAN PENARIKAN LENGKAP 104 LPSE
    log = DummyLog()
    target_years = [2026, 2025, 2024, 2023, 2022]
    max_pages = 10
    
    run_scraper(DAFTAR_LPSE, target_years, max_pages, log)
    
    print("\n===================================================")
    print(" 🎉 PROSES SCRAPING & SINKRONISASI SELESAI PERFECT!")
    print("===================================================")