@echo off
cd /d "%~dp0"
color 0A
echo ===================================================
echo    RADAR TENDER LPSE - UPDATE KILAT (15 UTAMA)
echo ===================================================
echo.
python -c "from app import DAFTAR_LPSE, run_scraper, FILE_EXCEL_OUTPUT, save_and_update_excel; from robot import DummyLog; lpse_utama = DAFTAR_LPSE[:15]; df = run_scraper(lpse_utama, [2026, 2025, 2024], 5, DummyLog()); save_and_update_excel(df, FILE_EXCEL_OUTPUT)"
echo.
"C:\Program Files\Git\cmd\git.exe" add .
"C:\Program Files\Git\cmd\git.exe" commit -m "Update Kilat LPSE Utama: %date%"
"C:\Program Files\Git\cmd\git.exe" push origin main
echo.
echo ===================================================
echo  BINGO! UPDATE KILAT SELESAI.
echo ===================================================
pause