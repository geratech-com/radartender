@echo off
cd /d "%~dp0"
color 0A
echo ===================================================
echo    RADAR TENDER LPSE - UPDATE KILAT (15 UTAMA)
echo ===================================================
echo.
python -c "from app import DAFTAR_LPSE, run_scraper; from robot import DummyLog; lpse_utama = DAFTAR_LPSE[:15]; run_scraper(lpse_utama, [2026, 2025], 3, DummyLog())"
echo.
"C:\Program Files\Git\cmd\git.exe" add .
"C:\Program Files\Git\cmd\git.exe" commit -m "Update Kilat LPSE Utama: %date%"
"C:\Program Files\Git\cmd\git.exe" push origin main
echo.
echo ===================================================
echo  BINGO! UPDATE KILAT SELESAI.
echo ===================================================
pause