@echo off
cd /d "%~dp0"
color 0B
echo ===================================================
echo    RADAR TENDER LPSE - UPDATE FULL (104 PORTAL)
echo ===================================================
echo.
python robot.py
echo.
"C:\Program Files\Git\cmd\git.exe" add Hasil_Penarikan_LPSE_Nasional.xlsx
"C:\Program Files\Git\cmd\git.exe" commit -m "Update Penuh 104 LPSE: %date%"
"C:\Program Files\Git\cmd\git.exe" push origin main
echo.
echo ===================================================
echo  BINGO! UPDATE FULL SELESAI.
echo ===================================================
pause