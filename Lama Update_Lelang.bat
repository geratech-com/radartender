@echo off
color 0B
echo ===================================================
echo     ROBOT RADAR TENDER LPSE (AUTO-UPDATE)
echo ===================================================
echo.
echo [1/3] Menjalankan scraping data LPSE...
python robot.py

echo.
echo [2/3] Mendaftarkan file Excel baru...
git add Hasil_Penarikan_LPSE_Nasional.xlsx

echo.
echo [3/3] Mengirim file ke website Streamlit...
git commit -m "Update otomatis dari PC Lokal"
git push origin main

echo.
echo ===================================================
echo  BINGO! WEBSITE STREAMLIT SUDAH TER-UPDATE.
echo ===================================================
pause