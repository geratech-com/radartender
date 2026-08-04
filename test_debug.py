import requests
import re

# Uji ambil data dari Kementerian PUPR
url = "https://spse.inaproc.id/pu/dt/lelang?draw=1&start=0&length=10&tahun=2026&kategoriId=4"
headers = {'X-Requested-With': 'XMLHttpRequest', 'User-Agent': 'Mozilla/5.0'}

print("🔍 Menghubungkan ke server LPSE PUPR...")
try:
    resp = requests.get(url, headers=headers, timeout=15)
    print(f"Status Koneksi: {resp.status_code}")
    
    json_data = resp.json()
    total_data = json_data.get('recordsTotal', 0)
    rows = json_data.get('data', [])
    
    print(f"📊 Total Data Ditemukan di Server: {total_data}")
    print(f"📦 Jumlah Baris yang Ditarik: {len(rows)}\n")
    
    if rows:
        print("--- CONTOH MENTAH BARIS PERTAMA ---")
        for i, col in enumerate(rows[0]):
            print(f"Kolom [{i}]: {str(col)[:100]}...")
    else:
        print("⚠️ Data kosong dari server untuk kategori ini.")

except Exception as e:
    print(f"❌ Terjadi Kesalahan: {e}")