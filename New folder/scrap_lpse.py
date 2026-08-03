import datetime
import os
import re
import time
import openpyxl
import pandas as pd
from playwright.sync_api import sync_playwright

# ==============================================================================
# 1. KONFIGURASI DAFTAR TARGET LPSE SE-INDONESIA & KRITERIA
# ==============================================================================
DAFTAR_LPSE = [
    # --- KEMENTERIAN & PUSAT ---
    {"nama": "LKPP Pusat", "url": "https://spse.inaproc.id/lkpp/lelang"},
    {"nama": "Kementerian PUPR", "url": "https://spse.inaproc.id/pu/lelang"},
    {
        "nama": "Kementerian Keuangan",
        "url": "https://spse.inaproc.id/kemenkeu/lelang",
    },
    {
        "nama": "Kementerian Perhub",
        "url": "https://spse.inaproc.id/dephub/lelang",
    },
    {
        "nama": "Kementerian Kesehatan",
        "url": "https://spse.inaproc.id/kemkes/lelang",
    },
    {
        "nama": "Kemendikbudristek",
        "url": "https://spse.inaproc.id/kemdikbud/lelang",
    },
    {
        "nama": "Kementerian Pertanian",
        "url": "https://spse.inaproc.id/pertanian/lelang",
    },
    {
        "nama": "Kementerian ESDM",
        "url": "https://spse.inaproc.id/esdm/lelang",
    },
    {
        "nama": "Kemenkumham",
        "url": "https://spse.inaproc.id/kemenkumham/lelang",
    },
    {
        "nama": "Kementerian Pertahanan",
        "url": "https://spse.inaproc.id/kemhan/lelang",
    },
    {
        "nama": "Kementerian Luar Negeri",
        "url": "https://spse.inaproc.id/kemlu/lelang",
    },
    {
        "nama": "Kemendagri",
        "url": "https://spse.inaproc.id/kemendagri/lelang",
    },
    {
        "nama": "Kementerian Agama",
        "url": "https://spse.inaproc.id/kemenag/lelang",
    },
    {
        "nama": "Kemnaker",
        "url": "https://spse.inaproc.id/kemnaker/lelang",
    },
    {
        "nama": "Kemenperin",
        "url": "https://spse.inaproc.id/kemenperin/lelang",
    },
    {
        "nama": "Kemendag",
        "url": "https://spse.inaproc.id/kemendag/lelang",
    },
    {
        "nama": "Kementerian LHK",
        "url": "https://spse.inaproc.id/menlhk/lelang",
    },
    {"nama": "KKP", "url": "https://spse.inaproc.id/kkp/lelang"},
    {
        "nama": "Kemendesa",
        "url": "https://spse.inaproc.id/kemendesa/lelang",
    },
    {
        "nama": "Kominfo",
        "url": "https://spse.inaproc.id/kominfo/lelang",
    },
    {
        "nama": "Kementerian BUMN",
        "url": "https://spse.inaproc.id/bumn/lelang",
    },
    {
        "nama": "Kemenkop UKM",
        "url": "https://spse.inaproc.id/kemenkop/lelang",
    },
    {
        "nama": "Kemenparekraf",
        "url": "https://spse.inaproc.id/kemenparekraf/lelang",
    },
    {
        "nama": "Kemensos",
        "url": "https://spse.inaproc.id/kemensos/lelang",
    },
    {
        "nama": "Bappenas",
        "url": "https://spse.inaproc.id/bappenas/lelang",
    },
    {
        "nama": "KemenPANRB",
        "url": "https://spse.inaproc.id/menpan/lelang",
    },
    {
        "nama": "ATR / BPN",
        "url": "https://spse.inaproc.id/atrbpn/lelang",
    },
    {
        "nama": "Kemenpora",
        "url": "https://spse.inaproc.id/kemenpora/lelang",
    },
    {
        "nama": "Kementerian PPPA",
        "url": "https://spse.inaproc.id/kemenpppa/lelang",
    },
    {
        "nama": "BKPM / Investasi",
        "url": "https://spse.inaproc.id/bkpm/lelang",
    },
    # --- LEMBAGA & POLHUKAM ---
    {"nama": "Mabes TNI", "url": "https://spse.inaproc.id/tni/lelang"},
    {"nama": "Mabes Polri", "url": "https://spse.inaproc.id/polri/lelang"},
    {
        "nama": "Kejaksaan Agung",
        "url": "https://spse.inaproc.id/kejaksaan/lelang",
    },
    {
        "nama": "Mahkamah Agung",
        "url": "https://spse.inaproc.id/mahkamahagung/lelang",
    },
    {"nama": "BPK RI", "url": "https://spse.inaproc.id/bpk/lelang"},
    {"nama": "BPKP", "url": "https://spse.inaproc.id/bpkp/lelang"},
    # --- PROVINSI & KOTA UTAMA ---
    {"nama": "DKI Jakarta", "url": "https://spse.inaproc.id/jakarta/lelang"},
    {"nama": "Jawa Barat", "url": "https://spse.inaproc.id/jabar/lelang"},
    {"nama": "Jawa Tengah", "url": "https://spse.inaproc.id/jateng/lelang"},
    {"nama": "Jawa Timur", "url": "https://spse.inaproc.id/jatim/lelang"},
    {"nama": "Banten", "url": "https://spse.inaproc.id/banten/lelang"},
    {"nama": "D.I. Yogyakarta", "url": "https://spse.inaproc.id/jogjaprov/lelang"},
    {"nama": "Sumatera Utara", "url": "https://spse.inaproc.id/sumut/lelang"},
    {"nama": "Sumatera Selatan", "url": "https://spse.inaproc.id/sumsel/lelang"},
    {"nama": "Bali", "url": "https://spse.inaproc.id/baliprov/lelang"},
    {
        "nama": "Sulawesi Selatan",
        "url": "https://spse.inaproc.id/sulsel/lelang",
    },
    {"nama": "Kota Surabaya", "url": "https://spse.inaproc.id/surabaya/lelang"},
    {"nama": "Kota Medan", "url": "https://spse.inaproc.id/medan/lelang"},
    {"nama": "Kota Makassar", "url": "https://spse.inaproc.id/makassar/lelang"},
]

TAHUN_TARGET = [2026, 2025, 2024, 2023, 2022]

KRITERIA_JENIS = [
    "Jasa Konsultansi Badan Usaha Konstruksi",
    "Jasa Konsultansi Badan Usaha Non Konstruksi",
    "Pekerjaan Konstruksi Terintegrasi",
]

KOLOM_TARGET = [
    "Sumber LPSE",
    "ID LPSE",
    "Instansi",
    "Nama Paket",
    "Tahapan",
    "Status",
    "HPS",
    "Metode",
    "Jenis Pemilihan",
    "Evaluasi",
    "Jenis Pengadaan",
    "Tahun Anggaran",
    "Pemenang Kontrak",
    "Nilai Kontrak",
    "Link",
    "Waktu Download",
]


# ==============================================================================
# 2. FUNGSI PARSING & HEADER MAPPING DINAMIS
# ==============================================================================
def parse_rupiah_pintar(text):
  if not text or text in ["-", "0", "Nilai Kontrak belum dibuat"]:
    return 0
  text_str = str(text).strip()

  match_unit = re.search(
      r"([\d\.,]+)\s*(M|Miliar|Jt|Juta|Rb|Ribu|T|Triliun)",
      text_str,
      re.IGNORECASE,
  )
  if match_unit:
    num_part = match_unit.group(1).replace(".", "").replace(",", ".")
    unit = match_unit.group(2).lower()
    multiplier = 1
    if unit in ["m", "miliar"]:
      multiplier = 1_000_000_000
    elif unit in ["jt", "juta"]:
      multiplier = 1_000_000
    elif unit in ["rb", "ribu"]:
      multiplier = 1_000
    elif unit in ["t", "triliun"]:
      multiplier = 1_000_000_000_000
    try:
      return int(float(num_part) * multiplier)
    except ValueError:
      pass

  clean = text_str.replace("Rp.", "").replace("Rp", "").strip()
  if "," in clean:
    parts = clean.split(",")
    if len(parts[-1]) <= 2 and parts[-1].isdigit():
      clean = parts[0]
    else:
      clean = clean.replace(",", "")

  clean_digits = re.sub(r"[^\d]", "", clean)
  try:
    return int(clean_digits) if clean_digits else 0
  except ValueError:
    return 0


def fetch_detail_evaluasi_pemenang(context, base_domain, kode_id, base_referer):
  """Mengekstrak HPS, Nama Pemenang Presisi, dan Nilai Kontrak dengan Header Mapping."""
  exact_hps = 0
  pemenang = "Belum Ditetapkan"
  nilai_kontrak = 0

  endpoints = [
      f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak",
      f"{base_domain}/evaluasi/{kode_id}/pemenang",
  ]

  try:
    dp = context.new_page()

    for url in endpoints:
      try:
        dp.goto(
            url,
            referer=base_referer,
            wait_until="domcontentloaded",
            timeout=8000,
        )
        body_text = dp.inner_text("body")

        if "Akses Ditolak" in body_text or "OPPS!" in body_text:
          continue

        # 1. Ekstraksi HPS Presisi dari Form / Table
        if exact_hps == 0:
          match_hps = re.search(
              r"HPS\s*(?:Rp\.?)?\s*([\d.,]+)", body_text, re.IGNORECASE
          )
          if match_hps:
            exact_hps = parse_rupiah_pintar(match_hps.group(1))

        # 2. Parsing Tabel Menggunakan Mapping Kolom Header
        tables = dp.query_selector_all("table")
        for table in tables:
          t_text = table.inner_text()
          if "Nama Pemenang" not in t_text:
            continue

          rows = table.query_selector_all("tr")
          header_col_map = {}

          for row in rows:
            cells = row.query_selector_all("th, td")
            cell_texts = [c.inner_text().strip() for c in cells]

            # Deteksi Baris Header Tabel
            if any("nama pemenang" in c.lower() for c in cell_texts):
              for idx, text in enumerate(cell_texts):
                t_low = text.lower()
                if "nama pemenang" in t_low:
                  header_col_map["pemenang"] = idx
                elif any(
                    k in t_low
                    for k in [
                        "harga kontrak",
                        "harga negosiasi",
                        "harga terkoreksi",
                    ]
                ):
                  header_col_map["nilai"] = idx
              continue

            # Parsing Baris Data Menggunakan Posisi Indeks Header
            if "pemenang" in header_col_map and len(cell_texts) > header_col_map[
                "pemenang"
            ]:
              cand_pemenang = (
                  cell_texts[header_col_map["pemenang"]].split("\n")[0].strip()
              )

              # Pastikan bukan baris header atau teks kosong
              if cand_pemenang and not any(
                  h in cand_pemenang.lower()
                  for h in ["nama pemenang", "alamat", "npwp"]
              ):
                pemenang = cand_pemenang

                # Ambil Nilai Kontrak dari indeks kolom yang tepat
                if "nilai" in header_col_map and len(cell_texts) > header_col_map[
                    "nilai"
                ]:
                  nilai_kontrak = parse_rupiah_pintar(
                      cell_texts[header_col_map["nilai"]]
                  )
                else:
                  for cell_val in reversed(cell_texts[1:]):
                    pv = parse_rupiah_pintar(cell_val)
                    if pv > 0:
                      nilai_kontrak = pv
                      break
                break

          if pemenang != "Belum Ditetapkan":
            break

        if pemenang != "Belum Ditetapkan":
          break

      except Exception:
        continue

    dp.close()
  except Exception:
    pass

  return exact_hps, pemenang, nilai_kontrak


# ==============================================================================
# 3. FUNGSI MERGE, DEDUPLICATE & AUTO-SAVE KE EXCEL
# ==============================================================================
def save_and_update_excel(df_new, file_output):
  if df_new.empty:
    return

  if os.path.exists(file_output):
    try:
      df_existing = pd.read_excel(file_output)
      df_combined = pd.concat([df_existing, df_new], ignore_index=True)
      df_final = df_combined.drop_duplicates(
          subset=["Sumber LPSE", "ID LPSE"], keep="last"
      )
    except Exception as e:
      print(f"⚠️ Gagal membaca file lama ({e}), membuat file baru.")
      df_final = df_new
  else:
    df_final = df_new

  with pd.ExcelWriter(file_output, engine="openpyxl") as writer:
    df_final.to_excel(writer, index=False, sheet_name="Data LPSE Nasional")
    worksheet = writer.sheets["Data LPSE Nasional"]

    num_format = "#,##0"
    for row in range(2, len(df_final) + 2):
      worksheet[f"G{row}"].number_format = num_format  # Kolom G = HPS
      worksheet[f"N{row}"].number_format = num_format  # Kolom N = Nilai Kontrak


# ==============================================================================
# 4. FUNGSI UTAMA PENARIKAN DATA
# ==============================================================================
def scrape_lpse_nasional(file_output, max_pages_per_year=10):
  print(
      "🚀 MEMULAI PENARIKAN DATA LPSE SE-INDONESIA (HEADER MAPPING DINAMIS"
      " AKTIF)..."
  )

  with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    for lpse in DAFTAR_LPSE:
      lpse_nama = lpse["nama"]
      lpse_url = lpse["url"]
      base_domain = lpse_url.replace("/lelang", "")
      data_lpse_ini = []

      print("\n" + "=" * 60)
      print(f"🏛️ PORTAL: {lpse_nama.upper()}")
      print(f"🔗 URL: {lpse_url}")
      print("=" * 60)

      try:
        page.goto(lpse_url, wait_until="networkidle", timeout=45000)
      except Exception as e:
        print(f"❌ Server Down / Timeout ({lpse_nama}): {e}")
        continue

      for tahun in TAHUN_TARGET:
        print(f"\n  📅 Tahun Anggaran: {tahun}")

        try:
          select_tahun = page.query_selector(
              "select:has(option[value='2026']),"
              " select:has(option:has-text('2026'))"
          )
          if select_tahun:
            select_tahun.select_option(label=str(tahun))
          else:
            selects = page.query_selector_all("select")
            for s in selects:
              if "2026" in s.inner_text() or str(tahun) in s.inner_text():
                s.select_option(label=str(tahun))
                break
          time.sleep(2)
        except Exception as e:
          print(f"     ⚠️ Gagal memilih tahun {tahun}: {e}")

        page_num = 1
        while page_num <= max_pages_per_year:
          print(f"     📄 Memproses Halaman {page_num}...")

          try:
            page.wait_for_selector("table tbody tr", timeout=8000)
          except Exception:
            print(f"     ⚠️ Tidak ada data ditemukan untuk tahun {tahun}.")
            break

          time.sleep(1)
          rows = page.query_selector_all("table tbody tr")
          ditemukan_di_halaman = 0

          for row in rows:
            text_content = row.inner_text()

            matching_kriteria = None
            for kriteria in KRITERIA_JENIS:
              if kriteria.lower() in text_content.lower():
                matching_kriteria = kriteria
                break

            if matching_kriteria:
              cells = row.query_selector_all("td")
              if len(cells) >= 5:
                kode_id = (
                    cells[0].inner_text().strip() if len(cells) > 0 else "-"
                )
                cell_nama_raw = (
                    cells[1].inner_text().strip() if len(cells) > 1 else "-"
                )
                nama_paket_clean = cell_nama_raw.split("\n")[0].strip()

                instansi = (
                    cells[2].inner_text().strip() if len(cells) > 2 else "-"
                )
                tahapan = (
                    cells[3].inner_text().strip() if len(cells) > 3 else "-"
                )
                hps_raw_tabel = (
                    cells[4].inner_text().strip() if len(cells) > 4 else "0"
                )

                link_evaluasi_pemenang = (
                    f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak"
                )
                hps_tabel_val = parse_rupiah_pintar(hps_raw_tabel)

                # Ekstraksi Pemenang Murni Via Mapping
                print(
                    f"        🔎 Mengekstrak Pemenang & HPS Paket ID"
                    f" {kode_id}..."
                )
                (
                    hps_detail,
                    pemenang,
                    nilai_kontrak,
                ) = fetch_detail_evaluasi_pemenang(
                    context, base_domain, kode_id, lpse_url
                )

                hps_final = (
                    hps_detail if hps_detail > 0 else hps_tabel_val
                )

                status_paket = (
                    "Tender Batal"
                    if "Batal" in tahapan
                    else (
                        "Tender Selesai"
                        if "Selesai" in tahapan
                        or pemenang != "Belum Ditetapkan"
                        else "Proses Tender"
                    )
                )

                waktu_download_now = datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

                item = {
                    "Sumber LPSE": lpse_nama,
                    "ID LPSE": kode_id,
                    "Instansi": instansi,
                    "Nama Paket": nama_paket_clean,
                    "Tahapan": tahapan,
                    "Status": status_paket,
                    "HPS": hps_final,
                    "Metode": "Seleksi / Tender",
                    "Jenis Pemilihan": "Prakualifikasi / Pascakualifikasi",
                    "Evaluasi": "Kualitas & Biaya",
                    "Jenis Pengadaan": matching_kriteria,
                    "Tahun Anggaran": tahun,
                    "Pemenang Kontrak": pemenang,
                    "Nilai Kontrak": nilai_kontrak,
                    "Link": link_evaluasi_pemenang,
                    "Waktu Download": waktu_download_now,
                }
                data_lpse_ini.append(item)
                ditemukan_di_halaman += 1

          print(f"        └─ Ditemukan {ditemukan_di_halaman} paket cocok.")

          next_button = page.query_selector(
              "li.paginate_button.next:not(.disabled) a,"
              " button:has-text('Berikutnya'):not([disabled]),"
              " .next:not(.disabled) a"
          )

          if not next_button:
            break

          try:
            next_button.click()
            page.wait_for_timeout(1800)
            page_num += 1
          except Exception:
            break

      # AUTO-SAVE PER PORTAL LPSE
      if data_lpse_ini:
        df_lpse = pd.DataFrame(data_lpse_ini, columns=KOLOM_TARGET)
        save_and_update_excel(df_lpse, file_output)
        print(
            f"💾 [AUTO-SAVE] Data {lpse_nama} ({len(data_lpse_ini)} paket)"
            " berhasil diamankan ke Excel."
        )

    browser.close()


# ==============================================================================
# 5. EKSEKUSI PROGRAM
# ==============================================================================
if __name__ == "__main__":
  FILE_EXCEL_OUTPUT = "Hasil_Penarikan_LPSE_Nasional.xlsx"

  # Eksekusi Penarikan
  scrape_lpse_nasional(FILE_EXCEL_OUTPUT, max_pages_per_year=10)

  print("\n" + "=" * 70)
  print("✅ PROSES SELESAI SELURUHNYA!")
  print(f"📁 Seluruh data tersimpan rapat di: {FILE_EXCEL_OUTPUT}")
  print("=" * 70)