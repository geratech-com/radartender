import asyncio
import datetime
import io
import os
import random
import re
import sys
import time
import openpyxl
import pandas as pd
import plotly.express as px
import requests
from playwright.sync_api import sync_playwright
import streamlit as st

# ==============================================================================
# FIX KHUSUS WINDOWS + STREAMLIT + PLAYWRIGHT SUBPROCESS
# ==============================================================================
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# ==============================================================================
# 1. KONFIGURASI HALAMAN & CUSTOM CSS
# ==============================================================================
st.set_page_config(
    page_title="LPSE Market Intelligence (HPS >= 2.5M)",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .metric-card {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        border: 1px solid #334155;
        border-radius: 16px;
        padding: 20px;
        color: white;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-3px);
        border-color: #38BDF8;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #94A3B8;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #F8FAFC;
        margin-top: 5px;
    }
    .metric-sub {
        font-size: 0.8rem;
        color: #38BDF8;
        margin-top: 4px;
        font-weight: 500;
    }
    .hot-leads-header {
        background: linear-gradient(90deg, #B91C1C 0%, #C2410C 100%);
        color: white;
        padding: 12px 20px;
        border-radius: 12px;
        font-weight: bold;
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 15px;
    }
    .stAppHeader {background-color: rgba(0,0,0,0);}
    div[data-testid="stSidebar"] {
        background-color: #0F172A;
        border-right: 1px solid #1E293B;
    }
    </style>
""",
    unsafe_allow_html=True,
)

FILE_EXCEL_OUTPUT = "Hasil_Penarikan_LPSE_Nasional.xlsx"
GITHUB_RAW_URL = "https://raw.githubusercontent.com/geratech-com/radartender/main/Hasil_Penarikan_LPSE_Nasional.xlsx"
BATAS_MINIMAL_HPS = 2_500_000_000  # Rp 2,5 Miliar

DAFTAR_LPSE = [
    {"nama": "LPSE Nasional", "url": "https://spse.inaproc.id/nasional/lelang"},
    {"nama": "Kementerian PUPR", "url": "https://spse.inaproc.id/pu/lelang"},
    {"nama": "Otorita IKN (IKN)", "url": "https://spse.inaproc.id/ikn/lelang"},
    {"nama": "Otorita IKN (OIKN)", "url": "https://spse.inaproc.id/oikn/lelang"},
    {"nama": "LKPP Pusat", "url": "https://spse.inaproc.id/lkpp/lelang"},
    {"nama": "Kementerian Keuangan", "url": "https://spse.inaproc.id/kemenkeu/lelang"},
    {"nama": "Kementerian Perhub", "url": "https://spse.inaproc.id/dephub/lelang"},
    {"nama": "Kementerian Kesehatan", "url": "https://spse.inaproc.id/kemkes/lelang"},
    {"nama": "Kemendikbudristek", "url": "https://spse.inaproc.id/kemdikbud/lelang"},
    {"nama": "Kementerian Pertanian", "url": "https://spse.inaproc.id/pertanian/lelang"},
    {"nama": "Kementerian ESDM", "url": "https://spse.inaproc.id/esdm/lelang"},
    {"nama": "Kemenkumham", "url": "https://spse.inaproc.id/kemenkumham/lelang"},
    {"nama": "Kementerian Pertahanan", "url": "https://spse.inaproc.id/kemhan/lelang"},
    {"nama": "Kementerian Luar Negeri", "url": "https://spse.inaproc.id/kemlu/lelang"},
    {"nama": "Kemendagri", "url": "https://spse.inaproc.id/kemendagri/lelang"},
    {"nama": "Kementerian Agama", "url": "https://spse.inaproc.id/kemenag/lelang"},
    {"nama": "Kemnaker", "url": "https://spse.inaproc.id/kemnaker/lelang"},
    {"nama": "Kemenperin", "url": "https://spse.inaproc.id/kemenperin/lelang"},
    {"nama": "Kemendag", "url": "https://spse.inaproc.id/kemendag/lelang"},
    {"nama": "Kementerian LHK", "url": "https://spse.inaproc.id/menlhk/lelang"},
    {"nama": "KKP", "url": "https://spse.inaproc.id/kkp/lelang"},
    {"nama": "Kemendesa", "url": "https://spse.inaproc.id/kemendesa/lelang"},
    {"nama": "Kominfo", "url": "https://spse.inaproc.id/kominfo/lelang"},
    {"nama": "Kementerian BUMN", "url": "https://spse.inaproc.id/bumn/lelang"},
    {"nama": "Kemenkop UKM", "url": "https://spse.inaproc.id/kemenkop/lelang"},
    {"nama": "Kemenparekraf", "url": "https://spse.inaproc.id/kemenparekraf/lelang"},
    {"nama": "Kemensos", "url": "https://spse.inaproc.id/kemensos/lelang"},
    {"nama": "Bappenas", "url": "https://spse.inaproc.id/bappenas/lelang"},
    {"nama": "KemenPANRB", "url": "https://spse.inaproc.id/menpan/lelang"},
    {"nama": "ATR / BPN", "url": "https://spse.inaproc.id/atrbpn/lelang"},
    {"nama": "Kemenpora", "url": "https://spse.inaproc.id/kemenpora/lelang"},
    {"nama": "Kementerian PPPA", "url": "https://spse.inaproc.id/kemenpppa/lelang"},
    {"nama": "BKPM / Investasi", "url": "https://spse.inaproc.id/bkpm/lelang"},
    {"nama": "Mabes TNI", "url": "https://spse.inaproc.id/tni/lelang"},
    {"nama": "Mabes Polri", "url": "https://spse.inaproc.id/polri/lelang"},
    {"nama": "Kejaksaan Agung", "url": "https://spse.inaproc.id/kejaksaan/lelang"},
    {"nama": "Mahkamah Agung", "url": "https://spse.inaproc.id/mahkamahagung/lelang"},
    {"nama": "BPK RI", "url": "https://spse.inaproc.id/bpk/lelang"},
    {"nama": "BPKP", "url": "https://spse.inaproc.id/bpkp/lelang"},
    {"nama": "BSSN", "url": "https://spse.inaproc.id/bssn/lelang"},
    {"nama": "BMKG", "url": "https://spse.inaproc.id/bmkg/lelang"},
    {"nama": "BNPB", "url": "https://spse.inaproc.id/bnpb/lelang"},
    {"nama": "BNN", "url": "https://spse.inaproc.id/bnn/lelang"},
    {"nama": "KPU RI", "url": "https://spse.inaproc.id/kpu/lelang"},
    {"nama": "BRIN", "url": "https://spse.inaproc.id/brin/lelang"},
]

ALL_3_CATEGORIES = [
    "Jasa Konsultansi Badan Usaha Konstruksi",
    "Jasa Konsultansi Badan Usaha Non Konstruksi",
    "Pekerjaan Konstruksi Terintegrasi"
]

TAHAPAN_SPSE_RESMI = [
    "1. Pengumuman Prakualifikasi", "2. Download Dokumen Kualifikasi", "3. Penjelasan Dokumen Prakualifikasi",
    "4. Kirim Persyaratan Kualifikasi", "5. Evaluasi Dokumen Kualifikasi", "6. Pembuktian Kualifikasi",
    "7. Penetapan Hasil Kualifikasi", "8. Pengumuman Hasil Prakualifikasi", "9. Masa Sanggah Prakualifikasi",
    "10. Download Dokumen Pemilihan", "11. Pemberian Penjelasan", "12. Upload Dokumen Penawaran",
    "13. Pembukaan dan Evaluasi Penawaran File I: Administrasi dan Teknis", "14. Pengumuman Hasil Evaluasi Administrasi dan Teknis",
    "15. Pembukaan dan Evaluasi Penawaran File II: Harga", "16. Penetapan Pemenang", "17. Pengumuman Pemenang",
    "18. Masa Sanggah", "19. Surat Penunjukan Penyedia Barang/Jasa", "20. Penandatanganan Kontrak",
    "Tender Sudah Selesai", "Tender Batal", "Tender Gagal", "Seleksi Batal", "Seleksi Gagal", "Evaluasi Ulang", "Tender Ulang"
]

KOLOM_TARGET = [
    "Sumber LPSE", "ID LPSE", "Tanggal Pembuatan", "Instansi", "Nama Paket",
    "Tahapan", "HPS", "Metode", "Jenis Pemilihan", "Evaluasi",
    "Jenis Pengadaan", "Tahun Anggaran", "Pemenang Kontrak", "Nilai Kontrak",
    "Link", "Waktu Download"
]

INDONESIAN_MONTHS = {
    "januari": "Jan", "februari": "Feb", "maret": "Mar", "april": "Apr",
    "mei": "May", "juni": "Jun", "juli": "Jul", "agustus": "Aug",
    "september": "Sep", "oktober": "Oct", "november": "Nov", "desember": "Dec",
}

# ==============================================================================
# 3. HELPER NORMALISASI & CLEANER KETAT
# ==============================================================================
def normalize_jenis_pengadaan(raw_cat, nama_paket=""):
    c = str(raw_cat).lower().strip()
    np = str(nama_paket).lower().strip()
    
    if "lainnya" in c or "barang" in c or c == "pekerjaan konstruksi":
        return "INVALID"
        
    if any(k in np for k in ["software", "developer", "aplikasi", "sistem informasi", "tata naskah", "lisensi"]):
        if "konsultan" in np or "konsultansi" in c: return "Jasa Konsultansi Badan Usaha Non Konstruksi"
        return "INVALID"
        
    if "terintegrasi" in c or "terintegrasi" in np or "design & build" in np:
        return "Pekerjaan Konstruksi Terintegrasi"
        
    if "konsult" in c and "konstruksi" in c and "non" not in c:
        return "Jasa Konsultansi Badan Usaha Konstruksi"
        
    if "non" in c and "konsult" in c:
        return "Jasa Konsultansi Badan Usaha Non Konstruksi"
        
    if "konsult" in c:
        if any(k in np for k in ["konstruksi", "pembangunan", "gedung", "jalan", "jembatan", "irigasi", "renovasi"]):
            return "Jasa Konsultansi Badan Usaha Konstruksi"
        return "Jasa Konsultansi Badan Usaha Non Konstruksi"
        
    return "INVALID"

def parse_rupiah_pintar(text, target_keyword=None):
    if not text or str(text).strip() in ["-", "0", "Nilai Kontrak belum dibuat"]: return 0.0
    text_str = str(text).strip()
    
    if target_keyword:
        pattern = re.compile(rf"{target_keyword}\s*(?::|Rp\.?)?\s*([\d\.,]+(?:\s*(?:Miliar|Juta|Triliun|M|Jt|Rb|Ribu|T)\b)?)", re.IGNORECASE)
        match = pattern.search(text_str)
        if match: return parse_rupiah_pintar(match.group(1))

    match_unit = re.search(r"([\d\.,]+)\s*(Miliar|Juta|Triliun|M|Jt|Rb|Ribu|T)\b", text_str, re.IGNORECASE)
    if match_unit:
        raw_num = match_unit.group(1).replace(".", "").replace(",", ".")
        unit = match_unit.group(2).lower()
        mult = 1
        if "jt" in unit or "juta" in unit: mult = 1_000_000
        elif "m" in unit: mult = 1_000_000_000
        elif "t" in unit: mult = 1_000_000_000_000
        elif "rb" in unit or "ribu" in unit: mult = 1_000
        try:
            val = float(raw_num) * mult
            if val >= 1_000_000: return val
        except ValueError: pass

    matches = re.finditer(r"(?:Rp\s*\.?\s*)?(\d{1,3}(?:\.\d{3})+(?:,\d+)?)", text_str, re.IGNORECASE)
    valid_vals = []
    for m in matches:
        clean = m.group(1).split(",")[0].replace(".", "")
        try:
            val = float(clean)
            if val >= 1_000_000: valid_vals.append(val)
        except ValueError: pass
    
    if valid_vals: return min(valid_vals)
    return 0.0

def clean_df_master(df):
    if df.empty or "Jenis Pengadaan" not in df.columns: return df
    df = df.copy()
    df["Jenis Pengadaan"] = df.apply(lambda r: normalize_jenis_pengadaan(r["Jenis Pengadaan"], r.get("Nama Paket", "")), axis=1)
    return df[df["Jenis Pengadaan"].isin(ALL_3_CATEGORIES)].reset_index(drop=True)

def normalize_tahapan(tahapan_raw):
    t_clean = " ".join(re.sub(r"<[^>]+>", " ", str(tahapan_raw)).strip().split())
    t_lower = t_clean.lower()
    for t_resmi in TAHAPAN_SPSE_RESMI:
        if re.sub(r"^\d+\.\s*", "", t_resmi).lower() in t_lower: return t_resmi
    return t_clean

def format_rupiah_tabel(val):
    try:
        val = float(val)
        return "Rp 0,00" if val == 0 else f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception: return "Rp 0,00"

def format_rupiah_eksekutif(val):
    try:
        val = float(val)
        if val >= 1_000_000_000_000:
            return f"Rp {val/1_000_000_000_000:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " <span style='font-size: 1rem;'>Triliun</span>"
        elif val >= 1_000_000_000:
            return f"Rp {val/1_000_000_000:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " <span style='font-size: 1rem;'>Miliar</span>"
        return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception: return "Rp 0,00"

def standardize_lpse_link(raw_url, kode_lpse_default="pu"):
    if pd.isna(raw_url) or not str(raw_url).strip() or str(raw_url) == "-": return "-"
    url_str = str(raw_url).strip()
    match_id = re.search(r"/(?:lelang|evaluasi)/(\d+)", url_str) or re.search(r"(\d{7,10})", url_str)
    if not match_id: return "-"
    kode_id = match_id.group(1)
    match_instansi = re.search(r"spse\.inaproc\.id/([^/]+)", url_str)
    instansi = match_instansi.group(1) if match_instansi else kode_lpse_default
    return f"https://spse.inaproc.id/{instansi}/evaluasi/{kode_id}/pemenangberkontrak"

def fetch_detail_paket(context, base_domain, kode_id, base_referer):
    exact_hps, pemenang, nilai_kontrak, tgl_pembuatan = 0.0, "Belum Ditetapkan", 0.0, "-"
    
    js_pemenang_extractor = """
    () => {
        let p_name = "Belum Ditetapkan";
        let p_kontrak = "";
        let ths = document.querySelectorAll('th, td');
        for(let th of ths){
            let txt = th.innerText.toLowerCase().trim();
            if(txt === 'nama penyedia' || txt === 'nama pemenang' || txt.includes('pemenang berkontrak')){
                let td = th.nextElementSibling;
                if(td && td.tagName.toLowerCase() === 'td') p_name = td.innerText.split('\\n')[0].trim();
            }
            if(txt === 'harga kontrak' || txt === 'nilai kontrak' || txt === 'hasil negosiasi' || txt === 'harga penawaran'){
                let td = th.nextElementSibling;
                if(td && td.tagName.toLowerCase() === 'td') p_kontrak = td.innerText.trim();
            }
        }
        return {pemenang: p_name, kontrak: p_kontrak};
    }
    """

    try:
        dp = context.new_page()
        # LANGSUNG tembak ke Pemenang untuk hemat waktu!
        endpoints = [
            f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak",
            f"{base_domain}/evaluasi/{kode_id}/pemenang",
        ]
        for url in endpoints:
            try:
                time.sleep(1) # Jeda sopan agar tidak 429
                dp.goto(url, referer=base_referer, wait_until="domcontentloaded", timeout=15000)
                body_text = dp.inner_text("body")
                
                if "Terlalu Banyak Permintaan" in body_text:
                    time.sleep(20)
                    dp.goto(url, referer=base_referer, wait_until="domcontentloaded", timeout=15000)
                    body_text = dp.inner_text("body")
                    
                if "Akses Ditolak" in body_text or "OPPS!" in body_text: continue

                res_js = dp.evaluate(js_pemenang_extractor)
                if res_js['pemenang'] and res_js['pemenang'] != "Belum Ditetapkan":
                    pemenang = res_js['pemenang']
                if res_js['kontrak']:
                    val_k = parse_rupiah_pintar(str(res_js['kontrak']))
                    if val_k >= 1_000_000: nilai_kontrak = val_k
                        
                if pemenang != "Belum Ditetapkan" or nilai_kontrak > 0: break
            except Exception: continue
        dp.close()
    except Exception: pass

    return exact_hps, pemenang, nilai_kontrak, tgl_pembuatan

def save_and_update_excel(df_new, file_output):
    if not df_new.empty:
        df_new["ID LPSE"] = df_new["ID LPSE"].astype(str).str.strip()
        df_new["Sumber LPSE"] = df_new["Sumber LPSE"].astype(str).str.strip()

    if os.path.exists(file_output):
        try:
            df_existing = pd.read_excel(file_output)
            df_existing["ID LPSE"] = df_existing["ID LPSE"].astype(str).str.strip()
            df_existing["Sumber LPSE"] = df_existing["Sumber LPSE"].astype(str).str.strip()
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_final = df_combined.drop_duplicates(subset=["Sumber LPSE", "ID LPSE"], keep="last")
        except Exception: df_final = df_new
    else:
        df_final = df_new if not df_new.empty else pd.DataFrame(columns=KOLOM_TARGET)

    df_final = clean_df_master(df_final)

    with pd.ExcelWriter(file_output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Data LPSE Nasional")
        worksheet = writer.sheets["Data LPSE Nasional"]
        for row in range(2, len(df_final) + 2):
            worksheet[f"G{row}"].number_format = "#,##0.00"
            worksheet[f"N{row}"].number_format = "#,##0.00"

# ==============================================================================
# 4. SCRAPER ENGINE (MEMBACA LANGSUNG DARI LAYAR BROWSER)
# ==============================================================================
def run_scraper(selected_lpse, target_years, max_pages, log_container):
    if sys.platform == "win32": asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    try: loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    all_scraped_data = []

    with sync_playwright() as p:
        # BROWSER ASLI: Terlihat oleh layar (Bypass Cloudflare Block)
        browser = p.chromium.launch(headless=False, args=["--disable-dev-shm-usage", "--no-sandbox", "--disable-blink-features=AutomationControlled"])
        context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

        KAT_MAP = {
            "3": "Jasa Konsultansi Badan Usaha Non Konstruksi",
            "4": "Jasa Konsultansi Badan Usaha Konstruksi",
            "8": "Pekerjaan Konstruksi Terintegrasi"
        }

        for idx, lpse in enumerate(selected_lpse, 1):
            lpse_nama, lpse_url = lpse["nama"], lpse["url"]
            base_domain = lpse_url.replace("/lelang", "")
            candidates_dict = {}

            log_container.info(f"⚡ [{idx}/{len(selected_lpse)}] Membaca Layar SPSE secara 'Sopan': **{lpse_nama}**...")
            page = context.new_page()

            for tahun in target_years:
                for kat_id, kat_label in KAT_MAP.items():
                    url_query = f"{lpse_url}?tahun={tahun}&kategoriId={kat_id}"
                    
                    # LOGIKA JEDA ANTI-429
                    while True:
                        try:
                            time.sleep(1.5)
                            page.goto(url_query, wait_until="domcontentloaded", timeout=30000)
                            body_text = page.inner_text("body")
                            if "Terlalu Banyak Permintaan" in body_text or "Rate Limit" in body_text:
                                log_container.warning(f"⚠️ Terkena Limit (429) di {lpse_nama}. Istirahat 20 detik...")
                                time.sleep(20)
                                continue 
                            break 
                        except Exception:
                            break
                            
                    try: page.wait_for_selector('table tbody tr', timeout=10000)
                    except: continue
                        
                    try:
                        page.select_option('select[name$="_length"]', '100', timeout=3000)
                        time.sleep(2)
                    except: pass
                        
                    # LOGIKA KLIK HALAMAN NEXT
                    while True:
                        try: page.wait_for_selector('div.dataTables_processing', state='hidden', timeout=10000)
                        except: pass
                        
                        rows = page.evaluate("""
                            () => {
                                let results = [];
                                let trs = document.querySelectorAll('table tbody tr');
                                for (let tr of trs) {
                                    let tds = tr.querySelectorAll('td');
                                    if (tds.length >= 4) {
                                        if (tds[0].innerText.toLowerCase().includes('tidak ada')) continue;
                                        results.push({
                                            kode: tds[0].innerText.trim(),
                                            nama: tds[1].innerHTML.trim(),
                                            instansi: tds[2].innerText.trim(),
                                            tahapan: tds[3].innerText.trim(),
                                            hps: tds.length > 4 ? tds[4].innerText.trim() : ""
                                        });
                                    }
                                }
                                return results;
                            }
                        """)

                        if rows:
                            for row in rows:
                                kode_id = row.get("kode", "")
                                if not kode_id or not kode_id.isdigit() or len(kode_id) < 7: continue
                                if kode_id in candidates_dict: continue

                                cell_nama_raw = row.get("nama", "")
                                instansi = row.get("instansi", "")
                                tahapan_raw = row.get("tahapan", "")
                                hps_raw = row.get("hps", "")
                                
                                tabel_cat = ""
                                match_cat = re.search(r'(?:<br\s*/?>|<\/a>)\s*([A-Za-z\s&]+?)\s*-\s*TA', cell_nama_raw, re.IGNORECASE)
                                if match_cat: tabel_cat = match_cat.group(1).strip()
                                
                                match_link = re.search(r"<a[^>]*>(.*?)</a>", cell_nama_raw, re.DOTALL | re.IGNORECASE)
                                nama_paket = re.sub(r"<[^>]+>", "", match_link.group(1)).strip() if match_link else re.sub(r"<[^>]+>", " ", cell_nama_raw).strip()

                                final_cat = normalize_jenis_pengadaan(tabel_cat, nama_paket)
                                if final_cat == "INVALID": continue
                                    
                                hps_val = parse_rupiah_pintar(hps_raw, "HPS")
                                if hps_val < BATAS_MINIMAL_HPS: 
                                    hps_val = parse_rupiah_pintar(re.sub(r"<[^>]+>", " ", cell_nama_raw).strip(), "HPS")

                                if hps_val >= BATAS_MINIMAL_HPS:
                                    tahapan_clean = normalize_tahapan(tahapan_raw)
                                    nilai_kontrak_tabel = extract_nilai_kontrak_from_text(cell_nama_raw)
                                    link_evaluasi = f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak"

                                    candidates_dict[kode_id] = {
                                        "Sumber LPSE": lpse_nama, "ID LPSE": kode_id, "Tanggal Pembuatan": "-",
                                        "Instansi": instansi, "Nama Paket": nama_paket, "Tahapan": tahapan_clean,
                                        "HPS": float(hps_val), "Metode": "Seleksi / Tender",
                                        "Jenis Pemilihan": "Prakualifikasi", "Evaluasi": "Kualitas & Biaya",
                                        "Jenis Pengadaan": final_cat, "Tahun Anggaran": tahun, 
                                        "Pemenang Kontrak": "Belum Ditetapkan", "Nilai Kontrak": float(nilai_kontrak_tabel), 
                                        "Link": link_evaluasi, "Waktu Download": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    }
                        
                        next_btn = page.locator('.paginate_button.next')
                        if next_btn.count() > 0:
                            class_attr = next_btn.first.get_attribute("class") or ""
                            if "disabled" in class_attr:
                                break
                            else:
                                try:
                                    next_btn.first.click(timeout=5000)
                                    time.sleep(random.uniform(1.0, 2.0))
                                except:
                                    break
                        else:
                            break

            page.close()

            # Buka detail hanya untuk yang lolos saringan (Cepat)
            candidates_lpse = list(candidates_dict.values())
            if candidates_lpse:
                status_text = st.empty()
                log_container.info(f"🔍 Mencari Pemenang untuk {len(candidates_lpse)} Kandidat Valid ({lpse_nama})...")
                
                valid_candidates = []
                for i, cand in enumerate(candidates_lpse):
                    status_text.caption(f"Memproses {i+1}/{len(candidates_lpse)}: ID {cand['ID LPSE']}...")
                    
                    _, pemenang, nilai_kontrak_detail, _ = fetch_detail_paket(context, base_domain, cand["ID LPSE"], lpse_url)

                    if pemenang != "Belum Ditetapkan": cand["Pemenang Kontrak"] = pemenang
                    if nilai_kontrak_detail >= 1_000_000: cand["Nilai Kontrak"] = float(nilai_kontrak_detail)
                    valid_candidates.append(cand)

                status_text.empty()
                if valid_candidates:
                    df_lpse = pd.DataFrame(valid_candidates, columns=KOLOM_TARGET)
                    save_and_update_excel(df_lpse, FILE_EXCEL_OUTPUT)
                    all_scraped_data.extend(valid_candidates)
                    log_container.success(f"🎉 [{lpse_nama}] Selesai! ({len(valid_candidates)} Paket Diamankan)")

        browser.close()

    if all_scraped_data: save_and_update_excel(pd.DataFrame(all_scraped_data, columns=KOLOM_TARGET), FILE_EXCEL_OUTPUT)
    if not os.path.exists(FILE_EXCEL_OUTPUT): save_and_update_excel(pd.DataFrame(columns=KOLOM_TARGET), FILE_EXCEL_OUTPUT)
    return pd.read_excel(FILE_EXCEL_OUTPUT) if os.path.exists(FILE_EXCEL_OUTPUT) else pd.DataFrame(columns=KOLOM_TARGET)


# ==============================================================================
# 5. USER INTERFACE DASHBOARD EXECUTIVE
# ==============================================================================
st.title("💎 LPSE Market Intelligence (Khusus HPS ≥ Rp 2,5 Miliar)")
st.caption("Platform Radar Lelang Prioritas Tinggi — Memantau Hanya Proyek Bernilai Rp 2,5 Miliar Ke Atas.")

st.sidebar.markdown("### 🤖 Status Sistem")
st.sidebar.success("✅ Pembaruan Otomatis Aktif")
st.sidebar.info(
    "Database diperbarui secara otomatis setiap hari via **GitHub Actions**.\n\n"
    "Gunakan filter di halaman utama untuk menganalisis proyek."
)

st.sidebar.warning(
    "⚠️ **Disclaimer:**\n\n"
    "Data yang tersaji di platform ini merupakan hasil penarikan otomatis dari situs resmi SPSE. "
    "Mohon untuk tetap melakukan konfirmasi dan verifikasi ulang secara langsung pada portal resmi SPSE terkait ya, Sobat! 🙏"
)

st.sidebar.markdown("---")

# ==============================================================================
# 6. DISPLAY DATA & VISUALISASI MEWAH
# ==============================================================================
@st.cache_data(ttl=60)
def load_lpse_data():
    cache_buster_url = f"{GITHUB_RAW_URL}?v={int(time.time())}"
    try:
        response = requests.get(cache_buster_url, timeout=15)
        if response.status_code == 200: return pd.read_excel(io.BytesIO(response.content), engine="openpyxl")
    except Exception as e: st.warning(f"⚠️ Gagal memuat data dari GitHub: {e}")
    if os.path.exists(FILE_EXCEL_OUTPUT): return pd.read_excel(FILE_EXCEL_OUTPUT)
    return pd.DataFrame()

df_master = load_lpse_data()

if not df_master.empty:
    df_master["ID LPSE"] = df_master["ID LPSE"].astype(str).str.strip()
    df_master["Sumber LPSE"] = df_master["Sumber LPSE"].astype(str).str.strip()
    df_master = df_master.drop_duplicates(subset=["Sumber LPSE", "ID LPSE"], keep="last")
    df_master = clean_df_master(df_master)
    if "Tanggal Pembuatan" not in df_master.columns: df_master["Tanggal Pembuatan"] = "-"
    if "Link" in df_master.columns: df_master["Link"] = df_master["Link"].apply(standardize_lpse_link)

    df_master["HPS"] = pd.to_numeric(df_master["HPS"], errors="coerce").fillna(0.0)
    df_master["Nilai Kontrak"] = pd.to_numeric(df_master["Nilai Kontrak"], errors="coerce").fillna(0.0)
    df_master = df_master[df_master["HPS"] >= BATAS_MINIMAL_HPS]

    df_master["Tahapan"] = df_master["Tahapan"].apply(normalize_tahapan)
    df_master["Tgl_Pembuatan_Sort"] = df_master["Tanggal Pembuatan"].apply(parse_tgl_pembuatan)
    df_master["Waktu_Download_Sort"] = pd.to_datetime(df_master["Waktu Download"], errors="coerce")

    df_master = df_master.sort_values(
        by=["Tgl_Pembuatan_Sort", "Waktu_Download_Sort", "ID LPSE"],
        ascending=[False, False, False], na_position="last"
    ).drop(columns=["Tgl_Pembuatan_Sort", "Waktu_Download_Sort"], errors="ignore")

df_aktif = df_master[~df_master["Tahapan"].str.contains("Selesai|Batal|Gagal|Ulang", case=False, na=False)] if not df_master.empty else pd.DataFrame()

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.markdown(
        f"""<div class="metric-card"><div class="metric-label">🚨 Tender Aktif</div>
        <div class="metric-value">{len(df_aktif):,} <span style="font-size: 1rem;">Paket</span></div></div>""", unsafe_allow_html=True)
with col_m2:
    st.markdown(
        f"""<div class="metric-card"><div class="metric-label">💰 Potensi Pasar Aktif</div>
        <div class="metric-value">{format_rupiah_eksekutif(df_aktif["HPS"].sum() if not df_aktif.empty else 0)}</div></div>""", unsafe_allow_html=True)
with col_m3:
    st.markdown(
        f"""<div class="metric-card"><div class="metric-label">🏆 Total Kontrak</div>
        <div class="metric-value">{format_rupiah_eksekutif(df_master["Nilai Kontrak"].sum() if not df_master.empty else 0)}</div></div>""", unsafe_allow_html=True)
with col_m4:
    st.markdown(
        f"""<div class="metric-card"><div class="metric-label">📦 Total Paket</div>
        <div class="metric-value">{len(df_master):,} <span style="font-size: 1rem;">Paket</span></div></div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
KOLOM_DISPLAY = ["Sumber LPSE", "ID LPSE", "Tanggal Pembuatan", "Instansi", "Nama Paket", "Tahapan", "HPS", "Jenis Pengadaan", "Tahun Anggaran", "Pemenang Kontrak", "Nilai Kontrak", "Link"]

st.markdown('<div class="hot-leads-header">🔥 HOT OPPORTUNITIES (Tender Aktif HPS ≥ Rp 2,5 Miliar)</div>', unsafe_allow_html=True)
if not df_aktif.empty:
    df_hot_show = df_aktif.head(10).copy()
    cols_exist = [c for c in KOLOM_DISPLAY if c in df_hot_show.columns]
    df_hot_show = df_hot_show[cols_exist]
    df_hot_show["HPS"] = df_hot_show["HPS"].apply(format_rupiah_tabel)
    df_hot_show["Nilai Kontrak"] = df_hot_show["Nilai Kontrak"].apply(format_rupiah_tabel)
    st.dataframe(df_hot_show, column_config={"Link": st.column_config.LinkColumn("⚡ Akses LPSE")}, use_container_width=True, height=320)
else:
    st.info("Belum ditemukan tender aktif berkategori HPS ≥ Rp 2,5 Miliar.")

st.markdown("---")

st.subheader("📊 Analisis Pasar Proyek Besar (≥ Rp 2,5 Miliar)")
if not df_master.empty:
    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        fig_status = px.pie(df_master, names="Tahapan", hole=0.4, title="<b>Distribusi Tahapan Tender</b>")
        fig_status.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#F8FAFC"))
        st.plotly_chart(fig_status, use_container_width=True)
    with chart_col2:
        df_lpse_sum = df_master.groupby("Sumber LPSE")["HPS"].sum().reset_index().sort_values(by="HPS", ascending=False).head(7)
        fig_lpse = px.bar(df_lpse_sum, x="HPS", y="Sumber LPSE", orientation="h", title="<b>Top 7 LPSE</b>", color="HPS", color_continuous_scale="Blues")
        fig_lpse.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#F8FAFC"), xaxis_title="Total Nilai HPS (Rupiah)", yaxis_title="")
        st.plotly_chart(fig_lpse, use_container_width=True)
else:
    st.info("Visualisasi grafik belum tersedia karena data masih kosong.")

st.markdown("---")
st.subheader("🔍 Filter Database Proyek Utama")
if not df_master.empty:
    f_c1, f_c2, f_c3, f_c4 = st.columns(4)
    with f_c1: filter_lpse = st.multiselect("Filter Portal LPSE:", options=sorted(list(set([lpse["nama"] for lpse in DAFTAR_LPSE] + df_master["Sumber LPSE"].unique().tolist()))))
    with f_c2: filter_jenis = st.multiselect("Filter Jenis Pengadaan:", options=ALL_3_CATEGORIES)
    with f_c3: filter_tahapan = st.multiselect("Filter Tahapan Detail:", options=TAHAPAN_SPSE_RESMI)
    with f_c4: search_keyword = st.text_input("Cari Nama Paket / PT Pemenang:")

    df_filtered = df_master.copy()
    if filter_lpse: df_filtered = df_filtered[df_filtered["Sumber LPSE"].isin(filter_lpse)]
    if filter_jenis: df_filtered = df_filtered[df_filtered["Jenis Pengadaan"].isin(filter_jenis)]
    if filter_tahapan: df_filtered = df_filtered[df_filtered["Tahapan"].isin(filter_tahapan)]
    if search_keyword:
        df_filtered = df_filtered[df_filtered["Nama Paket"].astype(str).str.contains(search_keyword, case=False) | df_filtered["Pemenang Kontrak"].astype(str).str.contains(search_keyword, case=False)]

    cols_exist_filtered = [c for c in KOLOM_DISPLAY if c in df_filtered.columns]
    df_display_filtered = df_filtered[cols_exist_filtered].copy()
    df_display_filtered["HPS"] = df_display_filtered["HPS"].apply(format_rupiah_tabel)
    df_display_filtered["Nilai Kontrak"] = df_display_filtered["Nilai Kontrak"].apply(format_rupiah_tabel)
    st.dataframe(df_display_filtered, column_config={"Link": st.column_config.LinkColumn("Detail SPSE")}, use_container_width=True, height=450)
else:
    st.info("Belum ada data untuk ditampilkan dalam tabel.")

try:
    resp = requests.get(GITHUB_RAW_URL)
    excel_bytes = resp.content
except Exception:
    if os.path.exists(FILE_EXCEL_OUTPUT):
        with open(FILE_EXCEL_OUTPUT, "rb") as f: excel_bytes = f.read()
    else: excel_bytes = b""

st.download_button(
    label="📥 Download Master Excel Lelang (HPS ≥ 2.5M) (.xlsx)",
    data=excel_bytes,
    file_name="Hasil_Penarikan_LPSE_Nasional_Final.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)