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

# ==============================================================================
# 2. DAFTAR TARGET LPSE LENGKAP SE-INDONESIA & MASTER DATA
# ==============================================================================
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
    {"nama": "BKN Pusat", "url": "https://spse.inaproc.id/bkn/lelang"},
    {"nama": "Prov DKI Jakarta", "url": "https://spse.inaproc.id/jakarta/lelang"},
    {"nama": "Prov Jawa Barat", "url": "https://spse.inaproc.id/jabar/lelang"},
    {"nama": "Prov Jawa Tengah", "url": "https://spse.inaproc.id/jateng/lelang"},
    {"nama": "Prov D.I. Yogyakarta", "url": "https://spse.inaproc.id/jogjaprov/lelang"},
    {"nama": "Prov Jawa Timur", "url": "https://spse.inaproc.id/jatim/lelang"},
    {"nama": "Prov Banten", "url": "https://spse.inaproc.id/banten/lelang"},
    {"nama": "Prov Bali", "url": "https://spse.inaproc.id/baliprov/lelang"},
    {"nama": "Prov Aceh", "url": "https://spse.inaproc.id/aceh/lelang"},
    {"nama": "Prov Sumatera Utara", "url": "https://spse.inaproc.id/sumut/lelang"},
    {"nama": "Prov Sumatera Barat", "url": "https://spse.inaproc.id/sumbar/lelang"},
    {"nama": "Prov Riau", "url": "https://spse.inaproc.id/riau/lelang"},
    {"nama": "Prov Kepulauan Riau", "url": "https://spse.inaproc.id/kepri/lelang"},
    {"nama": "Prov Jambi", "url": "https://spse.inaproc.id/jambi/lelang"},
    {"nama": "Prov Sumatera Selatan", "url": "https://spse.inaproc.id/sumsel/lelang"},
    {"nama": "Prov Bangka Belitung", "url": "https://spse.inaproc.id/babel/lelang"},
    {"nama": "Prov Bengkulu", "url": "https://spse.inaproc.id/bengkulu/lelang"},
    {"nama": "Prov Lampung", "url": "https://spse.inaproc.id/lampung/lelang"},
    {"nama": "Prov Kalimantan Barat", "url": "https://spse.inaproc.id/kalbar/lelang"},
    {"nama": "Prov Kalimantan Tengah", "url": "https://spse.inaproc.id/kalteng/lelang"},
    {"nama": "Prov Kalimantan Selatan", "url": "https://spse.inaproc.id/kalsel/lelang"},
    {"nama": "Prov Kalimantan Timur", "url": "https://spse.inaproc.id/kaltim/lelang"},
    {"nama": "Prov Kalimantan Utara", "url": "https://spse.inaproc.id/kaltara/lelang"},
    {"nama": "Prov NTB", "url": "https://spse.inaproc.id/ntb/lelang"},
    {"nama": "Prov NTT", "url": "https://spse.inaproc.id/ntt/lelang"},
    {"nama": "Prov Sulawesi Utara", "url": "https://spse.inaproc.id/sulut/lelang"},
    {"nama": "Prov Sulawesi Tengah", "url": "https://spse.inaproc.id/sulteng/lelang"},
    {"nama": "Prov Sulawesi Selatan", "url": "https://spse.inaproc.id/sulsel/lelang"},
    {"nama": "Prov Sulawesi Tenggara", "url": "https://spse.inaproc.id/sultra/lelang"},
    {"nama": "Prov Gorontalo", "url": "https://spse.inaproc.id/gorontaloprov/lelang"},
    {"nama": "Prov Sulawesi Barat", "url": "https://spse.inaproc.id/sulbar/lelang"},
    {"nama": "Prov Maluku", "url": "https://spse.inaproc.id/maluku/lelang"},
    {"nama": "Prov Maluku Utara", "url": "https://spse.inaproc.id/malut/lelang"},
    {"nama": "Prov Papua", "url": "https://spse.inaproc.id/papua/lelang"},
    {"nama": "Prov Papua Barat", "url": "https://spse.inaproc.id/papuabarat/lelang"},
    {"nama": "Prov Papua Selatan", "url": "https://spse.inaproc.id/papuaselatan/lelang"},
    {"nama": "Prov Papua Tengah", "url": "https://spse.inaproc.id/papuatengah/lelang"},
    {"nama": "Prov Papua Pegunungan", "url": "https://spse.inaproc.id/papuapegunungan/lelang"},
    {"nama": "Prov Papua Barat Daya", "url": "https://spse.inaproc.id/papuabaratdaya/lelang"},
    {"nama": "Kota Surabaya", "url": "https://spse.inaproc.id/surabaya/lelang"},
    {"nama": "Kota Medan", "url": "https://spse.inaproc.id/medan/lelang"},
    {"nama": "Kota Makassar", "url": "https://spse.inaproc.id/makassar/lelang"},
    {"nama": "Kota Bandung", "url": "https://spse.inaproc.id/bandung/lelang"},
    {"nama": "Kota Semarang", "url": "https://spse.inaproc.id/semarang/lelang"},
    {"nama": "Kota Palembang", "url": "https://spse.inaproc.id/palembang/lelang"},
    {"nama": "Kota Tangerang", "url": "https://spse.inaproc.id/tangerangkota/lelang"},
    {"nama": "Kota Tangerang Selatan", "url": "https://spse.inaproc.id/tangerangselatankota/lelang"},
    {"nama": "Kota Bekasi", "url": "https://spse.inaproc.id/bekasikota/lelang"},
    {"nama": "Kota Depok", "url": "https://spse.inaproc.id/depok/lelang"},
    {"nama": "Kota Bogor", "url": "https://spse.inaproc.id/bogorkota/lelang"},
    {"nama": "Kota Batam", "url": "https://spse.inaproc.id/batam/lelang"},
    {"nama": "Kota Pekanbaru", "url": "https://spse.inaproc.id/pekanbaru/lelang"},
    {"nama": "Kota Bandar Lampung", "url": "https://spse.inaproc.id/bandarlampung/lelang"},
    {"nama": "Kota Denpasar", "url": "https://spse.inaproc.id/denpasarkota/lelang"},
    {"nama": "Kota Balikpapan", "url": "https://spse.inaproc.id/balikpapan/lelang"},
    {"nama": "Kota Samarinda", "url": "https://spse.inaproc.id/samarinda/lelang"},
    {"nama": "Kota Banjarmasin", "url": "https://spse.inaproc.id/banjarmasin/lelang"},
    {"nama": "Kota Manado", "url": "https://spse.inaproc.id/manado/lelang"},
    {"nama": "Kota Jayapura", "url": "https://spse.inaproc.id/jayapura/lelang"},
]

# 3 KATEGORI RESMI MUTLAK
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
def normalize_jenis_pengadaan(raw_cat, real_jenis="", nama_paket=""):
    """
    STANDARISASI PRESISI 3 KATEGORI RESMI TERPROTEKSI
    """
    rj_lower = str(real_jenis).lower().strip()
    np_lower = str(nama_paket).lower().strip()
    raw_lower = str(raw_cat).lower().strip()
    
    if rj_lower in ["jasa lainnya", "pengadaan barang", "pekerjaan konstruksi"] or "lainnya" in rj_lower or "barang" in rj_lower:
        return "INVALID"

    if "terintegrasi" in raw_lower or "terintegrasi" in rj_lower:
        if any(k in np_lower for k in ["software", "developer", "aplikasi", "sistem informasi", "tata naskah", "lisensi"]):
            if "konsultan" in np_lower or "konsultansi" in rj_lower:
                return "Jasa Konsultansi Badan Usaha Non Konstruksi"
            return "INVALID"
        return "Pekerjaan Konstruksi Terintegrasi"

    if "konstruksi" in raw_lower and "non" not in raw_lower and "konsult" in raw_lower:
        return "Jasa Konsultansi Badan Usaha Konstruksi"
    if "konstruksi" in rj_lower and "non" not in rj_lower and "konsult" in rj_lower:
        return "Jasa Konsultansi Badan Usaha Konstruksi"

    if "non" in raw_lower or "non" in rj_lower:
        return "Jasa Konsultansi Badan Usaha Non Konstruksi"

    full_text = f"{raw_lower} {rj_lower} {np_lower}"
    if any(k in full_text for k in ["konstruksi", "pembangunan", "gedung", "jalan", "jembatan", "irigasi", "renovasi", "pasram", "rumkit", "tpst", "spam", "dinas", "kantor", "rumah", "showroom", "ded", "supervisi", "manajemen konstruksi"]):
        if not any(k in np_lower for k in ["software", "developer", "aplikasi", "sistem informasi", "tata naskah"]):
            return "Jasa Konsultansi Badan Usaha Konstruksi"

    if "konsult" in raw_lower or "konsult" in rj_lower:
        return "Jasa Konsultansi Badan Usaha Non Konstruksi"

    return "INVALID"

def parse_rupiah_pintar(text, target_keyword=None):
    if not text or str(text).strip() in ["-", "0", "Nilai Kontrak belum dibuat"]: return 0.0
    text_str = str(text).strip()
    
    if target_keyword:
        pattern = re.compile(rf"{target_keyword}\s*(?::|Rp\.?)?\s*([\d\.,]+(?:\s*(?:Miliar|Juta|Triliun|M|Jt|Rb|Ribu|T)\b)?)", re.IGNORECASE)
        match = pattern.search(text_str)
        if match:
            return parse_rupiah_pintar(match.group(1))

    # SOLUSI BUG TRILIUNAN: Cek 'jt' sebelum 't'
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
        clean = m.group(1)
        if "," in clean: clean = clean.split(",")[0]
        clean = clean.replace(".", "")
        try:
            val = float(clean)
            if val >= 1_000_000:
                valid_vals.append(val)
        except ValueError: pass
    
    if valid_vals:
        return min(valid_vals)

    raw_matches = re.finditer(r"\b(\d{7,})\b", text_str)
    for m in raw_matches:
        try:
            val = float(m.group(1))
            return val
        except ValueError: pass
    
    return 0.0

def extract_nilai_kontrak_from_text(text):
    match = re.search(r"Nilai\s+Kontrak\s*:\s*([^\n]+)", text, re.IGNORECASE)
    if match: return parse_rupiah_pintar(match.group(1))
    return 0.0

def clean_df_master(df):
    if df.empty or "Jenis Pengadaan" not in df.columns:
        return df

    df = df.copy()
    df["Jenis Pengadaan"] = df.apply(
        lambda r: normalize_jenis_pengadaan(r["Jenis Pengadaan"], r.get("Jenis Pengadaan", ""), r.get("Nama Paket", "")), axis=1
    )

    return df[df["Jenis Pengadaan"].isin(ALL_3_CATEGORIES)].reset_index(drop=True)

def normalize_tahapan(tahapan_raw):
    t_clean = re.sub(r"<[^>]+>", " ", str(tahapan_raw)).strip()
    t_clean = " ".join(t_clean.split())
    t_lower = t_clean.lower()
    
    for t_resmi in TAHAPAN_SPSE_RESMI:
        t_keyword = re.sub(r"^\d+\.\s*", "", t_resmi).lower()
        if t_keyword in t_lower:
            return t_resmi
    return t_clean

def parse_tgl_pembuatan(tgl_str):
    if pd.isna(tgl_str) or str(tgl_str).strip() in ["-", ""]: return pd.NaT
    s = str(tgl_str).lower().strip()
    for id_m, en_m in INDONESIAN_MONTHS.items(): s = s.replace(id_m, en_m)
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def format_rupiah_tabel(val):
    try:
        val = float(val)
        if val == 0: return "Rp 0,00"
        return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception: return "Rp 0,00"

def format_rupiah_eksekutif(val):
    try:
        val = float(val)
        if val >= 1_000_000_000_000:
            v_triliun = val / 1_000_000_000_000
            str_val = f"{v_triliun:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"Rp {str_val} <span style='font-size: 1rem;'>Triliun</span>"
        elif val >= 1_000_000_000:
            v_miliar = val / 1_000_000_000
            str_val = f"{v_miliar:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"Rp {str_val} <span style='font-size: 1rem;'>Miliar</span>"
        else:
            str_val = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"Rp {str_val}"
    except Exception: return "Rp 0,00"

def standardize_lpse_link(raw_url, kode_lpse_default="pu"):
    if pd.isna(raw_url) or not str(raw_url).strip() or str(raw_url) == "-": return "-"
    url_str = str(raw_url).strip()
    match_id = re.search(r"/(?:lelang|evaluasi)/(\d+)", url_str)
    if not match_id: match_id = re.search(r"(\d{7,10})", url_str)
    if not match_id: return "-"
    kode_id = match_id.group(1)
    match_instansi = re.search(r"spse\.inaproc\.id/([^/]+)", url_str)
    instansi = match_instansi.group(1) if match_instansi else kode_lpse_default
    return f"https://spse.inaproc.id/{instansi}/evaluasi/{kode_id}/pemenangberkontrak"


def fetch_detail_paket(context, base_domain, kode_id, base_referer):
    exact_hps = 0.0
    pemenang = "Belum Ditetapkan"
    nilai_kontrak = 0.0
    tgl_pembuatan = "-"
    real_jenis_pengadaan = ""

    js_pengumuman_extractor = """
    () => {
        let hps = "";
        let tgl = "-";
        let jenis = "";
        
        let trs = document.querySelectorAll('table tr');
        for (let tr of trs) {
            let cells = tr.querySelectorAll('th, td');
            if (cells.length >= 2) {
                let label = cells[0].innerText.toLowerCase().trim();
                let val = cells[1].innerText.trim();
                if (label.includes('jenis pengadaan')) {
                    jenis = val;
                } else if (label.includes('tanggal pembuatan')) {
                    tgl = val;
                } else if (label.includes('hps paket') || label.includes('nilai hps')) {
                    hps = val;
                }
            }
        }
        return {tgl: tgl, hps: hps, jenis: jenis};
    }
    """

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
        
        if (p_name === "Belum Ditetapkan" || p_name === "") {
            let trs = document.querySelectorAll('table tbody tr');
            for(let tr of trs) {
                if(tr.innerHTML.includes('fa-star') || tr.innerHTML.includes('Pemenang') || tr.innerHTML.includes('icon-star')) {
                    let tds = tr.querySelectorAll('td');
                    if(tds.length >= 2) {
                        p_name = tds[1].innerText.split('\\n')[0].trim();
                        p_kontrak = tds[tds.length-1].innerText.trim();
                    }
                }
            }
        }
        return {pemenang: p_name, kontrak: p_kontrak};
    }
    """

    try:
        dp = context.new_page()
        url_p = f"{base_domain}/lelang/{kode_id}/pengumumanlelang"
        try:
            dp.goto(url_p, referer=base_referer, wait_until="domcontentloaded", timeout=20000)
            res_pengumuman = dp.evaluate(js_pengumuman_extractor)
            
            if res_pengumuman.get('jenis'):
                real_jenis_pengadaan = str(res_pengumuman['jenis']).strip()

            if res_pengumuman.get('tgl') and res_pengumuman['tgl'] != "-":
                m = re.search(r"([\d]{1,2}\s+[A-Za-z]+\s+[\d]{4})", res_pengumuman['tgl'])
                if m: tgl_pembuatan = m.group(1).strip()
                
            if res_pengumuman.get('hps'):
                val = parse_rupiah_pintar(str(res_pengumuman['hps']), "HPS")
                if val > 0: exact_hps = val
        except Exception: pass

        endpoints = [
            f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak",
            f"{base_domain}/evaluasi/{kode_id}/pemenang",
        ]
        for url in endpoints:
            try:
                dp.goto(url, referer=base_referer, wait_until="domcontentloaded", timeout=20000)
                body_text = dp.inner_text("body")
                if "Akses Ditolak" in body_text or "OPPS!" in body_text: continue

                res_js = dp.evaluate(js_pemenang_extractor)
                if res_js['pemenang'] and res_js['pemenang'] != "Belum Ditetapkan":
                    pem = res_js['pemenang']
                    if "belum" not in pem.lower() and "nama pemenang" not in pem.lower():
                        pemenang = pem
                
                if res_js['kontrak']:
                    val_k = parse_rupiah_pintar(str(res_js['kontrak']))
                    if val_k >= 1_000_000:
                        nilai_kontrak = val_k
                        
                if pemenang != "Belum Ditetapkan" or nilai_kontrak > 0: break
            except Exception: continue

        dp.close()
    except Exception: pass

    return exact_hps, pemenang, nilai_kontrak, tgl_pembuatan, real_jenis_pengadaan


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
        except Exception: 
            df_final = df_new
    else:
        df_final = df_new if not df_new.empty else pd.DataFrame(columns=KOLOM_TARGET)

    df_final = clean_df_master(df_final)

    with pd.ExcelWriter(file_output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Data LPSE Nasional")
        worksheet = writer.sheets["Data LPSE Nasional"]
        num_format = "#,##0.00"
        for row in range(2, len(df_final) + 2):
            worksheet[f"G{row}"].number_format = num_format
            worksheet[f"N{row}"].number_format = num_format

# ==============================================================================
# 4. SCRAPER ENGINE (NON-HEADLESS ANTI CLOUDFLARE)
# ==============================================================================
def run_scraper(selected_lpse, target_years, max_pages, log_container):
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    try: loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    all_scraped_data = []

    with sync_playwright() as p:
        # ----------------------------------------------------------------------
        # SILVER BULLET: headless=False agar wujud aslinya keluar & bypass Cloudflare!
        # ----------------------------------------------------------------------
        browser = p.chromium.launch(
            headless=False,
            args=[
                "--disable-dev-shm-usage", 
                "--no-sandbox", 
                "--disable-setuid-sandbox",
                "--disable-blink-features=AutomationControlled"
            ],
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )

        KAT_MAP = {
            "3": "Jasa Konsultansi Badan Usaha Non Konstruksi",
            "4": "Jasa Konsultansi Badan Usaha Konstruksi",
            "8": "Pekerjaan Konstruksi Terintegrasi"
        }

        js_auto_fetcher = """
        async (args) => {
            const { baseDomain, tahun, katId } = args;
            const endpoints = [
                `${baseDomain}/dt/lelang?draw=1&start=0&length=500&tahun=${tahun}&kategoriId=${katId}`,
                `${baseDomain}/dt/lelang?draw=1&start=0&length=500&kategoriId=${katId}`
            ];
            for (let url of endpoints) {
                try {
                    const resp = await fetch(url, { headers: { 'X-Requested-With': 'XMLHttpRequest' } });
                    if (resp.ok) {
                        const json = await resp.json();
                        if (json && json.data && Array.isArray(json.data) && json.data.length > 0) {
                            return json.data;
                        }
                    }
                } catch(e) {}
            }
            return [];
        }
        """

        for idx, lpse in enumerate(selected_lpse, 1):
            lpse_nama = lpse["nama"]
            lpse_url = lpse["url"]
            base_domain = lpse_url.replace("/lelang", "")
            candidates_dict = {}

            log_container.info(f"⚡ [{idx}/{len(selected_lpse)}] Menyedot Data API SPSE: **{lpse_nama}**...")

            page = context.new_page()

            try:
                page.goto(lpse_url, wait_until="domcontentloaded", timeout=25000)
            except Exception:
                pass

            for tahun in target_years:
                for kat_id, kat_label in KAT_MAP.items():
                    rows_data = []
                    try:
                        rows_data = page.evaluate(js_auto_fetcher, {"baseDomain": base_domain, "tahun": tahun, "katId": kat_id})
                    except Exception:
                        pass
                        
                    if rows_data:
                        for row in rows_data:
                            if isinstance(row, list) and len(row) >= 4:
                                kode_id = str(row[0]).strip()
                                if not kode_id.isdigit() or len(kode_id) < 7: continue
                                if kode_id in candidates_dict: continue

                                cell_nama_raw = str(row[1]).strip()
                                instansi = str(row[2]).strip()
                                tahapan_raw = str(row[3]).strip()
                                hps_raw = str(row[4]).strip() if len(row) > 4 else cell_nama_raw

                                clean_text = re.sub(r"<[^>]+>", " ", cell_nama_raw).strip()
                                tahapan_clean = normalize_tahapan(tahapan_raw)

                                hps_val = parse_rupiah_pintar(hps_raw, "HPS")
                                if hps_val < BATAS_MINIMAL_HPS: hps_val = parse_rupiah_pintar(clean_text, "HPS")

                                if hps_val >= BATAS_MINIMAL_HPS:
                                    match_link = re.search(r"<a[^>]*>(.*?)</a>", cell_nama_raw, re.DOTALL | re.IGNORECASE)
                                    if match_link: nama_paket = re.sub(r"<[^>]+>", "", match_link.group(1)).strip()
                                    else: nama_paket = clean_text.split("spse")[0].split("TA 20")[0].strip()

                                    nilai_kontrak_tabel = extract_nilai_kontrak_from_text(cell_nama_raw)
                                    link_evaluasi = f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak"

                                    candidates_dict[kode_id] = {
                                        "Sumber LPSE": lpse_nama, "ID LPSE": kode_id, "Tanggal Pembuatan": "-",
                                        "Instansi": instansi, "Nama Paket": nama_paket, "Tahapan": tahapan_clean,
                                        "HPS": float(hps_val), "Metode": "Seleksi / Tender",
                                        "Jenis Pemilihan": "Prakualifikasi / Pascakualifikasi", "Evaluasi": "Kualitas & Biaya",
                                        "Jenis Pengadaan": kat_label,
                                        "Tahun Anggaran": tahun, "Pemenang Kontrak": "Belum Ditetapkan",
                                        "Nilai Kontrak": float(nilai_kontrak_tabel), "Link": link_evaluasi,
                                        "Waktu Download": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    }

            page.close()

            candidates_lpse = list(candidates_dict.values())
            if candidates_lpse:
                status_text = st.empty()
                log_container.info(f"🔍 Memeriksa Detail {len(candidates_lpse)} Paket (≥ 2.5M) dari ({lpse_nama})...")
                
                valid_candidates = []
                for i, cand in enumerate(candidates_lpse):
                    status_text.caption(f"Memproses {i+1}/{len(candidates_lpse)}: ID {cand['ID LPSE']}...")
                    
                    (exact_hps, pemenang, nilai_kontrak_detail, tgl_pembuatan, real_jenis) = fetch_detail_paket(
                        context, base_domain, cand["ID LPSE"], lpse_url
                    )

                    # NORMALISASI KATEGORI RESMI SECARA PRESISI 100%
                    final_cat = normalize_jenis_pengadaan(cand["Jenis Pengadaan"], real_jenis, cand["Nama Paket"])

                    if final_cat not in ALL_3_CATEGORIES:
                        continue

                    cand["Jenis Pengadaan"] = final_cat
                    cand["Tanggal Pembuatan"] = tgl_pembuatan if tgl_pembuatan != "-" else "-"
                    if pemenang != "Belum Ditetapkan": cand["Pemenang Kontrak"] = pemenang
                    if nilai_kontrak_detail >= 1_000_000: cand["Nilai Kontrak"] = float(nilai_kontrak_detail)
                    if exact_hps > 0: cand["HPS"] = float(exact_hps)
                    
                    valid_candidates.append(cand)

                status_text.empty()
                if valid_candidates:
                    df_lpse = pd.DataFrame(valid_candidates, columns=KOLOM_TARGET)
                    save_and_update_excel(df_lpse, FILE_EXCEL_OUTPUT)
                    all_scraped_data.extend(valid_candidates)
                    log_container.success(f"🎉 [{lpse_nama}] Berhasil Menyimpan {len(valid_candidates)} Paket Valid!")

        browser.close()

    if all_scraped_data:
        df_all = pd.DataFrame(all_scraped_data, columns=KOLOM_TARGET)
        save_and_update_excel(df_all, FILE_EXCEL_OUTPUT)

    if not os.path.exists(FILE_EXCEL_OUTPUT):
        df_empty = pd.DataFrame(columns=KOLOM_TARGET)
        save_and_update_excel(df_empty, FILE_EXCEL_OUTPUT)

    if os.path.exists(FILE_EXCEL_OUTPUT):
        return pd.read_excel(FILE_EXCEL_OUTPUT)
    return pd.DataFrame(columns=KOLOM_TARGET)


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
# 6. DISPLAY DATA & VISUALISASI MEWAH (LANGSUNG DARI GITHUB RAW URL)
# ==============================================================================
@st.cache_data(ttl=60)
def load_lpse_data():
    cache_buster_url = f"{GITHUB_RAW_URL}?v={int(time.time())}"
    try:
        response = requests.get(cache_buster_url, timeout=15)
        if response.status_code == 200:
            return pd.read_excel(io.BytesIO(response.content), engine="openpyxl")
    except Exception as e:
        st.warning(f"⚠️ Gagal memuat data dari GitHub: {e}")
    
    if os.path.exists(FILE_EXCEL_OUTPUT):
        return pd.read_excel(FILE_EXCEL_OUTPUT)
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
        ascending=[False, False, False],
        na_position="last",
    ).drop(columns=["Tgl_Pembuatan_Sort", "Waktu_Download_Sort"], errors="ignore")

df_aktif = (
    df_master[~df_master["Tahapan"].str.contains("Selesai|Batal|Gagal|Ulang", case=False, na=False)]
    if not df_master.empty
    else pd.DataFrame()
)

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">🚨 Tender Aktif (Pipeline ≥ 2.5M)</div>
            <div class="metric-value">{len(df_aktif):,} <span style="font-size: 1rem;">Paket</span></div>
            <div class="metric-sub">Peluang Siap Diikuti</div>
        </div>
        """, unsafe_allow_html=True,
    )
with col_m2:
    val_hps_aktif = df_aktif["HPS"].sum() if not df_aktif.empty else 0.0
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">💰 Potensi Pasar Aktif</div>
            <div class="metric-value">{format_rupiah_eksekutif(val_hps_aktif)}</div>
            <div class="metric-sub">Estimasi HPS Berjalan</div>
        </div>
        """, unsafe_allow_html=True,
    )
with col_m3:
    val_kontrak = df_master["Nilai Kontrak"].sum() if not df_master.empty else 0.0
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">🏆 Total Kontrak Tersebar</div>
            <div class="metric-value">{format_rupiah_eksekutif(val_kontrak)}</div>
            <div class="metric-sub">Pemenang Ditetapkan</div>
        </div>
        """, unsafe_allow_html=True,
    )
with col_m4:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">📦 Total Paket (≥ 2,5M)</div>
            <div class="metric-value">{len(df_master):,} <span style="font-size: 1rem;">Paket</span></div>
            <div class="metric-sub">Lintas LPSE Nasional</div>
        </div>
        """, unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)
KOLOM_DISPLAY = [
    "Sumber LPSE", "ID LPSE", "Tanggal Pembuatan", "Instansi", "Nama Paket",
    "Tahapan", "HPS", "Jenis Pengadaan", "Tahun Anggaran", "Pemenang Kontrak",
    "Nilai Kontrak", "Link"
]

st.markdown(
    '<div class="hot-leads-header">🔥 HOT OPPORTUNITIES (Tender Aktif HPS'
    ' ≥ Rp 2,5 Miliar — Urut Tanggal Pembuatan Terbaru)</div>', unsafe_allow_html=True,
)
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
        fig_status = px.pie(
            df_master, names="Tahapan", hole=0.4, title="<b>Distribusi Tahapan Tender (≥ 2,5M)</b>"
        )
        fig_status.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#F8FAFC"))
        st.plotly_chart(fig_status, use_container_width=True)
    with chart_col2:
        df_lpse_sum = df_master.groupby("Sumber LPSE")["HPS"].sum().reset_index().sort_values(by="HPS", ascending=False).head(7)
        fig_lpse = px.bar(
            df_lpse_sum, x="HPS", y="Sumber LPSE", orientation="h",
            title="<b>Top 7 LPSE dengan Nilai Pasar Proyek Besar Terbesar</b>",
            color="HPS", color_continuous_scale="Blues"
        )
        fig_lpse.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#F8FAFC"), xaxis_title="Total Nilai HPS (Rupiah)", yaxis_title="")
        st.plotly_chart(fig_lpse, use_container_width=True)
else:
    st.info("Visualisasi grafik belum tersedia karena data masih kosong.")

st.markdown("---")

st.subheader("🔍 Filter Database Proyek Utama")
if not df_master.empty:
    f_c1, f_c2, f_c3, f_c4 = st.columns(4)
    with f_c1:
        semua_nama_lpse = sorted(list(set([lpse["nama"] for lpse in DAFTAR_LPSE] + df_master["Sumber LPSE"].unique().tolist())))
        filter_lpse = st.multiselect("Filter Portal LPSE:", options=semua_nama_lpse)
    with f_c2:
        filter_jenis = st.multiselect("Filter Jenis Pengadaan:", options=ALL_3_CATEGORIES)
    with f_c3:
        filter_tahapan = st.multiselect("Filter Tahapan Detail:", options=TAHAPAN_SPSE_RESMI)
    with f_c4:
        search_keyword = st.text_input("Cari Nama Paket / PT Pemenang:", placeholder="Ketik kata kunci...")

    df_filtered = df_master.copy()
    if filter_lpse: df_filtered = df_filtered[df_filtered["Sumber LPSE"].isin(filter_lpse)]
    if filter_jenis: df_filtered = df_filtered[df_filtered["Jenis Pengadaan"].isin(filter_jenis)]
    if filter_tahapan: df_filtered = df_filtered[df_filtered["Tahapan"].isin(filter_tahapan)]
    if search_keyword:
        df_filtered = df_filtered[
            df_filtered["Nama Paket"].astype(str).str.contains(search_keyword, case=False) |
            df_filtered["Pemenang Kontrak"].astype(str).str.contains(search_keyword, case=False)
        ]

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
        with open(FILE_EXCEL_OUTPUT, "rb") as f:
            excel_bytes = f.read()
    else:
        excel_bytes = b""

st.download_button(
    label="📥 Download Master Excel Lelang (HPS ≥ 2.5M) (.xlsx)",
    data=excel_bytes,
    file_name="Hasil_Penarikan_LPSE_Nasional_Final.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)