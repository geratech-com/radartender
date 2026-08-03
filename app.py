import asyncio
import datetime
import os
os.system("playwright install chromium")
os.system("playwright install-deps chromium")
import re
import sys
import openpyxl
import pandas as pd
import plotly.express as px
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

CATEGORIES_MAP = {
    3: "Jasa Konsultansi Badan Usaha Non Konstruksi",
    4: "Jasa Konsultansi Badan Usaha Konstruksi",
    8: "Pekerjaan Konstruksi Terintegrasi"
}

ALL_3_CATEGORIES = [
    "Jasa Konsultansi Badan Usaha Konstruksi",
    "Jasa Konsultansi Badan Usaha Non Konstruksi",
    "Pekerjaan Konstruksi Terintegrasi"
]

TAHAPAN_SPSE_RESMI = [
    "1. Pengumuman Prakualifikasi",
    "2. Download Dokumen Kualifikasi",
    "3. Penjelasan Dokumen Prakualifikasi",
    "4. Kirim Persyaratan Kualifikasi",
    "5. Evaluasi Dokumen Kualifikasi",
    "6. Pembuktian Kualifikasi",
    "7. Penetapan Hasil Kualifikasi",
    "8. Pengumuman Hasil Prakualifikasi",
    "9. Masa Sanggah Prakualifikasi",
    "10. Download Dokumen Pemilihan",
    "11. Pemberian Penjelasan",
    "12. Upload Dokumen Penawaran",
    "13. Pembukaan dan Evaluasi Penawaran File I: Administrasi dan Teknis",
    "14. Pengumuman Hasil Evaluasi Administrasi dan Teknis",
    "15. Pembukaan dan Evaluasi Penawaran File II: Harga",
    "16. Penetapan Pemenang",
    "17. Pengumuman Pemenang",
    "18. Masa Sanggah",
    "19. Surat Penunjukan Penyedia Barang/Jasa",
    "20. Penandatanganan Kontrak",
    "Tender Sudah Selesai",
    "Tender Batal",
    "Tender Gagal",
    "Seleksi Batal",
    "Seleksi Gagal",
    "Evaluasi Ulang",
    "Tender Ulang"
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
# 3. HELPER PARSING PINTAR
# ==============================================================================
def parse_rupiah_pintar(text, target_keyword=None):
    if not text or str(text).strip() in ["-", "0", "Nilai Kontrak belum dibuat"]: return 0.0
    text_str = str(text).strip()
    
    if target_keyword:
        pattern = re.compile(rf"{target_keyword}\s*(?::|Rp\.?)?\s*([\d\.,]+(?:\s*(?:Miliar|Juta|Triliun|M|Jt|Rb|Ribu|T)\b)?)", re.IGNORECASE)
        match = pattern.search(text_str)
        if match:
            return parse_rupiah_pintar(match.group(1))

    match_unit = re.search(r"([\d\.,]+)\s*(Miliar|Juta|Triliun|M|Jt|Rb|Ribu|T)\b", text_str, re.IGNORECASE)
    if match_unit:
        raw_num = match_unit.group(1).replace(".", "").replace(",", ".")
        unit = match_unit.group(2).lower()
        mult = 1
        if "m" in unit: mult = 1_000_000_000
        elif "t" in unit: mult = 1_000_000_000_000
        elif "jt" in unit or "juta" in unit: mult = 1_000_000
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

def identify_jenis_pengadaan_flexible(row_text):
    text = row_text.lower()
    if "perorangan" in text: return None
    
    if "terintegrasi" in text:
        return "Pekerjaan Konstruksi Terintegrasi"
    elif "non konstruksi" in text or "non-konstruksi" in text:
        return "Jasa Konsultansi Badan Usaha Non Konstruksi"
    elif "konstruksi" in text or "supervisi" in text or "pengawasan" in text or "manajemen" in text or "konsultansi" in text:
        return "Jasa Konsultansi Badan Usaha Konstruksi"
    return None

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

    js_pengumuman_extractor = """
    () => {
        let hps = "";
        let tgl = "-";
        let ths = document.querySelectorAll('th');
        for(let th of ths){
            let txt = th.innerText.toLowerCase();
            if(txt.includes('tanggal pembuatan')){
                let td = th.nextElementSibling;
                if(td && td.tagName.toLowerCase() === 'td') tgl = td.innerText.trim();
            }
            if(txt.includes('hps paket')){
                let td = th.nextElementSibling;
                if(td && td.tagName.toLowerCase() === 'td') hps = td.innerText.trim();
            }
        }
        return {tgl: tgl, hps: hps};
    }
    """

    js_pemenang_extractor = """
    () => {
        let p_name = "Belum Ditetapkan";
        let p_kontrak = "";
        let ths = document.querySelectorAll('th');
        for(let th of ths){
            let txt = th.innerText.toLowerCase();
            if(txt.includes('nama pemenang') || txt.includes('pemenang berkontrak')){
                let td = th.nextElementSibling;
                if(td && td.tagName.toLowerCase() === 'td') p_name = td.innerText.split('\\n')[0].trim();
            }
            if(txt.includes('harga kontrak') || txt.includes('nilai kontrak') || txt.includes('hasil negosiasi')){
                let td = th.nextElementSibling;
                if(td && td.tagName.toLowerCase() === 'td') p_kontrak = td.innerText.trim();
            }
        }
        if(p_name === "Belum Ditetapkan" || p_name.toLowerCase().includes('nama pemenang')){
            let tables = document.querySelectorAll('table');
            for(let tbl of tables){
                let headers = Array.from(tbl.querySelectorAll('th'));
                let idxPem = headers.findIndex(h => h.innerText.toLowerCase().includes('pemenang'));
                let idxKon = headers.findIndex(h => h.innerText.toLowerCase().includes('kontrak') || h.innerText.toLowerCase().includes('negosiasi'));
                if(idxPem > -1){
                    let trs = tbl.querySelectorAll('tbody tr, tr');
                    for(let tr of trs){
                        let tds = tr.querySelectorAll('td');
                        if(tds.length > idxPem){
                            p_name = tds[idxPem].innerText.split('\\n')[0].trim();
                            if(idxKon > -1 && tds.length > idxKon){
                                p_kontrak = tds[idxKon].innerText.trim();
                            }
                            break; 
                        }
                    }
                }
            }
        }
        return {pemenang: p_name, kontrak: p_kontrak};
    }
    """

    try:
        dp = context.new_page()

        # 1. Halaman Pengumuman
        url_p = f"{base_domain}/lelang/{kode_id}/pengumumanlelang"
        try:
            dp.goto(url_p, referer=base_referer, wait_until="domcontentloaded", timeout=20000)
            res_pengumuman = dp.evaluate(js_pengumuman_extractor)
            
            if res_pengumuman['tgl'] and res_pengumuman['tgl'] != "-":
                m = re.search(r"([\d]{1,2}\s+[A-Za-z]+\s+[\d]{4})", res_pengumuman['tgl'])
                if m: tgl_pembuatan = m.group(1).strip()
                
            if res_pengumuman['hps']:
                val = parse_rupiah_pintar(str(res_pengumuman['hps']), "HPS")
                if val > 0: exact_hps = val
        except Exception: pass

        # 2. Halaman Evaluasi / Pemenang
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

    return exact_hps, pemenang, nilai_kontrak, tgl_pembuatan


def save_and_update_excel(df_new, file_output):
    if df_new.empty and not os.path.exists(file_output): return
    
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
    else: df_final = df_new

    with pd.ExcelWriter(file_output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Data LPSE Nasional")
        worksheet = writer.sheets["Data LPSE Nasional"]
        num_format = "#,##0.00"
        for row in range(2, len(df_final) + 2):
            worksheet[f"G{row}"].number_format = num_format
            worksheet[f"N{row}"].number_format = num_format

# ==============================================================================
# 4. SCRAPER ENGINE
# ==============================================================================
def run_scraper(selected_lpse, target_years, max_pages, log_container):
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    try: loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-dev-shm-usage", "--no-sandbox", "--disable-setuid-sandbox"],
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        # BANYAK MEMPERCEPAT: Blokir gambar, CSS, dan font yang tidak dibutuhkan
        context.route("**/*", lambda route: route.abort() if route.request.resource_type in ["image", "stylesheet", "font", "media"] else route.continue_())
        page = context.new_page()

        js_auto_fetcher = """
        async (args) => {
            const { baseDomain, tahun, maxPages } = args;
            let allRows = [];
            
            let targetCats = [];
            const selects = document.querySelectorAll('select');
            for (let sel of selects) {
                for (let opt of sel.options) {
                    let txt = opt.innerText.trim().toLowerCase();
                    if ((txt.includes('konsultansi') && txt.includes('badan usaha')) || txt.includes('terintegrasi')) {
                        if (opt.value && opt.value !== "") {
                            targetCats.push({ id: opt.value, label: opt.innerText.trim() });
                        }
                    }
                }
            }
            
            if (targetCats.length === 0) {
                targetCats = [
                    { id: "", label: "ALL" },
                    { id: "3", label: "Jasa Konsultansi Badan Usaha Non Konstruksi" },
                    { id: "4", label: "Jasa Konsultansi Badan Usaha Konstruksi" },
                    { id: "8", label: "Pekerjaan Konstruksi Terintegrasi" }
                ];
            } else {
                targetCats.unshift({ id: "", label: "ALL" });
            }
            
            for (const catObj of targetCats) {
                let start = 0; 
                const length = 100;
                let total = 1;
                let pagesFetched = 0;
                
                while (start < total && pagesFetched < maxPages) {
                    let url = `${baseDomain}/dt/lelang?draw=1&start=${start}&length=${length}&tahun=${tahun}`;
                    if (catObj.id !== "") {
                        url += `&kategoriId=${catObj.id}`;
                    }
                    try {
                        const resp = await fetch(url, { headers: { 'X-Requested-With': 'XMLHttpRequest' } });
                        if (!resp.ok) break;
                        const json = await resp.json();
                        if (json && json.data && Array.isArray(json.data)) {
                            const rowsWithCat = json.data.map(r => { 
                                return { rowData: r, categoryId: catObj.id, categoryLabel: catObj.label }; 
                            });
                            allRows = allRows.concat(rowsWithCat);
                            total = json.recordsFiltered || json.recordsTotal || json.data.length;
                            if (json.data.length === 0) break;
                        } else { break; }
                    } catch(e) { break; }
                    start += length;
                    pagesFetched++;
                    await new Promise(r => setTimeout(r, 300));
                }
            }
            return allRows;
        }
        """

        for idx, lpse in enumerate(selected_lpse, 1):
            lpse_nama = lpse["nama"]
            lpse_url = lpse["url"]
            base_domain = lpse_url.replace("/lelang", "")
            candidates_dict = {}

            log_container.info(f"⚡ [{idx}/{len(selected_lpse)}] Membedah Kategori & Menarik Data: **{lpse_nama}**...")

            try: page.goto(lpse_url, wait_until="domcontentloaded", timeout=30000)
            except Exception:
                log_container.warning(f"⚠️ Server Tidak Merespon ({lpse_nama})")
                continue

            for tahun in target_years:
                try: raw_dt_data = page.evaluate(js_auto_fetcher, {"baseDomain": base_domain, "tahun": tahun, "maxPages": max_pages})
                except Exception: raw_dt_data = []

                if raw_dt_data:
                    for item in raw_dt_data:
                        cat_label = item.get("categoryLabel", "")
                        row = item.get("rowData", [])
                        
                        if isinstance(row, list) and len(row) >= 4:
                            kode_id = str(row[0]).strip()
                            if kode_id in candidates_dict: continue
                                
                            cell_nama_raw = str(row[1]).strip()
                            instansi = str(row[2]).strip()
                            tahapan_raw = str(row[3]).strip()
                            hps_raw = str(row[4]).strip() if len(row) > 4 else cell_nama_raw
                            
                            clean_text = re.sub(r"<[^>]+>", " ", cell_nama_raw).strip()
                            tahapan_clean = normalize_tahapan(tahapan_raw)
                            
                            jenis_matched = None
                            if cat_label and cat_label != "ALL":
                                if "terintegrasi" in cat_label.lower():
                                    jenis_matched = "Pekerjaan Konstruksi Terintegrasi"
                                elif "non konstruksi" in cat_label.lower() or "non-konstruksi" in cat_label.lower():
                                    jenis_matched = "Jasa Konsultansi Badan Usaha Non Konstruksi"
                                elif "konstruksi" in cat_label.lower() and "non" not in cat_label.lower():
                                    jenis_matched = "Jasa Konsultansi Badan Usaha Konstruksi"
                            
                            if not jenis_matched:
                                jenis_matched = identify_jenis_pengadaan_flexible(clean_text)
                                if not jenis_matched: continue

                            hps_val = parse_rupiah_pintar(hps_raw, "HPS")
                            if hps_val < BATAS_MINIMAL_HPS: hps_val = parse_rupiah_pintar(clean_text, "HPS")

                            if hps_val >= BATAS_MINIMAL_HPS:
                                match_link = re.search(r"<a[^>]*>(.*?)</a>", cell_nama_raw, re.DOTALL | re.IGNORECASE)
                                if match_link: nama_paket = re.sub(r"<[^>]+>", "", match_link.group(1)).strip()
                                else: nama_paket = clean_text.split("spse")[0].split("TA 20")[0].strip()

                                nilai_kontrak_tabel = extract_nilai_kontrak_from_text(clean_text)
                                link_evaluasi = f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak"

                                candidates_dict[kode_id] = {
                                    "Sumber LPSE": lpse_nama, "ID LPSE": kode_id, "Tanggal Pembuatan": "-",
                                    "Instansi": instansi, "Nama Paket": nama_paket, "Tahapan": tahapan_clean,
                                    "HPS": float(hps_val), "Metode": "Seleksi / Tender",
                                    "Jenis Pemilihan": "Prakualifikasi / Pascakualifikasi", "Evaluasi": "Kualitas & Biaya",
                                    "Jenis Pengadaan": jenis_matched, "Tahun Anggaran": tahun, "Pemenang Kontrak": "Belum Ditetapkan",
                                    "Nilai Kontrak": float(nilai_kontrak_tabel), "Link": link_evaluasi,
                                    "Waktu Download": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                }

                # TAHAP 1B: DOM FALLBACK
                if not candidates_dict:
                    try:
                        target_url = f"{base_domain}/lelang?tahun={tahun}"
                        page.goto(target_url, wait_until="domcontentloaded", timeout=30000)
                        page.wait_for_timeout(1000)
                        select_jenis = page.query_selector("select[name*='kategori']")
                        if select_jenis:
                            select_jenis.select_option(value="") 
                            page.wait_for_timeout(2000)

                        select_length = page.query_selector("select[name*='length']")
                        if select_length:
                            select_length.select_option(value="100")
                            page.wait_for_timeout(1500)
                        
                        page_num = 1
                        while page_num <= max_pages:
                            page.wait_for_selector("table tbody tr", timeout=8000)
                            rows = page.query_selector_all("table tbody tr")
                            if not rows: break
                            
                            first_id = rows[0].inner_text().strip()
                            for row in rows:
                                r_txt = row.inner_text()
                                if "tidak ada data" in r_txt.lower() or "processing" in r_txt.lower(): continue
                                
                                cells = row.query_selector_all("td")
                                if len(cells) < 4: continue
                                
                                kode_id = cells[0].inner_text().strip()
                                if kode_id in candidates_dict: continue
                                
                                cell_nama = cells[1].inner_text().strip()
                                instansi = cells[2].inner_text().strip()
                                tahapan_raw = cells[3].inner_text().strip()
                                tahapan_clean = normalize_tahapan(tahapan_raw)
                                
                                hps_str = cells[4].inner_text().strip() if len(cells) > 4 else r_txt
                                
                                jenis_matched = identify_jenis_pengadaan_flexible(r_txt)
                                if not jenis_matched: continue

                                hps_val = parse_rupiah_pintar(hps_str, "HPS")
                                if hps_val < BATAS_MINIMAL_HPS: hps_val = parse_rupiah_pintar(cell_nama, "HPS")
                                
                                if hps_val >= BATAS_MINIMAL_HPS:
                                    match_link = re.search(r"<a[^>]*>(.*?)</a>", cell_nama, re.DOTALL | re.IGNORECASE)
                                    if match_link: nama_paket = re.sub(r"<[^>]+>", "", match_link.group(1)).strip()
                                    else: nama_paket = cell_nama.split("spse")[0].split("TA 20")[0].strip()

                                    nilai_kontrak_tabel = extract_nilai_kontrak_from_text(cell_nama)
                                    link_evaluasi = f"{base_domain}/evaluasi/{kode_id}/pemenangberkontrak"

                                    candidates_dict[kode_id] = {
                                        "Sumber LPSE": lpse_nama, "ID LPSE": kode_id, "Tanggal Pembuatan": "-",
                                        "Instansi": instansi, "Nama Paket": nama_paket, "Tahapan": tahapan_clean,
                                        "HPS": float(hps_val), "Metode": "Seleksi / Tender",
                                        "Jenis Pemilihan": "Prakualifikasi / Pascakualifikasi", "Evaluasi": "Kualitas & Biaya",
                                        "Jenis Pengadaan": jenis_matched, "Tahun Anggaran": tahun, "Pemenang Kontrak": "Belum Ditetapkan",
                                        "Nilai Kontrak": float(nilai_kontrak_tabel), "Link": link_evaluasi,
                                        "Waktu Download": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    }
                            
                            next_li = page.query_selector("#tbllelang_next, li.paginate_button.next, li.next")
                            if not next_li or "disabled" in (next_li.get_attribute("class") or "").lower(): break
                            next_link = next_li.query_selector("a") or next_li
                            try:
                                next_link.click()
                                page.wait_for_timeout(1000)
                                for _ in range(20):
                                    page.wait_for_timeout(500)
                                    curr_rows = page.query_selector_all("table tbody tr")
                                    if curr_rows and curr_rows[0].inner_text().strip() != first_id and "processing" not in curr_rows[0].inner_text().lower():
                                        break
                                page_num += 1
                            except: break
                    except Exception: pass

            # TAHAP 2: PENGAYAAN DETAIL
            candidates_lpse = list(candidates_dict.values())
            if candidates_lpse:
                status_text = st.empty()
                log_container.info(f"🔍 Mengambil Detail Presisi untuk {len(candidates_lpse)} paket dari ({lpse_nama})...")
                
                for i, cand in enumerate(candidates_lpse):
                    status_text.caption(f"Memproses {i+1}/{len(candidates_lpse)}: ID {cand['ID LPSE']}...")
                    
                    (exact_hps, pemenang, nilai_kontrak_detail, tgl_pembuatan) = fetch_detail_paket(
                        context, base_domain, cand["ID LPSE"], lpse_url
                    )

                    cand["Tanggal Pembuatan"] = tgl_pembuatan if tgl_pembuatan != "-" else "-"
                    if pemenang != "Belum Ditetapkan": cand["Pemenang Kontrak"] = pemenang
                    if nilai_kontrak_detail >= 1_000_000: cand["Nilai Kontrak"] = float(nilai_kontrak_detail)
                    if exact_hps > 0: cand["HPS"] = float(exact_hps)
                    
                status_text.empty()
                df_lpse = pd.DataFrame(candidates_lpse, columns=KOLOM_TARGET)
                save_and_update_excel(df_lpse, FILE_EXCEL_OUTPUT)
                log_container.success(f"🎉 [{lpse_nama}] Sukses Menemukan {len(candidates_lpse)} Paket Lengkap!")

        browser.close()

    if not os.path.exists(FILE_EXCEL_OUTPUT):
        df_empty = pd.DataFrame(columns=KOLOM_TARGET)
        save_and_update_excel(df_empty, FILE_EXCEL_OUTPUT)


# ==============================================================================
# 5. USER INTERFACE DASHBOARD EXECUTIVE
# ==============================================================================
st.title("💎 LPSE Market Intelligence (Khusus HPS ≥ Rp 2,5 Miliar)")
st.caption(
    "Platform Radar Lelang Prioritas Tinggi — Memantau Hanya Proyek Bernilai"
    " Rp 2,5 Miliar Ke Atas."
)

# KUNCI KEAMANAN (PIN) UNTUK MODE ADMINISTRATOR
st.sidebar.markdown("### 🔐 Otorisasi Akses")
mode_akses = st.sidebar.radio("Pilih Mode Pengguna:", ["Viewer (Hanya Lihat Data)", "Administrator (Update Data)"])

if mode_akses == "Administrator (Update Data)":
    pin_input = st.sidebar.text_input("Masukkan PIN Rahasia:", type="password")
    
    # 📌 PIN DEFAULT: admin123 (Silakan Anda ganti sesuka hati)
    if pin_input == "apriganggas":
        st.sidebar.success("✅ Akses Administrator Terbuka!")
        st.sidebar.markdown("---")
        
        # --- MENU KONTROL DATA (HANYA MUNCUL JIKA PIN BENAR) ---
        st.sidebar.markdown("### 🎛️ Pusat Kontrol Data (Admin)")
        list_nama_lpse = [item["nama"] for item in DAFTAR_LPSE]
        options_lpse_ui = ["ALL (Pilih Semua LPSE)"] + list_nama_lpse
        selected_lpse_input = st.sidebar.multiselect("Target Portal LPSE:", options=options_lpse_ui, default=["Kementerian PUPR"])

        if "ALL (Pilih Semua LPSE)" in selected_lpse_input: selected_names = list_nama_lpse
        else: selected_names = selected_lpse_input

        options_tahun_ui = ["ALL (Pilih Semua Tahun)", 2026, 2025, 2024, 2023, 2022]
        selected_tahun_input = st.sidebar.multiselect("Tahun Anggaran:", options=options_tahun_ui, default=[2025])

        if "ALL (Pilih Semua Tahun)" in selected_tahun_input: tahun_pilihan = [2026, 2025, 2024, 2023, 2022]
        else: tahun_pilihan = selected_tahun_input

        max_halaman = st.sidebar.number_input("Max Halaman Scan:", min_value=1, max_value=50, value=10)
        target_lpse_objects = [item for item in DAFTAR_LPSE if item["nama"] in selected_names]

        if st.sidebar.button("⚡ Sync & Update Database", type="primary"):
            if not target_lpse_objects or not tahun_pilihan:
                st.sidebar.error("Pilih minimal 1 LPSE & Tahun Anggaran!")
            else:
                log_box = st.empty()
                with st.spinner(f"Scanning {len(target_lpse_objects)} LPSE untuk Tahun {tahun_pilihan}..."):
                    run_scraper(target_lpse_objects, tahun_pilihan, max_halaman, log_box)
                st.sidebar.success("🎉 Penarikan Data Selesai!")
                st.rerun()
    elif pin_input != "":
        st.sidebar.error("❌ PIN Salah!")
else:
    st.sidebar.info("👋 **Mode Viewer Aktif**. Anda dapat menggunakan filter pencarian dan melihat visualisasi grafik di layar utama. Namun, akses untuk menarik atau mengubah data (Scraping) dikunci.")

st.sidebar.markdown("---")

# ==============================================================================
# 6. DISPLAY DATA & VISUALISASI MEWAH
# ==============================================================================
if os.path.exists(FILE_EXCEL_OUTPUT):
    df_master = pd.read_excel(FILE_EXCEL_OUTPUT)

    if not df_master.empty:
        # PEMBERSIHAN OTOMATIS TIPE DATA
        df_master["ID LPSE"] = df_master["ID LPSE"].astype(str).str.strip()
        df_master["Sumber LPSE"] = df_master["Sumber LPSE"].astype(str).str.strip()
        df_master = df_master.drop_duplicates(subset=["Sumber LPSE", "ID LPSE"], keep="last")

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

    # KARTU METRIK UTAMA
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

    # BANNER HIGHLIGHT: HOT SALES OPPORTUNITIES
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

    # VISUALISASI GRAFIK INTERAKTIF
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

    # FILTER & DATABASE TABEL INTERAKTIF
    st.subheader("🔍 Filter Database Proyek Utama")
    if not df_master.empty:
        f_c1, f_c2, f_c3, f_c4 = st.columns(4)
        with f_c1:
            filter_lpse = st.multiselect("Filter Portal LPSE:", options=df_master["Sumber LPSE"].unique())
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

    # TOMBOL DOWNLOAD EXCEL LENGKAP
    with open(FILE_EXCEL_OUTPUT, "rb") as f:
        excel_bytes = f.read()
    st.download_button(
        label="📥 Download Master Excel Lelang (HPS ≥ 2.5M) (.xlsx)",
        data=excel_bytes,
        file_name="Hasil_Penarikan_LPSE_Nasional_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("👋 Database belum tersedia. Silakan masuk sebagai Administrator untuk menarik data baru.")