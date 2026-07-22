import streamlit as st
import pandas as pd
from io import BytesIO
from num2words import num2words
from xhtml2pdf import pisa
from datetime import timedelta, date, datetime
import base64
import json
import os
import gspread # TAMBAHAN
from google.oauth2.service_account import Credentials # TAMBAHAN

# --- ATURAN HAK AKSES KODE REKENING ---
ATURAN_REKENING = {
    "Lugas Khalid Maulana": ("1",),
    "Puspita Sari Handayani, SE": ("2",),
    "Agus Setiawan, S.Pd": ("3",),
    "Ramli Nur Aziza": ("4",),
    "Galur Rismawati, SE": ("5",),
    "Herman Dwi": ("1", "2", "3", "4", "5")
}

@st.cache_data
def load_data():
    try:
        url_sheet = "https://docs.google.com/spreadsheets/d/1pqEUk3sjkrQKeuyWYpCGp_PU-iKfgDYtMu2WN7lixjM/export?format=xlsx"
        
        df_user = pd.read_excel(url_sheet, sheet_name="User")
        df_dpa = pd.read_excel(url_sheet, sheet_name="Master_DPA")
                
        df_user.columns = df_user.columns.astype(str).str.strip()
        df_dpa.columns = df_dpa.columns.astype(str).str.strip()
        
        return df_user, df_dpa
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        return pd.DataFrame(), pd.DataFrame()

# --- FUNGSI KONEKSI GOOGLE SHEET UNTUK MENYIMPAN ---
def connect_gsheet():
    scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    try:
        creds = Credentials.from_service_account_file('credentials.json', scopes=scopes)
        client = gspread.authorize(creds)
        # Buka berdasarkan URL atau ID
        sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1pqEUk3sjkrQKeuyWYpCGp_PU-iKfgDYtMu2WN7lixjM/edit")
        worksheet = sheet.worksheet("Memory_PBJ")
        return worksheet
    except Exception as e:
        st.error(f"Gagal koneksi ke Google Sheet. Pastikan credentials.json ada. Error: {e}")
        return None

def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        return ""

logo_ngawi_b64 = get_base64_image("logo_ngawi.png")

st.set_page_config(page_title="PBJ SUMENGKO Generator", layout="wide")
st.title("Aplikasi Pembuat Dokumen PBJ Desa Sumengko")

# ---------------------------------------------------------
# FITUR MEMORI CLOUD (GOOGLE SHEET) DI SIDEBAR
# ---------------------------------------------------------
with st.sidebar:
    st.header("☁️ Memori Cloud (G-Sheet)")
    st.write("Simpan/muat progres isian form ke Sheet 'Memory_PBJ'.")
    
    st.markdown("---")
    st.subheader("📥 Simpan Data Baru")
    nama_file_baru = st.text_input("Ketik nama / ID Draft:", value="draft_pbj_001")
    
    if st.button("Simpan Data ke Cloud", type="primary"):
        state_to_save = {}
        for k, v in st.session_state.items():
            if k.startswith("cetak_"):
                continue
            if isinstance(v, (str, int, float, bool, list)):
                state_to_save[k] = v
            elif isinstance(v, date):
                state_to_save[k] = v.isoformat()
        
        json_string = json.dumps(state_to_save)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        ws = connect_gsheet()
        if ws:
            # Simpan 3 kolom: Waktu, Nama_File, Data_JSON
            ws.append_row([timestamp, nama_file_baru, json_string])
            st.success(f"Data {nama_file_baru} berhasil disimpan ke Google Sheet!")
            # st.rerun() dihapus agar tidak me-refresh paksa sebelum sukses dibaca

    st.markdown("---")
    st.subheader("📂 Muat Data Tersimpan")
    
    ws = connect_gsheet()
    if ws:
        try:
            # Ambil semua data (Baris 1 = Header, Baris 2 dst = Data)
            all_values = ws.get_all_values()
            if len(all_values) > 1:
                # Ambil kolom ke-2 (index 1) sebagai daftar nama file
                list_file_backup = [row[1] for row in all_values[1:]]
                
                file_terpilih = st.selectbox("Pilih file yang ingin dimuat:", list_file_backup[::-1]) # Dibalik agar terbaru di atas
                
                if st.button("Muat Data Terpilih"):
                    # Cari string JSON dari baris yang dipilih
                    json_data = None
                    for row in all_values[1:]:
                        if row[1] == file_terpilih:
                            json_data = row[2]
                    
                    if json_data:
                        loaded_state = json.loads(json_data)
                        for k, v in loaded_state.items():
                            if k.startswith("cetak_"):
                                continue
                            if isinstance(v, str) and ("tgl" in k.lower()):
                                try:
                                    st.session_state[k] = date.fromisoformat(v)
                                except ValueError:
                                    st.session_state[k] = v
                            else:
                                st.session_state[k] = v
                        
                        st.success(f"Data dari {file_terpilih} berhasil dimuat!")
                        st.rerun()
            else:
                st.info("Belum ada data di Sheet Memory_PBJ.")
        except Exception as e:
            st.warning("Pastikan Sheet bernama 'Memory_PBJ' sudah dibuat.")

def get_tanggal_indo(tgl):
    if pd.isna(tgl) or tgl is None:
        return "-", "-", "-", "-", "-"
    
    nama_bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                  "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    nama_hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    
    hari = nama_hari[tgl.weekday()]
    tanggal = tgl.day
    bulan = nama_bulan[tgl.month - 1]
    tahun = tgl.year
    
    format_lengkap = f"{tanggal} {bulan} {tahun}"
    return format_lengkap, hari, str(tanggal), bulan, str(tahun)

df_user, df_dpa = load_data()

if not df_user.empty and not df_dpa.empty:
    tab_utama, tab_paket, tab_cetak, tab_pelaksanaan, tab_pelaporan1, tab_pelaporan2 = st.tabs([
        "1. Persiapan - Data Utama", 
        "2. Persiapan - Paket Belanja", 
        "3. Persiapan - Cetak",
        "4. Pelaksanaan - Dokumen",
        "5. Pelaporan 1 - Penyedia",
        "6. Pelaporan 2 - Kades"
    ])

    with tab_utama:
        st.header("A. Pilih Pelaksana (User)")
        daftar_user = df_user['Nama_Lengkap_KaurKasi'].tolist()
        pilih_user = st.selectbox("Pilih Pelaksana Kegiatan:", daftar_user, key="pilih_user")
        
        data_terpilih = df_user[df_user['Nama_Lengkap_KaurKasi'] == pilih_user].iloc[0]
        role_user = data_terpilih['Role_KaurKasi']
        nik_user = data_terpilih['NIK_KaurKasi']
        alamat_user = data_terpilih.get('Alamat', 'Desa Sumengko')
        
        st.info(f"**Jabatan:** {role_user} | **NIK:** {nik_user}")

        st.header("B. Data Kegiatan & RAB Siskeudes")
        col1, col2 = st.columns(2)
        with col1:
            try:
                bidang = st.selectbox("Bidang:", df_dpa['Bidang'].unique(), key="bidang")
                kegiatan = st.selectbox("Kegiatan:", df_dpa[df_dpa['Bidang'] == bidang]['Kegiatan'].unique(), key="kegiatan")
            except KeyError:
                st.error("Pastikan Excel memiliki kolom 'Bidang' dan 'Kegiatan'")
                bidang, kegiatan = "Kosong", "Kosong"
                
            tgl_persiapan = st.date_input("Tanggal Awal Persiapan", key="tgl_persiapan")
        
        with col2:
            hari_persiapan = st.number_input("Jumlah Hari Persiapan", min_value=1, key="hari_persiapan")
            hari_pelaksanaan = st.number_input("Jumlah Hari Pelaksanaan", min_value=1, key="hari_pelaksanaan")
            hari_pelaporan = st.number_input("Jumlah Hari Pelaporan", min_value=1, key="hari_pelaporan")
            total_hari = hari_persiapan + hari_pelaksanaan + hari_pelaporan
            st.write(f"**Total Waktu Pelaksanaan:** {total_hari} Hari")

        st.subheader("RAB Kegiatan (Telah Difilter Berdasarkan Akses User)")
        rab_kegiatan = pd.DataFrame()
        daftar_uraian = [] 
        
        if 'Kegiatan' in df_dpa.columns:
            # Ambil RAB mentah
            rab_mentah = df_dpa[df_dpa['Kegiatan'] == kegiatan].copy()
            
            # --- FILTER BERDASARKAN HAK AKSES USER ---
            rab_mentah['Kode_Rekening'] = rab_mentah['Kode_Rekening'].astype(str)
            awalan_diizinkan = ATURAN_REKENING.get(pilih_user)
            
            if awalan_diizinkan:
                rab_kegiatan = rab_mentah[rab_mentah['Kode_Rekening'].str.startswith(awalan_diizinkan)]
            else:
                rab_kegiatan = rab_mentah # Jika nama user tidak ada di dict, tampilkan semua (bisa diubah jadi kosong jika ingin blokir total)
            # ----------------------------------------
            
            if rab_kegiatan.empty:
                st.warning(f"Tidak ada data RAB untuk kegiatan ini yang sesuai dengan Hak Akses Kode Rekening Anda ({awalan_diizinkan}).")
            else:
                st.dataframe(rab_kegiatan, use_container_width=True)
                if 'Uraian' in rab_kegiatan.columns:
                    daftar_uraian = rab_kegiatan['Uraian'].tolist()
                else:
                    st.warning("Kolom 'Uraian' tidak ditemukan di Excel.")
        
        # (Kode Generate PDF RAB SISKEUDES dsb tetap SAMA PERSIS seperti milik Anda)
        st.markdown("---")
        st.subheader("Cetak Dokumen RAB Siskeudes")
        if st.button("Generate PDF - RAB Siskeudes", type="primary"):
            total_anggaran_rab = 0
            if not rab_kegiatan.empty:
                kolom_pencarian = [col for col in rab_kegiatan.columns if 'jumlah' in col.lower() or 'anggaran' in col.lower()]
                if kolom_pencarian:
                    total_anggaran_rab = pd.to_numeric(rab_kegiatan[kolom_pencarian[-1]], errors='coerce').sum() 
            
            terbilang_rab = num2words(total_anggaran_rab, lang='id').title() + " Rupiah"
            tanggal_format, _, _, _, _ = get_tanggal_indo(tgl_persiapan)

            baris_tabel_rab_html = ""
            for index, row in rab_kegiatan.iterrows():
                kode = row.get('Kode_Rekening', '-')
                uraian = row.get('Uraian', '-')
                vol = row.get('Vol_Rencana', '-')
                satuan = row.get('Satuan_Rencana', '-')
                harga = row.get('Harga_Rencana', '-')
                jumlah = row.get(kolom_pencarian[-1] if kolom_pencarian else 'Anggaran_Rencana', '-')
                
                harga_val = pd.to_numeric(harga, errors='coerce')
                jumlah_val = pd.to_numeric(jumlah, errors='coerce')
                harga_str = f"Rp {harga_val:,.0f}".replace(",", ".") if not pd.isna(harga_val) else "-"
                jumlah_str = f"Rp {jumlah_val:,.0f}".replace(",", ".") if not pd.isna(jumlah_val) else "-"

                baris_tabel_rab_html += f"""
                <tr>
                    <td>{kode}</td>
                    <td>{uraian}</td>
                    <td style="text-align: center;">{vol}</td>
                    <td style="text-align: center;">{satuan}</td>
                    <td>{harga_str}</td>
                    <td>{jumlah_str}</td>
                </tr>
                """

            html_rab = f"""
            <html>
            <head>
                <style>
                    @page {{ size: a4 portrait; margin: 2cm; }}
                    body {{ font-family: Helvetica, Arial, sans-serif; font-size: 11px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
                    th, td {{ border: 1px solid black; padding: 6px; text-align: left; vertical-align: middle; }}
                    th {{ background-color: #f2f2f2; text-align: center; font-weight: bold; }}
                    .judul {{ text-align: center; font-weight: bold; font-size: 16px; margin-bottom: 20px; }}
                    .tabel-header td {{ border: none; padding: 3px; font-weight: bold; }}
                    .tabel-ttd td {{ border: none; text-align: center; width: 33.33%; padding-top: 40px; }}
                </style>
            </head>
            <body>
                <div class="judul">RAB DPA (SISKEUDES)</div>
                <table class="tabel-header">
                    <tr><td style="width: 150px;">Bidang</td><td>: {bidang}</td></tr>
                    <tr><td>Kegiatan</td><td>: {kegiatan}</td></tr>
                    <tr><td>Waktu Pelaksanaan</td><td>: {total_hari} Hari</td></tr>
                </table>
                <table style="margin-top: 15px;">
                    <tr>
                        <th rowspan="2">Kode</th>
                        <th rowspan="2">Uraian</th>
                        <th colspan="4">Rincian</th>
                    </tr>
                    <tr>
                        <th>Volume</th>
                        <th>Satuan</th>
                        <th>Harga Satuan</th>
                        <th>Jumlah (Anggaran Rencana)</th>
                    </tr>
                    {baris_tabel_rab_html}
                    <tr>
                        <td colspan="5" style="text-align: center;"><b>JUMLAH</b></td>
                        <td><b>Rp {total_anggaran_rab:,.0f}</b></td>
                    </tr>
                </table>
                <p><b>Terbilang : </b> <i>{terbilang_rab}</i></p>
                <table class="tabel-ttd">
                    <tr>
                        <td>
                            Disetujui<br>Kepala Desa<br><br><br><br><br>
                            <b>Drs.H. SUMARLI, M.Ag</b>
                        </td>
                        <td>
                            Telah diverifikasi<br>Sekretaris Desa<br><br><br><br><br>
                            <b>Herman Dwi N,ST</b>
                        </td>
                        <td>
                            Sumengko, {tanggal_format}<br>Pelaksana Kegiatan Anggaran<br><br><br><br><br>
                            <b><u>{pilih_user}</u></b>
                        </td>
                    </tr>
                </table>
            </body>
            </html>
            """
            
            pdf_rab_output = BytesIO()
            pisa.CreatePDF(BytesIO(html_rab.encode('utf-8')), dest=pdf_rab_output)
            
            st.success("PDF RAB Siskeudes berhasil dibuat!")
            st.download_button(
                label="Unduh PDF RAB Siskeudes",
                data=pdf_rab_output.getvalue(),
                file_name=f"RAB_Siskeudes_{kegiatan}.pdf",
                mime="application/pdf"
            )

    # === TAB 2 SAMPAI BAWAH TIDAK ADA PERUBAHAN LOGIKA, KARENA DATA 'daftar_uraian' SUDAH TERFILTER DARI ATAS ===
    # KODE UNTUK TAB 2, 3, 4, 5, 6 SAMA SEPERTI KODE ASLI ANDA
    with tab_paket:
        st.header("Pembagian HPS dan Spesifikasi per Supplier")
        jml_paket = st.number_input("Berapa Paket Belanja / Supplier?", min_value=1, max_value=10, step=1, key="jml_paket")
        tabs_supplier = st.tabs([f"Paket Belanja {i+1}" for i in range(jml_paket)])
        paket_data = [] 

        for i, tab in enumerate(tabs_supplier):
            with tab:
                nama_paket = st.text_input(f"Nama Paket {i+1}", key=f"nama_paket_{i}")
                col_kiri, col_kanan = st.columns(2)
                with col_kiri:
                    latar_belakang = st.text_area("Latar Belakang (KAK)", key=f"latar_{i}")
                    penerima = st.text_area("Penerima Manfaat", key=f"penerima_{i}")
                with col_kanan:
                    mekanisme_bayar = st.selectbox("Mekanisme Pembayaran KAK", ["Sekaligus (100%)", "Termin 2", "Termin 3"], key=f"bayar_{i}")
                    spesifikasi_umum = st.text_area("Spesifikasi Teknis (Paragraf Umum KAK)", key=f"spek_umum_{i}")

                st.write("**Pilih Barang dari RAB untuk Paket ini:**")
                barang_paket = st.multiselect(f"Barang Paket {i+1}", daftar_uraian, key=f"barang_{i}")
                
                detail_barang = []
                if barang_paket:
                    st.markdown("---")
                    st.write("📝 **Sesuaikan Harga HPS & Input Spesifikasi Per Barang:**")
                    for j, barang in enumerate(barang_paket):
                        harga_default, vol_default, satuan_default, kode_rek = 0, 0, "-", "-"
                        
                        if not rab_kegiatan.empty:
                            baris_rab = rab_kegiatan[rab_kegiatan['Uraian'] == barang]
                            if not baris_rab.empty:
                                row_rab = baris_rab.iloc[0]
                                kode_rek = row_rab.get('Kode_Rekening', '-')
                                satuan_default = row_rab.get('Satuan_Rencana', '-')
                                
                                val_harga = row_rab.get('Harga_Rencana', 0)
                                val_harga_num = pd.to_numeric(val_harga, errors='coerce')
                                harga_default = int(val_harga_num) if pd.notna(val_harga_num) else 0
                                
                                val_vol = row_rab.get('Vol_Rencana', 0)
                                val_vol_num = pd.to_numeric(val_vol, errors='coerce')
                                vol_default = float(val_vol_num) if pd.notna(val_vol_num) else 0.0

                        col_hps1, col_hps2 = st.columns(2)
                        with col_hps1:
                            input_hps = st.number_input(f"Harga HPS - {barang}", min_value=0, value=harga_default, step=1000, key=f"hps_{i}_{j}")
                        with col_hps2:
                            input_spek = st.text_input(f"Spesifikasi - {barang}", placeholder="Contoh: SNI, Garansi 1 Tahun, dll", key=f"spek_brg_{i}_{j}")
                        
                        detail_barang.append({
                            "kode": kode_rek,
                            "uraian": barang,
                            "vol_rencana": vol_default,
                            "satuan": satuan_default,
                            "harga_hps": input_hps,
                            "spesifikasi_khusus": input_spek
                        })

                paket_data.append({
                    "id_paket": i+1,
                    "nama_paket": nama_paket,
                    "latar_belakang": latar_belakang,
                    "penerima": penerima,
                    "mekanisme_bayar": mekanisme_bayar,
                    "spesifikasi_umum": spesifikasi_umum,
                    "daftar_barang": barang_paket,
                    "detail_barang": detail_barang 
                })

    # (Lanjutkan dengan Paste kode tab_cetak, tab_pelaksanaan, tab_pelaporan1, tab_pelaporan2 Anda yang lama tanpa perlu mengubah apapun)
    # Karena logic filternya bekerja di "daftar_uraian", secara otomatis paket barang di Tab selanjutnya akan menyesuaikan user.
else:
    st.warning("Menunggu data Master_APBDES.xlsx...")