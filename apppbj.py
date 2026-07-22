import streamlit as st
import pandas as pd
from io import BytesIO
from num2words import num2words
from xhtml2pdf import pisa
from datetime import timedelta, date
import base64
import json
import os

# 1. Pastikan blok fungsi ini ADA dan ditulis LEBIH DULU
@st.cache_data
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        return ""

# 2. BARU dipanggil di bawahnya
logo_ngawi_b64 = get_base64_image("logo_ngawi.png")

st.set_page_config(page_title="PBJ SUMENGKO Generator", layout="wide")
st.title("Aplikasi Pembuat Dokumen PBJ Desa Sumengko")

# ---------------------------------------------------------
# FITUR MEMORI LOKAL (SIDEBAR)
# ---------------------------------------------------------
with st.sidebar:
    st.header("💾 Memori Lokal (Uji Coba)")
    st.write("Simpan atau muat progres isian form ke/dari file lokal (JSON) di laptop Anda.")
    
    st.markdown("---")
    st.subheader("📥 Simpan Data Baru")
    # Form untuk menyimpan data baru
    nama_file_baru = st.text_input("Ketik nama file (tanpa .json):", value="draft_pbj_001")
    
    if st.button("Simpan Data", type="primary"):
        state_to_save = {}
        for k, v in st.session_state.items():
            # Lewati key milik tombol cetak agar tidak memicu error
            if k.startswith("cetak_"):
                continue
                
            # Hanya simpan tipe data yang aman untuk JSON (string, angka, boolean, list)
            if isinstance(v, (str, int, float, bool, list)):
                state_to_save[k] = v
            # Jika formatnya tanggal, ubah jadi string (teks)
            elif isinstance(v, date):
                state_to_save[k] = v.isoformat()
        
        with open(f"{nama_file_baru}.json", "w") as f:
            json.dump(state_to_save, f)
        st.success(f"Data berhasil disimpan ke {nama_file_baru}.json!")
        # Rerun agar daftar dropdown file di bawah langsung ter-update
        st.rerun()

    st.markdown("---")
    st.subheader("📂 Muat Data Tersimpan")
    
    # Mencari semua file dengan ekstensi .json di folder/direktori saat ini
    list_file_backup = [f for f in os.listdir('.') if f.endswith('.json')]
    
    # Cek apakah ada file json yang ditemukan
    if list_file_backup:
        # Tampilkan dropdown pilihan file
        file_terpilih = st.selectbox("Pilih file yang ingin dimuat:", list_file_backup)
        
        if st.button("Muat Data Terpilih"):
            if os.path.exists(file_terpilih):
                with open(file_terpilih, "r") as f:
                    loaded_state = json.load(f)
                
                for k, v in loaded_state.items():
                    # Mencegah error jika memuat file JSON lama yang terlanjur menyimpan key button
                    if k.startswith("cetak_"):
                        continue
                        
                    # Cek jika input aslinya adalah tanggal, kembalikan ke format Date
                    if isinstance(v, str) and ("tgl" in k.lower()):
                        try:
                            st.session_state[k] = date.fromisoformat(v)
                        except ValueError:
                            st.session_state[k] = v
                    else:
                        st.session_state[k] = v
                
                st.success(f"Data dari {file_terpilih} berhasil dimuat!")
                st.rerun() # Refresh halaman untuk menampilkan data yang baru dimuat
            else:
                st.error("File tidak ditemukan!")
    else:
        st.info("Belum ada file data yang tersimpan di folder ini.")

# ---------------------------------------------------------
# FUNGSI BANTUAN TANGGAL INDONESIA
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# BACA DATA MASTER EXCEL
# ---------------------------------------------------------
@st.cache_data
def load_data():
    try:
        df_user = pd.read_excel("Master_APBDES.xlsx", sheet_name="User")
        df_dpa = pd.read_excel("Master_APBDES.xlsx", sheet_name="Master_DPA")
                
        df_user.columns = df_user.columns.astype(str).str.strip()
        df_dpa.columns = df_dpa.columns.astype(str).str.strip()
        
        return df_user, df_dpa
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        return pd.DataFrame(), pd.DataFrame()

df_user, df_dpa = load_data()

# ---------------------------------------------------------
# UI & LOGIC
# ---------------------------------------------------------
if not df_user.empty and not df_dpa.empty:
    tab_utama, tab_paket, tab_cetak, tab_pelaksanaan, tab_pelaporan1, tab_pelaporan2 = st.tabs([
        "1. Tahap Persiapan - Data Utama", 
        "2. Tahap Persiapan - Paket Belanja", 
        "3. Tahap Persiapan - Cetak",
        "4. Tahap Pelaksanaan - Dokumen",
        "5. Tahap Pelaporan 1 - Penyedia",
        "6. Tahap Pelaporan 2 - Kades"
    ])

    # === TAB 1: Tahap Persiapan - DATA UTAMA & CETAK RAB ===
    with tab_utama:
        st.header("A. Pilih Pelaksana (User)")
        daftar_user = df_user['Nama_Lengkap_KaurKasi'].tolist()
        # Penambahan key="pilih_user" agar bisa tersimpan di session_state
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

        st.subheader("RAB Kegiatan")
        rab_kegiatan = pd.DataFrame()
        daftar_uraian = [] 
        
        if 'Kegiatan' in df_dpa.columns:
            rab_kegiatan = df_dpa[df_dpa['Kegiatan'] == kegiatan]
            st.dataframe(rab_kegiatan, use_container_width=True)
            if 'Uraian' in rab_kegiatan.columns:
                daftar_uraian = rab_kegiatan['Uraian'].tolist()
            else:
                st.warning("Kolom 'Uraian' tidak ditemukan di Excel.")
        
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

    # === TAB 2: Tahap Persiapan - PEMBAGIAN PAKET BELANJA ===
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

    # === TAB 3: Tahap Persiapan - CETAK DOKUMEN PERSIAPAN ===
    with tab_cetak:
        st.header("Cetak Dokumen Persiapan Supplier (PDF)")
        st.write("Klik tombol di bawah untuk membuat dokumen HPS, Pakta Integritas, dan KAK per Paket/Supplier.")
        
        for p in paket_data:
            if st.button(f"Generate PDF Dokumen Persiapan - {p['nama_paket']}", type="primary", key=f"cetak_{p['id_paket']}"):
                
                tanggal_format, _, _, _, _ = get_tanggal_indo(tgl_persiapan)
                
                mulai_persiapan = tgl_persiapan
                selesai_persiapan = mulai_persiapan + timedelta(days=max(0, hari_persiapan - 1))
                mulai_pelaksanaan = selesai_persiapan + timedelta(days=1)
                selesai_pelaksanaan = mulai_pelaksanaan + timedelta(days=max(0, hari_pelaksanaan - 1))
                mulai_pelaporan = selesai_pelaksanaan + timedelta(days=1)
                selesai_pelaporan = mulai_pelaporan + timedelta(days=max(0, hari_pelaporan - 1))
                
                str_m_persiapan, _, _, _, _ = get_tanggal_indo(mulai_persiapan)
                str_s_persiapan, _, _, _, _ = get_tanggal_indo(selesai_persiapan)
                str_m_pelaksanaan, _, _, _, _ = get_tanggal_indo(mulai_pelaksanaan)
                str_s_pelaksanaan, _, _, _, _ = get_tanggal_indo(selesai_pelaksanaan)
                str_m_pelaporan, _, _, _, _ = get_tanggal_indo(mulai_pelaporan)
                str_s_pelaporan, _, _, _, _ = get_tanggal_indo(selesai_pelaporan)

                baris_tabel_hps_html = ""
                baris_tabel_spek_html = ""
                total_hps_paket = 0

                for idx, dt_brg in enumerate(p['detail_barang']):
                    vol = dt_brg['vol_rencana']
                    harga_hps = dt_brg['harga_hps']
                    jumlah_hps = vol * harga_hps
                    total_hps_paket += jumlah_hps
                    
                    harga_str = f"Rp {harga_hps:,.0f}".replace(",", ".")
                    jumlah_str = f"Rp {jumlah_hps:,.0f}".replace(",", ".")

                    baris_tabel_hps_html += f"""
                    <tr>
                        <td>{dt_brg['kode']}</td>
                        <td>{dt_brg['uraian']}</td>
                        <td style="text-align: center;">{vol}</td>
                        <td style="text-align: center;">{dt_brg['satuan']}</td>
                        <td>{harga_str}</td>
                        <td>{jumlah_str}</td>
                    </tr>
                    """

                    baris_tabel_spek_html += f"""
                    <tr>
                        <td style="text-align: center;">{idx + 1}</td>
                        <td>{dt_brg['uraian']}</td>
                        <td>{dt_brg['spesifikasi_khusus']}</td>
                    </tr>
                    """
                
                teks_terbilang = num2words(total_hps_paket, lang='id').title() + " Rupiah"

                html_persiapan = f"""
                <html>
                <head>
                    <style>
                        @page {{ size: a4 portrait; margin: 2cm; }}
                        body {{ font-family: Helvetica, Arial, sans-serif; font-size: 11px; text-align: justify; line-height: 1.5; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                        th, td {{ border: 1px solid black; padding: 5px; text-align: left; vertical-align: middle; }}
                        th {{ background-color: #f2f2f2; text-align: center; font-weight: bold; }}
                        .judul {{ text-align: center; font-weight: bold; font-size: 14px; margin-bottom: 10px; }}
                        .sub-judul {{ font-weight: bold; margin-top: 15px; margin-bottom: 2px; font-size: 12px; }}
                        .ttd-table {{ border: none; width: 100%; margin-top: 30px; }}
                        .ttd-td {{ border: none; width: 50%; text-align: center; }}
                        .no-margin {{ margin: 0px; padding: 0px; }}
                    </style>
                </head>
                <body>
                    <div class="judul">HARGA PERKIRAAN SENDIRI (HPS)</div>
                    <table style="border: none;">
                        <tr style="border: none;"><td style="border: none; width: 20%;"><b>Kegiatan</b></td><td style="border: none;">: {kegiatan}</td></tr>
                        <tr style="border: none;"><td style="border: none;"><b>Paket Belanja</b></td><td style="border: none;">: {p['nama_paket']}</td></tr>
                    </table>
                    <table>
                        <tr>
                            <th>Kode</th>
                            <th>Uraian</th>
                            <th>Volume</th>
                            <th>Satuan</th>
                            <th>Harga Satuan</th>
                            <th>Jumlah</th>
                        </tr>
                        {baris_tabel_hps_html}
                        <tr>
                            <td colspan="5" style="text-align: right;"><b>JUMLAH TOTAL</b></td>
                            <td><b>Rp {total_hps_paket:,.0f}</b></td>
                        </tr>
                    </table>
                    <p><b>Terbilang:</b> <i>{teks_terbilang}</i></p>
                    <p><i>NB: Harga sudah termasuk pajak.</i></p>
                    
                    <table class="ttd-table">
                        <tr style="border: none;">
                            <td class="ttd-td"></td>
                            <td class="ttd-td">
                                Sumengko, {tanggal_format}<br>
                                Kasi/Kaur {role_user}<br><br><br><br><br>
                                <b><u>{pilih_user}</u></b>
                            </td>
                        </tr>
                    </table>

                    <pdf:nextpage />
                    <div class="judul">PAKTA INTEGRITAS</div>
                    <p>Saya yang bertanda tangan dibawah ini:</p>
                    <table style="border: none; margin-left: 20px;">
                        <tr style="border: none;"><td style="border: none; width: 150px;">Nama</td><td style="border: none;">: {pilih_user}</td></tr>
                        <tr style="border: none;"><td style="border: none;">Nomor Identitas</td><td style="border: none;">: {nik_user}</td></tr>
                        <tr style="border: none;"><td style="border: none;">Alamat</td><td style="border: none;">: {alamat_user}</td></tr>
                        <tr style="border: none;"><td style="border: none;">Bertindak untuk/atas nama</td><td style="border: none;">: Kasi/Kaur {role_user}</td></tr>
                    </table>
                    <p>Dalam rangka pengadaan pekerjaan <b>{kegiatan}</b> Paket Belanja <b>{p['nama_paket']}</b> yang tertuang dalam DPA/DPPA pada Desa Sumengko Kecamatan Kwadungan, dengan ini menyatakan bahwa:</p>
                    <ol>
                        <li>Tidak akan melakukan praktek korupsi, kolusi dan nepotisme (KKN);</li>
                        <li>Akan melaksanakan proses pengadaan secara bersih, transparan dan profesional untuk memberikan hasil kerja terbaik sesuai ketentuan peraturan perundang-undangan; dan</li>
                        <li>Apabila melanggar hal-hal yang dinyatakan dalam PAKTA INTEGRITAS ini, bersedia menerima sanksi sesuai dengan perundang-undangan yang berlaku.</li>
                    </ol>
                    
                    <table class="ttd-table">
                        <tr style="border: none;">
                            <td class="ttd-td"></td>
                            <td class="ttd-td">
                                Sumengko, {tanggal_format}<br>
                                Kasi/Kaur {role_user}<br><br><br><br><br>
                                <b><u>{pilih_user}</u></b>
                            </td>
                        </tr>
                    </table>

                    <pdf:nextpage />
                    <div class="judul">KERANGKA ACUAN KERJA (KAK)</div>
                    <div class="sub-judul">Kegiatan/Pekerjaan</div>
                    <p class="no-margin">{kegiatan}</p>
                    <p class="no-margin">{p['nama_paket']}</p>
                    <div class="sub-judul">Latar Belakang</div>
                    <p class="no-margin">{p['latar_belakang']}</p>
                    <div class="sub-judul">Penerima Manfaat</div>
                    <p class="no-margin">{p['penerima']}</p>
                    <div class="sub-judul">Cara melaksanakan</div>
                    <p class="no-margin">Pelaksanaan pengadaan dilaksanakan melalui Swakelola/melalui Penyedia dengan metode Pembelian Langsung sesuai dengan ketentuan yang diatur dalam Peraturan Bupati Ngawi Nomor 177 Tahun 2021 tentang Tata Cara Pengadaan Barang dan Jasa di Desa.</p>
                    <div class="sub-judul">Jangka Waktu Pelaksana Kegiatan</div>
                    <p class="no-margin">Pelaksana kegiatan pengadaan barang/jasa di Desa adalah {role_user} pada kegiatan {kegiatan} yang dilaksanakan selama {total_hari} hari kalender. Dimana :<br>
                    Tahap Persiapan = {hari_persiapan} mulai ({str_m_persiapan} sampai {str_s_persiapan})<br>
                    Tahap Pelaksanaan = {hari_pelaksanaan} mulai ({str_m_pelaksanaan} sampai {str_s_pelaksanaan})<br>
                    Tahap Pelaporan = {hari_pelaporan} mulai ({str_m_pelaporan} sampai {str_s_pelaporan})
                    </p>
                    <div class="sub-judul">Spesifikasi Teknis</div>
                    <p class="no-margin">Spesifikasi teknis Pelaksanaan {kegiatan} {p['nama_paket']} sebagaimana terlampir. {p['spesifikasi_umum']}</p>
                    <div class="sub-judul">Pembayaran</div>
                    <p class="no-margin">Mekanisme pembayaran kegiatan {kegiatan} {p['nama_paket']} dilaksanakan secara {p['mekanisme_bayar']} sebesar Rp {total_hps_paket:,.0f} (<i>{teks_terbilang}</i>).</p>

                    <table class="ttd-table">
                        <tr style="border: none;">
                            <td class="ttd-td"></td>
                            <td class="ttd-td">
                                Sumengko, {tanggal_format}<br>
                                Kasi/Kaur<br><br><br><br><br>
                                <b><u>{pilih_user}</u></b>
                            </td>
                        </tr>
                    </table>
                    
                    <pdf:nextpage />
                    <p>Lampiran KAK {p['nama_paket']}</p>
                    <div class="judul">SPESIFIKASI TEKNIS</div>
                    <p class="no-margin">{kegiatan}</p>
                    <p class="no-margin">{p['nama_paket']}</p>
                    
                    <table>
                        <tr>
                            <th style="width: 5%;">No</th>
                            <th style="width: 45%;">Nama Barang</th>
                            <th style="width: 50%;">Spesifikasi</th>
                        </tr>
                        {baris_tabel_spek_html}
                    </table>
                    
                    <table class="ttd-table">
                        <tr style="border: none;">
                            <td class="ttd-td"></td>
                            <td class="ttd-td">
                                Sumengko, {tanggal_format}<br>
                                Kasi/Kaur<br><br><br><br><br>
                                <b><u>{pilih_user}</u></b>
                            </td>
                        </tr>
                    </table>
                </body>
                </html>
                """

                pdf_output = BytesIO()
                pisa_status = pisa.CreatePDF(BytesIO(html_persiapan.encode('utf-8')), dest=pdf_output)

                if pisa_status.err:
                    st.error("Terjadi kesalahan saat merender PDF.")
                else:
                    st.success(f"PDF Dokumen Persiapan {p['nama_paket']} siap diunduh!")
                    st.download_button(
                        label=f"Unduh Dokumen {p['nama_paket']}",
                        data=pdf_output.getvalue(),
                        file_name=f"Persiapan_{kegiatan}_{p['nama_paket']}.pdf",
                        mime="application/pdf"
                    )

    # === TAB 4: TAHAP PELAKSANAAN - INPUT & CETAK DOKUMEN ===
    with tab_pelaksanaan:
        st.header("Cetak Dokumen Pelaksanaan (Surat Pesanan & BA Negosiasi)")
        
        if not paket_data:
            st.warning("Silakan buat Paket Belanja terlebih dahulu di Tab 2.")
        else:
            for i, p in enumerate(paket_data):
                with st.expander(f"Pelaksanaan: {p['nama_paket']}", expanded=True):
                    
                    st.subheader("1. Data Surat & Tanggal Pelaksanaan")
                    col_doc1, col_doc2 = st.columns(2)
                    with col_doc1:
                        nomor_banego = st.text_input("Nomor Surat BA Negosiasi", value="001/BAN", key=f"no_banego_{i}")
                        nomor_spesanan = st.text_input("Nomor Surat Pesanan", value="001/SP", key=f"no_sp_{i}")
                        lingkup_pekerjaan = st.text_input("Lingkup Pekerjaan", value=f"Pengadaan barang {p['nama_paket']}", key=f"lingkup_{i}")
                    with col_doc2:
                        tgl_pelaksanaan = st.date_input("Tanggal Pelaksanaan (BA Nego & SP)", tgl_persiapan + timedelta(days=hari_persiapan), key=f"tgl_pel_{i}")
                        metode_bayar_pelaksanaan = st.selectbox("Metode Pembayaran (Surat Pesanan)", ["Tunai", "Non Tunai / Transfer"], key=f"metode_pel_{i}")

                    st.subheader("2. Identitas Penyedia (Supplier)")
                    col_sup1, col_sup2 = st.columns(2)
                    with col_sup1:
                        nama_perusahaan = st.text_input("Nama Perusahaan / Toko", key=f"pt_nama_{i}")
                        alamat_perusahaan = st.text_input("Alamat Perusahaan", key=f"pt_alamat_{i}")
                        npwp_perusahaan = st.text_input("NPWP Perusahaan", key=f"pt_npwp_{i}")
                        rek_bank = st.text_input("Rekening Bank Jatim (Nomor)", key=f"pt_rek_{i}")
                        atas_nama_rek = st.text_input("Atas Nama Rekening", key=f"pt_an_{i}")
                    with col_sup2:
                        nama_pic = st.text_input("Nama PIC / Pemilik", key=f"pic_nama_{i}")
                        nik_pic = st.text_input("NIK PIC", key=f"pic_nik_{i}")
                        jabatan_pic = st.text_input("Jabatan PIC", value="Direktur/Pemilik", key=f"pic_jabatan_{i}")

                    st.subheader("3. Negosiasi Harga & Volume")
                    nego_data = []
                    
                    for j, dt_brg in enumerate(p['detail_barang']):
                        st.write(f"**Barang:** {dt_brg['uraian']} | **HPS:** Rp {dt_brg['harga_hps']:,.0f} | **Vol:** {dt_brg['vol_rencana']} {dt_brg['satuan']}")
                        col_neg1, col_neg2 = st.columns(2)
                        with col_neg1:
                            input_vol_nego = st.number_input(f"Volume Negosiasi - {dt_brg['uraian']}", min_value=0.0, value=float(dt_brg['vol_rencana']), step=1.0, key=f"nego_vol_{i}_{j}")
                        with col_neg2:
                            input_harga_nego = st.number_input(f"Harga Negosiasi - {dt_brg['uraian']}", min_value=0, value=int(dt_brg['harga_hps']), step=1000, key=f"nego_harga_{i}_{j}")
                        
                        nego_data.append({
                            "kode": dt_brg['kode'],
                            "uraian": dt_brg['uraian'],
                            "satuan": dt_brg['satuan'],
                            "vol_sebelum": dt_brg['vol_rencana'],
                            "harga_sebelum": dt_brg['harga_hps'],
                            "jumlah_sebelum": dt_brg['vol_rencana'] * dt_brg['harga_hps'],
                            "vol_nego": input_vol_nego,
                            "harga_nego": input_harga_nego,
                            "jumlah_nego": input_vol_nego * input_harga_nego
                        })
                    
                    st.markdown("---")
                    if st.button(f"Generate PDF Dokumen Pelaksanaan - {p['nama_paket']}", type="primary", key=f"cetak_pel_{i}"):
                        
                        tgl_lengkap, hari_nama, tgl_angka, bulan_nama, tahun_angka = get_tanggal_indo(tgl_pelaksanaan)
                        total_hps_pelaksanaan = sum(item['jumlah_sebelum'] for item in nego_data)
                        total_nego_pelaksanaan = sum(item['jumlah_nego'] for item in nego_data)
                        terbilang_hps = num2words(total_hps_pelaksanaan, lang='id').title() + " Rupiah"
                        terbilang_nego = num2words(total_nego_pelaksanaan, lang='id').title() + " Rupiah"

                        baris_tabel_nego_html = ""
                        for idx, item in enumerate(nego_data):
                            h_seb = f"Rp {item['harga_sebelum']:,.0f}".replace(",", ".")
                            j_seb = f"Rp {item['jumlah_sebelum']:,.0f}".replace(",", ".")
                            h_neg = f"Rp {item['harga_nego']:,.0f}".replace(",", ".")
                            j_neg = f"Rp {item['jumlah_nego']:,.0f}".replace(",", ".")

                            baris_tabel_nego_html += f"""
                            <tr>
                                <td style="text-align: center;">{idx + 1}</td>
                                <td>{item['uraian']}</td>
                                <td style="text-align: center;">{item['vol_sebelum']}</td>
                                <td style="text-align: center;">{item['satuan']}</td>
                                <td>{h_seb}</td>
                                <td>{j_seb}</td>
                                <td style="text-align: center;">{item['vol_nego']}</td>
                                <td style="text-align: center;">{item['satuan']}</td>
                                <td>{h_neg}</td>
                                <td>{j_neg}</td>
                            </tr>
                            """
                        
                        baris_tabel_sp_html = ""
                        for idx, item in enumerate(nego_data):
                            h_neg = f"Rp {item['harga_nego']:,.0f}".replace(",", ".")
                            j_neg = f"Rp {item['jumlah_nego']:,.0f}".replace(",", ".")
                            
                            baris_tabel_sp_html += f"""
                            <tr>
                                <td style="text-align: center;">{idx + 1}</td>
                                <td>{item['uraian']}</td>
                                <td style="text-align: center;">{item['vol_nego']}</td>
                                <td style="text-align: center;">{item['satuan']}</td>
                                <td>{h_neg}</td>
                                <td>{j_neg}</td>
                            </tr>
                            """

                        html_pelaksanaan = f"""
                        <html>
                        <head>
                            <style>
                                @page {{ size: a4 portrait; margin: 2cm; }}
                                body {{ font-family: Helvetica, Arial, sans-serif; font-size: 11px; text-align: justify; line-height: 1.5; }}
                                table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                                th, td {{ border: 1px solid black; padding: 5px; text-align: left; vertical-align: middle; }}
                                th {{ background-color: #f2f2f2; text-align: center; font-weight: bold; font-size: 10px; }}
                                .judul {{ text-align: center; font-weight: bold; font-size: 14px; margin-bottom: 5px; text-decoration: underline; }}
                                .sub-judul {{ text-align: center; font-size: 12px; margin-bottom: 15px; }}
                                .ttd-table {{ border: none; width: 100%; margin-top: 30px; }}
                                .ttd-td {{ border: none; width: 50%; text-align: center; vertical-align: bottom; }}
                                .kop-surat {{ text-align: center; border-bottom: 3px solid black; padding-bottom: 10px; margin-bottom: 15px; }}
                            </style>
                        </head>
                        <body>
                            <!-- HALAMAN 1: BERITA ACARA NEGOSIASI -->
                            <div class="judul">BERITA ACARA NEGOSIASI</div>
                            <div class="sub-judul">Nomor : 100.3.3/{nomor_banego}/404.608.8/{tahun_angka}</div>
                            
                            <table style="border: none; margin-bottom: 10px;">
                                <tr style="border: none;"><td style="border: none; width: 150px; padding:2px;"><b>Nama Kegiatan/Pekerjaan</b></td><td style="border: none; padding:2px;">: {kegiatan}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;"><b>Paket Belanja</b></td><td style="border: none; padding:2px;">: {p['nama_paket']}</td></tr>
                            </table>

                            <p>Pada hari ini <b>{hari_nama}</b> tanggal <b>{tgl_angka}</b> bulan <b>{bulan_nama}</b> tahun <b>{tahun_angka}</b> kami Kasi/Kaur {role_user} Desa Sumengko telah melaksanakan negosiasi harga untuk :</p>
                            
                            <table style="border: none; margin-left: 20px; margin-bottom: 10px;">
                                <tr style="border: none;"><td style="border: none; width: 130px; padding:2px;">Kegiatan</td><td style="border: none; padding:2px;">: {kegiatan}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Lingkup Pekerjaan</td><td style="border: none; padding:2px;">: {lingkup_pekerjaan}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Paket Belanja</td><td style="border: none; padding:2px;">: {p['nama_paket']}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Lokasi</td><td style="border: none; padding:2px;">: Desa Sumengko</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">HPS</td><td style="border: none; padding:2px;">: Rp {total_hps_pelaksanaan:,.0f} <i>(Terbilang: {terbilang_hps})</i></td></tr>
                            </table>
                            
                            <p>Setelah dilakukan negosiasi, harga tersebut sepakat menjadi sebesar : <b>Rp {total_nego_pelaksanaan:,.0f}</b><br>
                            Terbilang : <i>{terbilang_nego}</i><br>
                            yang dalam pelaksanaan pengadaannya dilakukan sesuai dengan surat pesanan. Untuk rincian hasil negosiasi sebagaimana terlampir.</p>
                            
                            <p><b>Penyedia : {nama_perusahaan}</b></p>
                            <table>
                                <tr>
                                    <th rowspan="2">No</th>
                                    <th rowspan="2">Nama Barang</th>
                                    <th colspan="4">Rincian Sebelum</th>
                                    <th colspan="4">Rincian Sesudah Negosiasi</th>
                                </tr>
                                <tr>
                                    <th>Volume</th>
                                    <th>Satuan</th>
                                    <th>Harga Satuan</th>
                                    <th>Jumlah</th>
                                    <th>Volume</th>
                                    <th>Satuan</th>
                                    <th>Harga Satuan</th>
                                    <th>Jumlah</th>
                                </tr>
                                {baris_tabel_nego_html}
                                <tr>
                                    <td colspan="5" style="text-align: right;"><b>JUMLAH</b></td>
                                    <td><b>Rp {total_hps_pelaksanaan:,.0f}</b></td>
                                    <td colspan="3" style="text-align: right;"><b>JUMLAH</b></td>
                                    <td><b>Rp {total_nego_pelaksanaan:,.0f}</b></td>
                                </tr>
                            </table>
                            
                            <p>Selanjutnya untuk proses transaksi akan dituangkan dalam bentuk pembelian langsung antara Kasi/Kaur {role_user} sebagai pelaksana kegiatan anggaran dengan penyedia jasa.</p>
                            <p>Demikian berita acara negosiasi ini dibuat untuk diketahui dan dipergunakan mestinya.</p>
                            
                            <table class="ttd-table">
                                <tr style="border: none;">
                                    <td class="ttd-td">
                                        Nama Kasi/Kaur {role_user}<br><br><br><br><br><br>
                                        <b><u>{pilih_user}</u></b>
                                    </td>
                                    <td class="ttd-td">
                                        Nama Penyedia Jasa<br>
                                        <b>{nama_perusahaan}</b><br><br><br><br><br>
                                        <b><u>{nama_pic}</u></b>
                                    </td>
                                </tr>
                            </table>

                            <!-- HALAMAN 2: SURAT PESANAN -->
                            <pdf:nextpage />
                            
                            <table style="border: none; margin-bottom: 15px; border-bottom: 3px solid black; padding-bottom: 5px;">
                                <tr style="border: none;">
                                    <td style="border: none; width: 15%; text-align: center; vertical-align: middle;">
                                        <img src="data:image/png;base64,{logo_ngawi_b64}" width="60">
                                    </td>
                                    <td style="border: none; width: 85%; text-align: center; vertical-align: middle;">
                                        <b style="font-size: 14px;">PEMERINTAH KABUPATEN NGAWI</b><br>
                                        <b style="font-size: 14px;">KECAMATAN KWADUNGAN</b><br>
                                        <b style="font-size: 16px;">DESA SUMENGKO</b><br>
                                        Jalan Raya Desa Sumengko No. 08 Telp. - Kode Pos 63283
                                    </td>
                                </tr>
                            </table>
                            
                            <div class="judul">SURAT PESANAN</div>
                            <div class="sub-judul">Nomor : 400.3.3/{nomor_spesanan}/404.608.8/{tahun_angka}</div>
                            
                            <p style="text-align: right;">Tanggal: {tgl_lengkap}</p>
                            
                            <table style="border: none; margin-bottom: 10px;">
                                <tr style="border: none;"><td style="border: none; width: 150px; padding:2px;"><b>Untuk melaksanakan</b></td><td style="border: none; padding:2px;">: {kegiatan}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;"><b>Paket Belanja</b></td><td style="border: none; padding:2px;">: {p['nama_paket']}</td></tr>
                            </table>

                            <p>Yang bertanda tangan dibawah ini :</p>
                            <table style="border: none; margin-left: 20px; margin-bottom: 10px;">
                                <tr style="border: none;"><td style="border: none; width: 130px; padding:2px;">Nama</td><td style="border: none; padding:2px;">: {pilih_user}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Jabatan</td><td style="border: none; padding:2px;">: Kasi/Kaur {role_user}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Alamat</td><td style="border: none; padding:2px;">: {alamat_user}</td></tr>
                            </table>
                            <p>Selanjutnya disebut sebagai <b>Pelaksana Kegiatan Anggaran (PKA)</b></p>
                            
                            <table style="border: none; margin-left: 20px; margin-bottom: 10px;">
                                <tr style="border: none;"><td style="border: none; width: 130px; padding:2px;">Nama</td><td style="border: none; padding:2px;">: {nama_pic}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">NIK</td><td style="border: none; padding:2px;">: {nik_pic}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Jabatan</td><td style="border: none; padding:2px;">: {jabatan_pic}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Nama Penyedia</td><td style="border: none; padding:2px;">: {nama_perusahaan}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Alamat Perusahaan</td><td style="border: none; padding:2px;">: {alamat_perusahaan}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Rekening Bank Jatim</td><td style="border: none; padding:2px;">: {rek_bank}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Nama Rekening</td><td style="border: none; padding:2px;">: {atas_nama_rek}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">NPWP Perusahaan</td><td style="border: none; padding:2px;">: {npwp_perusahaan}</td></tr>
                            </table>
                            <p>Yang bertindak dan atas nama perusahaan/toko <b>{nama_perusahaan}</b><br>
                            Selanjutnya disebut sebagai <b>PENYEDIA</b></p>

                            <p>Untuk mengirim barang dengan memperhatikan ketentuan-ketentuan sebagai berikut:</p>
                            <table>
                                <tr>
                                    <th style="width: 5%;">No</th>
                                    <th style="width: 40%;">Nama Barang</th>
                                    <th style="width: 10%;">Vol</th>
                                    <th style="width: 10%;">Satuan</th>
                                    <th style="width: 15%;">Harga Satuan</th>
                                    <th style="width: 20%;">Jumlah</th>
                                </tr>
                                {baris_tabel_sp_html}
                                <tr>
                                    <td colspan="5" style="text-align: right;"><b>JUMLAH</b></td>
                                    <td><b>Rp {total_nego_pelaksanaan:,.0f}</b></td>
                                </tr>
                            </table>
                            <p><b>Terbilang:</b> <i>{terbilang_nego}</i></p>
                            <p><i>NB* harga sudah termasuk pajak</i></p>
                            
                            <p>Surat pesanan ini berlaku sejak tanggal ditanda tangani oleh para pihak sampai dengan selesainya pelaksanaan. Adapun mekanisme pembayaran dilakukan secara <b>{metode_bayar_pelaksanaan}</b> dan setelah barang diperiksa oleh PKA yang dituangkan dalam berita acara pemeriksaan barang/jasa dan berita acara Serah Terima.</p>
                            <p>Demikian Surat Pesanan ini dibuat dan ditandatangani oleh para pihak untuk dipergunakan sebagaimana mestinya.</p>
                            
                            <table class="ttd-table">
                                <tr style="border: none;">
                                    <td class="ttd-td">
                                        Untuk dan atas nama<br>
                                        Kasi/Kaur {role_user} Desa Sumengko<br>
                                        Selaku Pelaksana Kegiatan Anggaran<br><br><br><br><br>
                                        <b><u>{pilih_user}</u></b>
                                    </td>
                                    <td class="ttd-td">
                                        Untuk dan atas nama<br>
                                        Penyedia<br>
                                        <b>{nama_perusahaan}</b><br><br><br><br><br>
                                        <b><u>{nama_pic}</u></b>
                                    </td>
                                </tr>
                            </table>
                        </body>
                        </html>
                        """

                        pdf_output_pelaksanaan = BytesIO()
                        pisa_status_pelaksanaan = pisa.CreatePDF(BytesIO(html_pelaksanaan.encode('utf-8')), dest=pdf_output_pelaksanaan)

                        if pisa_status_pelaksanaan.err:
                            st.error("Terjadi kesalahan saat merender PDF Pelaksanaan.")
                        else:
                            st.success(f"PDF Dokumen Pelaksanaan {p['nama_paket']} siap diunduh!")
                            st.download_button(
                                label=f"Unduh Dokumen Pelaksanaan {p['nama_paket']}",
                                data=pdf_output_pelaksanaan.getvalue(),
                                file_name=f"Pelaksanaan_{kegiatan}_{p['nama_paket']}.pdf",
                                mime="application/pdf"
                            )

    # === TAB 5: TAHAP PELAPORAN 1 - PENYEDIA ===
    with tab_pelaporan1:
        st.header("Cetak Laporan: BA Pemeriksaan & BA Serah Terima (Penyedia ke PKA)")
        
        if not paket_data:
            st.warning("Silakan selesaikan pembuatan Paket Belanja di Tab 2 dan Pelaksanaan di Tab 4 terlebih dahulu.")
        else:
            for i, p in enumerate(paket_data):
                with st.expander(f"Pelaporan: {p['nama_paket']}", expanded=True):
                    
                    no_sp_tarik = st.session_state.get(f"no_sp_{i}", "001/SP")
                    tgl_pelaksanaan_tarik = st.session_state.get(f"tgl_pel_{i}", tgl_persiapan + timedelta(days=hari_persiapan))
                    nama_perusahaan_tarik = st.session_state.get(f"pt_nama_{i}", "-")
                    pic_tarik = st.session_state.get(f"pic_nama_{i}", "-")
                    alamat_pt_tarik = st.session_state.get(f"pt_alamat_{i}", "-")
                    
                    st.info(f"📌 **Info Terkait (Dari Tab 4):**\n"
                            f"- Penyedia: **{nama_perusahaan_tarik}** (PIC: {pic_tarik})\n"
                            f"- No. SP: **{no_sp_tarik}**")

                    st.subheader("Input Data Pelaporan")
                    col_lapor1, col_lapor2 = st.columns(2)
                    
                    with col_lapor1:
                        no_bapemeriksaan = st.text_input("Nomor BA Pemeriksaan Barang", value="001/BAP", key=f"no_bapem_{i}")
                        no_bast = st.text_input("Nomor BAST Penyedia -> PKA", value="001/BAST", key=f"no_bast_{i}")
                        
                    with col_lapor2:
                        prosentase_akhir = st.number_input("Prosentase Akhir Kegiatan (%)", min_value=0, max_value=100, value=100, key=f"persen_{i}")
                        link_gambar = st.text_input("Link URL Gambar / Dokumentasi (Opsional)", placeholder="Contoh: https://link-gambar.com/foto.jpg", key=f"link_gambar_{i}")
                        tgl_pelaporan_input = st.date_input("Tanggal Pelaporan", tgl_pelaksanaan_tarik + timedelta(days=hari_pelaksanaan), key=f"tgl_lapor_{i}")

                    st.markdown("---")
                    if st.button(f"Generate PDF Dokumen Pelaporan Penyedia - {p['nama_paket']}", type="primary", key=f"cetak_lapor_{i}"):
                        
                        tgl_lengkap_sp, _, _, _, tahun_sp = get_tanggal_indo(tgl_pelaksanaan_tarik)
                        tgl_lengkap, hari_nama, tgl_angka, bulan_nama, tahun_angka = get_tanggal_indo(tgl_pelaporan_input)
                        
                        total_nego_lapor = 0
                        baris_tabel_lapor_html = ""
                        
                        for j, dt_brg in enumerate(p['detail_barang']):
                            v_nego = st.session_state.get(f"nego_vol_{i}_{j}", dt_brg['vol_rencana'])
                            h_nego = st.session_state.get(f"nego_harga_{i}_{j}", dt_brg['harga_hps'])
                            j_nego = v_nego * h_nego
                            total_nego_lapor += j_nego
                            
                            h_str = f"Rp {h_nego:,.0f}".replace(",", ".")
                            j_str = f"Rp {j_nego:,.0f}".replace(",", ".")
                            
                            baris_tabel_lapor_html += f"""
                            <tr>
                                <td style="text-align: center;">{j + 1}</td>
                                <td>{dt_brg['uraian']}</td>
                                <td style="text-align: center;">{v_nego}</td>
                                <td style="text-align: center;">{dt_brg['satuan']}</td>
                                <td>{h_str}</td>
                                <td>{j_str}</td>
                            </tr>
                            """
                            
                        terbilang_lapor = num2words(total_nego_lapor, lang='id').title() + " Rupiah"
                        
                        html_lampiran_gambar = ""
                        if link_gambar:
                            if link_gambar.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
                                html_lampiran_gambar = f'<div style="text-align:center;"><img src="{link_gambar}" width="400" /></div>'
                            else:
                                html_lampiran_gambar = f'<div style="text-align:center; padding: 20px; border: 1px dashed black;">Lampiran URL Dokumen:<br><a href="{link_gambar}">{link_gambar}</a></div>'
                        else:
                            html_lampiran_gambar = '<div style="text-align:center; padding: 100px; border: 1px dashed black;">(Tempat melampirkan/menempelkan foto dokumentasi secara manual)</div>'

                        html_pelaporan = f"""
                        <html>
                        <head>
                            <style>
                                @page {{ size: a4 portrait; margin: 2cm; }}
                                body {{ font-family: Helvetica, Arial, sans-serif; font-size: 11px; text-align: justify; line-height: 1.5; }}
                                table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                                th, td {{ border: 1px solid black; padding: 5px; text-align: left; vertical-align: middle; }}
                                th {{ background-color: #f2f2f2; text-align: center; font-weight: bold; font-size: 10px; }}
                                .judul {{ text-align: center; font-weight: bold; font-size: 14px; margin-bottom: 5px; text-decoration: underline; }}
                                .sub-judul {{ text-align: center; font-size: 12px; margin-bottom: 15px; }}
                                .ttd-table {{ border: none; width: 100%; margin-top: 30px; }}
                                .ttd-td {{ border: none; width: 50%; text-align: center; vertical-align: bottom; }}
                                .kop-surat {{ text-align: center; border-bottom: 3px solid black; padding-bottom: 10px; margin-bottom: 15px; }}
                            </style>
                        </head>
                        <body>
                            <!-- HALAMAN 1: BA PEMERIKSAAN -->
                            <div class="judul" style="text-decoration: none;">BERITA ACARA PEMERIKSAAN BARANG</div>
                            <div class="sub-judul" style="margin-bottom: 10px;">Nomor : 100.3.3/{no_bapemeriksaan}/404.608.8/{tahun_angka}</div>
                            
                            <p>Pada hari ini <b>{hari_nama}</b> tanggal <b>{tgl_angka}</b> bulan <b>{bulan_nama}</b> tahun <b>{tahun_angka}</b> kami yang bertanda tangan dibawah ini:</p>
                            
                            <table style="border: none; margin-left: 20px; margin-bottom: 10px;">
                                <tr style="border: none;"><td style="border: none; width: 130px; padding:2px;">Nama</td><td style="border: none; padding:2px;">: {pilih_user}</td></tr>
                                <tr style="border: none;"><td style="border: none; padding:2px;">Jabatan</td><td style="border: none; padding:2px;">: {role_user}</td></tr>
                            </table>
                            
                            <p>Dengan ini menyatakan :<br>
                            Berdasarkan surat pesanan nomor 400.3.3/{no_sp_tarik}/404.608.8/{tahun_sp} tanggal {tgl_lengkap_sp} :</p>
                            
                            <p>Pelaksanaan pekerjaan <b>{p['nama_paket']}</b> kegiatan <b>{kegiatan}</b> pada Pemerintah Desa Sumengko mencapai progres <b>{prosentase_akhir} %</b>.</p>
                            
                            <p>Penyedia Jasa telah menyerahkan pekerjaan Paket Belanja <b>{p['nama_paket']}</b> telah diserahkan kepada {role_user}.</p>
                            
                            <p>{role_user} telah memeriksa hasil Paket Belanja <b>{p['nama_paket']}</b> dan menerima hasil pekerjaan sebagaimana tersebut diatas sesuai spesifikasi yang tertera di dalam Surat Pesanan 400.3.3/{no_sp_tarik}/404.608.8/{tahun_sp} tanggal {tgl_lengkap_sp}.</p>
                            
                            <p>Berita Acara ini dibuat untuk kelengkapan proses administrasi dan dasar pembayaran kepada penyedia.<br>
                            Demikian Berita Acara Pemeriksaan Barang ini dibuat dipergunakan sebagaimana mestinya.</p>
                            
                            <table class="ttd-table">
                                <tr style="border: none;">
                                    <td class="ttd-td">
                                        Penyedia Jasa {nama_perusahaan_tarik}<br><br><br><br><br>
                                        <b><u>{pic_tarik}</u></b>
                                    </td>
                                    <td class="ttd-td">
                                        {role_user}<br><br><br><br><br>
                                        <b><u>{pilih_user}</u></b>
                                    </td>
                                </tr>
                            </table>

                            <!-- HALAMAN 2: LAMPIRAN GAMBAR -->
                            <pdf:nextpage />
                            <div class="judul" style="text-decoration: none; text-align: left; font-size: 11px;">Lampiran Foto/Dokumentasi barang yang telah dikirim.</div>
                            <br>
                            {html_lampiran_gambar}

                            <!-- HALAMAN 3: BAST PENYEDIA KE PKA -->
                            <pdf:nextpage />
                            <div class="judul" style="text-decoration: none;">BERITA ACARA SERAH TERIMA</div>
                            <div class="sub-judul" style="margin-bottom: 10px;">Nomor : 100.3.3/{no_bast}/404.608.8/{tahun_angka}</div>
                            
                            <p>Pada hari ini <b>{hari_nama}</b> tanggal <b>{tgl_angka}</b> bulan <b>{bulan_nama}</b> tahun <b>{tahun_angka}</b> penyedia jasa <b>{nama_perusahaan_tarik}</b> yang beralamat di <b>{alamat_pt_tarik}</b> telah menyelesaikan 100 % (seratus persen) pekerjaan dan menyerahkan hasil pekerjaan pengadaan dengan baik sesuai dengan berita acara pemeriksaaan 100.3.3/{no_bapemeriksaan}/404.608.8/{tahun_angka} tanggal {tgl_lengkap} selanjutnya Kasi/kaur {role_user} Sumengko menerima hasil pelaksanaan pekerjaan berupa:</p>
                            
                            <table>
                                <tr>
                                    <th style="width: 5%;">No</th>
                                    <th style="width: 40%;">Nama Barang</th>
                                    <th style="width: 10%;">Vol</th>
                                    <th style="width: 10%;">Satuan</th>
                                    <th style="width: 15%;">Harga Satuan</th>
                                    <th style="width: 20%;">Jumlah</th>
                                </tr>
                                {baris_tabel_lapor_html}
                                <tr>
                                    <td colspan="5" style="text-align: right;"><b>JUMLAH</b></td>
                                    <td><b>Rp {total_nego_lapor:,.0f}</b></td>
                                </tr>
                            </table>
                            
                            <p>Terbilang : <i>{terbilang_lapor}</i></p>
                            
                            <p>Demikian Berita Acara Penyerahan Hasil Pengadaan dibuat untuk diketahui dan dipergunakan sebagaimana mestinya.</p>
                            
                            <table class="ttd-table">
                                <tr style="border: none;">
                                    <td class="ttd-td">
                                        {role_user}<br>
                                        Desa Sumengko<br><br><br><br><br>
                                        <b><u>{pilih_user}</u></b>
                                    </td>
                                    <td class="ttd-td">
                                        Penyedia Jasa<br>
                                        {nama_perusahaan_tarik}<br><br><br><br><br>
                                        <b><u>{pic_tarik}</u></b>
                                    </td>
                                </tr>
                            </table>
                        </body>
                        </html>
                        """

                        pdf_output_lapor = BytesIO()
                        pisa_status_lapor = pisa.CreatePDF(BytesIO(html_pelaporan.encode('utf-8')), dest=pdf_output_lapor)

                        if pisa_status_lapor.err:
                            st.error("Terjadi kesalahan saat merender PDF Pelaporan Penyedia.")
                        else:
                            st.success(f"PDF Dokumen Pelaporan Penyedia {p['nama_paket']} siap diunduh!")
                            st.download_button(
                                label=f"Unduh Dokumen Pelaporan Penyedia {p['nama_paket']}",
                                data=pdf_output_lapor.getvalue(),
                                file_name=f"Pelaporan_Penyedia_{kegiatan}_{p['nama_paket']}.pdf",
                                mime="application/pdf"
                            )
                            
    # === TAB 6: TAHAP PELAPORAN 2 - KADES (KONSOLIDASI) ===
    with tab_pelaporan2:
        st.header("Cetak Laporan Penyerahan ke Kepala Desa (Konsolidasi)")
        st.write("Dokumen ini akan merangkum dan menggabungkan seluruh paket belanja yang telah diselesaikan ke dalam satu Surat dan BAST.")
        
        if not paket_data:
            st.warning("Silakan selesaikan Tahap 1 sampai 5 terlebih dahulu.")
        else:
            col_kds1, col_kds2 = st.columns(2)
            with col_kds1:
                no_surat_penyerahan_kades = st.text_input("Nomor Surat Penyerahan Pekerjaan (PKA -> Kades)", value="001/SP-KADES", key="no_sp_kades_all")
                no_bast_kades = st.text_input("Nomor BAST (PKA -> Kades)", value="001/BAST-KADES", key="no_bast_kades_all")
            with col_kds2:
                tgl_pelaporan_kades = st.date_input("Tanggal Pelaporan ke Kades", tgl_persiapan + timedelta(days=total_hari), key="tgl_lapor_kades_all")
                
            st.markdown("---")
            if st.button("Generate PDF Dokumen Pelaporan - Kades", type="primary"):
                tgl_lengkap_kades, hari_kades, tgl_angka_kades, bulan_kades, tahun_kades = get_tanggal_indo(tgl_pelaporan_kades)
                
                teks_rujukan_bast = ""
                total_semua_paket = 0
                
                konsolidasi_items = {}
                
                for i, p in enumerate(paket_data):
                    no_bast_paket = st.session_state.get(f"no_bast_{i}", f"001/BAST-{i+1}")
                    tgl_lapor_paket = st.session_state.get(f"tgl_lapor_{i}", tgl_pelaporan_kades)
                    tgl_l_pkt, _, _, _, _ = get_tanggal_indo(tgl_lapor_paket)
                    
                    nama_pkt = p['nama_paket'] if p['nama_paket'] else f"Paket {i+1}"
                    teks_rujukan_bast += f"<li style='margin-left: 20px;'>BAST {nama_pkt} Nomor: 100.3.3/{no_bast_paket}/404.608.8/{tahun_kades} tanggal {tgl_l_pkt}</li>"
                    
                    for j, dt_brg in enumerate(p['detail_barang']):
                        v_nego = st.session_state.get(f"nego_vol_{i}_{j}", dt_brg['vol_rencana'])
                        h_nego = st.session_state.get(f"nego_harga_{i}_{j}", dt_brg['harga_hps'])
                        
                        item_key = (dt_brg['uraian'], dt_brg['satuan'], h_nego)
                        
                        if item_key not in konsolidasi_items:
                            konsolidasi_items[item_key] = {
                                'uraian': dt_brg['uraian'],
                                'satuan': dt_brg['satuan'],
                                'harga': h_nego,
                                'vol': 0.0,
                                'jumlah': 0.0,
                                'paket_asal': []
                            }
                            
                        konsolidasi_items[item_key]['vol'] += float(v_nego)
                        konsolidasi_items[item_key]['jumlah'] += (float(v_nego) * float(h_nego))
                        
                        if nama_pkt not in konsolidasi_items[item_key]['paket_asal']:
                            konsolidasi_items[item_key]['paket_asal'].append(nama_pkt)
                
                baris_tabel_konsolidasi = ""
                idx_global = 1
                
                for key, data in konsolidasi_items.items():
                    total_semua_paket += data['jumlah']
                    
                    h_str = f"Rp {data['harga']:,.0f}".replace(",", ".")
                    j_str = f"Rp {data['jumlah']:,.0f}".replace(",", ".")
                    
                    vol_val = data['vol']
                    vol_display = int(vol_val) if vol_val.is_integer() else vol_val
                    
                    paket_info = ", ".join(data['paket_asal'])
                    
                    baris_tabel_konsolidasi += f"""
                    <tr>
                        <td style="text-align: center;">{idx_global}</td>
                        <td>{data['uraian']} <br><i style="font-size:9px;">(Asal: {paket_info})</i></td>
                        <td style="text-align: center;">{vol_display}</td>
                        <td style="text-align: center;">{data['satuan']}</td>
                        <td>{h_str}</td>
                        <td>{j_str}</td>
                    </tr>
                    """
                    idx_global += 1
                        
                terbilang_semua = num2words(total_semua_paket, lang='id').title() + " Rupiah"
                
                html_kades = f"""
                <html>
                <head>
                    <style>
                        @page {{ size: a4 portrait; margin: 2cm; }}
                        body {{ font-family: Helvetica, Arial, sans-serif; font-size: 11px; text-align: justify; line-height: 1.5; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }}
                        th, td {{ border: 1px solid black; padding: 5px; text-align: left; vertical-align: middle; }}
                        th {{ background-color: #f2f2f2; text-align: center; font-weight: bold; font-size: 10px; }}
                        .judul {{ text-align: center; font-weight: bold; font-size: 14px; margin-bottom: 5px; text-decoration: underline; }}
                        .sub-judul {{ text-align: center; font-size: 12px; margin-bottom: 15px; }}
                        .ttd-table {{ border: none; width: 100%; margin-top: 30px; }}
                        .ttd-td {{ border: none; width: 50%; text-align: center; vertical-align: bottom; }}
                        .kop-surat {{ text-align: center; border-bottom: 3px solid black; padding-bottom: 10px; margin-bottom: 15px; }}
                    </style>
                </head>
                <body>
                    <!-- HALAMAN 1: SURAT PENYERAHAN HASIL PEKERJAAN (PKA KE KADES) -->
                    <table style="border: none; margin-bottom: 15px; border-bottom: 3px solid black; padding-bottom: 5px;">
                        <tr style="border: none;">
                            <td style="border: none; width: 15%; text-align: center; vertical-align: middle;">
                                <img src="data:image/png;base64,{logo_ngawi_b64}" width="60">
                            </td>
                            <td style="border: none; width: 85%; text-align: center; vertical-align: middle;">
                                <b style="font-size: 14px;">PEMERINTAH KABUPATEN NGAWI</b><br>
                                <b style="font-size: 14px;">KECAMATAN KWADUNGAN</b><br>
                                <b style="font-size: 16px;">DESA SUMENGKO</b><br>
                                Jalan Raya Desa Sumengko No. 08 Telp. - Kode Pos 63283
                            </td>
                        </tr>
                    </table>
                    
                    <table style="border: none; margin-bottom: 20px;">
                        <tr style="border: none;">
                            <td style="border: none; width: 60%; vertical-align: top;">
                                Nomor : 400.3.3/{no_surat_penyerahan_kades}/404.608.8/{tahun_kades}<br>
                                Perihal : Penyerahan hasil Pekerjaan
                            </td>
                            <td style="border: none; width: 40%; vertical-align: top;">
                                Sumengko, {tgl_lengkap_kades}<br><br>
                                Kepada Yth<br>
                                Kepala Desa<br>
                                Drs. H. SUMARLI, M.Ag<br>
                                Di<br>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;TEMPAT
                            </td>
                        </tr>
                    </table>
                    
                    <p>Sehubungan dengan telah selesainya dilaksanakan pemeriksaan pengadaan barang/jasa serta serah terima pekerjaan dari Penyedia Jasa kepada Kasi/kaur {role_user} berdasarkan berita acara serah terima kegiatan/pekerjaan {kegiatan} antara lain:</p>
                    
                    <ul>
                        {teks_rujukan_bast}
                    </ul>
                    
                    <p>Dengan ini kami {role_user} selaku Pelaksana Kegiatan Anggaran menyerahkan hasil kegiatan/pekerjaan {kegiatan} sepenuhnya kepada Kepala Desa.</p>
                    <p>Demikian surat penyerahan hasil pekerjaan ini atas perhatiannya dan kami ucapkan terima kasih.</p>
                    
                    <table class="ttd-table" style="margin-top: 50px;">
                        <tr style="border: none;">
                            <td class="ttd-td"></td>
                            <td class="ttd-td">
                                {role_user}<br><br><br><br><br>
                                <b><u>{pilih_user}</u></b>
                            </td>
                        </tr>
                    </table>

                    <!-- HALAMAN 2: BA PENYERAHAN HASIL PENGADAAN (PKA KE KADES) -->
                    <pdf:nextpage />
                    <div class="judul" style="text-decoration: none;">BERITA ACARA PENYERAHAN HASIL PENGADAAN</div>
                    <div class="sub-judul" style="margin-bottom: 10px;">Nomor : 100.3.3/{no_bast_kades}/404.608.8/{tahun_kades}</div>
                    
                    <p>Pada hari ini <b>{hari_kades}</b> tanggal <b>{tgl_angka_kades}</b> bulan <b>{bulan_kades}</b> tahun <b>{tahun_kades}</b>, {role_user} selaku Pelaksana kegiatan Anggaran Desa Sumengko yang ditetapkan dengan surat keputusan kepala Desa Sumengko nomor 100.3.3/3/404.608.8/{tahun_kades} tanggal 3 Februari {tahun_kades} tahun anggaran {tahun_kades}, menyerahkan hasil pekerjaan pengadaan {kegiatan} dengan baik kepada Kepala Desa Sumengko berupa:</p>
                    
                    <table>
                        <tr>
                            <th style="width: 5%;">No</th>
                            <th style="width: 40%;">Nama Barang</th>
                            <th style="width: 10%;">Vol</th>
                            <th style="width: 10%;">Satuan</th>
                            <th style="width: 15%;">Harga Satuan</th>
                            <th style="width: 20%;">Jumlah</th>
                        </tr>
                        {baris_tabel_konsolidasi}
                        <tr>
                            <td colspan="5" style="text-align: right;"><b>JUMLAH</b></td>
                            <td><b>Rp {total_semua_paket:,.0f}</b></td>
                        </tr>
                    </table>
                    
                    <p>Terbilang : <i>{terbilang_semua}</i></p>
                    <p>Demikian Berita Acara Penyerahan Hasil Pengadaan dibuat untuk diketahui dan dipergunakan sebagaimana mestinya.</p>
                    
                    <table class="ttd-table">
                        <tr style="border: none;">
                            <td class="ttd-td">
                                <br>Yang menerima<br>
                                Kepala Desa Sumengko<br>
                                Selaku<br>
                                Pemegang Kekuasaan Pengelolaan<br>
                                Keuangan Desa (PKPKD)<br><br><br><br><br>
                                <b>Drs. H. SUMARLI, M.Ag</b>
                            </td>
                            <td class="ttd-td">
                                Sumengko, {tgl_lengkap_kades}<br>
                                Yang menyerahkan<br>
                                Kasi/kaur {role_user}<br>
                                Desa Sumengko<br>
                                selaku Pelaksana Kegiatan Anggaran<br><br><br><br><br>
                                <b><u>{pilih_user}</u></b>
                            </td>
                        </tr>
                    </table>
                </body>
                </html>
                """
                
                pdf_output_kades = BytesIO()
                pisa_status_kades = pisa.CreatePDF(BytesIO(html_kades.encode('utf-8')), dest=pdf_output_kades)

                if pisa_status_kades.err:
                    st.error("Terjadi kesalahan saat merender PDF Pelaporan Kades.")
                else:
                    st.success("PDF Dokumen Konsolidasi Pelaporan Kades siap diunduh!")
                    st.download_button(
                        label="Unduh Dokumen Pelaporan Kades",
                        data=pdf_output_kades.getvalue(),
                        file_name=f"Pelaporan_Kades_{kegiatan}.pdf",
                        mime="application/pdf"
                    )
else:
    st.warning("Menunggu data Master_APBDES.xlsx...")