# --- costum_excel_UI10.py (Streamlit App Lengkap - Minimalis UI & Logika Kode 4) ---

import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
import os
import io
import re 

# >>> PENTING: IMPORT FUNGSI KONVERSI DI AWAL FILE <<<
try:
    # Asumsi: File ini ada di direktori yang sama
    from converter_ui_13 import convert_to_xml 
except ImportError:
    # Menampilkan peringatan, namun aplikasi tetap berjalan tanpa fungsi XML
    st.warning("⚠️ Peringatan: File 'converter_ui_13.py' tidak ditemukan. Konversi XML dinonaktifkan.")
    def convert_to_xml(excel_path, output_xml_path, company_name):
        return False, "Fungsi konversi XML tidak tersedia (converter_ui_13.py hilang)."


# --- DATA PERUSAHAAN (Harus SAMA dengan converter_ui_7.py) ---
COMPANY_DATA = {
    "Pilih Perusahaan": {
        "TIN": "",
        "IDTKU": ""
    },
    "PT. Citraguna Lestari": {
        "TIN": "0313555997451000",
        "IDTKU": "0313555997451000000000"
    },
    "PT. Efran Berkat Aditama": {
        "TIN": "0933679607416000",
        "IDTKU": "0933679607416000000000"
    }
}
# --- KONFIGURASI FILE ---
NAMA_FILE_DASAR = 'Custom_Column.xlsx'
NAMA_SHEET_FAKTUR = 'Faktur'
NAMA_SHEET_DETAIL = 'DetailFaktur'
HEADER_ROW_DATA = 6
FOLDER_OUTPUT_LOKAL = "D:/HasilFaktur"
JUMLAH_PER_HALAMAN = 20 

# Daftar kolom yang harus dipastikan berformat TEXT (string)
KOLOM_STRING = ['ID TKU PENJUAL', 'NPWP', 'NITKU PEMBELI', 'KODE BARANG/ JASA (CORETAX)']

# --- PEMETAAN KOLOM FAKTUR ---
MAPPING_KOLOM_FAKTUR = {
    'Baris': 'Baris',
    'Date': 'Tanggal Faktur',
    'Jenis Faktur': 'Jenis Faktur',
    'KODE TRANSAKSI': 'Kode Transaksi',
    'KETERANGAN TAMBAHAN': 'Keterangan Tambahan',
    'Document Number': 'Dokumen Pendukung',
    'Period Dok Pendukung': 'Period Dok Pendukung',
    'Referensi': 'Referensi',
    'CAP FASILITAS': 'Cap Fasilitas',
    'ID TKU PENJUAL': 'ID TKU Penjual',
    'NPWP': 'NPWP/NIK Pembeli',
    'JENIS ID': 'Jenis ID Pembeli',
    'Negara Pembeli': 'Negara Pembeli',
    'Nomor Dokumen Pembeli': 'Nomor Dokumen Pembeli',
    'NAMA NPWP CUSTOMER': 'Nama Pembeli',
    'ALAMAT': 'Alamat Pembeli',
    'Email Pembeli': 'Email Pembeli',
    'NITKU PEMBELI': 'ID TKU Pembeli'
}

# --- PEMETAAN KOLOM DETAIL FAKTUR ---
MAPPING_KOLOM_DETAIL = {
    'Baris': 'Baris',
    'JENIS BARANG/ JASA (CORETAX)': 'Barang/Jasa',
    'KODE BARANG/ JASA (CORETAX)': 'Kode Barang Jasa',
    'Description': 'Nama Barang/Jasa',
    'SATUAN BARANG (CORETAX)': 'Nama Satuan Ukur',
    'Sales Price': 'Harga Satuan',
    'Qty Net': 'Jumlah Barang Jasa',
    'NET DISKON': 'Total Diskon',
    'DPP': 'DPP',
    'DPP Nilai Lain': 'DPP Nilai Lain',
    'Tarif PPN': 'Tarif PPN',
    'PPN': 'PPN',
    'Tarif PPnBM': 'Tarif PPnBM',
    'PPnBM': 'PPnBM'
}

# Helper function untuk format bersyarat desimal
def format_conditional_decimals(series):
    def custom_format(x):
        if pd.isna(x):
            return "" 
        rounded_x = round(x, 2)
        if abs(rounded_x - round(rounded_x)) < 1e-9: 
            return str(int(round(rounded_x))) 
        else:
            # Gunakan f-string untuk memastikan 2 desimal jika bukan integer utuh
            return f"{rounded_x:.2f}"
            
    return series.apply(custom_format)

# --- FUNGSI PEMROSESAN DATA ---
def hitung_nomor_baris_unik_dengan_kode_transaksi(df):
    # 1. Pastikan kolom-kolom kunci bertipe string
    df['Document Number'] = df['Document Number'].astype(str)
    # Pastikan KODE TRANSAKSI adalah string (penting untuk filter '4')
    df['KODE TRANSAKSI'] = df['KODE TRANSAKSI'].astype(str) 

    # 2. Buat kolom kombinasi unik
    df['Kombinasi_Unik'] = df['Document Number'] + '-' + df['KODE TRANSAKSI']
    
    # 3. Identifikasi kombinasi unik
    unique_combinations = df['Kombinasi_Unik'].unique()
    
    # 4. Buat pemetaan (mapping) ID
    combo_to_row_num = pd.Series(
        data=np.arange(1, len(unique_combinations) + 1), 
        index=unique_combinations
    )
    
    # 5. Terapkan nomor baris unik
    df['Nomor_Baris_Unik'] = df['Kombinasi_Unik'].map(combo_to_row_num)
    
    # 6. Hapus kolom bantuan (opsional, untuk menjaga kebersihan DF)
    df = df.drop(columns=['Kombinasi_Unik'])
    
    return df

@st.cache_data
def proses_faktur(df, company_info):
    df_header = df.drop_duplicates(subset=['Nomor_Baris_Unik'], keep='first').copy()
    df_output = pd.DataFrame()
    df_output['Baris'] = df_header['Nomor_Baris_Unik']
    df_header['Date'] = pd.to_datetime(df_header['Date'], errors='coerce')
    df_output['Date'] = df_header['Date'].dt.strftime('%d/%m/%Y').fillna(np.nan)
    df_output['Jenis Faktur'] = 'Normal'
    
    # Kolom Kunci untuk Logika
    df_output['KODE TRANSAKSI'] = df_header['KODE TRANSAKSI'].astype(str) # Pastikan string
    df_output['KETERANGAN TAMBAHAN'] = df_header['KETERANGAN TAMBAHAN']
    df_output['CAP FASILITAS'] = df_header['CAP FASILITAS']
    
    # --- START: LOGIKA BARU KODE TRANSAKSI = 4 ---
    kode_transaksi_4_mask = (df_output['KODE TRANSAKSI'] == '4')
    
    # Jika KODE TRANSAKSI adalah '4', set KETERANGAN TAMBAHAN menjadi NaN
    df_output.loc[kode_transaksi_4_mask, 'KETERANGAN TAMBAHAN'] = np.nan 
    
    # Jika KODE TRANSAKSI adalah '4', set CAP FASILITAS menjadi NaN
    df_output.loc[kode_transaksi_4_mask, 'CAP FASILITAS'] = np.nan 
    # --- END: LOGIKA BARU KODE TRANSAKSI = 4 ---
    
    df_output['Document Number'] = df_header['Document Number']
    df_output['Period Dok Pendukung'] = np.nan
    df_output['Referensi'] = df_header['Document Number'].astype(str)
    
    # --- Menggunakan data perusahaan yang dipilih ---
    df_output['ID TKU PENJUAL'] = company_info['IDTKU']
    df_output['NPWP'] = df_header['NPWP'].astype(str) 

    df_output['JENIS ID'] = df_header['JENIS ID']
    df_output['Negara Pembeli'] = 'IDN'
    df_output['Nomor Dokumen Pembeli'] = df_output['NPWP']
    df_output['NAMA NPWP CUSTOMER'] = df_header['NAMA NPWP CUSTOMER']
    df_output['ALAMAT'] = df_header['ALAMAT']
    df_output['Email Pembeli'] = np.nan
    df_output['NITKU PEMBELI'] = df_header['NITKU PEMBELI'].astype(str)
    
    df_output = df_output.rename(columns=MAPPING_KOLOM_FAKTUR)
    df_output = df_output[list(MAPPING_KOLOM_FAKTUR.values())]
    return df_output

@st.cache_data
def proses_detail_faktur(df):
    df_detail = pd.DataFrame()
    df_detail['Baris'] = df['Nomor_Baris_Unik']
    
    # --- START: LOGIKA BARU BERDASARKAN UNIT ---
    
    # 1. Tentukan kondisi: unit mengandung 'Ekor', 'Pcs', atau 'Pack' (tidak case-sensitive)
    condition_units = ['EKOR', 'Pcs', 'Pack', 'Ekor/Pcs/Pack']
    mask_custom_price = df['Unit'].astype(str).str.contains('|'.join(condition_units), case=False, na=False)

    # 2. Buat kolom sumber akhir
    df['Final_Sales_Price_Source'] = df['Sales Price']
    df['Final_Qty_Net_Source'] = df['Qty Net']

    # 3. Terapkan logika dinamis: jika unit 'Ekor/Pcs/Pack', ambil dari kolom EKR
    df.loc[mask_custom_price, 'Final_Sales_Price_Source'] = df['Harga EKR']
    df.loc[mask_custom_price, 'Final_Qty_Net_Source'] = df['Qty EKR']

    # 4. Konversi kolom sumber akhir ke numerik
    df['Final_Sales_Price_Num'] = pd.to_numeric(df['Final_Sales_Price_Source'], errors='coerce')
    df['Final_Qty_Net_Num'] = pd.to_numeric(df['Final_Qty_Net_Source'], errors='coerce')
    
    # Konversi NET DISKON
    df['NET DISKON'] = pd.to_numeric(df['NET DISKON'], errors='coerce').fillna(0)
    
    # --- END: LOGIKA BARU BERDASARKAN UNIT ---

    # 1. Kolom non-numerik awal
    df_detail['JENIS BARANG/ JASA (CORETAX)'] = df['JENIS BARANG/ JASA (CORETAX)']
    df_detail['KODE BARANG/ JASA (CORETAX)'] = df['KODE BARANG/ JASA (CORETAX)'].astype(str).str.zfill(6)
    df_detail['Description'] = df['Description']
    df_detail['SATUAN BARANG (CORETAX)'] = df['SATUAN BARANG (CORETAX)']
    
    # Menggunakan kolom kuantitas yang baru
    df_detail['Qty Net'] = df['Final_Qty_Net_Num'] # Di sini seharusnya 60
    df_detail['NET DISKON'] = df['NET DISKON']
    
    # ... (lanjutan perhitungan dan finalisasi tetap sama)
    # 2. Perhitungan: Simpan Sales Price dan DPP dalam bentuk numerik
    df_detail['Sales Price_Num'] = df['Final_Sales_Price_Num']
    
    # Kalkulasi DPP menggunakan nilai numerik yang sudah dipilih
    df_detail['DPP_Num'] = df_detail['Sales Price_Num'] * df_detail['Qty Net']
    
    # 3. Kalkulasi turunan menggunakan nilai numerik
    net_total = df_detail['DPP_Num'] - df_detail['NET DISKON']
    
    # Perhitungan DPP Nilai Lain dan PPN/PPnBM (Pembulatan ke integer terdekat)
    df_detail['DPP Nilai Lain'] = np.round(net_total * 11 / 12, 0).astype('Int64').fillna(0) 
    df_detail['Tarif PPN'] = 12
    df_detail['PPN'] = np.round(df_detail['DPP Nilai Lain'] * 12 / 100, 0).astype('Int64').fillna(0)
    df_detail['Tarif PPnBM'] = 0
    df_detail['PPnBM'] = 0

    # 4. Terapkan format string ke Sales Price dan DPP untuk output final
    df_detail['Sales Price'] = format_conditional_decimals(df_detail['Sales Price_Num'])
    df_detail['DPP'] = format_conditional_decimals(df_detail['DPP_Num'])
    
    # 5. Hapus kolom bantuan numerik
    cols_to_drop = ['Sales Price_Num', 'DPP_Num', 'Final_Sales_Price_Source', 'Final_Qty_Net_Source']
    df_detail = df_detail.drop(columns=cols_to_drop, errors='ignore')

    # 6. Finalisasi
    df_detail = df_detail.rename(columns=MAPPING_KOLOM_DETAIL)
    df_detail = df_detail[list(MAPPING_KOLOM_DETAIL.values())]
    return df_detail

def to_excel_bytes(df_faktur, df_detail):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_faktur.to_excel(writer, sheet_name=NAMA_SHEET_FAKTUR, index=False)
        df_detail.to_excel(writer, sheet_name=NAMA_SHEET_DETAIL, index=False)
    return output.getvalue()

# Fungsi untuk membaca file ke bytes
def to_file_bytes(file_path):
    try:
        if not os.path.exists(file_path):
             return None
             
        with open(file_path, 'rb') as f:
            return f.read()
    except Exception:
        return None

# Fungsi untuk mengambil daftar file
def ambil_daftar_file(folder_path):
    import re
    if not os.path.exists(folder_path):
        return pd.DataFrame(columns=["Tanggal", "Waktu", "Nama File", "Alamat Folder", "Tanggal_Date", "Tanggal_Waktu_Sort"])
        
    file_list = []
    for file_name in os.listdir(folder_path):
        # Memastikan hanya file output yang diproses
        if file_name.lower().endswith(NAMA_FILE_DASAR.lower()) and "output" not in file_name.lower():
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):
                tanggal_display = "Tidak Diketahui"
                waktu_display = "00:00:00"
                tanggal_obj = None

                try:
                    match = re.search(r'(\d{6}[\s_]\d{6})', file_name) 
                    if match:
                        timestamp_raw = match.group(1)
                        timestamp_str = timestamp_raw.replace('_', ' ')
                        
                        tanggal_obj = datetime.strptime(timestamp_str, "%d%m%y %H%M%S")
                        tanggal_display = tanggal_obj.strftime("%d/%m/%Y")
                        waktu_display = tanggal_obj.strftime("%H:%M:%S") 
                except Exception:
                    pass
                
                sort_value = tanggal_obj if tanggal_obj else datetime.min
                    
                file_list.append({
                    "Tanggal": tanggal_display,
                    "Waktu": waktu_display, 
                    "Nama File": file_name,
                    "Alamat Folder": folder_path, 
                    "Tanggal_Date": tanggal_obj.date() if tanggal_obj else datetime.min.date(),
                    "Tanggal_Waktu_Sort": sort_value
                })
                
    df_files = pd.DataFrame(file_list)
    
    if not df_files.empty:
        df_files = df_files.sort_values(by="Tanggal_Waktu_Sort", ascending=False).reset_index(drop=True)
        
    return df_files

# --- Fungsi Paginasi dengan Tampilan Minimalis ---
def tampilkan_tabel_dengan_paginasi(df, jumlah_per_halaman, selected_company_name):
    if 'current_page' not in st.session_state:
        st.session_state.current_page = 1
        
    total_rows = len(df)
    total_halaman = (total_rows - 1) // jumlah_per_halaman + 1
    
    if total_halaman > 0:
        if st.session_state.current_page > total_halaman:
            st.session_state.current_page = total_halaman
        if st.session_state.current_page < 1:
            st.session_state.current_page = 1

    halaman = st.session_state.current_page 
    
    if total_rows == 0:
        st.info("Tidak ada data untuk ditampilkan.")
        return

    st.caption(f"Menampilkan halaman {halaman} dari {total_halaman} (Total {total_rows} File)")
    
    # Navigasi paginasi di atas tabel
    col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 8])
    
    with col_nav1:
        if st.button("⬅️", disabled=halaman == 1, key="prev_btn", help="Halaman Sebelumnya"):
            st.session_state.current_page -= 1
            st.rerun()
            
    with col_nav2:
        if st.button("➡️", disabled=halaman == total_halaman, key="next_btn", help="Halaman Berikutnya"):
            st.session_state.current_page += 1
            st.rerun()
            
    # Tampilkan Header Tabel (CSS Minimalis)
    st.markdown("""
        <style>
        /* Gaya baru untuk tampilan lebih bersih */
        .header-row { 
            display: flex; 
            font-weight: bold; 
            border-bottom: 1px solid #ddd; /* Garis pemisah yang tipis */
            padding: 8px 0; 
            margin-bottom: 0px; 
            font-size: 14px; 
            color: #444; /* Warna teks yang tidak terlalu mencolok */
        }
        .col-no { flex: 0.8; text-align: center; } 
        .col-tgl { flex: 1.5; } 
        .col-waktu { flex: 1; } 
        .col-nama { flex: 4; } 
        .col-btn1 { flex: 1.5; text-align: center; } /* Konversi */
        .col-btn2 { flex: 1.5; text-align: center; } /* Unduh XML */
        .col-btn3 { flex: 1.5; text-align: center; } /* Unduh XLSX */
        .row-item { 
            display: flex; 
            padding: 5px 0; 
            border-bottom: 1px dashed #eee; /* Garis putus-putus */
            align-items: center;
        }
        </style>
        <div class="header-row">
            <div class="col-no">No.</div>
            <div class="col-tgl">Tanggal</div>
            <div class="col-waktu">Waktu</div>
            <div class="col-nama">Nama File XLSX</div>
            <div class="col-btn1">Konversi</div>
            <div class="col-btn2">Unduh XML</div>
            <div class="col-btn3">Unduh XLSX</div>
        </div>
    """, unsafe_allow_html=True)
    
    # Data baris
    mulai = (halaman - 1) * jumlah_per_halaman
    df_slice = df.iloc[mulai: mulai + jumlah_per_halaman]
    
    for index_in_slice, (original_index, row) in enumerate(df_slice.iterrows()):
        global_row_num = mulai + index_in_slice + 1
        
        # Gunakan st.markdown untuk membungkus baris agar CSS kustom diterapkan
        st.markdown('<div class="row-item">', unsafe_allow_html=True)
        
        col0, col1, col2, col3, col4, col5, col6 = st.columns([0.8, 1.5, 1, 4, 1.5, 1.5, 1.5])
        
        col0.markdown(f'<div class="col-no">{global_row_num}</div>', unsafe_allow_html=True)
        col1.markdown(f'<div class="col-tgl">{row["Tanggal"]}</div>', unsafe_allow_html=True)
        col2.markdown(f'<div class="col-waktu">{row["Waktu"]}</div>', unsafe_allow_html=True)
        col3.markdown(f'<div class="col-nama" style="font-size: 13px;">{row["Nama File"]}</div>', unsafe_allow_html=True)
        
        # Path File
        xlsx_name = row["Nama File"]
        excel_path = os.path.join(row["Alamat Folder"], xlsx_name)
        xlsx_exists = os.path.exists(excel_path)
        
        xlsx_name_without_ext = os.path.splitext(xlsx_name)[0]
        PREDICTABLE_XML_NAME = f"{xlsx_name_without_ext}.xml"
        output_xml_path = os.path.join(row["Alamat Folder"], PREDICTABLE_XML_NAME)
        xml_exists = os.path.exists(output_xml_path)

        # Tombol Konversi ke XML (col4)
        with col4:
             if st.button("⚙️ Konversi", key=f"konversi_{original_index}", disabled=xml_exists or 'convert_to_xml' not in globals()):
                try:
                    with st.spinner('Memproses konversi ke XML...'):
                        success, message = convert_to_xml(excel_path, output_xml_path, selected_company_name)
                    
                    if success:
                        st.toast(f"✅ Konversi berhasil!", icon='⚙️')
                        st.rerun()
                    else:
                        st.error(f"❌ Gagal konversi: {message}")
                except Exception as e:
                    st.error(f"Terjadi kesalahan konversi XML: {e}")
        
        # Tombol Unduh XML (col5)
        with col5:
            if xml_exists:
                xml_data = to_file_bytes(output_xml_path) 
                st.download_button(
                    label="⬇️ XML", # Label yang lebih minimalis
                    data=xml_data if xml_data is not None else "",
                    file_name=PREDICTABLE_XML_NAME,
                    mime="application/xml",
                    key=f"download_xml_{original_index}", 
                    disabled=xml_data is None
                )
            else:
                st.button("Belum Ada", disabled=True, key=f"xml_na_{original_index}")
             
        # Tombol Unduh XLSX (col6) - Trigger/Download
        with col6:
            if xlsx_exists:
                btn_key = f"trigger_xlsx_download_{original_index}"
                
                if btn_key not in st.session_state:
                    st.session_state[btn_key] = False
                
                # Tombol Download sesungguhnya
                if st.session_state[btn_key] and f"xlsx_data_{original_index}" in st.session_state:
                    st.download_button(
                        label="⬇️ Unduh", # Label minimalis
                        data=st.session_state[f"xlsx_data_{original_index}"],
                        file_name=xlsx_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"xlsx_btn_2_{original_index}"
                    )
                    # Reset status
                    del st.session_state[f"xlsx_data_{original_index}"]
                    st.session_state[btn_key] = False
                else:
                    # Tombol Trigger
                    if st.button("Muat File", key=f"xlsx_btn_1_{original_index}", help="Klik ini untuk memuat data ke memori, kemudian klik tombol unduh yang muncul."):
                        xlsx_data = to_file_bytes(excel_path)
                        st.session_state[f"xlsx_data_{original_index}"] = xlsx_data
                        st.session_state[btn_key] = True
                        st.rerun()
            else:
                 st.button("Hilang", disabled=True, key=f"xlsx_na_{original_index}")
                 
        st.markdown('</div>', unsafe_allow_html=True) # Penutup row-item


# --- STREAMLIT UTAMA ---
def main():
    # Menggunakan layout wide untuk tampilan 'fresh'
    st.set_page_config(page_title="Alat Transformasi Data", layout="wide") 
    
    # 🌟 Judul Minimalis
    st.title("📄 Transformasi Data Faktur")
    st.markdown("Aplikasi untuk konversi file Excel kustom ke format E-Faktur yang terstruktur (XLSX & XML).")
    
    st.divider() 
    
    # --- 1. Konfigurasi Input (Minimalis dengan Kolom) ---
    st.subheader("1. Konfigurasi Data Penjual & Unggah File")
    
    col_comp, col_file = st.columns([1, 1.5]) 
    
    with col_comp:
        selected_company_name = st.selectbox(
            "Pilih Data Penjual (Perusahaan Anda):",
            options=list(COMPANY_DATA.keys()),
            index=0,
            key="company_select_minimal"
        )
        company_info = COMPANY_DATA[selected_company_name]
        
        # Menggunakan st.caption untuk informasi yang lebih minimalis
        st.caption(f"""
            **NPWP Penjual (TIN):** `{company_info['TIN']}`
            **ID TKU Penjual:** `{company_info['IDTKU']}`
        """)
        
        is_company_valid = selected_company_name != "Pilih Perusahaan"
        if not is_company_valid:
            st.error("⚠️ Mohon pilih nama perusahaan yang valid.")

    with col_file:
        uploaded_file = st.file_uploader("Unggah File Data Mentah (*.xlsx, *.xls, *.csv)", type=['xlsx', 'xls', 'csv'])
        
        if uploaded_file is not None:
            st.success(f"File **{uploaded_file.name}** berhasil diunggah.")
            
    # --- Tombol Proses di bawah area upload ---
    st.markdown("---")
    if is_company_valid and uploaded_file is not None:
        if st.button("🚀 PROSES & SIMPAN FILE", key="proses_button_minimal", type="primary"):
            try:
                # Baca file data
                dtype_khusus = {k: str for k in KOLOM_STRING}
                dtype_khusus.update({'Sales Price': np.float64, 'Qty Net': np.float64, 'NET DISKON': np.float64})
                if uploaded_file.name.lower().endswith('.csv'):
                    # PENTING: Baca file yang diunggah ulang agar tidak ada masalah cache saat st.rerun
                    uploaded_file.seek(0)
                    df_data = pd.read_csv(uploaded_file, header=HEADER_ROW_DATA, dtype=dtype_khusus, na_values=['-', ''])
                else:
                    uploaded_file.seek(0)
                    df_data = pd.read_excel(uploaded_file, header=HEADER_ROW_DATA, dtype=dtype_khusus, na_values=['-', ''])
                
                with st.spinner('Sedang memproses data dan menyimpan lokal...'):
                    
                    df_data_processed = hitung_nomor_baris_unik_dengan_kode_transaksi(df_data.copy())
                    df_faktur = proses_faktur(df_data_processed.copy(), company_info)
                    df_detail = proses_detail_faktur(df_data_processed.copy())
                    
                    timestamp = datetime.now().strftime("%d%m%y %H%M%S")
                    
                    # LOGIKA UNTUK NAMA FILE XLSX (SUDAH TERMASUK NAMA PERUSAHAAN)
                    company_prefix = re.sub(r'[\.\s]', '_', selected_company_name).replace('__', '_')
                    NAMA_FILE_OUTPUT = f"{company_prefix}_{timestamp} {NAMA_FILE_DASAR}"
                    
                    os.makedirs(FOLDER_OUTPUT_LOKAL, exist_ok=True)
                        
                    file_path = os.path.join(FOLDER_OUTPUT_LOKAL, NAMA_FILE_OUTPUT)
                    
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        df_faktur.to_excel(writer, sheet_name=NAMA_SHEET_FAKTUR, index=False)
                        df_detail.to_excel(writer, sheet_name=NAMA_SHEET_DETAIL, index=False)
                        
                    # Tampilkan feedback minimalis
                    st.toast('✅ PROSES BERHASIL!', icon='🎉')
                    
                    with st.container(border=True): 
                        st.success(f"🎉 **File BERHASIL DIPROSES!**")
                        st.markdown(f"File output **{NAMA_FILE_OUTPUT}** telah disimpan ke **`{FOLDER_OUTPUT_LOKAL}`** dan tersedia di bagian Riwayat.")
                        
                        # Tampilkan preview singkat hasil
                        st.markdown("**Preview Hasil Faktur:**")
                        st.dataframe(df_faktur.head(1).style.set_properties(**{'font-size': '8pt'}), hide_index=True)
                        st.rerun() # Rerun untuk update riwayat

            except KeyError as e:
                st.error(f"❌ ERROR: Kolom sumber '**{e.args[0]}**' tidak ditemukan di file data Anda.")
            except Exception as e:
                st.error(f"❌ Terjadi Kesalahan umum saat memproses data: {e}")
                

    # --- 2. Riwayat File (Minimalis & Filter) ---
    st.divider()
    st.subheader("2. Riwayat File, Konversi & Unduh")
    
    df_file_history = ambil_daftar_file(FOLDER_OUTPUT_LOKAL)
    
    if df_file_history.empty:
        st.info(f"Belum ada file yang diproses di folder `{FOLDER_OUTPUT_LOKAL}`.")
    else:
        # --- Bagian Filter Tanggal ---
        min_date = df_file_history['Tanggal_Date'].min()
        max_date = df_file_history['Tanggal_Date'].max()
        
        col_date1, col_date2, _ = st.columns([1, 1, 3])
        
        with col_date1:
            start_date = st.date_input("Filter Tanggal Mulai", value=min_date, min_value=min_date, max_value=max_date, key="date_start_min")
        
        with col_date2:
            end_date = st.date_input("Filter Tanggal Akhir", value=max_date, min_value=min_date, max_value=max_date, key="date_end_min")
            
        df_filtered = df_file_history[
            (df_file_history['Tanggal_Date'] >= start_date) & 
            (df_file_history['Tanggal_Date'] <= end_date)
        ].copy()
        
        # Reset paginasi jika filter berubah
        if 'last_filter_len' not in st.session_state or st.session_state.last_filter_len != len(df_filtered):
            st.session_state.current_page = 1
            st.session_state.last_filter_len = len(df_filtered)

        if df_filtered.empty:
            st.warning("Tidak ada file yang ditemukan dalam rentang tanggal yang dipilih.")
        else:
            tampilkan_tabel_dengan_paginasi(df_filtered, JUMLAH_PER_HALAMAN, selected_company_name)

if __name__ == "__main__":
    main()