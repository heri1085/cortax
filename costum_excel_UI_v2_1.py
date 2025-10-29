# --- costum_excel_UI10.py (Streamlit App Lengkap - Minimalis UI & Logika Kode 4) ---

# costum_excel_UI_v2_1.py

import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
import os
import io
import re 

# --- IMPORTS GOOGLE DRIVE BARU ---
# Diperlukan untuk otentikasi dan interaksi API Google
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import json # Diperlukan untuk membaca JSON secrets
# -----------------------------------

# >>> PENTING: IMPORT FUNGSI KONVERSI DI AWAL FILE <<<
# ... (Sisanya dari blok try/except import converter_ui_13) ...

# --- DATA PERUSAHAAN ... (kode COMPANY_DATA) ---

# --- DATA KONFIGURASI GOOGLE DRIVE ---
# Mengambil kredensial dari file .streamlit/secrets.toml
# Ini adalah pengganti FOLDER_OUTPUT_LOKAL

if "google_drive_folder_id" not in st.secrets:
    st.error("❌ Kunci `google_drive_folder_id` tidak ditemukan di secrets.toml.")
else:
    GOOGLE_DRIVE_FOLDER_ID = st.secrets['google_drive_folder_id']

SCOPES = ['https://www.googleapis.com/auth/drive']

@st.cache_resource
def authenticate_gdrive():
    """Mengotentikasi ke Google Drive menggunakan Service Account dari Streamlit Secrets."""
    try:
        # Mengambil credentials dari st.secrets["gcp_service_account"]
        info = dict(st.secrets["gcp_service_account"])
        credentials = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        
        # Membangun service Drive API
        return build('drive', 'v3', credentials=credentials)
    except Exception as e:
        st.error(f"❌ Gagal otentikasi Google Drive. Pastikan `secrets.toml` benar. Error: {e}")
        return None

def upload_file_to_drive(service, file_bytes, filename, folder_id):
    """Mengupload data file (bytes) ke Google Drive."""
    try:
        # Metadata file - HAPUS 'parents' untuk upload ke Shared Drive
        file_metadata = {'name': filename}
        
        # Untuk Shared Drive, tambahkan supportsAllDrives=True
        media = MediaIoBaseUpload(io.BytesIO(file_bytes),
                                  mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                  resumable=True)

        # Upload ke Shared Drive
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='webViewLink',
            supportsAllDrives=True  # ← PENTING untuk Shared Drive
        ).execute()
                          
        return True, f"File berhasil diupload. [Lihat di Drive]({file.get('webViewLink')})"
    except Exception as e:
        return False, f"Gagal mengupload file ke Google Drive. Error: {e}"

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
#FOLDER_OUTPUT_LOKAL = "D:/HasilFaktur"
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
                    
                    # KODE BARU (PENGGANTI KODE LAMA)

                    # 1. Simpan ke buffer memory (bytes)
                    # File harus diubah menjadi bytes agar dapat ditransfer (di-upload atau di-download)
                    excel_data_bytes = to_excel_bytes(df_faktur, df_detail) # Asumsi fungsi ini sudah didefinisikan untuk membuat bytes

                    # 2. Inisiasi Service dan Upload ke Google Drive
                    drive_service = authenticate_gdrive() # Memanggil fungsi otentikasi GDrive dari st.secrets
                    st.info("Sedang mengupload file ke Google Drive...")

                    if drive_service:
                        # Memanggil fungsi upload, menggunakan bytes (bukan path) dan ID folder dari secrets
                        success, message = upload_file_to_drive(
                            drive_service,
                            excel_data_bytes, # Data file dalam bentuk bytes (memori)
                            NAMA_FILE_OUTPUT, # Nama file yang akan terlihat di Drive
                            GOOGLE_DRIVE_FOLDER_ID # ID folder tujuan di Drive
                        )
                        
                        with st.container(border=True): 
                            if success:
                                st.balloons()
                                st.success(f"🎉 **File BERHASIL DIPROSES dan diupload ke Google Drive!**")
                                st.markdown(message) # Menampilkan link ke file yang sudah terupload
                            else:
                                st.error(f"❌ Proses Gagal Upload ke Google Drive. {message}")

                    # 3. Sediakan tombol download lokal sebagai backup (menggunakan bytes dari memori)
                    st.download_button(
                        label="⬇️ Unduh File E-Faktur XLSX (Lokal)",
                        data=excel_data_bytes,
                        file_name=NAMA_FILE_OUTPUT,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="final_xlsx_download"
                    )

            except KeyError as e:
                st.error(f"❌ ERROR: Kolom sumber '**{e.args[0]}**' tidak ditemukan di file data Anda.")
            except Exception as e:
                st.error(f"❌ Terjadi Kesalahan umum saat memproses data: {e}")
                

if __name__ == "__main__":
    main()