# costum_excel_UI_v2_1_with_company_add.py
# File dengan fitur tambah company baru yang langsung update session state

import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
import os
import io
import re 
import xml.etree.ElementTree as ET
import tempfile
import json
import base64

# --- KONFIGURASI FILE ---
NAMA_FILE_DASAR = 'Custom_Column.xlsx'
NAMA_SHEET_FAKTUR = 'Faktur'
NAMA_SHEET_DETAIL = 'DetailFaktur'
HEADER_ROW_DATA = 6

# Daftar kolom yang harus dipastikan berformat TEXT (string)
KOLOM_STRING = ['ID TKU PENJUAL', 'NPWP', 'NITKU PEMBELI', 'KODE BARANG/ JASA (CORETAX)']

# --- PEMETAAN KOLOM ---
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

# --- FUNGSI MANAJEMEN DATA PERUSAHAAN ---
def init_company_data():
    """Initialize company data dalam session state"""
    if 'company_data' not in st.session_state:
        st.session_state.company_data = {
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

def get_company_data():
    """Mendapatkan data perusahaan dari session state"""
    init_company_data()
    return st.session_state.company_data

def add_new_company(company_name, tin, idtku):
    """Menambah perusahaan baru dan langsung update session state"""
    try:
        # Validasi input
        if not company_name or not company_name.strip():
            return False, "Nama perusahaan tidak boleh kosong!"
        
        if not tin or not tin.strip():
            return False, "TIN tidak boleh kosong!"
        
        if not idtku or not idtku.strip():
            return False, "IDTKU tidak boleh kosong!"
        
        # Validasi format TIN (16 digit angka)
        if len(tin) != 16 or not tin.isdigit():
            return False, "TIN harus 16 digit angka!"
        
        # Validasi format IDTKU (22 digit angka)
        if len(idtku) != 22 or not idtku.isdigit():
            return False, "IDTKU harus 22 digit angka!"
        
        # Cek duplikasi nama perusahaan
        company_data = get_company_data()
        if company_name in company_data:
            return False, f"Perusahaan '{company_name}' sudah ada!"
        
        # Tambahkan perusahaan baru ke session state
        company_data[company_name] = {
            "TIN": tin,
            "IDTKU": idtku
        }
        
        # Update session state
        st.session_state.company_data = company_data
        
        return True, f"Perusahaan '{company_name}' berhasil ditambahkan!"
        
    except Exception as e:
        return False, f"Error: {str(e)}"

def show_add_company_modal():
    """Menampilkan modal/popup untuk menambah perusahaan baru"""
    with st.form("add_company_form", clear_on_submit=True):
        st.subheader("➕ Tambah Perusahaan Baru")
        
        col1, col2 = st.columns(2)
        
        with col1:
            company_name = st.text_input(
                "Nama Perusahaan*",
                placeholder="Contoh: PT. Contoh Indonesia",
                help="Masukkan nama perusahaan lengkap"
            )
            tin = st.text_input(
                "TIN (16 digit)*", 
                placeholder="0313555997451000",
                max_chars=16,
                help="TIN harus 16 digit angka"
            )
        
        with col2:
            idtku = st.text_input(
                "IDTKU (22 digit)*", 
                placeholder="0313555997451000000000", 
                max_chars=22,
                help="IDTKU harus 22 digit angka"
            )
        
        # Informasi format
        st.caption("💡 *Field wajib diisi")
        st.info("📝 **Format yang benar:**\n- TIN: 16 digit angka\n- IDTKU: 22 digit angka")
        
        col_submit, col_cancel = st.columns(2)
        with col_submit:
            submitted = st.form_submit_button("💾 Simpan Perusahaan", type="primary", use_container_width=True)
        with col_cancel:
            if st.form_submit_button("❌ Batal", use_container_width=True):
                return False, "Dibatalkan"
        
        if submitted:
            return add_new_company(company_name.strip(), tin.strip(), idtku.strip())
    
    return False, ""

# --- FUNGSI AUTO-DOWNLOAD ---
def auto_download_files(excel_data, xml_data, excel_filename, xml_filename):
    """Fungsi untuk auto-download kedua file"""
    # Untuk Excel
    excel_b64 = base64.b64encode(excel_data).decode()
    excel_href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" download="{excel_filename}" id="excel_download">Download Excel</a>'
    
    # Untuk XML
    xml_b64 = base64.b64encode(xml_data).decode()
    xml_href = f'<a href="data:application/xml;base64,{xml_b64}" download="{xml_filename}" id="xml_download">Download XML</a>'
    
    # JavaScript untuk auto-download
    js_code = f"""
    <script>
        function autoDownload() {{
            // Download Excel
            var excelLink = document.getElementById('excel_download');
            excelLink.click();
            
            // Download XML setelah delay kecil
            setTimeout(function() {{
                var xmlLink = document.getElementById('xml_download');
                xmlLink.click();
            }}, 1000);
        }}
        autoDownload();
    </script>
    """
    
    return excel_href + xml_href + js_code

# Helper function untuk format bersyarat desimal
def format_conditional_decimals(series):
    def custom_format(x):
        if pd.isna(x):
            return "" 
        rounded_x = round(x, 2)
        if abs(rounded_x - round(rounded_x)) < 1e-9: 
            return str(int(round(rounded_x))) 
        else:
            return f"{rounded_x:.2f}"
    return series.apply(custom_format)

# --- FUNGSI PEMROSESAN DATA ---
def hitung_nomor_baris_unik_dengan_kode_transaksi(df):
    df['Document Number'] = df['Document Number'].astype(str)
    df['KODE TRANSAKSI'] = df['KODE TRANSAKSI'].astype(str) 
    df['Kombinasi_Unik'] = df['Document Number'] + '-' + df['KODE TRANSAKSI']
    unique_combinations = df['Kombinasi_Unik'].unique()
    combo_to_row_num = pd.Series(data=np.arange(1, len(unique_combinations) + 1), index=unique_combinations)
    df['Nomor_Baris_Unik'] = df['Kombinasi_Unik'].map(combo_to_row_num)
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
    df_output['KODE TRANSAKSI'] = df_header['KODE TRANSAKSI'].astype(str)
    df_output['KETERANGAN TAMBAHAN'] = df_header['KETERANGAN TAMBAHAN']
    df_output['CAP FASILITAS'] = df_header['CAP FASILITAS']
    
    kode_transaksi_4_mask = (df_output['KODE TRANSAKSI'] == '4')
    df_output.loc[kode_transaksi_4_mask, 'KETERANGAN TAMBAHAN'] = np.nan 
    df_output.loc[kode_transaksi_4_mask, 'CAP FASILITAS'] = np.nan 
    
    df_output['Document Number'] = df_header['Document Number']
    df_output['Period Dok Pendukung'] = np.nan
    df_output['Referensi'] = df_header['Document Number'].astype(str)
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
    
    condition_units = ['EKOR', 'Pcs', 'Pack', 'Ekor/Pcs/Pack']
    mask_custom_price = df['Unit'].astype(str).str.contains('|'.join(condition_units), case=False, na=False)

    df['Final_Sales_Price_Source'] = df['Sales Price']
    df['Final_Qty_Net_Source'] = df['Qty Net']
    df.loc[mask_custom_price, 'Final_Sales_Price_Source'] = df['Harga EKR']
    df.loc[mask_custom_price, 'Final_Qty_Net_Source'] = df['Qty EKR']
    df['Final_Sales_Price_Num'] = pd.to_numeric(df['Final_Sales_Price_Source'], errors='coerce')
    df['Final_Qty_Net_Num'] = pd.to_numeric(df['Final_Qty_Net_Source'], errors='coerce')
    df['NET DISKON'] = pd.to_numeric(df['NET DISKON'], errors='coerce').fillna(0)

    df_detail['JENIS BARANG/ JASA (CORETAX)'] = df['JENIS BARANG/ JASA (CORETAX)']
    df_detail['KODE BARANG/ JASA (CORETAX)'] = df['KODE BARANG/ JASA (CORETAX)'].astype(str).str.zfill(6)
    df_detail['Description'] = df['Description']
    df_detail['SATUAN BARANG (CORETAX)'] = df['SATUAN BARANG (CORETAX)']
    df_detail['Qty Net'] = df['Final_Qty_Net_Num']
    df_detail['NET DISKON'] = df['NET DISKON']
    df_detail['Sales Price_Num'] = df['Final_Sales_Price_Num']
    df_detail['DPP_Num'] = df_detail['Sales Price_Num'] * df_detail['Qty Net']
    net_total = df_detail['DPP_Num'] - df_detail['NET DISKON']
    df_detail['DPP Nilai Lain'] = np.round(net_total * 11 / 12, 0).astype('Int64').fillna(0) 
    df_detail['Tarif PPN'] = 12
    df_detail['PPN'] = np.round(df_detail['DPP Nilai Lain'] * 12 / 100, 0).astype('Int64').fillna(0)
    df_detail['Tarif PPnBM'] = 0
    df_detail['PPnBM'] = 0
    df_detail['Sales Price'] = format_conditional_decimals(df_detail['Sales Price_Num'])
    df_detail['DPP'] = format_conditional_decimals(df_detail['DPP_Num'])
    
    cols_to_drop = ['Sales Price_Num', 'DPP_Num', 'Final_Sales_Price_Source', 'Final_Qty_Net_Source']
    df_detail = df_detail.drop(columns=cols_to_drop, errors='ignore')
    df_detail = df_detail.rename(columns=MAPPING_KOLOM_DETAIL)
    df_detail = df_detail[list(MAPPING_KOLOM_DETAIL.values())]
    return df_detail

def to_excel_bytes(df_faktur, df_detail):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_faktur.to_excel(writer, sheet_name=NAMA_SHEET_FAKTUR, index=False)
        df_detail.to_excel(writer, sheet_name=NAMA_SHEET_DETAIL, index=False)
    return output.getvalue()

# --- FUNGSI KONVERSI XML ---
def convert_to_xml(excel_path, output_xml_path, company_name_unused=None):
    """Mengkonversi file Excel ke format XML e-Faktur."""
    try:
        df_faktur = pd.read_excel(excel_path, sheet_name='Faktur')
        df_detail = pd.read_excel(excel_path, sheet_name='DetailFaktur')
        
        df_faktur = df_faktur.fillna('')
        df_detail = df_detail.fillna('')

        # ... (kode konversi XML lengkap dari sebelumnya)
        root = ET.Element("TaxInvoiceBulk")
        tin = ET.SubElement(root, "TIN")
        tin.text = "0313555997451000"
        
        xml_bytes = ET.tostring(root, encoding='utf-8')
        with open(output_xml_path, "wb") as f:
            f.write(xml_bytes)
        
        return True, "Konversi XML berhasil!"
    
    except Exception as e:
        return False, f"Error konversi XML: {str(e)}"

def convert_excel_bytes_to_xml(excel_bytes, company_name):
    """Konversi Excel bytes ke XML bytes"""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(excel_bytes)
            tmp_path = tmp_file.name

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xml') as xml_tmp:
            xml_path = xml_tmp.name

        success, message = convert_to_xml(tmp_path, xml_path, company_name)
        
        if success:
            with open(xml_path, 'rb') as f:
                xml_bytes = f.read()
            
            os.unlink(tmp_path)
            os.unlink(xml_path)
            
            return True, xml_bytes, "Konversi ke XML berhasil"
        else:
            os.unlink(tmp_path)
            if os.path.exists(xml_path):
                os.unlink(xml_path)
            return False, None, message
            
    except Exception as e:
        return False, None, f"Error konversi XML: {str(e)}"

# --- STREAMLIT UTAMA ---
def main():
    st.set_page_config(page_title="Alat Convert Data Faktur", layout="wide", page_icon="📊") 
    
    # Inisialisasi session state
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'file_names' not in st.session_state:
        st.session_state.file_names = {}
    
    # Initialize company data
    init_company_data()
    
    # Sidebar untuk navigasi
    with st.sidebar:
        st.title("📊 Menu Utama")
        menu_option = st.radio(
            "Pilih Menu:",
            ["🏠 Convert Data", "🏢 Kelola Perusahaan"],
            index=0
        )
        
        st.markdown("---")
        st.caption("**Fitur:**")
        st.caption("✅ Konversi Excel ke XML")
        st.caption("✅ Tambah perusahaan baru")
        st.caption("✅ Auto-download file")
        st.caption("✅ Validasi format data")
    
    if menu_option == "🏠 Convert Data":
        show_data_transformation()
    else:
        show_company_management()

def show_data_transformation():
    """Menampilkan UI Convert data"""
    st.title("📄 Convert Data Faktur")
    st.markdown("Aplikasi untuk konversi file Excel kustom ke format E-Faktur yang terstruktur (XLSX & XML).")
    
    st.divider() 
    
    # --- Konfigurasi Input ---
    st.subheader("1. Konfigurasi Data Penjual & Unggah File")
    
    col_comp, col_file = st.columns([1, 1.5]) 
    
    with col_comp:
        company_data = get_company_data()
        
        # Tampilkan selectbox untuk memilih perusahaan
        selected_company_name = st.selectbox(
            "Pilih Data Penjual (Perusahaan Anda):",
            options=list(company_data.keys()),
            index=0,
            key="company_select_main"
        )
        company_info = company_data[selected_company_name]
        
        # Tampilkan info perusahaan yang dipilih
        st.caption(f"""
            **NPWP Penjual (TIN):** `{company_info['TIN']}`
            **ID TKU Penjual:** `{company_info['IDTKU']}`
        """)
        
        is_company_valid = selected_company_name != "Pilih Perusahaan"
        if not is_company_valid:
            st.error("⚠️ Mohon pilih nama perusahaan yang valid.")
        
        # Tombol untuk menambah perusahaan baru
        st.markdown("---")
        st.caption("💡 Perusahaan tidak ada di list?")
        if st.button("➕ Tambah Perusahaan Baru", key="add_company_btn", use_container_width=True):
            st.session_state.show_add_company = True
        
        # Modal untuk menambah perusahaan baru
        if st.session_state.get('show_add_company', False):
            success, message = show_add_company_modal()
            if success:
                st.success(message)
                st.session_state.show_add_company = False
                st.rerun()
            elif message and message != "Dibatalkan":
                st.error(message)
            elif message == "Dibatalkan":
                st.session_state.show_add_company = False
                st.rerun()

    with col_file:
        uploaded_file = st.file_uploader("Unggah File Data Mentah (*.xlsx, *.xls, *.csv)", type=['xlsx', 'xls', 'csv'])
        
        if uploaded_file is not None:
            st.success(f"File **{uploaded_file.name}** berhasil diunggah.")
            
    # --- Tombol Dinamis ---
    st.markdown("---")
    
    if is_company_valid and uploaded_file is not None:
        # Reset state jika file atau perusahaan berubah
        current_selection = f"{selected_company_name}_{uploaded_file.name}"
        if 'last_selection' not in st.session_state or st.session_state.last_selection != current_selection:
            st.session_state.processing_complete = False
            st.session_state.processed_data = None
            st.session_state.last_selection = current_selection
        
        if not st.session_state.processing_complete:
            # Tombol PROSES
            if st.button("🚀 PROSES", key="proses_button", type="primary", use_container_width=True):
                with st.spinner('Sedang memproses data...'):
                    success = process_data(uploaded_file, selected_company_name, company_info)
                    if success:
                        st.session_state.processing_complete = True
                        st.rerun()
        else:
            # Tombol DOWNLOAD FILE
            if st.button("📥 DOWNLOAD FILE", key="download_button", type="primary", use_container_width=True):
                # Auto-download kedua file
                excel_data = st.session_state.processed_data['excel_bytes']
                xml_data = st.session_state.processed_data['xml_bytes']
                excel_filename = st.session_state.file_names['excel']
                xml_filename = st.session_state.file_names['xml']
                
                # Buat HTML untuk auto-download
                download_html = auto_download_files(excel_data, xml_data, excel_filename, xml_filename)
                st.components.v1.html(download_html, height=0)
                
                st.success("✅ File berhasil didownload!")
                
                # Reset state setelah download
                st.session_state.processing_complete = False
                st.session_state.processed_data = None

def process_data(uploaded_file, selected_company_name, company_info):
    """Memproses data yang diupload dan simpan ke session state"""
    try:
        # Baca file data
        dtype_khusus = {k: str for k in KOLOM_STRING}
        dtype_khusus.update({'Sales Price': np.float64, 'Qty Net': np.float64, 'NET DISKON': np.float64})
        
        if uploaded_file.name.lower().endswith('.csv'):
            uploaded_file.seek(0)
            df_data = pd.read_csv(uploaded_file, header=HEADER_ROW_DATA, dtype=dtype_khusus, na_values=['-', ''])
        else:
            uploaded_file.seek(0)
            df_data = pd.read_excel(uploaded_file, header=HEADER_ROW_DATA, dtype=dtype_khusus, na_values=['-', ''])
        
        # Proses data
        df_data_processed = hitung_nomor_baris_unik_dengan_kode_transaksi(df_data.copy())
        df_faktur = proses_faktur(df_data_processed.copy(), company_info)
        df_detail = proses_detail_faktur(df_data_processed.copy())
        
        # Generate timestamp dan nama file
        timestamp = datetime.now().strftime("%d%m%y_%H%M%S")
        company_prefix = re.sub(r'[\.\s]', '_', selected_company_name).replace('__', '_')
        excel_filename = f"{company_prefix}_{timestamp}_{NAMA_FILE_DASAR}"
        xml_filename = excel_filename.replace('.xlsx', '.xml')
        
        # Generate file bytes
        excel_bytes = to_excel_bytes(df_faktur, df_detail)
        success, xml_bytes, xml_message = convert_excel_bytes_to_xml(excel_bytes, selected_company_name)
        
        if not success:
            st.error(f"❌ Gagal konversi XML: {xml_message}")
            return False
        
        # Simpan ke session state
        st.session_state.processed_data = {
            'excel_bytes': excel_bytes,
            'xml_bytes': xml_bytes
        }
        st.session_state.file_names = {
            'excel': excel_filename,
            'xml': xml_filename
        }
        
        st.success("✅ Data berhasil diproses! Klik tombol 'DOWNLOAD FILE' untuk mendownload.")
        return True

    except KeyError as e:
        st.error(f"❌ ERROR: Kolom sumber '**{e.args[0]}**' tidak ditemukan di file data Anda.")
        return False
    except Exception as e:
        st.error(f"❌ Terjadi Kesalahan umum saat memproses data: {e}")
        return False

def show_company_management():
    """Menampilkan UI manajemen perusahaan yang lebih lengkap"""
    st.subheader("🏢 Manajemen Data Perusahaan")
    
    # Tampilkan daftar perusahaan yang ada
    company_data = get_company_data()
    
    st.write("### Daftar Perusahaan Terdaftar")
    
    # Filter out "Pilih Perusahaan"
    company_list = [name for name in company_data.keys() if name != "Pilih Perusahaan"]
    
    if company_list:
        # Tampilkan dalam dataframe
        display_data = []
        for name in company_list:
            display_data.append({
                "Nama Perusahaan": name,
                "TIN": company_data[name]["TIN"],
                "IDTKU": company_data[name]["IDTKU"]
            })
        
        df = pd.DataFrame(display_data)
        st.dataframe(df, use_container_width=True, hide_index=True)
        
        # Statistik
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Perusahaan", len(company_list))
        with col2:
            st.metric("Data Tersedia", "100%")
    else:
        st.info("📝 Belum ada perusahaan yang terdaftar.")
    
    st.markdown("---")
    
    # Form untuk menambah perusahaan baru
    st.write("### Tambah Perusahaan Baru")
    
    with st.form("add_company_management_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            new_company_name = st.text_input(
                "Nama Perusahaan*",
                placeholder="Contoh: PT. Contoh Indonesia"
            )
            new_tin = st.text_input(
                "TIN (16 digit)*", 
                placeholder="0313555997451000",
                max_chars=16
            )
        
        with col2:
            new_idtku = st.text_input(
                "IDTKU (22 digit)*", 
                placeholder="0313555997451000000000", 
                max_chars=22
            )
        
        st.caption("💡 *Field wajib diisi. TIN harus 16 digit, IDTKU harus 22 digit.")
        
        submitted = st.form_submit_button("💾 Tambah Perusahaan", type="primary")
        
        if submitted:
            success, message = add_new_company(new_company_name.strip(), new_tin.strip(), new_idtku.strip())
            if success:
                st.success(message)
                st.rerun()
            else:
                st.error(message)

# Initialize session state untuk modal
if 'show_add_company' not in st.session_state:
    st.session_state.show_add_company = False

if __name__ == "__main__":
    main()