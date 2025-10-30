# costum_excel_UI_v2_1_complete.py
# File gabungan lengkap dengan Google Drive integration dan XML converter

import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
import os
import io
import re 
import xml.etree.ElementTree as ET

# --- IMPORTS GOOGLE DRIVE ---
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- GOOGLE DRIVE CREDENTIALS (LANGSUNG DI DALAM KODE) ---
GOOGLE_DRIVE_CREDENTIALS = {
    "type": "service_account",
    "project_id": "deep-wares-476608-k6",
    "private_key_id": "d8493ad42b1550b48bd44229b1e27fcafb19198b",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDrtmbnGcGsiXPF\noqibYtd+dNdOnBbo8n1AOkM/nKfPi/WXQmFI4qVAQoGh8xigPprRdTYgzSkoDFpZ\nA4qDXqpoCkWf7mK1lHwEWVhgq7wbz/8F7jwai3e0q67y8+BOV5nK5ECVSo2i7I/e\naBJTf4FoOdPl1TQ4oKloIfAUZn9myhbOoNG7uDMqJyVrbE3unPwbpLWcLlgs47iN\nb7NA6MpYVdI/G5JNxwdccHoFxUHKJDOlXaoJvDRjvgYY+PnhrtFR+y7I/5J6DoJu\nSI4mb0ZFiqSPkp5N46OAmlvTSvc9oagRfeNhF+3HK9clO9l+UxVz9p/7vuQO3d+A\nz+xhPm6tAgMBAAECggEAFCDo664+5Dl+qe9of5Byv4lBw6Q/rpfai0wOZxW9HNOi\nXs7b2OoF25VyVX6ZaUG/XcipHSb370QSyQws3WgNnEHYTT6BrmjmKEbq7aJst2FN\nMVf1SP0E9Wg90eBylK4VY4aXX0hVDQcvh/oSKpA2YEyj5DoPMhDCNrLDOjLAiWnp\nio/COIzCmYH4npKOcy/thwiIDxc1oPMdJ76PhKB7hL43widoY8911uzmkFjocs/z\niBcFtcqB5A0yMWktWFy6pnT6CUqbu1tGLtjPSszcd5C4SvkkU2QuZfBBSls4syes\nL5Snw6cnArzAOZPwdD3Zbs50R1OeqKsXPSgQm6BiDwKBgQD+iQjzC6H5Kt3TAJT7\n1Ui0U/Grh2EDFnsVFWetCIznlFatnvGDpb5imTeXLxlaNbIsJ1+Vy6/xzHSfoDbC\nF4XMG3nbmLYO7lYwjWJpP/pXTRLOxlU3brrJ1MqpV+qT/AvLXfrJgID5SA6p/B9R\nXQZBnG9v8D9Ba1kFD/sZs0ifRwKBgQDtEaNz6OxXlKI3p+B9ufvZfut1gLhZZygd\nCXvbKuM7eT3yridszuBWV3bA+jzlX8bvaBog6/PNX0VFsSpit7MYO8KCRMX6qW4C\nPBCcK3PwEl917E4EBRSBG4InDXSx6q4Fibu3CSUAiC6VUy/4lKZWnRjCrwi0Wdax\nCDnyyENEawKBgQDNmk+6U2kw5els6YluqughkbUjAyXf4qblbvbiFIgNAAvMDxmz\nZ4yUBIOdAZSPA9S6ScuyKmgi/y4tSrMwsOOI7aDPO1nEUc0ZRvt5XQ6Qulc444gu\nUKldIx/ZDoUTqIK2cWGEpuKHeicuM9PwWx4UnZJ+bygJZfbwWepQ5jsBPQKBgCgb\n49uofNQUgZ2I+aJL5SBXKN8oeAgoO9Et2ZarGublDGYSsPjnx2zUP3NG4yarnazD\noCgRw53rkHzKDLHHArYG5lHD8PsHz2TxZJ4HfAYSFbSN1mU3ttBhJxQUYrn3SpoF\nQ40kLu54G+mHuYYFIF9nFNlbY82fpYMuMR+yYi8hAoGAZXcBT77pfWgNAKLDCsX3\nOkpIn5ilLtpyiNbMhnVbQihfcLXl41056W2z+I4JRcMxKJADjr/Yv+/0viT5KLdv\nqGt6pDX3J1FxlgIInj96yzBwrAX5d0ZudZFcLl1qrARTaE2B5YDsjhe938IGsOD3\nfok0XM8v/kUAbqtuQzs7zXQ=\n-----END PRIVATE KEY-----",
    "client_email": "streamlit-drive-uploader@deep-wares-476608-k6.iam.gserviceaccount.com",
    "client_id": "117217413653054694547",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/streamlit-drive-uploader%40deep-wares-476608-k6.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

# --- KONFIGURASI GOOGLE DRIVE ---
GOOGLE_DRIVE_FOLDER_ID = "19x97EHOzCD5L0hXrBEgLIa-T80IHIxp_"
SCOPES = ['https://www.googleapis.com/auth/drive']

@st.cache_resource
def authenticate_gdrive():
    """Mengotentikasi ke Google Drive menggunakan credentials yang embedded."""
    try:
        credentials = service_account.Credentials.from_service_account_info(
            GOOGLE_DRIVE_CREDENTIALS, 
            scopes=SCOPES
        )
        return build('drive', 'v3', credentials=credentials)
    except Exception as e:
        st.error(f"❌ Gagal otentikasi Google Drive. Error: {e}")
        return None

def upload_file_to_drive(service, file_bytes, filename, folder_id):
    """Mengupload data file (bytes) ke Google Drive."""
    try:
        # Metadata file dengan parents untuk menentukan folder tujuan
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        media = MediaIoBaseUpload(
            io.BytesIO(file_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            resumable=True
        )

        # Upload ke Google Drive
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='webViewLink',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
                          
        return True, f"File berhasil diupload. [Lihat di Drive]({file.get('webViewLink')})"
    except Exception as e:
        return False, f"Gagal mengupload file ke Google Drive. Error: {e}"

# --- FUNGSI KONVERSI XML (Dari converter_ui_13.py) ---
def convert_to_xml(excel_path, output_xml_path, company_name_unused=None):
    """
    Mengkonversi file Excel ke format XML e-Faktur.
    """
    try:
        # Baca sheet yang diperlukan dari file Excel
        df_faktur = pd.read_excel(excel_path, sheet_name='Faktur')
        df_detail = pd.read_excel(excel_path, sheet_name='DetailFaktur')
        
        # Isi semua nilai kosong dengan string kosong
        df_faktur = df_faktur.fillna('')
        df_detail = df_detail.fillna('')

        # Perbaikan pengolahan tanggal
        df_faktur['Tanggal Faktur'] = pd.to_datetime(
            df_faktur['Tanggal Faktur'], 
            format='%d/%m/%Y',
            errors='coerce'
        )
        df_faktur['Tanggal Faktur'] = df_faktur['Tanggal Faktur'].dt.strftime('%Y-%m-%d')

        # Perbaikan kuantitas
        df_detail['Jumlah Barang Jasa'] = pd.to_numeric(
            df_detail['Jumlah Barang Jasa'], 
            errors='coerce'
        ).round(2)

        # Konversi dan format ulang kolom-kolom
        df_faktur['Baris'] = df_faktur['Baris'].astype(str)
        df_faktur['Kode Transaksi'] = df_faktur['Kode Transaksi'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
        df_faktur['NPWP/NIK Pembeli'] = df_faktur['NPWP/NIK Pembeli'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(16)
        df_faktur['Nomor Dokumen Pembeli'] = df_faktur['Nomor Dokumen Pembeli'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(16)
        df_faktur['ID TKU Pembeli'] = df_faktur['ID TKU Pembeli'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(22)
        
        df_detail['Baris'] = df_detail['Baris'].astype(str)
        df_detail['Kode Barang Jasa'] = df_detail['Kode Barang Jasa'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(6)

        # Perbaikan untuk kolom numerik
        numeric_cols_no_decimal = ['Total Diskon', 'DPP Nilai Lain', 'Tarif PPN', 'PPN', 'Tarif PPnBM', 'PPnBM']
        for col in numeric_cols_no_decimal:
            df_detail[col] = df_detail[col].replace('', '0').astype(float).round(0).astype(int).astype(str)

        # Format kolom dengan desimal
        decimal_cols = ['Harga Satuan', 'DPP']
        for col in decimal_cols:
            series = df_detail[col].replace('', '0').astype(float)
            df_detail[col] = series.apply(
                lambda x: str(int(x)) if x == round(x) 
                else '{:.2f}'.format(round(x, 2))
            )

        df_detail['Jumlah Barang Jasa'] = df_detail['Jumlah Barang Jasa'].astype(str).str.replace(r'\.00', '', regex=True)
        
        # Bersihkan data
        df_faktur = df_faktur[df_faktur['Baris'] != 'END']
        df_detail = df_detail[df_detail['Baris'] != 'END']

        # Ambil NPWP penjual dari ID TKU Penjual
        id_tku_penjual_str = str(df_faktur['ID TKU Penjual'].iloc[0]).strip()
        npwp_penjual = id_tku_penjual_str[:16].zfill(16)

        # Buat root element
        root = ET.Element("TaxInvoiceBulk")
        root.set("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

        tin = ET.SubElement(root, "TIN")
        tin.text = npwp_penjual 
        
        list_of_tax_invoice = ET.SubElement(root, "ListOfTaxInvoice")

        # Iterasi setiap baris di Faktur
        for _, row in df_faktur.iterrows():
            tax_invoice = ET.SubElement(list_of_tax_invoice, "TaxInvoice")

            # Format tanggal
            tanggal_faktur = row['Tanggal Faktur']
            if pd.isna(tanggal_faktur):
                formatted_date = ""
            elif isinstance(tanggal_faktur, datetime):
                formatted_date = tanggal_faktur.strftime('%Y-%m-%d')
            else:
                try:
                    date_obj = pd.to_datetime(tanggal_faktur)
                    formatted_date = date_obj.strftime('%Y-%m-%d')
                except:
                    formatted_date = str(tanggal_faktur).split()[0]
                    
            ET.SubElement(tax_invoice, "TaxInvoiceDate").text = formatted_date
            
            ET.SubElement(tax_invoice, "TaxInvoiceOpt").text = row['Jenis Faktur']
            ET.SubElement(tax_invoice, "TrxCode").text = str(row['Kode Transaksi'])
            ET.SubElement(tax_invoice, "AddInfo").text = str(row['Keterangan Tambahan'])
            ET.SubElement(tax_invoice, "CustomDoc").text = str(row['Dokumen Pendukung'])
            ET.SubElement(tax_invoice, "CustomDocMonthYear").text = str(row['Period Dok Pendukung'])
            ET.SubElement(tax_invoice, "RefDesc").text = str(row['Referensi'])
            ET.SubElement(tax_invoice, "FacilityStamp").text = str(row['Cap Fasilitas'])
            ET.SubElement(tax_invoice, "SellerIDTKU").text = str(row['ID TKU Penjual'])
            ET.SubElement(tax_invoice, "BuyerTin").text = str(row['NPWP/NIK Pembeli'])
            ET.SubElement(tax_invoice, "BuyerDocument").text = str(row['Jenis ID Pembeli'])
            ET.SubElement(tax_invoice, "BuyerCountry").text = str(row['Negara Pembeli'])
            ET.SubElement(tax_invoice, "BuyerDocumentNumber").text = str(row['Nomor Dokumen Pembeli'])
            ET.SubElement(tax_invoice, "BuyerName").text = str(row['Nama Pembeli'])
            ET.SubElement(tax_invoice, "BuyerAdress").text = str(row['Alamat Pembeli'])
            ET.SubElement(tax_invoice, "BuyerEmail").text = str(row['Email Pembeli'])
            ET.SubElement(tax_invoice, "BuyerIDTKU").text = str(row['ID TKU Pembeli'])

            # Ambil detail barang/jasa yang sesuai
            list_of_good_service = ET.SubElement(tax_invoice, "ListOfGoodService")
            detail_rows = df_detail[df_detail['Baris'] == row['Baris']]

            for _, detail_row in detail_rows.iterrows():
                good_service = ET.SubElement(list_of_good_service, "GoodService")
                ET.SubElement(good_service, "Opt").text = str(detail_row['Barang/Jasa'])
                ET.SubElement(good_service, "Code").text = str(detail_row['Kode Barang Jasa'])
                ET.SubElement(good_service, "Name").text = str(detail_row['Nama Barang/Jasa'])
                ET.SubElement(good_service, "Unit").text = str(detail_row['Nama Satuan Ukur'])
                ET.SubElement(good_service, "Price").text = str(detail_row['Harga Satuan'])
                ET.SubElement(good_service, "Qty").text = str(detail_row['Jumlah Barang Jasa'])
                ET.SubElement(good_service, "TotalDiscount").text = str(detail_row['Total Diskon'])
                ET.SubElement(good_service, "TaxBase").text = str(detail_row['DPP'])
                ET.SubElement(good_service, "OtherTaxBase").text = str(detail_row['DPP Nilai Lain'])
                ET.SubElement(good_service, "VATRate").text = str(detail_row['Tarif PPN'])
                ET.SubElement(good_service, "VAT").text = str(detail_row['PPN'])
                ET.SubElement(good_service, "STLGRate").text = str(detail_row['Tarif PPnBM'])
                ET.SubElement(good_service, "STLG").text = str(detail_row['PPnBM'])
        
        # Konversi ke XML dan simpan
        xml_bytes = ET.tostring(root, encoding='utf-8')
        with open(output_xml_path, "wb") as f:
            f.write(xml_bytes)
        
        return True, f"Konversi Berhasil! File telah dibuat di:\n{output_xml_path}"
    
    except FileNotFoundError:
        return False, "Error: File tidak ditemukan. Pastikan nama sheet sudah benar (Faktur dan DetailFaktur)."
    except KeyError as e:
        return False, f"Error: Kolom '{e.args[0]}' tidak ditemukan. Periksa nama kolom di Excel."
    except Exception as e:
        return False, f"Terjadi kesalahan: {e}"

# --- FUNGSI KONVERSI EXCEL BYTES KE XML BYTES ---
def convert_excel_bytes_to_xml(excel_bytes, company_name):
    """
    Konversi Excel bytes langsung ke XML bytes tanpa menyimpan file.
    """
    try:
        # Simpan Excel bytes ke file sementara
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(excel_bytes)
            tmp_path = tmp_file.name

        # Konversi ke XML menggunakan fungsi yang ada
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xml') as xml_tmp:
            xml_path = xml_tmp.name

        success, message = convert_to_xml(tmp_path, xml_path, company_name)
        
        if success:
            # Baca XML bytes
            with open(xml_path, 'rb') as f:
                xml_bytes = f.read()
            
            # Hapus file temporary
            os.unlink(tmp_path)
            os.unlink(xml_path)
            
            return True, xml_bytes, "Konversi ke XML berhasil"
        else:
            # Hapus file temporary
            os.unlink(tmp_path)
            if os.path.exists(xml_path):
                os.unlink(xml_path)
            return False, None, message
            
    except Exception as e:
        return False, None, f"Error konversi XML: {str(e)}"

# --- DATA PERUSAHAAN ---
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
            return f"{rounded_x:.2f}"
            
    return series.apply(custom_format)

# --- FUNGSI PEMROSESAN DATA ---
def hitung_nomor_baris_unik_dengan_kode_transaksi(df):
    df['Document Number'] = df['Document Number'].astype(str)
    df['KODE TRANSAKSI'] = df['KODE TRANSAKSI'].astype(str) 

    df['Kombinasi_Unik'] = df['Document Number'] + '-' + df['KODE TRANSAKSI']
    
    unique_combinations = df['Kombinasi_Unik'].unique()
    
    combo_to_row_num = pd.Series(
        data=np.arange(1, len(unique_combinations) + 1), 
        index=unique_combinations
    )
    
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
    
    # Kolom Kunci untuk Logika
    df_output['KODE TRANSAKSI'] = df_header['KODE TRANSAKSI'].astype(str)
    df_output['KETERANGAN TAMBAHAN'] = df_header['KETERANGAN TAMBAHAN']
    df_output['CAP FASILITAS'] = df_header['CAP FASILITAS']
    
    # Logika KODE TRANSAKSI = 4
    kode_transaksi_4_mask = (df_output['KODE TRANSAKSI'] == '4')
    df_output.loc[kode_transaksi_4_mask, 'KETERANGAN TAMBAHAN'] = np.nan 
    df_output.loc[kode_transaksi_4_mask, 'CAP FASILITAS'] = np.nan 
    
    df_output['Document Number'] = df_header['Document Number']
    df_output['Period Dok Pendukung'] = np.nan
    df_output['Referensi'] = df_header['Document Number'].astype(str)
    
    # Menggunakan data perusahaan yang dipilih
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
    
    # Logika berdasarkan UNIT
    condition_units = ['EKOR', 'Pcs', 'Pack', 'Ekor/Pcs/Pack']
    mask_custom_price = df['Unit'].astype(str).str.contains('|'.join(condition_units), case=False, na=False)

    df['Final_Sales_Price_Source'] = df['Sales Price']
    df['Final_Qty_Net_Source'] = df['Qty Net']

    df.loc[mask_custom_price, 'Final_Sales_Price_Source'] = df['Harga EKR']
    df.loc[mask_custom_price, 'Final_Qty_Net_Source'] = df['Qty EKR']

    df['Final_Sales_Price_Num'] = pd.to_numeric(df['Final_Sales_Price_Source'], errors='coerce')
    df['Final_Qty_Net_Num'] = pd.to_numeric(df['Final_Qty_Net_Source'], errors='coerce')
    
    df['NET DISKON'] = pd.to_numeric(df['NET DISKON'], errors='coerce').fillna(0)

    # Kolom non-numerik awal
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

# --- STREAMLIT UTAMA ---
def main():
    st.set_page_config(page_title="Alat Transformasi Data", layout="wide") 
    
    st.title("📄 Transformasi Data Faktur")
    st.markdown("Aplikasi untuk konversi file Excel kustom ke format E-Faktur yang terstruktur (XLSX & XML).")
    
    st.divider() 
    
    # --- 1. Konfigurasi Input ---
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
            
    # --- Tombol Proses ---
    st.markdown("---")
    if is_company_valid and uploaded_file is not None:
        if st.button("🚀 PROSES & SIMPAN FILE", key="proses_button_minimal", type="primary"):
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
                
                with st.spinner('Sedang memproses data...'):
                    
                    df_data_processed = hitung_nomor_baris_unik_dengan_kode_transaksi(df_data.copy())
                    df_faktur = proses_faktur(df_data_processed.copy(), company_info)
                    df_detail = proses_detail_faktur(df_data_processed.copy())
                    
                    timestamp = datetime.now().strftime("%d%m%y_%H%M%S")
                    company_prefix = re.sub(r'[\.\s]', '_', selected_company_name).replace('__', '_')
                    NAMA_FILE_OUTPUT = f"{company_prefix}_{timestamp}_{NAMA_FILE_DASAR}"
                    
                    # Generate Excel file
                    excel_data_bytes = to_excel_bytes(df_faktur, df_detail)

                    # Upload ke Google Drive
                    drive_service = authenticate_gdrive()
                    st.info("Sedang mengupload file ke Google Drive...")

                    if drive_service:
                        success, message = upload_file_to_drive(
                            drive_service,
                            excel_data_bytes,
                            NAMA_FILE_OUTPUT,
                            GOOGLE_DRIVE_FOLDER_ID
                        )
                        
                        with st.container(border=True): 
                            if success:
                                st.balloons()
                                st.success(f"🎉 **File BERHASIL DIPROSES dan diupload ke Google Drive!**")
                                st.markdown(message)
                            else:
                                st.error(f"❌ Proses Gagal Upload ke Google Drive. {message}")

                    # Tombol download Excel
                    st.download_button(
                        label="⬇️ Unduh File E-Faktur XLSX (Lokal)",
                        data=excel_data_bytes,
                        file_name=NAMA_FILE_OUTPUT,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="final_xlsx_download"
                    )

                    # Tombol download XML
                    import tempfile
                    success, xml_bytes, xml_message = convert_excel_bytes_to_xml(excel_data_bytes, selected_company_name)
                    if success:
                        xml_filename = NAMA_FILE_OUTPUT.replace('.xlsx', '.xml')
                        st.download_button(
                            label="📄 Unduh File E-Faktur XML",
                            data=xml_bytes,
                            file_name=xml_filename,
                            mime="application/xml",
                            key="final_xml_download"
                        )
                        st.success("✅ File XML juga tersedia untuk diunduh!")
                    else:
                        st.warning(f"⚠️ Konversi XML: {xml_message}")

            except KeyError as e:
                st.error(f"❌ ERROR: Kolom sumber '**{e.args[0]}**' tidak ditemukan di file data Anda.")
            except Exception as e:
                st.error(f"❌ Terjadi Kesalahan umum saat memproses data: {e}")

if __name__ == "__main__":
    main()