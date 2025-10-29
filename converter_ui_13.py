import datetime
import pandas as pd
import xml.etree.ElementTree as ET
# from xml.dom import minidom # <--- Modul ini sudah TIDAK DIPERLUKAN lagi
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Fungsi Konversi (kode yang sudah diperbaiki) ---
def convert_to_xml(excel_path, output_xml_path, company_name_unused=None):
    try:
        # Baca sheet yang diperlukan dari file Excel
        df_faktur = pd.read_excel(excel_path, sheet_name='Faktur')
        df_detail = pd.read_excel(excel_path, sheet_name='DetailFaktur')
        
        # === PERBAIKAN FORMAT DATA SECARA KESELURUHAN ===
        # Isi semua nilai kosong dengan string kosong
        df_faktur = df_faktur.fillna('')
        df_detail = df_detail.fillna('')

        # ==========================================================
        # 📌 MASUKKAN KODE PERBAIKAN TANGGAL DI SINI
        # ==========================================================

        # --- PERBAIKAN PENGOLAHAN TANGGAL: MENGATASI DD/MM/YYYY DIANGGAP MM/DD/YYYY ---

        # 1. Konversi kolom 'Tanggal Faktur' dari format string yang ada (asumsi DD/MM/YYYY)
        #    menjadi objek datetime.
        df_faktur['Tanggal Faktur'] = pd.to_datetime(
            df_faktur['Tanggal Faktur'], 
            format='%d/%m/%Y', # Secara eksplisit menentukan format input Hari/Bulan/Tahun
            errors='coerce'    # Menangani nilai yang tidak valid (jika ada)
        )

        # 2. Format ulang objek datetime tersebut menjadi string YYYY-MM-DD, 
        #    sesuai standar yang dibutuhkan untuk XML/e-Faktur.
        df_faktur['Tanggal Faktur'] = df_faktur['Tanggal Faktur'].dt.strftime('%Y-%m-%d')

        # --- AKHIR PERBAIKAN PENGOLAHAN TANGGAL ---

        # --- PERBAIKAN KUANTITAS (QTY) menjadi maksimal 2 desimal ---
        # 1. Konversi ke numerik (float)
        df_detail['Jumlah Barang Jasa'] = pd.to_numeric(
            df_detail['Jumlah Barang Jasa'], 
            errors='coerce'
        )
        # 2. BULATKAN ke 2 angka di belakang koma.
        df_detail['Jumlah Barang Jasa'] = df_detail['Jumlah Barang Jasa'].round(2)
        
        # 3. Opsional, ubah ke string dengan format 2 desimal tetap (jika diperlukan)
        # Jika Anda ingin memastikan output selalu 2 desimal (cth: 990.20)
        # df_detail['Jumlah Barang Jasa'] = df_detail['Jumlah Barang Jasa'].apply(lambda x: '{:.2f}'.format(x) if pd.notna(x) else '')
        
        # Untuk kasus Anda, pembulatan .round(2) saja sudah cukup mengatasi masalah floating-point 
        # dan menghasilkan 990.2 saat konversi ke string XML.
        # --- AKHIR PERBAIKAN KUANTITAS ---

        # Konversi dan format ulang kolom-kolom di sheet 'Faktur'
        df_faktur['Baris'] = df_faktur['Baris'].astype(str)
        # Menghilangkan '.0' dan memastikan 2 digit dengan leading zero
        df_faktur['Kode Transaksi'] = df_faktur['Kode Transaksi'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
        # Menghilangkan '.0' dan memastikan 16 digit dengan leading zero
        df_faktur['NPWP/NIK Pembeli'] = df_faktur['NPWP/NIK Pembeli'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(16)
        # Menghilangkan '.0' dan memastikan 16 digit dengan leading zero
        df_faktur['Nomor Dokumen Pembeli'] = df_faktur['Nomor Dokumen Pembeli'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(16)
        # Menghilangkan '.0' dan memastikan 22 digit dengan leading zero
        df_faktur['ID TKU Pembeli'] = df_faktur['ID TKU Pembeli'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(22)
        
        # Konversi dan format ulang kolom-kolom di sheet 'DetailFaktur'
        df_detail['Baris'] = df_detail['Baris'].astype(str)
        # Menghilangkan '.0' dan memastikan 6 digit dengan leading zero
        df_detail['Kode Barang Jasa'] = df_detail['Kode Barang Jasa'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(6)

        # Perbaikan untuk kolom numerik: bulatkan tanpa koma, atasi nilai kosong, dan ubah ke string
        numeric_cols_no_decimal = ['Total Diskon', 'DPP Nilai Lain', 'Tarif PPN', 'PPN', 'Tarif PPnBM', 'PPnBM']
        for col in numeric_cols_no_decimal:
            # Mengganti '' dengan '0', konversi ke float, bulatkan, konversi ke integer, lalu ke string
            df_detail[col] = df_detail[col].replace('', '0').astype(float).round(0).astype(int).astype(str)
        # --- FORMAT KOLOM DENGAN DESIMAL (2 Angka di belakang koma) ---

        # Kolom yang MEMERLUKAN desimal: Harga Satuan dan DPP
        decimal_cols = ['Harga Satuan', 'DPP']
        for col in decimal_cols:
            # 1. Mengubah ke float, mengganti NaN/kosong dengan 0.00
            series = df_detail[col].replace('', '0').astype(float)

            # 2. Menerapkan format fleksibel: Jika bilangan bulat -> tampilkan tanpa koma; jika desimal -> tampilkan 2 angka koma
            df_detail[col] = series.apply(
                lambda x: str(int(x)) if x == round(x) 
                else '{:.2f}'.format(round(x, 2))
            )

        # Perbaikan khusus untuk 'Jumlah Barang Jasa' 
        # (Asumsi: Qty mungkin memiliki desimal. Kode asli menggunakan str.replace. 
        # Jika Qty harus 2 desimal, gunakan format(..., '.2f'))
        # Jika tujuannya hanya membersihkan string '.00' saja:
        df_detail['Jumlah Barang Jasa'] = df_detail['Jumlah Barang Jasa'].astype(str).str.replace(r'\.00', '', regex=True)
        
        # Bersihkan data
        df_faktur = df_faktur[df_faktur['Baris'] != 'END']
        df_detail = df_detail[df_detail['Baris'] != 'END']

        # --- AMBIL NPWP PENJUAL DARI ID TKU PENJUAL ---
        # 1. Ambil nilai 'ID TKU Penjual' dari baris pertama (asumsi ini seragam)
        # 2. Pastikan nilainya string dan ambil 16 karakter pertamanya
        # 3. Handle kasus nilai kosong/NaN dengan fallback ke string kosong (atau default NPWP)
        
        id_tku_penjual_str = str(df_faktur['ID TKU Penjual'].iloc[0]).strip()
        
        # Ambil 16 digit pertama (NPWP)
        npwp_penjual = id_tku_penjual_str[:16].zfill(16)
        
        # Jika nilai 'ID TKU Penjual' di Excel kosong, 'npwp_penjual' akan menjadi "0000000000000000"
        # Kita bisa menambahkan validasi di sini jika diperlukan.
        # -------------------------------------------------

        # Buat root element
        root = ET.Element("TaxInvoiceBulk")
        root.set("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

        # GANTI BAGIAN INI:
        tin = ET.SubElement(root, "TIN")
        # GUNAKAN NILAI DINAMIS DARI EXCEL
        tin.text = npwp_penjual 
        
        list_of_tax_invoice = ET.SubElement(root, "ListOfTaxInvoice")

        # Iterasi setiap baris di Faktur
        for _, row in df_faktur.iterrows():
            tax_invoice = ET.SubElement(list_of_tax_invoice, "TaxInvoice")

            # --- Pemformatan Tanggal ---
            tanggal_faktur = row['Tanggal Faktur']
            if pd.isna(tanggal_faktur):
                formatted_date = ""
            elif isinstance(tanggal_faktur, datetime.datetime):
                formatted_date = tanggal_faktur.strftime('%Y-%m-%d')
            else:
                try:
                    date_obj = pd.to_datetime(tanggal_faktur)
                    formatted_date = date_obj.strftime('%Y-%m-%d')
                except:
                    formatted_date = str(tanggal_faktur).split()[0]
                    
            ET.SubElement(tax_invoice, "TaxInvoiceDate").text = formatted_date
            # ---------------------------
            
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
        
        # --------------------------------------------------------------------
        # KODE PERBAIKAN UTAMA UNTUK MENGHILANGKAN WHITESPACE DAN MEMPERBAIKI INDENTASI
        # --------------------------------------------------------------------
        
        # 1. Konversi root element ke string (bytes) yang padat (minified)
        #    Ini menggantikan xml_str = ET.tostring(root, encoding='utf-8') 
        #    dan menghilangkan fungsi minidom.parseString() dan toprettyxml()
        xml_bytes = ET.tostring(root, encoding='utf-8')

        # 2. Simpan ke file dengan memastikan f.write() ada di dalam blok 'with'
        with open(output_xml_path, "wb") as f:
            # Tulis bytes XML yang padat ke file
            f.write(xml_bytes) 
        
        # --------------------------------------------------------------------
        
        return True, f"Konversi Berhasil! File telah dibuat di:\n{output_xml_path}"
    
    except FileNotFoundError:
        return False, "Error: File tidak ditemukan. Pastikan nama sheet sudah benar (Faktur dan DetailFaktur)."
    except KeyError as e:
        return False, f"Error: Kolom '{e.args[0]}' tidak ditemukan. Periksa nama kolom di Excel."
    except Exception as e:
        return False, f"Terjadi kesalahan: {e}"

# --- Fungsi untuk UI ---
class ExcelToXMLApp:
    def __init__(self, master):
        self.master = master
        master.title("Konversi Excel ke XML")
        master.geometry("450x180")
        master.resizable(False, False)

        self.file_path = ""
        
        # Frame untuk menampung widget
        self.main_frame = tk.Frame(master, padx=15, pady=15)
        self.main_frame.pack(fill="both", expand=True)

        # Label untuk menampilkan jalur file
        self.label_file = tk.Label(self.main_frame, text="Pilih file Excel Anda:", font=("Arial", 10))
        self.label_file.pack(pady=(0, 5))
        
        self.file_entry = tk.Entry(self.main_frame, width=50)
        self.file_entry.pack(side="left", padx=(0, 5), fill="x", expand=True)

        self.button_browse = tk.Button(self.main_frame, text="Telusuri...", command=self.browse_file)
        self.button_browse.pack(side="right")

        # Tombol konversi
        self.convert_button = tk.Button(master, text="Konversi ke XML", command=self.run_conversion, font=("Arial", 12), width=20, height=2)
        self.convert_button.pack(pady=10)

        # Label status
        self.status_label = tk.Label(master, text="", fg="black", font=("Arial", 10))
        self.status_label.pack()

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, self.file_path)
        self.status_label.config(text="")

    def run_conversion(self):
        if not self.file_path:
            self.status_label.config(text="Pilih file Excel terlebih dahulu!", fg="red")
            return

        # Buat nama file berdasarkan timestamp
        timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
        initial_filename = f"Output_Converter_{timestamp}.xml"

        output_path = filedialog.asksaveasfilename(
            defaultextension=".xml",
            filetypes=[("XML files", "*.xml")],
            initialfile=initial_filename
        )
                
        # Jika pengguna menekan 'Cancel', hentikan proses
        if not output_path:
            self.status_label.config(text="Operasi dibatalkan.", fg="red")
            return

        self.status_label.config(text="Proses konversi...", fg="blue")
        self.master.update_idletasks() # Memperbarui UI

        success, message = convert_to_xml(self.file_path, output_path)
        
        if success:
            self.status_label.config(text=message, fg="green")
            messagebox.showinfo("Konversi Berhasil", message)
        else:
            self.status_label.config(text=message, fg="red")
            messagebox.showerror("Konversi Gagal", message)

# --- Jalankan Aplikasi ---
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToXMLApp(root)
    root.mainloop()