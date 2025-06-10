import pandas as pd
import os
import glob

# --- Konfigurasi ---
# Masukkan path ke folder yang berisi file-file Excel Anda.
# Jika skrip ini ada di dalam folder yang sama dengan file Excel, cukup tulis '.'
folder_path = 'C:/Users/mk.arisman.pili/Downloads/DAFTAR PANGKALAN SAM MEDAN/SAM MEDAN' 

# Nama untuk file output setelah digabungkan
nama_file_output = 'Gabungan_Data_Pangkalan_Final.xlsx'

# Daftar field/kolom yang diinginkan sesuai urutan
kolom_yang_diinginkan = [
    'Sold To', 'Nama Agen', 'Wilayah', 'Tipe Pangkalan', 'ID Registrasi',
    'Nama Pangkalan', 'Email Pangkalan', 'No HP Pemilik', 'No KTP Pemilik',
    'Nama Pemilik', 'Nomor NIB', 'Qty Kontrak', 'Catatan', 'Tipe Pembayaran',
    'No Rekening', 'Nama Bank', 'Nama Akun Bank', 'Tanggal Diapprove SAM',
    'Provinsi', 'Kota', 'Kecamatan', 'Kelurahan', 'RT', 'RW', 'Kode Pos',
    'Alamat', 'Integrasi MAP', 'MID', 'Latitude', 'Longitude', 'Status',
    'Pembaharuan'
]
# --- Akhir Konfigurasi ---

def gabungkan_excel(path, file_output_name, columns):
    """
    Menggabungkan semua file Excel, mengurutkan, dan menyimpannya satu level
    folder di atas dari folder sumber.
    """
    # Mencari semua file dengan ekstensi .xlsx dan .xls di dalam folder
    search_path = os.path.join(path, "*.xlsx")
    file_list = glob.glob(search_path) + glob.glob(os.path.join(path, "*.xls"))

    if not file_list:
        print(f"⚠️ Tidak ada file Excel yang ditemukan di dalam folder: {os.path.abspath(path)}")
        return

    # List untuk menampung semua dataframe
    list_df = []

    print(f"Membaca {len(file_list)} file Excel dari '{os.path.abspath(path)}'...")
    for f in file_list:
        try:
            df = pd.read_excel(f, engine='openpyxl')
            list_df.append(df)
            print(f"  - Berhasil membaca file: {os.path.basename(f)}")
        except Exception as e:
            print(f"  - ❗️ Gagal membaca file: {os.path.basename(f)}. Error: {e}")

    if not list_df:
        print("Gagal memproses file Excel. Pastikan file tidak rusak.")
        return

    # Menggabungkan semua dataframe menjadi satu
    df_gabungan = pd.concat(list_df, ignore_index=True)

    # Memastikan semua kolom yang diinginkan ada
    for col in columns:
        if col not in df_gabungan.columns:
            df_gabungan[col] = None
    
    # Mengatur urutan kolom
    df_gabungan = df_gabungan[columns]

    # Mengurutkan data berdasarkan 'Wilayah', lalu 'Nama Agen'
    print("\nMengurutkan data berdasarkan 'Wilayah' dan 'Nama Agen'...")
    df_gabungan_sorted = df_gabungan.sort_values(by=['Wilayah', 'Nama Agen'])

    # --- PERUBAHAN LOKASI PENYIMPANAN ---
    # Menentukan path untuk menyimpan file output: satu folder di atas folder sumber
    try:
        # 1. Dapatkan path absolut dari folder sumber
        sumber_abs_path = os.path.abspath(path)
        # 2. Dapatkan path parent directory (satu level di atas)
        parent_dir = os.path.dirname(sumber_abs_path)
        # 3. Gabungkan path parent dengan nama file output
        final_output_path = os.path.join(parent_dir, file_output_name)

        # Menyimpan dataframe ke path yang telah ditentukan
        df_gabungan_sorted.to_excel(final_output_path, index=False)
        
        print("\n✅ Sukses! Data telah digabungkan dan diurutkan.")
        print(f"File disimpan di folder satu level di atas pada lokasi:")
        print(f"-> {final_output_path}")

    except Exception as e:
        print(f"\n❌ Gagal menyimpan file output. Pastikan Anda memiliki izin untuk menulis ke folder tujuan.")
        print(f"Error: {e}")

# Menjalankan fungsi utama
if __name__ == "__main__":
    gabungkan_excel(folder_path, nama_file_output, kolom_yang_diinginkan)