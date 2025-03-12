import os
import pandas as pd

# Path ke file Excel
excel_path = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\kti farmasi\KTI D III FARMASI TAHUN 2023\Gabungan\rename.xlsx"

# Path ke folder tempat file PDF berada
folder_path = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\kti farmasi\KTI D III FARMASI TAHUN 2023\Gabungan"

# Cek apakah folder ada
if not os.path.exists(folder_path):
    print("⚠️ Folder tidak ditemukan! Periksa kembali path.")
    exit()

# Baca file Excel
df = pd.read_excel(excel_path, engine="openpyxl")

# Ambil daftar file yang ada di folder
file_list = os.listdir(folder_path)
print("📂 Daftar file di folder:", file_list)

# Loop setiap baris dalam Excel
for index, row in df.iterrows():
    # Ambil nama lama dan baru, lalu hapus spasi tambahan
    old_name = str(row['nama_lama']).strip()
    new_name = str(row['nama_baru']).strip()

    # Pastikan ekstensi PDF
    if not old_name.endswith(".pdf"):
        old_name += ".pdf"
    if not new_name.endswith(".pdf"):
        new_name += ".pdf"

    old_path = os.path.join(folder_path, old_name)
    new_path = os.path.join(folder_path, new_name)

    # Cek apakah file ada sebelum rename
    if os.path.exists(old_path):
        os.rename(old_path, new_path)
        print(f"✅ Berhasil: {old_name} -> {new_name}")
    else:
        print(f"⚠️ File tidak ditemukan: {old_name}")

print("✅ Selesai mengganti nama file!")
