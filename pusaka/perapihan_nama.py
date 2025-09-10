import os
import re

# Merapihkan nama file pdf
# Membersihkan nama file PDF agar tidak ada karakter aneh, lalu menulis ulang dengan format Title Case.

# Ganti ini dengan folder tempat file PDF kamu berada
folder_path = "D:\Technical Support\Service\E-Library\Poltekes TNI AU\ebook keperawatan\ebook perawat-20250423T153045Z-001 23-04-2025"

def format_nama_file(nama_file):
    # Hilangkan ekstensi dulu
    nama, ekstensi = os.path.splitext(nama_file)

    # Hilangkan semua tanda baca kecuali spasi
    nama = re.sub(r"[^\w\s]", " ", nama)

    # Ganti underscore dan dash dengan spasi
    nama = nama.replace("_", " ").replace("-", " ")

    # Ubah ke format kapital tiap kata
    nama = " ".join(word.capitalize() for word in nama.split())

    # Gabungkan lagi dengan ekstensi
    return f"{nama}{ekstensi}"

# Proses semua file di folder
for nama_file in os.listdir(folder_path):
    if nama_file.lower().endswith(".pdf"):
        nama_baru = format_nama_file(nama_file)
        lama = os.path.join(folder_path, nama_file)
        baru = os.path.join(folder_path, nama_baru)
        
        os.rename(lama, baru)
        print(f"Renamed: {nama_file} -> {nama_baru}")
