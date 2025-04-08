import os
import pandas as pd
import unicodedata
import difflib
import fitz  # PyMuPDF untuk membaca & memperbaiki PDF
from PyPDF2 import PdfReader

# Path ke file Excel
excel_path = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\KTI PERAWAT\rename.xlsx"

# Path ke folder tempat file PDF berada
folder_path = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\KTI PERAWAT"

# Cek apakah folder ada
if not os.path.exists(folder_path):
    print("⚠️ Folder tidak ditemukan! Periksa kembali path.")
    exit()

# Baca file Excel
df = pd.read_excel(excel_path, engine="openpyxl")

# Ambil daftar file yang ada di folder (hanya PDF)
file_list = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

# Fungsi untuk membersihkan dan menormalisasi nama file
def clean_filename(name):
    return (
        unicodedata.normalize("NFKD", str(name))  # Normalisasi teks untuk menghilangkan karakter tersembunyi
        .strip()
        .replace("$", "")
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("  ", " ")  # Hapus spasi ganda
        .encode("utf-8", "ignore")
        .decode("utf-8")
    )

# Fungsi untuk mengecek apakah file PDF bisa dibaca (tidak corrupt)
def is_pdf_readable(file_path):
    try:
        with open(file_path, "rb") as f:
            reader = PdfReader(f)
            return bool(reader.pages)  # Jika bisa membaca halaman, berarti file OK
    except Exception:
        return False  # Jika error, berarti file bermasalah

# Fungsi untuk memperbaiki file PDF yang corrupt
def fix_corrupt_pdf(pdf_path):
    try:
        new_path = pdf_path.replace(".pdf", "_fixed.pdf")
        with fitz.open(pdf_path) as doc:
            doc.save(new_path)  # Simpan ulang file
        os.replace(new_path, pdf_path)  # Ganti file lama dengan yang diperbaiki
        return True
    except Exception:
        return False  # Jika gagal diperbaiki

# Bersihkan daftar nama file di folder
file_list_cleaned = [clean_filename(f) for f in file_list]

# Debugging daftar nama file setelah pembersihan
print("\n📂 Daftar file di folder (setelah pembersihan):")
for f in file_list_cleaned:
    print(f" - {f}")

# Inisialisasi daftar file yang tidak cocok dan file yang bermasalah
unmatched_files = []
corrupt_files = []

# Loop setiap baris dalam Excel
for index, row in df.iterrows():
    old_name = clean_filename(str(row["nama_lama"]))
    new_name = clean_filename(str(row["nama_baru"]))

    # Pastikan ekstensi PDF
    if not old_name.lower().endswith(".pdf"):
        old_name += ".pdf"
    if not new_name.lower().endswith(".pdf"):
        new_name += ".pdf"

    # Cari file dengan nama yang hampir sama
    best_match = difflib.get_close_matches(old_name, file_list_cleaned, n=1, cutoff=0.7)

    if best_match:
        old_path = os.path.join(folder_path, best_match[0])
        new_path = os.path.join(folder_path, new_name)

        # Cek apakah file PDF bisa dibaca
        if not is_pdf_readable(old_path):
            print(f"❌ File bermasalah (corrupt): {best_match[0]}, mencoba memperbaiki...")
            if fix_corrupt_pdf(old_path):
                print(f"✅ Berhasil memperbaiki: {best_match[0]}")
            else:
                print(f"❌ Gagal memperbaiki: {best_match[0]}")
                corrupt_files.append(best_match[0])
                continue  # Lewati file yang tidak bisa diperbaiki

        try:
            os.rename(old_path, new_path)
            print(f"✅ Berhasil: {best_match[0]} -> {new_name}")
        except Exception as e:
            print(f"❌ Gagal mengganti nama {best_match[0]}: {e}")
    else:
        print(f"⚠️ Tidak ditemukan: {old_name} (Cek karakter tersembunyi atau spasi tambahan!)")
        unmatched_files.append(old_name)  # Tambahkan ke daftar tidak cocok

# **📢 Tampilkan daftar file yang tidak cocok**
if unmatched_files:
    print("\n⚠️ Berikut file yang ada di Excel tetapi tidak ditemukan di folder:")
    for f in unmatched_files:
        print(f" - {f}")
else:
    print("\n✅ Semua file dari Excel berhasil ditemukan!")

# **📢 Tampilkan daftar file yang bermasalah**
if corrupt_files:
    print("\n❌ Berikut file yang rusak/corrupt dan tidak bisa diperbaiki:")
    for f in corrupt_files:
        print(f" - {f}")
else:
    print("\n✅ Semua file PDF bisa dibaca!")

print("\n✅ Proses selesai!")
