import os
from pypdf import PdfReader
import logging

# Matikan warning/info dari pypdf
logging.getLogger("pypdf").setLevel(logging.CRITICAL)

# Folder tempat file PDF berada
folder_pdf = r"D:\Technical Support\Service\E-Library\IAI Tanjung Pinang\E-BOOK PUSTAKA IAIMU\PAI"

# Ambil semua file PDF dalam folder
pdf_files = [f for f in os.listdir(folder_pdf) if f.endswith('.pdf')]

# Counter hasil pengecekan
ok_count = 0
bad_count = 0

# Cek satu per satu apakah bisa dibuka
for pdf in pdf_files:
    pdf_path = os.path.join(folder_pdf, pdf)
    try:
        reader = PdfReader(pdf_path)  # Coba baca file PDF
        reader.pages[0]  # Coba akses halaman pertama
        print(f"✅ File OK: {pdf}")
        ok_count += 1
    except Exception as e:
        print(f"❌ File rusak: {pdf} | Error: {e}")
        bad_count += 1

# Tampilkan ringkasan
print("\n===== RINGKASAN HASIL CEK PDF =====")
print(f"Total file dicek : {len(pdf_files)}")
print(f"✅ File OK       : {ok_count}")
print(f"❌ File rusak    : {bad_count}")
