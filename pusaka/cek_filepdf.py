import os
from pypdf import PdfReader

# Folder tempat file PDF berada
folder_pdf = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\jurnal bidan\INTERNASIONAL JOURNAL OF COMUNITY BASED NURSING AND MIDWIFERY SHIRAZ\UNIVERSITY OF MEDICAL IRAN VOL 10 (2022)"

# Ambil semua file PDF dalam folder
pdf_files = [f for f in os.listdir(folder_pdf) if f.endswith('.pdf')]

# Cek satu per satu apakah bisa dibuka
for pdf in pdf_files:
    pdf_path = os.path.join(folder_pdf, pdf)
    try:
        reader = PdfReader(pdf_path)  # Coba baca file PDF
        reader.pages[0]  # Coba akses halaman pertama
        print(f"✅ File OK: {pdf}")
    except Exception as e:
        print(f"❌ File rusak: {pdf} | Error: {e}");
