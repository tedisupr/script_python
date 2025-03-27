import os
import fitz  # PyMuPDF

# Mencari semua subfolder yang memiliki file PDF.
# Menggabungkan semua file PDF dalam setiap subfolder menjadi satu PDF.
# Menyimpan hasil dengan format: gabungan_NamaSubfolder.pdf
# Melewati file yang kosong dan mencatat file yang gagal dibaca.
# Tidak menimpa file gabungan yang sudah ada.

# Folder utama tempat file PDF berada
folder_pdf = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\KTI PERAWAT"

# Folder utama tempat menyimpan hasil
output_folder = folder_pdf  # Simpan hasil di main folder

# Cari semua subfolder yang memiliki file PDF
subfolders = {}

for root, _, files in os.walk(folder_pdf):
    pdf_files = [os.path.join(root, f) for f in files if f.lower().endswith('.pdf')]
    
    if pdf_files:
        subfolders[root] = sorted(pdf_files)  # Simpan PDF dalam subfolder

# Proses setiap subfolder yang memiliki file PDF
for subfolder, pdf_files in subfolders.items():
    subfolder_name = os.path.basename(subfolder)
    output_filename = f"gabungan_{subfolder_name}.pdf"
    output_file = os.path.join(output_folder, output_filename)

    # Cek apakah file hasil sudah ada
    if os.path.exists(output_file):
        print(f"❌ Gagal! File hasil sudah ada: {output_file}")
        continue  # Lewati subfolder ini

    # Buat objek PDF gabungan
    merged_pdf = fitz.open()
    corrupt_files = []  # Simpan daftar file PDF yang gagal dibaca
    processed_files = []  # Simpan daftar file PDF yang berhasil digabung

    for file_path in pdf_files:
        file_name = os.path.basename(file_path)

        # Cek apakah file kosong
        if os.path.getsize(file_path) == 0:
            print(f"⚠️ File kosong dilewati: {file_name}")
            continue

        # Coba buka PDF, jika error masukkan ke daftar corrupt_files
        try:
            with fitz.open(file_path) as pdf_doc:
                merged_pdf.insert_pdf(pdf_doc)
                processed_files.append(file_name)  # Simpan file yang berhasil diproses
        except Exception as e:
            corrupt_files.append(file_name)
            print(f"❌ Gagal membaca PDF: {file_name} - Error: {e}")

    # Simpan hanya jika ada PDF yang berhasil digabung
    if merged_pdf.page_count > 0:
        merged_pdf.save(output_file)
        print(f"✅ PDF berhasil digabung! Disimpan sebagai: {output_filename}")
    else:
        print(f"❌ Tidak ada file PDF yang bisa digabungkan di {subfolder}")

    merged_pdf.close()

    # Jika ada file PDF yang berhasil digabung, tampilkan daftar
    if processed_files:
        print("\n📂 Daftar file PDF yang berhasil digabung:")
        for f in processed_files:
            print(f"   - {f}")

    # Jika ada file PDF yang corrupt, tampilkan notifikasi
    if corrupt_files:
        print(f"\n🚨 {len(corrupt_files)} file PDF gagal digabungkan di {subfolder}:")
        for f in corrupt_files:
            print(f"   - {f}")

print("\n🎉 Semua proses selesai!")
