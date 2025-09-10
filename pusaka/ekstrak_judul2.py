import os
import fitz  # PyMuPDF untuk membaca PDF
import win32file  # Untuk menangani path panjang di Windows
import re  # Untuk membersihkan karakter tidak valid
from datetime import datetime

# Fungsi script ini
# Membersihkan & menstandarkan nama file PDF di folder,
# Menambahkan judul dari isi PDF agar nama lebih deskriptif,
# Menghindari masalah path panjang, nama duplikat, atau karakter ilegal,
# Otomatis menghapus prefix tertentu (“gabungan_Lab. ... - ”),
# Menyimpan ringkasan hasil (jumlah berhasil, gagal, daftar gagal) ke layar & log file.

# Karakter yang tidak diperbolehkan dalam nama file Windows
INVALID_CHARS = r'[<>:"/\\|?*]'

def clean_filename(filename):
    """Bersihkan karakter tidak valid dalam nama file dan hilangkan spasi/titik di akhir."""
    cleaned = re.sub(INVALID_CHARS, " ", filename)  # Hapus karakter ilegal
    cleaned = " ".join(cleaned.split()).strip()  # Hapus spasi ganda
    return cleaned.rstrip(" .")  # Pastikan tidak diakhiri spasi atau titik

def extract_title_from_pdf(pdf_path):
    """Ekstrak judul dari halaman pertama PDF"""
    try:
        with fitz.open(pdf_path) as doc:
            text = doc[0].get_text("text")
    except Exception:
        return "untitled"

    lines = [line.strip() for line in text.split("\n") if line.strip()]
    title = " ".join(lines[:4]) if lines else "untitled"

    return clean_filename(title)  # Bersihkan nama judul

def rename_pdfs_in_folder(folder_path):
    """Rename semua file PDF dalam folder"""
    folder_path = os.path.abspath(folder_path)
    long_folder_path = f"\\\\?\\{folder_path}"  # Gunakan path panjang

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

    total_files = len(pdf_files)
    if not pdf_files:
        print("⚠ Tidak ada file PDF yang ditemukan!")
        return  # Tidak ada file, keluar

    success_count = 0
    fail_count = 0
    failed_files = []

    log_lines = []
    log_lines.append(f"=== LOG RENAME PDF ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===\n")
    log_lines.append(f"Total file PDF ditemukan: {total_files}\n")

    for filename in pdf_files:
        pdf_path = os.path.join(folder_path, filename)
        long_pdf_path = f"\\\\?\\{pdf_path}"

        # Ambil judul dari PDF
        title = extract_title_from_pdf(long_pdf_path)

        # Bersihkan nama file lama
        base_filename, _ = os.path.splitext(filename)
        base_filename = clean_filename(base_filename)

        # Hapus prefix tertentu
        if base_filename.startswith("gabungan_Lab. Kimia - "):
            base_filename = base_filename.replace("gabungan_Lab. Kimia - ", "").strip()
        if base_filename.startswith("gabungan_Lab. TekFar - "):
            base_filename = base_filename.replace("gabungan_Lab. TekFar - ", "").strip()
        if base_filename.startswith("gabungan_Lab. Mikrobiologi - "):
            base_filename = base_filename.replace("gabungan_Lab. Mikrobiologi - ", "").strip()

        # Format nama baru dengan tambahan judul dari PDF
        new_filename = f"{base_filename} # {title}.pdf"

        # Jika nama file terlalu panjang, persingkat
        if len(new_filename) > 245:
            new_filename = new_filename[:245] + ".pdf"

        new_path = os.path.join(folder_path, new_filename)
        long_new_path = f"\\\\?\\{new_path}"

        # Coba rename, jika gagal buat nama alternatif
        try:
            win32file.MoveFileEx(long_pdf_path, long_new_path, win32file.MOVEFILE_REPLACE_EXISTING)
            msg = f"✅ '{filename}' berhasil diubah menjadi '{new_filename}'"
            print(msg)
            log_lines.append(msg)
            success_count += 1
        except:
            # Jika masih gagal, coba nama lebih pendek
            short_filename = f"{base_filename}.pdf"
            short_new_path = os.path.join(folder_path, short_filename)
            long_short_new_path = f"\\\\?\\{short_new_path}"
            try:
                win32file.MoveFileEx(long_pdf_path, long_short_new_path, win32file.MOVEFILE_REPLACE_EXISTING)
                msg = f"⚠️ '{filename}' hanya berhasil diubah menjadi '{short_filename}' (nama dipersingkat)"
                print(msg)
                log_lines.append(msg)
                fail_count += 1
                failed_files.append(filename)
            except:
                msg = f"❌ Gagal mengubah nama '{filename}', file tetap seperti semula."
                print(msg)
                log_lines.append(msg)
                fail_count += 1
                failed_files.append(filename)

    # === Resume ===
    resume = [
        "\n=== RESUME ===",
        f"Total file PDF ditemukan: {total_files}",
        f"Jumlah berhasil ✅: {success_count}",
        f"Jumlah gagal ❌: {fail_count}"
    ]
    if failed_files:
        resume.append("Daftar gagal:")
        resume.extend([f" - {f}" for f in failed_files])

    # Tampilkan resume ke layar
    print("\n".join(resume))

    # Simpan log ke file
    log_lines.append("\n" + "\n".join(resume))
    log_file = os.path.join(folder_path, "rename_log.txt")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write("\n".join(log_lines) + "\n\n")

    print(f"\n📝 Log hasil sudah disimpan di: {log_file}")

# Jalankan script
folder = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\KTI PERAWAT"
rename_pdfs_in_folder(folder)