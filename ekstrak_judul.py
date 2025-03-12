import os
import fitz  # PyMuPDF untuk membaca PDF
import win32file  # Untuk menangani path panjang di Windows

def extract_title_from_pdf(pdf_path):
    """Ekstrak judul dari halaman pertama PDF dan normalisasi formatnya"""
    doc = fitz.open(pdf_path)
    text = doc[0].get_text("text")
    lines = [line.strip() for line in text.split("\n") if line.strip()]

    # Ambil maksimal 3 baris pertama sebagai judul
    title = " ".join(lines[:4]) if lines else "untitled"

    # Bersihkan karakter yang tidak valid dalam nama file Windows
    title = "".join(c if c.isalnum() or c in " _-()" else "_" for c in title).strip()

    # Normalisasi format (Mengembalikan _ menjadi ())
    title = title.replace("_", " ").replace("  ", " ")

    return title

def rename_pdfs_in_folder(folder_path):
    """Rename semua PDF dalam folder dengan menggunakan Win32 API"""
    folder_path = os.path.abspath(folder_path)
    long_folder_path = f"\\\\?\\{folder_path}"  # Tambahkan prefix \\?\

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            long_pdf_path = f"\\\\?\\{pdf_path}"  # Path panjang untuk Windows API

            # Ambil judul dari halaman pertama PDF
            title = extract_title_from_pdf(long_pdf_path)
            new_filename = f"{title}.pdf" if title else "untitled.pdf"
            new_path = os.path.join(folder_path, new_filename)
            long_new_path = f"\\\\?\\{new_path}"  # Path panjang untuk tujuan rename

            if os.path.exists(long_new_path):
                print(f"⚠ File '{new_filename}' sudah ada. Lewati...")
            else:
                try:
                    # Gunakan MoveFileEx untuk rename dengan path panjang
                    win32file.MoveFileEx(long_pdf_path, long_new_path, win32file.MOVEFILE_REPLACE_EXISTING)
                    print(f"✅ '{filename}' diubah menjadi '{new_filename}'")
                except Exception as e:
                    print(f"❌ Gagal rename '{filename}': {e}")

# Tentukan folder tempat file PDF berada
folder = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\kti farmasi\LTA FARMASI 2024"  # Ganti dengan path folder PDF kamu
rename_pdfs_in_folder(folder)
