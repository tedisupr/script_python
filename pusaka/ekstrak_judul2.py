import os
import fitz  # PyMuPDF untuk membaca PDF
import win32file  # Untuk menangani path panjang di Windows
import re  # Untuk membersihkan karakter tidak valid

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

    if not pdf_files:
        return  # Tidak ada file, keluar

    for filename in pdf_files:
        pdf_path = os.path.join(folder_path, filename)
        long_pdf_path = f"\\\\?\\{pdf_path}"

        # Ambil judul dari PDF
        title = extract_title_from_pdf(long_pdf_path)

        # Gabungkan nama file lama dengan judul PDF
        base_filename, _ = os.path.splitext(filename)
        base_filename = clean_filename(base_filename)
        new_filename = f"{base_filename} $ {title}.pdf"

        # Jika nama file terlalu panjang, persingkat
        if len(new_filename) > 245:
            new_filename = new_filename[:245] + ".pdf"

        new_path = os.path.join(folder_path, new_filename)
        long_new_path = f"\\\\?\\{new_path}"

        # Coba rename, jika gagal buat nama alternatif
        try:
            win32file.MoveFileEx(long_pdf_path, long_new_path, win32file.MOVEFILE_REPLACE_EXISTING)
            print(f"✅ '{filename}' berhasil diubah menjadi '{new_filename}'")
        except:
            # Jika masih gagal, coba nama lebih pendek
            short_filename = f"{base_filename}.pdf"
            short_new_path = os.path.join(folder_path, short_filename)
            long_short_new_path = f"\\\\?\\{short_new_path}"
            try:
                win32file.MoveFileEx(long_pdf_path, long_short_new_path, win32file.MOVEFILE_REPLACE_EXISTING)
                print(f"✅ '{filename}' berhasil diubah menjadi '{short_filename}' (nama dipersingkat)")
            except:
                pass  # Jika masih gagal, abaikan tanpa menampilkan error

# Jalankan script
folder = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\kti farmasi\FARMASI JURNAL BARU 2021\Softfile Tk.3 2021\gabungan"
rename_pdfs_in_folder(folder)