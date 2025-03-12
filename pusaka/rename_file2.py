import os

def rename_pdfs_in_folder(folder_path):
    """Rename file PDF dengan aturan:
    1. Hapus 9 karakter pertama dari kiri.
    2. Hapus semua karakter dari '$' ke kanan (termasuk '$').
    """
    folder_path = os.path.abspath(folder_path)
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("⚠ Tidak ada file PDF yang ditemukan!")
        return

    for filename in pdf_files:
        old_path = os.path.join(folder_path, filename)

        # Hapus 9 karakter pertama
        new_name = filename[9:] if len(filename) > 9 else filename
        
        # Hapus semua dari karakter '$' ke kanan
        new_name = new_name.split("$")[0].strip() + ".pdf"

        new_path = os.path.join(folder_path, new_name)

        # Cek jika nama baru sama atau sudah ada
        if old_path == new_path or os.path.exists(new_path):
            print(f"⚠ Lewati: '{filename}' (nama sudah ada)")
            continue

        try:
            os.rename(old_path, new_path)
            print(f"✅ Rename: '{filename}' -> '{new_name}'")
        except Exception as e:
            print(f"❌ Gagal rename '{filename}': {e}")

# Tentukan folder tempat file PDF berada
folder = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\kti farmasi\KTI D III FARMASI TAHUN 2023\Gabungan"  # Ganti dengan path folder
rename_pdfs_in_folder(folder)
