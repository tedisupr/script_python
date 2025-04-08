import os

def rename_pdfs_in_folder(folder_path, keyword):
    """Rename file PDF dengan aturan:
    1. Hapus kata kunci tertentu dalam nama file.
    """
    folder_path = os.path.abspath(folder_path)
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("⚠ Tidak ada file PDF yang ditemukan!")
        return

    for filename in pdf_files:
        old_path = os.path.join(folder_path, filename)
        
        # Awali dengan nama file asli
        new_name = filename.replace(keyword, "").strip()

        # Tambahkan kembali ekstensi .pdf
        new_name = new_name + ".pdf"
        new_path = os.path.join(folder_path, new_name)

        # Cek jika nama baru sama atau sudah ada
        if old_path == new_path or os.path.exists(new_path):
            print(f"⚠ Lewati: '{filename}' (nama sudah ada atau tidak berubah)")
            continue

        try:
            os.rename(old_path, new_path)
            print(f"✅ Rename: '{filename}' -> '{new_name}'")
        except Exception as e:
            print(f"❌ Gagal rename '{filename}': {e}")

# Tentukan folder tempat file PDF berada dan kata kunci yang ingin dihapus
folder = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\KTI PERAWAT"  # Ganti dengan path folder
keyword = ".pdf"  # Ganti dengan kata kunci yang ingin dihapus
rename_pdfs_in_folder(folder, keyword)
