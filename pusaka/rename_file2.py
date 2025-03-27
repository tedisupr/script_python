import os

def rename_pdfs_in_folder(folder_path):
    """Rename file PDF dengan aturan:
    1. Hapus semua karakter dari '$' ke kiri (termasuk '$').
    """
    folder_path = os.path.abspath(folder_path)
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("⚠ Tidak ada file PDF yang ditemukan!")
        return

    for filename in pdf_files:
        old_path = os.path.join(folder_path, filename)
        
        # Awali dengan nama file asli
        new_name = filename  

        # Hapus semua karakter dari '$' ke kiri (termasuk '$')
        if "$" in new_name:
            new_name = new_name.split("$")[-1].strip()

        # Tambahkan kembali ekstensi .pdf
        new_name = new_name.strip() + ".pdf"
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

# Tentukan folder tempat file PDF berada
folder = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\kti farmasi\FARMASI JURNAL BARU 2021\Softfile Tk.3 2021\gabungan"  # Ganti dengan path folder
rename_pdfs_in_folder(folder)
