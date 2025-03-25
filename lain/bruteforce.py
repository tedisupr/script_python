import zipfile
import os
import itertools
import string
from tkinter import Tk, filedialog

def brute_force_zip(zip_filename, output_folder):
    if not os.path.exists(zip_filename):
        print("[❌] File ZIP tidak ditemukan!")
        return
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    characters = string.ascii_lowercase + string.ascii_uppercase + string.digits  # Huruf besar, kecil, dan angka
    
    try:
        with zipfile.ZipFile(zip_filename) as zf:
            for password_tuple in itertools.product(characters, repeat=8):  # Panjang password tepat 8 karakter
                password = ''.join(password_tuple)
                try:
                    zf.extractall(path=output_folder, pwd=bytes(password, "utf-8"))
                    print(f"\n[✅] Password ditemukan: {password}")
                    print(f"[📂] File diekstrak ke: {output_folder}")
                    return
                except:
                    print(f"[❌] Mencoba: {password}", end="\r")  # Update progress tanpa memenuhi layar
        
        print("\n[❌] Password tidak ditemukan dalam kombinasi 8 karakter!")
    except KeyboardInterrupt:
        print("\n[⚠] Brute force dihentikan oleh pengguna!")
    except Exception as e:
        print(f"\n[⚠] Terjadi kesalahan: {e}")

# Gunakan file dialog untuk memilih file ZIP dan folder tujuan
Tk().withdraw()
zip_filename = filedialog.askopenfilename(title="Pilih File ZIP", filetypes=[("ZIP Files", "*.zip")])
output_folder = filedialog.askdirectory(title="Pilih Folder Penyimpanan Hasil Ekstraksi")

print(f"ZIP File: {zip_filename}")
print(f"Output Folder: {output_folder}")

# Jalankan brute force dengan panjang password 8 karakter
brute_force_zip(zip_filename, output_folder)
