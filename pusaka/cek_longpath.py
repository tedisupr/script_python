import os

# untuk cek long path di aplikasi

# Path file yang mau dicek
path = r"\\?\C:\Windows\System32\drivers\etc\hosts"

try:
    with open(path, "r") as f:
        print("✅ Long path support bekerja!")
except Exception as e:
    print(f"❌ Long path support gagal: {e}")

# Tampilkan panjang path
print("\n===== INFO PATH =====")
print(f"Path: {path}")
print(f"Jumlah karakter path: {len(path)}")
