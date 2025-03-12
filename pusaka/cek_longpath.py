import os

try:
    with open(r"\\?\C:\Windows\System32\drivers\etc\hosts", "r") as f:
        print("Long path support bekerja!")
except Exception as e:
    print(f"Long path support gagal: {e}")
