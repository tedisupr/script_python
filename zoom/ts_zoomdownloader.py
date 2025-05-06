import os
import requests
import json
from datetime import datetime

# Gantilah dengan nilai yang sesuai dari Zoom App kamu
CLIENT_ID = 'IfIQY7GiQqep5EFIVUroWw'
CLIENT_SECRET = 'fvUEqUXaOFNboimdB2lES8PBNNoH7bWM'
REDIRECT_URI = 'YOUR_REDIRECT_URI'
ACCESS_TOKEN = 'YOUR_ACCESS_TOKEN'

# Fungsi untuk mendapatkan Access Token (jika kamu belum punya)
def get_access_token():
    url = 'https://zoom.us/oauth/token'
    headers = {
        'Authorization': f'Basic {CLIENT_ID}:{CLIENT_SECRET}',
    }
    data = {
        'grant_type': 'authorization_code',
        'code': 'AUTHORIZATION_CODE_FROM_ZOOM',  # Ganti dengan authorization code
        'redirect_uri': REDIRECT_URI,
    }

    response = requests.post(url, data=data, headers=headers)
    if response.status_code == 200:
        access_token = response.json()['access_token']
        print("Access token berhasil didapatkan.")
        return access_token
    else:
        print(f"Error saat mendapatkan access token: {response.status_code}")
        return None

# Fungsi untuk mengambil daftar user
def get_user_list(access_token):
    url = 'https://api.zoom.us/v2/users'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    users = []
    while url:
        response = requests.get(url, headers=headers)
        data = response.json()

        if 'users' in data:
            users.extend(data['users'])
        else:
            print("Error mengambil daftar user:", data)

        # Pagination (jika ada lebih banyak user)
        url = data.get('next_page_token')
        if url:
            url = f"{url}&page_token={url}"

    return users

# Fungsi untuk mengambil rekaman dari user tertentu
def get_recordings(user_email, access_token):
    url = f'https://api.zoom.us/v2/users/{user_email}/recordings'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    from_date = "2023-01-01"  # Sesuaikan dengan tanggal yang diinginkan
    to_date = datetime.today().strftime('%Y-%m-%d')
    params = {
        'from': from_date,
        'to': to_date
    }

    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        recordings = response.json().get('meetings', [])
        return recordings
    else:
        print(f"Error saat mengambil rekaman untuk {user_email}: {response.status_code}")
        return []

# Fungsi untuk mendownload rekaman
def download_recordings(meetings, access_token, user_email):
    folder = f"zoom_recordings/{user_email}"
    os.makedirs(folder, exist_ok=True)

    for meeting in meetings:
        topic = meeting['topic'].replace("/", "_").replace("\\", "_")
        print(f"🔍 Memproses rekaman untuk: {topic}")
        for file in meeting['recording_files']:
            file_type = file["file_type"].lower()
            file_url = file['download_url'] + f"?access_token={access_token}"
            filename = f"{topic}_{file['id']}.{file_type}"
            filepath = os.path.join(folder, filename)

            print(f"📥 URL Rekaman: {file_url}")
            print(f"📂 Menyimpan ke: {filepath}")
            
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            }

            try:
                r = requests.get(file_url, headers=headers)
                if r.status_code == 200:
                    with open(filepath, "wb") as f:
                        f.write(r.content)
                    print(f"✅ Rekaman berhasil diunduh: {filename}")
                else:
                    print(f"❌ Gagal mengunduh: {filename} (Status: {r.status_code})")
                    print(f"Response Text: {r.text}")
            except Exception as e:
                print(f"⚠️ Error saat mendownload {filename}: {e}")

# Main function untuk proses
def main():
    # Gantilah dengan access_token yang valid
    access_token = ACCESS_TOKEN

    if not access_token:
        print("Access Token tidak ditemukan, proses berhenti.")
        return

    # Ambil daftar user
    print("🎯 Mengambil daftar user...")
    users = get_user_list(access_token)

    # Loop untuk setiap user dan ambil rekaman
    for user in users:
        user_email = user['email']
        print(f"🔍 Mengecek rekaman untuk: {user_email}")

        # Ambil rekaman untuk user ini
        meetings = get_recordings(user_email, access_token)

        # Jika ada rekaman, coba unduh
        if meetings:
            print(f"🎞️ Jumlah rekaman ditemukan: {len(meetings)}")
            download_recordings(meetings, access_token, user_email)
        else:
            print(f"🎞️ Tidak ada rekaman untuk {user_email}.")

    print("✅ Semua rekaman selesai diproses.")

if __name__ == "__main__":
    main()
