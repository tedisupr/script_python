import os
import time
import win32com.client
from pathlib import Path

def restart_powerpoint():
    """Tutup semua instance PowerPoint sebelum memulai"""
    try:
        os.system("taskkill /F /IM POWERPNT.EXE")
        time.sleep(2)  # Tunggu sebelum melanjutkan
    except:
        pass  # Abaikan jika tidak ada PowerPoint yang berjalan

def convert_powerpoint_to_pdf(file_path, pdf_path):
    """Konversi PowerPoint ke PDF dengan penanganan error"""
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = 1  # Pastikan PowerPoint berjalan

        # Buka presentasi dalam mode Read-Only
        presentation = ppt.Presentations.Open(file_path, WithWindow=False, ReadOnly=True)

        # Tunggu beberapa detik untuk stabilitas
        time.sleep(2)

        # Coba simpan ulang ke PPTX jika terjadi masalah
        temp_pptx_path = str(Path(file_path).with_suffix(".temp.pptx"))
        presentation.SaveAs(temp_pptx_path, 24)  # 24 = ppSaveAsOpenXMLPresentation (PPTX)
        presentation.Close()
        
        # Buka ulang file yang sudah disimpan ulang
        presentation = ppt.Presentations.Open(temp_pptx_path, WithWindow=False, ReadOnly=True)

        # Simpan ke PDF
        presentation.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        presentation.Close()

        # Hapus file sementara
        os.remove(temp_pptx_path)
        os.remove(file_path)  # Hapus file asli setelah konversi
        print(f"✅ PowerPoint: '{os.path.basename(file_path)}' -> PDF (File asli dihapus)")

    except Exception as e:
        print(f"❌ Gagal mengonversi PowerPoint '{file_path}': {e}")

    finally:
        ppt.Quit()  # Pastikan PowerPoint ditutup setelah konversi

def convert_all_non_pdf_to_pdf(folder_path):
    """Konversi semua file non-PDF di dalam folder dan subfolder menjadi PDF jika belum ada file PDF"""
    folder_path = os.path.abspath(folder_path)

    # Restart PowerPoint sebelum mulai
    restart_powerpoint()

    for root, _, files in os.walk(folder_path):
        pdf_files = [f for f in files if f.lower().endswith(".pdf")]

        for filename in files:
            if filename.lower().endswith(".pdf"):
                continue  # Lewati file PDF yang sudah ada

            file_path = os.path.join(root, filename)
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"

            ext = filename.lower().split(".")[-1]

            if ext in ["ppt", "pptx"]:
                convert_powerpoint_to_pdf(file_path, pdf_path)

# Tentukan folder tempat file berada
folder = r"D:\Technical Support\Service\E-Library\Poltekes TNI AU\kti farmasi2\FARMASI JURNAL BARU 2021\Softfile Tk.3 2021"
convert_all_non_pdf_to_pdf(folder)
