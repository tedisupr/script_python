# konversi file ke pdf
import os
import time
import win32com.client
from pathlib import Path

# script untuk mengonversi file Microsoft Office (Word, Excel, PowerPoint, juga RTF) menjadi PDF secara otomatis dalam sebuah folder (termasuk subfolder).

def restart_office_apps():
    """Tutup semua instance PowerPoint, Word, dan Excel jika sedang berjalan"""
    apps = ["POWERPNT.EXE", "WINWORD.EXE", "EXCEL.EXE"]
    for app in apps:
        running = os.popen(f'tasklist | findstr {app}').read()
        if app in running:
            print(f"🔄 Menutup {app} yang sedang berjalan...")
            os.system(f"taskkill /F /IM {app}")
            time.sleep(2)  # Tunggu sebentar sebelum melanjutkan
        else:
            print(f"✅ Tidak ada {app} yang berjalan.")

def convert_office_to_pdf(file_path, pdf_path):
    """Konversi file Office (PowerPoint, Word, Excel, RTF) ke PDF"""
    ext = file_path.suffix.lower()
    
    if file_path.name.startswith("~$"):
        print(f"⚠️ Lewati file temporary: {file_path}")
        return  # Lewati file temporary dari Microsoft Office
    
    try:
        if ext in [".ppt", ".pptx"]:
            print(f"🚀 Mengonversi PowerPoint: {file_path} -> {pdf_path}")
            app = win32com.client.Dispatch("PowerPoint.Application")
            app.Visible = 1
            presentation = app.Presentations.Open(str(file_path), WithWindow=False, ReadOnly=True)
            presentation.SaveAs(str(pdf_path), 32)  # 32 = Save as PDF
            presentation.Close()
            app.Quit()
        
        elif ext in [".doc", ".docx", ".rtf"]:
            print(f"🚀 Mengonversi Word/RTF: {file_path} -> {pdf_path}")
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False
            try:
                doc = app.Documents.Open(str(file_path), ReadOnly=True)
                doc.SaveAs(str(pdf_path), 17)  # 17 = Save as PDF
                doc.Close()
            except Exception as e:
                print(f"❌ Gagal membuka Word/RTF: {file_path} -> {e}")
            app.Quit()
            
        elif ext in [".xls", ".xlsx"]:
            print(f"🚀 Mengonversi Excel: {file_path} -> {pdf_path}")
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            try:
                wb = app.Workbooks.Open(str(file_path), ReadOnly=True)
                wb.ExportAsFixedFormat(0, str(pdf_path))  # 0 = Save as PDF
                wb.Close()
            except Exception as e:
                print(f"❌ Gagal membuka Excel: {file_path} -> {e}")
            app.Quit()
        
        print(f"✅ Berhasil: {file_path} -> {pdf_path}")
    
    except Exception as e:
        print(f"❌ Gagal mengonversi {file_path}: {e}")

def convert_all_non_pdf_to_pdf(folder_path):
    """Konversi semua file non-PDF di dalam folder dan subfolder"""
    folder_path = Path(folder_path).resolve()
    
    restart_office_apps()
    
    for file_path in folder_path.rglob("*.*"):  # Scan semua file dalam folder dan subfolder
        if file_path.suffix.lower() == ".pdf":
            continue  # Lewati file PDF
        
        pdf_path = file_path.with_suffix(".pdf")
        convert_office_to_pdf(file_path, pdf_path)

# Jalankan konversi
folder = Path(r"D:\\Technical Support\\Service\\E-Library\\Poltekes TNI AU\\KTI PERAWAT")
convert_all_non_pdf_to_pdf(folder)
