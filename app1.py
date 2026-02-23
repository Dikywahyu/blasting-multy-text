import streamlit as st
import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import win32clipboard
from PIL import Image
from io import BytesIO

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="WA Bulk Sender ITICM", layout="wide")
st.title("🚀 WA Bulk Sender - Optimasi Kecepatan")

# --- FUNGSI HELPER CLIPBOARD ---
def send_image_to_clipboard(path):
    img = Image.open(path)
    output = BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def send_text_to_clipboard(text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()

# --- FUNGSI SETUP BROWSER (HANYA SEKALI) ---
def init_driver():
    chrome_options = Options()
    script_dir = os.getcwd()
    profile_path = os.path.join(script_dir, "WA_Bot_Profile")
    chrome_options.add_argument(f"--user-data-dir={profile_path}")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

# --- FUNGSI INTI PENGIRIMAN PER NOMOR ---
def kirim_per_nomor(driver, no_wa, jalur_gambar, pesan):
    wait = WebDriverWait(driver, 60)
    try:
        # Buka tab baru dan pindah ke tab tersebut
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        
        url = f"https://web.whatsapp.com/send?phone={no_wa}"
        driver.get(url)

        # 1. Tunggu Footer & Klik Kolom Chat Utama
        footer = wait.until(EC.presence_of_element_located((By.XPATH, '//footer[@class="_ak1i x1wiwyrm"]')))
        chat_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="textbox"][@aria-placeholder="Ketik pesan"]')))
        chat_box.click()

        # 2. PASTE GAMBAR & KIRIM
        send_image_to_clipboard(jalur_gambar)
        chat_box.send_keys(Keys.CONTROL + 'v')
        btn_kirim_gambar = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="button"][@aria-label="Kirim"]')))
        btn_kirim_gambar.click()
        
        # Tunggu gambar terkirim
        time.sleep(5) 

        # 3. PASTE TEXT & KIRIM
        chat_box_baru = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@aria-placeholder="Ketik pesan"]')))
        chat_box_baru.click()
        send_text_to_clipboard(pesan)
        chat_box_baru.send_keys(Keys.CONTROL + 'v')
        time.sleep(1)
        
        btn_kirim_teks = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Kirim"] | //div[@role="button"][@aria-label="Kirim"]')))
        btn_kirim_teks.click()
        
        time.sleep(3) 
        
        # Tutup tab saat ini dan kembali ke handle pertama (tab kosong/utama)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        return True
    except Exception as e:
        # Jika gagal, tutup tab yang bermasalah agar tidak menumpuk
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
        return False

# --- UI & LOGIKA ---
with st.sidebar:
    st.header("Konfigurasi")
    img_file = st.file_uploader("Upload Gambar", type=['jpg', 'jpeg', 'png'])
    cap_ganjil = st.text_area("Pesan Ganjil", "Halo {name}...", height=150)
    cap_genap = st.text_area("Pesan Genap", "Hai {name}...", height=150)

uploaded_file = st.file_uploader("Pilih Excel", type=['xlsx'])

if uploaded_file and img_file:
    temp_path = "temp_broadcast.jpg"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())

    df = pd.read_excel(uploaded_file)
    df['Status'] = "Pending"
    table_placeholder = st.empty()
    table_placeholder.dataframe(df)

    if st.button("🚀 Mulai Kirim Pesan"):
        # Inisialisasi driver hanya sekali
        driver = init_driver()
        progress_bar = st.progress(0)
        
        try:
            for index, row in df.iterrows():
                nama = row['Nama']
                no_wa = str(row['WA']).replace(".0", "").strip()
                if no_wa.startswith('0'): no_wa = "+62" + no_wa[1:]
                elif no_wa.startswith('8'): no_wa = "+62" + no_wa
                elif not no_wa.startswith('+'): no_wa = "+" + no_wa

                pesan_fix = cap_ganjil if (index + 1) % 2 != 0 else cap_genap
                pesan_fix = pesan_fix.replace("{name}", nama)

                df.at[index, 'Status'] = "⏳ Mengirim..."
                table_placeholder.dataframe(df)

                if kirim_per_nomor(driver, no_wa, temp_path, pesan_fix):
                    df.at[index, 'Status'] = "✅ Berhasil"
                else:
                    df.at[index, 'Status'] = "❌ Gagal"

                table_placeholder.dataframe(df)
                progress_bar.progress((index + 1) / len(df))
            
            st.success("Semua pesan selesai diproses!")
        finally:
            driver.quit() # Tutup browser sepenuhnya setelah loop selesai
else:
    st.info("Upload file Excel dan Gambar di sidebar.")