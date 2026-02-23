import streamlit as st
import pandas as pd
import time
import os
import random
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

st.set_page_config(page_title="WA Bulk Anti-Banned - ITICM", layout="wide")
st.title("🚀 WA Interactive Blasting - By Diky Wahyu")

# --- FUNGSI CLIPBOARD ---
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

# --- SETUP BROWSER PROFIL TERPISAH ---
def init_driver(profile_folder):
    chrome_options = Options()
    path = os.path.join(os.getcwd(), profile_folder)
    chrome_options.add_argument(f"--user-data-dir={path}")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

# --- FUNGSI KIRIM PESAN TANPA TAB BARU ---
def kirim_interaksi(driver, no_tujuan, pesan):
    wait = WebDriverWait(driver, 30)
    try:
        driver.get(f"https://web.whatsapp.com/send?phone={no_tujuan}")
        chat_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="textbox"][@aria-placeholder="Ketik pesan"]')))
        send_text_to_clipboard(pesan)
        chat_box.send_keys(Keys.CONTROL + 'v')
        chat_box.send_keys(Keys.ENTER)
        time.sleep(2)
        return True
    except:
        return False

# --- SIDEBAR KONFIGURASI ---
with st.sidebar:
    st.header("📲 Konfigurasi Dual HP")
    no_hp_1 = st.text_input("Nomor HP 1 (Blasting)", placeholder="628xxx")
    no_hp_2 = st.text_input("Nomor HP 2 (Chatter)", placeholder="628xxx")
    txt_file = st.file_uploader("Upload Script Chat (.txt)", type=['txt'])
    
    st.divider()
    st.header("📁 File & Pesan")
    img_file = st.file_uploader("Upload Gambar Utama", type=['jpg', 'jpeg', 'png'])
    
    cap_ganjil = st.text_area("Pesan Ganjil", "Halo {name}...")
    cap_genap = st.text_area("Pesan Genap", "Hai {name}...")

# --- LOGIKA UTAMA ---
uploaded_excel = st.file_uploader("Pilih Data Excel", type=['xlsx'])

if uploaded_excel and img_file and txt_file and no_hp_1 and no_hp_2:
    # Simpan file sementara
    temp_img = "temp_img.jpg"
    with open(temp_img, "wb") as f:
        f.write(img_file.getbuffer())
    
    # Baca Script Chat TXT
    chat_lines = txt_file.getvalue().decode("utf-8").splitlines()
    chat_lines = [line.strip() for line in chat_lines if line.strip()] # Bersihkan baris kosong
    
    df = pd.read_excel(uploaded_excel)
    df['Status'] = "Pending"
    tabel = st.empty()
    tabel.dataframe(df)

    if st.button("🚀 Jalankan Blasting + Infinite Chat"):
        d1 = init_driver("Profile_HP_1")
        d2 = init_driver("Profile_HP_2")
        
        st.warning("Scan QR di kedua browser! Tunggu sampai WhatsApp siap.")
        time.sleep(15)

        current_txt_index = 0
        progress_bar = st.progress(0)

        for index, row in df.iterrows():
            nama = row['Nama']
            target = str(row['WA']).replace(".0", "").strip()
            if target.startswith('0'): target = "62" + target[1:]

            # 1. HP 1 BLASTING KE TARGET (TAB BARU)
            d1.execute_script("window.open('');")
            d1.switch_to.window(d1.window_handles[-1])
            try:
                d1.get(f"https://web.whatsapp.com/send?phone={target}")
                wait_blast = WebDriverWait(d1, 60)
                box = wait_blast.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="textbox"][@aria-placeholder="Ketik pesan"]')))
                
                # Paste Gambar
                send_image_to_clipboard(temp_img)
                box.send_keys(Keys.CONTROL + 'v')
                wait_blast.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="button"][@aria-label="Kirim"]'))).click()
                time.sleep(4)
                
                # Paste Teks
                p = cap_ganjil if index % 2 == 0 else cap_genap
                send_text_to_clipboard(p.replace("{name}", nama))
                box_baru = wait_blast.until(EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@aria-placeholder="Ketik pesan"]')))
                box_baru.send_keys(Keys.CONTROL + 'v')
                box_baru.send_keys(Keys.ENTER)
                time.sleep(3)
            finally:
                d1.close()
                d1.switch_to.window(d1.window_handles[0])

            # 2. SIMULASI CHAT ANTAR HP (MENGGUNAKAN TXT)
            # Ambil baris dari TXT
            pesan_txt = chat_lines[current_txt_index]
            
            # Ganjil: HP 2 kirim ke HP 1 | Genap: HP 1 balas ke HP 2
            if index % 2 == 0:
                kirim_interaksi(d2, no_hp_1, pesan_txt)
            else:
                kirim_interaksi(d1, no_hp_2, pesan_txt)

            # Update index TXT (Reset ke 0 jika habis)
            current_txt_index += 1
            if current_txt_index >= len(chat_lines):
                current_txt_index = 0 # Kembali ke baris pertama TXT

            # Update Status
            df.at[index, 'Status'] = f"✅ Terkirim (Chat line {current_txt_index})"
            tabel.dataframe(df)
            progress_bar.progress((index + 1) / len(df))
            
            # Jeda acak biar natural
            time.sleep(random.randint(5, 12))

        st.success("Proses Blasting & Infinite Chat Selesai!")
        d1.quit()
        d2.quit()