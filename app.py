import streamlit as st
import pandas as pd
import time
import pyautogui
import keyboard
import os
import webbrowser
from PIL import Image
from io import BytesIO
import win32clipboard

# Konfigurasi Halaman
st.set_page_config(page_title="WA Bulk Sender ITICM", layout="wide")

st.title("🚀 WA Bulk Sender - By Diky Wahyu")

# Fungsi Kirim Gambar ke Clipboard
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

# Fungsi Kirim Teks ke Clipboard
def send_text_to_clipboard(text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()

def kirim_wa_flow_baru(no_wa, jalur_gambar, pesan):
    # Membuka WhatsApp Web
    url = f"https://web.whatsapp.com/send?phone={no_wa}"
    webbrowser.open(url)
    time.sleep(15) # Beri waktu loading (sesuaikan jika internet lambat)

    try:
        # 0. klik area
        width, height = pyautogui.size()
        pyautogui.click(width * 0.5, height * 0.9) 
        time.sleep(1)

        # 1. Copy & Paste Gambar
        send_image_to_clipboard(jalur_gambar)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(3) # Tunggu preview gambar muncul

        # 2. Copy & Paste Text (Caption)
        send_text_to_clipboard(pesan)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)

        # 3. Enter untuk kirim
        keyboard.press('enter')
        time.sleep(3) # Tunggu proses kirim selesai
        
        # 4. Tutup Tab
        pyautogui.hotkey('ctrl', 'w')
        return True
    except Exception as e:
        st.error(f"Gagal pada nomor {no_wa}: {e}")
        return False

# --- UI SIDEBAR ---
with st.sidebar:
    st.header("Konfigurasi")
    img_file = st.file_uploader("Upload Gambar", type=['jpg', 'jpeg', 'png'])
    caption_1 = st.text_area("Pesan 1 (Ganjil)", "Halo {name}...\nKUOTA TERBATAS!!\nTERBATAS!!", height=150)
    caption_2 = st.text_area("Pesan 2 (Genap)", "Hai Milennials {name}...\nBeasiswa S1 ITICM", height=150)
    st.warning("⚠️ Jangan pindah jendela saat aplikasi bekerja!")

# --- LOGIKA UTAMA ---
uploaded_file = st.file_uploader("Pilih File Excel (.xlsx)", type=['xlsx'])

if uploaded_file and img_file:
    temp_path = "temp_image.jpg"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())

    df = pd.read_excel(uploaded_file)
    df['Status'] = "Pending"
    
    table_placeholder = st.empty()
    table_placeholder.dataframe(df)

    if st.button("🚀 Mulai Kirim Pesan"):
        progress_bar = st.progress(0)
        
        for index, row in df.iterrows():
            nama = row['Nama']
            no_wa = str(row['WA']).replace(".0", "").strip()
            if not no_wa.startswith('+'):
                no_wa = "+62" + no_wa[1:] if no_wa.startswith('0') else "+" + no_wa

            # Logika Ganjil/Genap
            pesan_personal = caption_1 if (index + 1) % 2 != 0 else caption_2
            pesan_personal = pesan_personal.replace("{name}", nama)

            df.at[index, 'Status'] = "⏳ Proses..."
            table_placeholder.dataframe(df)

            if kirim_wa_flow_baru(no_wa, temp_path, pesan_personal):
                df.at[index, 'Status'] = "✅ Terhasil"
            else:
                df.at[index, 'Status'] = "❌ Gagal"

            table_placeholder.dataframe(df)
            progress_bar.progress((index + 1) / len(df))
            
        st.success("Selesai mengolah semua data!")
else:
    st.info("Upload file Excel dan Gambar di sidebar untuk memulai.")