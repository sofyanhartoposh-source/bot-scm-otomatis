import time
import requests
import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def jalankan_bot():
    # Setting agar browser jalan di server (tanpa layar)
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    
    # Simulasi Login & Ambil Data
    print("Membuka SCM...")
    # Di sini nanti kita isi perintah klik-klik tombolnya
    
    # Contoh Data yang berhasil diambil
    data_scm = ["2026-02-06", "Data SCM", "Export Sukses"]
    
    # KIRIM KE GAS (Ganti URL ini dengan URL Web App GAS kamu!)
    gas_url = "https://script.google.com/macros/s/AKfycbzuWeZNwdAqMDcg81bfkiOhIosljdswcR0tdRWeenQjwokhgyuM_0PZORvMt8IH5-E/exec"
    payload = {"data": data_scm}
    
    response = requests.post(gas_url, data=json.dumps(payload))
    print(f"Status kirim ke Sheet: {response.text}")

if __name__ == "__main__":
    jalankan_bot()
