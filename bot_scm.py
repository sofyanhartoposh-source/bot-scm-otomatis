import time
import requests
import json
import pandas as pd
import io
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def jalankan_bot():
    print("=== [LOG START] Memulai Operasi Bot SCM ===")
    
    options = uc.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    print("[1/6] Menyiapkan Browser Siluman...")
    driver = uc.Chrome(options=options, version_main=144)
    
    try:
        print("[2/6] Membuka Halaman Login: https://scm.nusadaya.net/login")
        driver.get("https://scm.nusadaya.net/login")
        
        wait = WebDriverWait(driver, 25)
        
        print("[3/6] Mengisi Form Login...")
        # Logika pencarian input
        email_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text' or @placeholder='Email atau NIP']")))
        email_input.send_keys("sofyan.hartopo@pln.co.id")
        
        pass_input = driver.find_element(By.XPATH, "//input[@type='password']")
        pass_input.send_keys("9314014DY")
        
        print("-> Mengetuk tombol Log in...")
        login_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]")
        login_btn.click()
        
        print("-> Login dikirim, menunggu dashboard (15 detik)...")
        time.sleep(15)
        
        # Cek apakah login berhasil dengan melihat cookie
        cookies = driver.get_cookies()
        if not cookies:
            print("(!) Peringatan: Tidak ada cookie terdeteksi. Login mungkin gagal.")
        else:
            print(f"-> Berhasil mendapatkan {len(cookies)} cookies.")

        print("[4/6] Menembak URL Export Langsung...")
        session_cookies = {c['name']: c['value'] for c in cookies}
        export_url = "https://scm.nusadaya.net/monitoring-kontrak-rinci/export?khs=all&bidang=all&tahun=2026&stage="
        
        response_dl = requests.get(export_url, cookies=session_cookies)
        
        if response_dl.status_code == 200:
            print(f"-> Download Sukses! Ukuran file: {len(response_dl.content)} bytes")
        else:
            print(f"(!) Gagal download. Status Code: {response_dl.status_code}")
            return

        print("[5/6] Membongkar isi Excel dengan Pandas...")
        df = pd.read_excel(io.BytesIO(response_dl.content))
        df = df.fillna("") # Bersihkan data kosong agar JSON tidak rusak
        data_rows = df.values.tolist()
        print(f"-> Total data ditemukan: {len(data_rows)} baris.")

        print("[6/6] Mengirim data ke Google Sheets...")
        gas_url = "https://script.google.com/macros/s/AKfycbzuWeZNwdAqMDcg81bfkiOhIosljdswcR0tdRWeenQjwokhgyuM_0PZORvMt8IH5-E/exec"
        
        payload = json.dumps({"rows": data_rows})
        res_gas = requests.post(gas_url, data=payload, headers={'Content-Type': 'application/json'})
        
        print(f"=== [HASIL AKHIR] Respon dari Google Sheets: {res_gas.text} ===")

    except Exception as e:
        print(f"!!! [ERROR] Terjadi kendala teknis: {str(e)}")
    finally:
        print("=== [LOG END] Menutup Browser & Selesai ===")
        driver.quit()

if __name__ == "__main__":
    jalankan_bot()
